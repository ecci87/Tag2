#!/usr/bin/env python3
"""
Image Captioning Tool - Backend Server
FastAPI-based server for browsing images and managing captions.
"""

import os
import sys
import json
import asyncio
import subprocess
import platform
from pathlib import Path
from io import BytesIO
from typing import Optional
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Query, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from PIL import Image

app = FastAPI(title="Image Captioning Tool")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thread pool for IO-bound thumbnail generation
executor = ThreadPoolExecutor(max_workers=8)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"}

# In-memory thumbnail cache: (path, mtime) -> bytes
thumbnail_cache: dict[tuple[str, float], bytes] = {}
THUMBNAIL_SIZES = [64, 128, 256, 400]
PREVIEW_MAX_SIZE = 2048  # Max edge for preview images

# ===== CONFIG FILE =====
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def _load_config() -> dict:
    """Load config from disk. Returns default structure if file missing."""
    default = {
        "last_folder": "",
        # Sentences shared across all folders when a folder has no own config yet
        "default_sentences": [],
        # Per-folder sentence lists. Key = absolute folder path.
        # To copy sentences to a new folder, duplicate a block here.
        "folders": {}
    }
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Merge with defaults so new keys are always present
            for k, v in default.items():
                data.setdefault(k, v)
            return data
        except Exception:
            pass
    return default


def _save_config(cfg: dict):
    """Persist config to disk with pretty formatting for easy hand-editing."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def _get_folder_sections(cfg: dict, folder: str) -> list[dict]:
    """Get sections for a folder, falling back to default_sections.
    Returns list of dicts: [{"name": "SectionName", "sentences": [...]}, ...]
    The unnamed section (name='') holds sentences without a section header.
    """
    folder_key = os.path.normpath(folder)
    folder_cfg = cfg.get("folders", {}).get(folder_key, None)
    if folder_cfg is not None:
        if "sections" in folder_cfg:
            return [dict(s) for s in folder_cfg["sections"]]
        # Migration: old flat sentences list -> single unnamed section
        if "sentences" in folder_cfg:
            return [{"name": "", "sentences": list(folder_cfg["sentences"])}]
    # Try default_sections
    default_sections = cfg.get("default_sections", [])
    if default_sections:
        return [dict(s) for s in default_sections]
    # Migration from old default_sentences
    default_sentences = cfg.get("default_sentences", [])
    if default_sentences:
        return [{"name": "", "sentences": list(default_sentences)}]
    return [{"name": "", "sentences": []}]


def _set_folder_sections(cfg: dict, folder: str, sections: list[dict]):
    """Set sections for a specific folder."""
    folder_key = os.path.normpath(folder)
    if "folders" not in cfg:
        cfg["folders"] = {}
    if folder_key not in cfg["folders"]:
        cfg["folders"][folder_key] = {}
    cfg["folders"][folder_key]["sections"] = sections
    # Remove legacy key if present
    cfg["folders"][folder_key].pop("sentences", None)


def _all_sentences_from_sections(sections: list[dict]) -> list[str]:
    """Flatten sections into a single list of all predefined sentences."""
    result = []
    for sec in sections:
        result.extend(sec.get("sentences", []))
    return result


def _generate_thumbnail(filepath: str, size: int) -> bytes:
    """Generate a JPEG thumbnail of the given size (longest edge)."""
    try:
        with Image.open(filepath) as img:
            img.thumbnail((size, size), Image.LANCZOS)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            buf = BytesIO()
            quality = 85 if size >= 1024 else 80
            img.save(buf, format="JPEG", quality=quality, optimize=True)
            return buf.getvalue()
    except Exception:
        return b""


def _get_thumbnail(filepath: str, size: int) -> bytes:
    """Get thumbnail from cache or generate it."""
    mtime = os.path.getmtime(filepath)
    key = (filepath, mtime, size)
    if key not in thumbnail_cache:
        thumbnail_cache[key] = _generate_thumbnail(filepath, size)
    return thumbnail_cache[key]


@app.get("/api/list-images")
async def list_images(folder: str = Query(...)):
    """List all image files in the given folder."""
    folder_path = Path(folder)
    if not folder_path.is_dir():
        raise HTTPException(status_code=400, detail="Not a valid directory")

    images = []
    try:
        for entry in sorted(folder_path.iterdir()):
            if entry.is_file() and entry.suffix.lower() in IMAGE_EXTENSIONS:
                stat = entry.stat()
                caption_file = entry.with_suffix(".txt")
                images.append({
                    "name": entry.name,
                    "path": str(entry),
                    "size": stat.st_size,
                    "mtime": stat.st_mtime,
                    "has_caption": caption_file.exists(),
                })
    except PermissionError:
        raise HTTPException(status_code=403, detail="Permission denied")

    return {"images": images, "folder": str(folder_path.resolve())}


@app.get("/api/thumbnail")
async def get_thumbnail(path: str = Query(...), size: int = Query(default=256)):
    """Return a thumbnail for the given image path."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")

    # Clamp size to nearest available
    actual_size = min(THUMBNAIL_SIZES, key=lambda s: abs(s - size))

    loop = asyncio.get_event_loop()
    data = await loop.run_in_executor(executor, _get_thumbnail, path, actual_size)

    if not data:
        raise HTTPException(status_code=500, detail="Failed to generate thumbnail")

    return StreamingResponse(BytesIO(data), media_type="image/jpeg")


@app.get("/api/image")
async def get_image(path: str = Query(...)):
    """Serve the full-resolution image."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path)


@app.get("/api/preview")
async def get_preview(path: str = Query(...)):
    """Serve a preview-quality image (downscaled to max 2048px edge, JPEG).
    Much faster than serving the full-res file for large images."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")

    loop = asyncio.get_event_loop()
    data = await loop.run_in_executor(executor, _get_thumbnail, path, PREVIEW_MAX_SIZE)

    if not data:
        # Fallback to serving full file
        return FileResponse(path)

    return StreamingResponse(
        BytesIO(data),
        media_type="image/jpeg",
        headers={"Cache-Control": "public, max-age=3600"},
    )


class BatchCaptionUpdate(BaseModel):
    image_paths: list[str]
    sentence: str
    enabled: bool


class BatchFreeTextUpdate(BaseModel):
    image_path: str
    free_text: str


def _get_caption_path(image_path: str) -> Path:
    """Get the corresponding .txt path for an image."""
    p = Path(image_path)
    return p.with_suffix(".txt")


def _read_caption_file(image_path: str, predefined_sentences: list[str],
                       section_headers: list[str] | None = None) -> dict:
    """Read caption file and separate predefined sentences from free text.
    section_headers is the list of known section header lines (freeform text).
    Handles both old flat format and new sectioned format with '- ' prefixes.
    """
    caption_path = _get_caption_path(image_path)
    enabled_sentences = []
    free_lines = []
    header_set = set(section_headers or [])

    if caption_path.exists():
        try:
            content = caption_path.read_text(encoding="utf-8")
            pred_set = set(predefined_sentences)
            in_structured = True

            for line in content.split("\n"):
                stripped = line.strip()

                if not stripped:
                    if not in_structured:
                        free_lines.append(line)
                    continue

                # Known section header - consume it (part of structured section)
                if in_structured and stripped in header_set:
                    continue

                # Sentence with "- " prefix (new format)
                if in_structured and stripped.startswith("- "):
                    sentence_text = stripped[2:]
                    if sentence_text in pred_set:
                        enabled_sentences.append(sentence_text)
                        continue

                # Backward compat: sentence without "- " prefix (old format)
                if in_structured and stripped in pred_set:
                    enabled_sentences.append(stripped)
                    continue

                # Not a predefined sentence - everything from here is free text
                in_structured = False
                free_lines.append(line)
        except Exception:
            pass

    return {
        "enabled_sentences": enabled_sentences,
        "free_text": "\n".join(free_lines),
    }


def _write_caption_file(image_path: str, enabled_sentences: list[str], free_text: str,
                         sections: list[dict] | None = None):
    """Write caption file with sectioned format.
    Section names are written as-is (freeform headers like '## Lighting', '**Shape**', etc.).
    Sentences get '- ' prefixed. The unnamed section (name='') has no header line.
    Free text goes at the end.
    """
    caption_path = _get_caption_path(image_path)

    if sections:
        enabled_set = set(enabled_sentences)
        blocks = []

        for section in sections:
            sec_name = section.get("name", "")
            sec_sentences = [s for s in section.get("sentences", []) if s in enabled_set]
            if not sec_sentences:
                continue

            lines = []
            if sec_name:
                lines.append(sec_name)
            for s in sec_sentences:
                lines.append(f"- {s}")
            blocks.append("\n".join(lines))

        content_parts = []
        if blocks:
            content_parts.append("\n\n".join(blocks))
        if free_text and free_text.strip():
            content_parts.append(free_text.strip())

        content = "\n\n".join(content_parts)
    else:
        # Fallback: simple format with '- ' prefix
        parts = []
        if enabled_sentences:
            parts.append("\n".join(f"- {s}" for s in enabled_sentences))
        if free_text and free_text.strip():
            parts.append(free_text.strip())
        content = "\n\n".join(parts)

    if content:
        caption_path.write_text(content + "\n", encoding="utf-8")
    elif caption_path.exists():
        caption_path.write_text("", encoding="utf-8")


@app.get("/api/caption")
async def get_caption(path: str = Query(...), sentences: str = Query(default="[]")):
    """Read caption data for an image. sentences is a JSON array of predefined sentences."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Image file not found")
    try:
        predefined = json.loads(sentences)
    except json.JSONDecodeError:
        predefined = []

    # Load section headers for proper parsing
    cfg = _load_config()
    folder = str(Path(path).parent)
    sections = _get_folder_sections(cfg, folder)
    headers = [s["name"] for s in sections if s.get("name")]

    data = _read_caption_file(path, predefined, headers)
    return data


@app.post("/api/caption/batch-toggle")
async def batch_toggle_sentence(update: BatchCaptionUpdate):
    """Toggle a predefined sentence on/off for multiple images."""
    results = []

    # Load sections config from the folder of the first image
    cfg = _load_config()
    folder = str(Path(update.image_paths[0]).parent) if update.image_paths else ""
    sections = _get_folder_sections(cfg, folder)
    all_sentences = _all_sentences_from_sections(sections)

    for img_path in update.image_paths:
        if not os.path.isfile(img_path):
            results.append({"path": img_path, "error": "File not found"})
            continue

        # Read existing caption data
        headers = [s["name"] for s in sections if s.get("name")]
        data = _read_caption_file(img_path, all_sentences, headers)
        enabled = list(data["enabled_sentences"])
        free_text = data["free_text"]

        if update.enabled:
            if update.sentence not in enabled:
                enabled.append(update.sentence)
        else:
            enabled = [s for s in enabled if s != update.sentence]

        _write_caption_file(img_path, enabled, free_text, sections)
        results.append({"path": img_path, "ok": True})

    return {"results": results}


@app.post("/api/caption/save-free-text")
async def save_free_text(data: BatchFreeTextUpdate):
    """Save free text for a single image, preserving predefined sentences."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image not found")

    # We need to know which lines are predefined to preserve them
    # Client will send the sentences query param
    # Actually let's accept sentences in the body too
    return {"ok": True}


class SaveCaptionFull(BaseModel):
    image_path: str
    enabled_sentences: list[str]
    free_text: str


@app.post("/api/caption/save")
async def save_caption(data: SaveCaptionFull):
    """Full save of caption data for one image."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image not found")
    # Load sections config for proper formatting
    cfg = _load_config()
    folder = str(Path(data.image_path).parent)
    sections = _get_folder_sections(cfg, folder)
    _write_caption_file(data.image_path, data.enabled_sentences, data.free_text, sections)
    return {"ok": True}


@app.get("/api/captions/bulk")
async def get_captions_bulk(paths: str = Query(...), sentences: str = Query(default="[]")):
    """Get caption status for multiple images at once.
    paths: JSON array of image paths
    sentences: JSON array of predefined sentences
    """
    try:
        image_paths = json.loads(paths)
        predefined = json.loads(sentences)
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid JSON")

    # Load section headers (assume all images in same folder)
    cfg = _load_config()
    if image_paths:
        folder = str(Path(image_paths[0]).parent)
        sections = _get_folder_sections(cfg, folder)
        headers = [s["name"] for s in sections if s.get("name")]
    else:
        headers = []

    results = {}
    for img_path in image_paths:
        if os.path.isfile(img_path):
            results[img_path] = _read_caption_file(img_path, predefined, headers)
        else:
            results[img_path] = {"enabled_sentences": [], "free_text": ""}

    return results


# ===== SETTINGS API =====

class SettingsUpdate(BaseModel):
    last_folder: Optional[str] = None
    sections: Optional[list[dict]] = None
    folder: Optional[str] = None  # which folder these sections belong to


@app.get("/api/settings")
async def get_settings(folder: Optional[str] = Query(default=None)):
    """Get full settings. If folder is specified, include sections for that folder."""
    cfg = _load_config()
    result = {
        "last_folder": cfg.get("last_folder", ""),
    }
    if folder:
        result["sections"] = _get_folder_sections(cfg, folder)
        result["folder"] = os.path.normpath(folder)
    return result


@app.post("/api/settings")
async def update_settings(data: SettingsUpdate):
    """Update settings. Saves last_folder and/or per-folder sections."""
    cfg = _load_config()
    if data.last_folder is not None:
        cfg["last_folder"] = data.last_folder
    if data.sections is not None and data.folder:
        _set_folder_sections(cfg, data.folder, data.sections)
    _save_config(cfg)
    return {"ok": True}


@app.get("/api/open-in-explorer")
async def open_in_explorer(path: str = Query(...)):
    """Open the OS file explorer with the given file selected."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    try:
        system = platform.system()
        if system == "Windows":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif system == "Darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            # Linux: open the containing folder
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# Serve the frontend
@app.get("/")
async def index():
    return FileResponse(os.path.join(os.path.dirname(__file__), "static", "index.html"))


# Mount static files
static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")


if __name__ == "__main__":
    import uvicorn
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8899
    uvicorn.run(app, host="0.0.0.0", port=port)
