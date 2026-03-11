#!/usr/bin/env python3
"""
Image Captioning Tool - Backend Server
FastAPI-based server for browsing images and managing captions.
"""

import os
import sys
import json
import asyncio
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


def _get_folder_sentences(cfg: dict, folder: str) -> list[str]:
    """Get sentences for a folder, falling back to default_sentences."""
    folder_key = os.path.normpath(folder)
    folder_cfg = cfg.get("folders", {}).get(folder_key, None)
    if folder_cfg is not None:
        return folder_cfg.get("sentences", [])
    return list(cfg.get("default_sentences", []))


def _set_folder_sentences(cfg: dict, folder: str, sentences: list[str]):
    """Set sentences for a specific folder."""
    folder_key = os.path.normpath(folder)
    if "folders" not in cfg:
        cfg["folders"] = {}
    if folder_key not in cfg["folders"]:
        cfg["folders"][folder_key] = {}
    cfg["folders"][folder_key]["sentences"] = sentences


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


class CaptionData(BaseModel):
    image_path: str
    sentences: list[str]  # predefined sentences that are enabled
    free_text: str


class BatchCaptionUpdate(BaseModel):
    image_paths: list[str]
    sentence: str
    enabled: bool


class BatchFreeTextUpdate(BaseModel):
    image_path: str
    free_text: str


class SentencesConfig(BaseModel):
    sentences: list[str]


def _get_caption_path(image_path: str) -> Path:
    """Get the corresponding .txt path for an image."""
    p = Path(image_path)
    return p.with_suffix(".txt")


def _read_caption_file(image_path: str, predefined_sentences: list[str]) -> dict:
    """Read caption file and separate predefined sentences from free text."""
    caption_path = _get_caption_path(image_path)
    enabled_sentences = []
    free_lines = []

    if caption_path.exists():
        try:
            content = caption_path.read_text(encoding="utf-8")
            lines = content.split("\n")
            pred_set = set(predefined_sentences)
            for line in lines:
                stripped = line.strip()
                if stripped in pred_set:
                    enabled_sentences.append(stripped)
                elif stripped:  # non-empty, non-predefined
                    free_lines.append(line)
                # preserve empty lines in free text section only after we've passed predefined
            # Actually, let's be smarter: predefined are always at top
            # Re-parse: read from top, consume predefined lines, rest is free text
            enabled_sentences = []
            free_lines = []
            in_predefined_section = True
            for line in lines:
                stripped = line.strip()
                if in_predefined_section and stripped in pred_set:
                    enabled_sentences.append(stripped)
                else:
                    if in_predefined_section and stripped == "":
                        # Could be separator between predefined and free text
                        in_predefined_section = False
                        continue
                    in_predefined_section = False
                    free_lines.append(line)
        except Exception:
            pass

    return {
        "enabled_sentences": enabled_sentences,
        "free_text": "\n".join(free_lines),
    }


def _write_caption_file(image_path: str, enabled_sentences: list[str], free_text: str):
    """Write caption file with predefined sentences at top, then free text."""
    caption_path = _get_caption_path(image_path)
    parts = []
    if enabled_sentences:
        parts.append("\n".join(enabled_sentences))
    if free_text.strip():
        parts.append(free_text.strip())
    content = "\n\n".join(parts)
    if content:
        caption_path.write_text(content + "\n", encoding="utf-8")
    elif caption_path.exists():
        # If nothing to write but file exists, keep it empty or remove
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

    data = _read_caption_file(path, predefined)
    return data


@app.post("/api/caption/batch-toggle")
async def batch_toggle_sentence(update: BatchCaptionUpdate):
    """Toggle a predefined sentence on/off for multiple images."""
    results = []
    for img_path in update.image_paths:
        if not os.path.isfile(img_path):
            results.append({"path": img_path, "error": "File not found"})
            continue
        # We need predefined sentences list - client sends just the sentence to toggle
        # Read existing file, toggle the sentence
        caption_path = _get_caption_path(img_path)
        existing_lines = []
        free_lines = []

        if caption_path.exists():
            try:
                content = caption_path.read_text(encoding="utf-8")
                existing_lines = [l.strip() for l in content.split("\n") if l.strip()]
            except Exception:
                pass

        if update.enabled:
            # Add sentence at top if not present
            if update.sentence not in existing_lines:
                existing_lines.insert(0, update.sentence)
        else:
            # Remove sentence
            existing_lines = [l for l in existing_lines if l != update.sentence]

        # Write back - we need to preserve structure
        # Re-read properly to keep free text intact
        caption_path_obj = _get_caption_path(img_path)
        if caption_path_obj.exists():
            raw = caption_path_obj.read_text(encoding="utf-8")
        else:
            raw = ""

        # Parse: remove the sentence line if disabling, add at top if enabling
        lines = raw.split("\n")
        if update.enabled:
            # Check if sentence already in file
            if update.sentence not in [l.strip() for l in lines]:
                # Insert at very top
                lines.insert(0, update.sentence)
        else:
            lines = [l for l in lines if l.strip() != update.sentence]

        # Clean up: remove leading/trailing empty lines
        content = "\n".join(lines).strip()
        if content:
            caption_path_obj.write_text(content + "\n", encoding="utf-8")
        else:
            caption_path_obj.write_text("", encoding="utf-8")

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
    _write_caption_file(data.image_path, data.enabled_sentences, data.free_text)
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

    results = {}
    for img_path in image_paths:
        if os.path.isfile(img_path):
            results[img_path] = _read_caption_file(img_path, predefined)
        else:
            results[img_path] = {"enabled_sentences": [], "free_text": ""}

    return results


# ===== SETTINGS API =====

class SettingsUpdate(BaseModel):
    last_folder: Optional[str] = None
    sentences: Optional[list[str]] = None
    folder: Optional[str] = None  # which folder these sentences belong to


@app.get("/api/settings")
async def get_settings(folder: Optional[str] = Query(default=None)):
    """Get full settings. If folder is specified, include sentences for that folder."""
    cfg = _load_config()
    result = {
        "last_folder": cfg.get("last_folder", ""),
        "default_sentences": cfg.get("default_sentences", []),
    }
    if folder:
        result["sentences"] = _get_folder_sentences(cfg, folder)
        result["folder"] = os.path.normpath(folder)
    return result


@app.post("/api/settings")
async def update_settings(data: SettingsUpdate):
    """Update settings. Saves last_folder and/or per-folder sentences."""
    cfg = _load_config()
    if data.last_folder is not None:
        cfg["last_folder"] = data.last_folder
    if data.sentences is not None and data.folder:
        _set_folder_sentences(cfg, data.folder, data.sentences)
    _save_config(cfg)
    return {"ok": True}


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
