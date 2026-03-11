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
import warnings
from pathlib import Path
from io import BytesIO
from typing import Optional
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Query, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from PIL import Image

from tag2_captions import (
    _build_caption_text,
    _get_caption_path,
    _read_caption_file,
    _write_caption_file,
)
from tag2_images import (
    IMAGE_EXTENSIONS,
    PREVIEW_MAX_SIZE,
    THUMBNAIL_SIZES,
    _apply_real_crop,
    _clear_thumbnail_cache_for_path,
    _encode_image_for_ollama,
    _generate_thumbnail,
    _get_crop_backup_dir,
    _get_crop_backup_path,
    _get_display_image_size,
    _get_image_crop,
    _get_thumbnail,
    _load_oriented_image,
    _normalize_crop_rect,
    _normalize_exif_bytes,
    _normalize_image_key,
    _remove_real_crop,
    _render_image_bytes,
    _rotate_image,
    _rotate_image_file,
    _save_image_file,
)
from tag2_ollama import (
    _auto_caption_sections as _tag2_auto_caption_sections,
    _auto_caption_sentences as _tag2_auto_caption_sentences,
    _compose_ollama_host,
    _extract_free_text_lines,
    _get_ollama_enable_free_text,
    _get_ollama_free_text_prompt_template,
    _get_ollama_group_prompt_template,
    _get_ollama_host,
    _get_ollama_model,
    _get_ollama_port,
    _get_ollama_prompt_template,
    _get_ollama_server,
    _get_ollama_timeout_seconds,
    _merge_free_text,
    _normalize_caption_line,
    _ollama_generate,
    _ollama_prompt_for_free_text,
    _ollama_prompt_for_group,
    _ollama_prompt_for_sentence,
    _parse_ollama_selection,
    _parse_ollama_yes_no,
    _split_ollama_host,
    _suggest_free_text as _tag2_suggest_free_text,
)
from tag2_sections import (
    _all_headers_from_sections,
    _all_sentences_from_sections,
    _apply_sentence_selection,
    _find_group_for_sentence,
    _get_crop_aspect_ratios,
    _get_folder_sections,
    _get_group_target,
    _group_hidden_sentences,
    _is_general_section_name,
    _is_hidden_group_sentence,
    _iter_caption_targets,
    _iter_caption_targets_with_indices,
    _normalize_enabled_sentences,
    _ordered_sections_for_output,
    _rename_sentence_in_sections,
    _set_folder_sections,
)

app = FastAPI(title="Image Captioning Tool")

# This tool is intended for local, user-selected image folders and should be able
# to open very large images without Pillow emitting decompression bomb warnings.
Image.MAX_IMAGE_PIXELS = None
warnings.simplefilter("ignore", Image.DecompressionBombWarning)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thread pool for IO-bound thumbnail generation
executor = ThreadPoolExecutor(max_workers=8)

# In-memory thumbnail cache lives in `tag2_images`.

# ===== CONFIG FILE =====
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
DEFAULT_OLLAMA_SERVER = "127.0.0.1"
DEFAULT_OLLAMA_PORT = 11434
DEFAULT_OLLAMA_TIMEOUT_SECONDS = 20
DEFAULT_OLLAMA_MODEL = "llava"
DEFAULT_CROP_ASPECT_RATIOS = ["4:3", "16:9", "3:4", "1:1", "9:16", "2:3", "3:2"]
DEFAULT_OLLAMA_PROMPT_TEMPLATE = (
    "You are verifying a single image caption. "
    "Reply with exactly one word: YES or NO. "
    "Reply YES only if the caption is clearly correct for the image. "
    "Reply NO if it is wrong, uncertain, too specific, or not clearly visible.\n\n"
    "Caption: {caption}\n"
    "Answer:"
)
DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE = (
    "You are selecting the single best caption for an image from a numbered list. "
    "Reply with exactly one number from 1 to {count}. "
    "Pick the most likely correct caption for the image.\n\n"
    "Group: {group_name}\n"
    "{options}\n\n"
    "Answer:"
)
DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE = (
    "You are improving an image caption file. The caption text below already covers known details and must not be repeated. "
    "Look at the image and return only notable, important visual details that are still missing. "
    "Return either NONE or one short line per missing detail, with no bullets or numbering.\n\n"
    "Current caption text:\n{caption_text}\n\n"
    "Answer:"
)


def _load_config() -> dict:
    """Load config from disk. Returns default structure if file missing."""
    default = {
        "last_folder": "",
        # Sentences shared across all folders when a folder has no own config yet
        "default_sentences": [],
        "crop_aspect_ratios": list(DEFAULT_CROP_ASPECT_RATIOS),
        "image_crops": {},
        # Local Ollama server settings
        "ollama_server": DEFAULT_OLLAMA_SERVER,
        "ollama_port": DEFAULT_OLLAMA_PORT,
        "ollama_timeout_seconds": DEFAULT_OLLAMA_TIMEOUT_SECONDS,
        "ollama_host": f"http://{DEFAULT_OLLAMA_SERVER}:{DEFAULT_OLLAMA_PORT}",
        "ollama_model": DEFAULT_OLLAMA_MODEL,
        "ollama_prompt_template": DEFAULT_OLLAMA_PROMPT_TEMPLATE,
        "ollama_group_prompt_template": DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE,
        "ollama_enable_free_text": True,
        "ollama_free_text_prompt_template": DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE,
        # Per-folder sentence lists. Key = absolute folder path.
        # To copy sentences to a new folder, duplicate a block here.
        "folders": {}
    }
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "ollama_server" not in data or "ollama_port" not in data:
                parsed_server, parsed_port = _split_ollama_host(
                    str(data.get("ollama_host") or default["ollama_host"]),
                    DEFAULT_OLLAMA_SERVER,
                    DEFAULT_OLLAMA_PORT,
                )
                data.setdefault("ollama_server", parsed_server)
                data.setdefault("ollama_port", parsed_port)
            data["ollama_host"] = _compose_ollama_host(
                data.get("ollama_server"),
                data.get("ollama_port"),
                DEFAULT_OLLAMA_SERVER,
                DEFAULT_OLLAMA_PORT,
            )
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


def _auto_caption_sentences(
    host: str,
    model: str,
    image_path: str,
    sentences: list[str],
    prompt_template: str | None = None,
    timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS,
) -> tuple[list[str], list[dict]]:
    """Compatibility adapter for sentence-level auto captioning."""
    return _tag2_auto_caption_sentences(
        host,
        model,
        image_path,
        sentences,
        encode_image_func=_encode_image_for_ollama,
        generate_func=_ollama_generate,
        prompt_template=prompt_template or DEFAULT_OLLAMA_PROMPT_TEMPLATE,
        timeout=timeout,
    )


def _auto_caption_sections(
    host: str,
    model: str,
    image_path: str,
    sections: list[dict],
    prompt_template: str | None = None,
    group_prompt_template: str | None = None,
    timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS,
) -> tuple[list[str], list[dict]]:
    """Compatibility adapter for section-aware auto captioning."""
    return _tag2_auto_caption_sections(
        host,
        model,
        image_path,
        sections,
        encode_image_func=_encode_image_for_ollama,
        generate_func=_ollama_generate,
        prompt_template=prompt_template or DEFAULT_OLLAMA_PROMPT_TEMPLATE,
        group_prompt_template=group_prompt_template or DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE,
        timeout=timeout,
    )


def _suggest_free_text(
    host: str,
    model: str,
    image_path: str,
    caption_text: str,
    prompt_template: str | None = None,
    timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS,
) -> str:
    """Compatibility adapter for Ollama free-text suggestions."""
    return _tag2_suggest_free_text(
        host,
        model,
        image_path,
        caption_text,
        encode_image_func=_encode_image_for_ollama,
        generate_func=_ollama_generate,
        prompt_template=prompt_template or DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE,
        timeout=timeout,
    )


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
    data = await loop.run_in_executor(executor, _get_thumbnail, path, actual_size, None)

    if not data:
        raise HTTPException(status_code=500, detail="Failed to generate thumbnail")

    return StreamingResponse(BytesIO(data), media_type="image/jpeg")


@app.get("/api/image")
async def get_image(path: str = Query(...)):
    """Serve the full-resolution image."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")

    loop = asyncio.get_event_loop()
    data, media_type = await loop.run_in_executor(executor, _render_image_bytes, path, None, None, False)
    return StreamingResponse(BytesIO(data), media_type=media_type)


@app.get("/api/preview")
async def get_preview(path: str = Query(...), ignore_crop: bool = Query(default=False)):
    """Serve a preview-quality image (downscaled to max 2048px edge, JPEG).
    Much faster than serving the full-res file for large images."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")

    loop = asyncio.get_event_loop()
    data = await loop.run_in_executor(executor, _get_thumbnail, path, PREVIEW_MAX_SIZE, None)

    if not data:
        rendered, media_type = await loop.run_in_executor(executor, _render_image_bytes, path, None, PREVIEW_MAX_SIZE, True)
        return StreamingResponse(BytesIO(rendered), media_type=media_type)

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


class RenameCaptionPresetUpdate(BaseModel):
    folder: str
    old_sentence: str
    new_sentence: str


class CropUpdate(BaseModel):
    image_path: str
    crop: Optional[dict] = None

class RotateUpdate(BaseModel):
    image_path: str
    direction: str


@app.get("/api/crop")
async def get_crop(path: str = Query(...)):
    """Get whether an image currently has a reversible real crop applied."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Image file not found")
    return {
        "path": path,
        "crop": _get_image_crop(path),
    }


@app.post("/api/crop")
async def save_crop(data: CropUpdate):
    """Apply or remove a real crop on the image file."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image file not found")
    cfg = _load_config()
    try:
        if data.crop is None:
            _remove_real_crop(data.image_path)
        else:
            _apply_real_crop(data.image_path, data.crop)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid crop: {e}") from e
    if isinstance(cfg.get("image_crops"), dict):
        cfg["image_crops"].pop(_normalize_image_key(data.image_path), None)
    _save_config(cfg)
    return {
        "ok": True,
        "crop": _get_image_crop(data.image_path),
    }

@app.post("/api/rotate")
async def rotate_image(data: RotateUpdate):
    """Rotate an image by 90 degrees left or right."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image file not found")
    try:
        crop_state = _rotate_image(data.image_path, data.direction)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid rotation: {e}") from e
    return {
        "ok": True,
        "crop": crop_state,
    }


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
    headers = _all_headers_from_sections(sections)

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
        headers = _all_headers_from_sections(sections)
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


@app.post("/api/caption/rename-preset")
async def rename_caption_preset(update: RenameCaptionPresetUpdate):
    """Rename a configured caption preset and migrate existing caption files."""
    folder = os.path.normpath(update.folder)
    old_sentence = (update.old_sentence or "").strip()
    new_sentence = (update.new_sentence or "").strip()
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")
    if not old_sentence or not new_sentence:
        raise HTTPException(status_code=400, detail="Both old and new caption text are required")
    if old_sentence == new_sentence:
        cfg = _load_config()
        return {"ok": True, "sections": _get_folder_sections(cfg, folder)}

    cfg = _load_config()
    sections = _get_folder_sections(cfg, folder)
    all_sentences_before = _all_sentences_from_sections(sections)
    if old_sentence not in all_sentences_before:
        raise HTTPException(status_code=404, detail="Caption not found")
    if new_sentence in all_sentences_before:
        raise HTTPException(status_code=400, detail="A caption with that text already exists")

    renamed = _rename_sentence_in_sections(sections, old_sentence, new_sentence)
    if not renamed:
        raise HTTPException(status_code=404, detail="Caption not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)

    headers = _all_headers_from_sections(sections)
    folder_path = Path(folder)
    for entry in folder_path.iterdir():
        if not entry.is_file() or entry.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_sentences_before, headers)
        enabled = [new_sentence if sentence == old_sentence else sentence for sentence in data.get("enabled_sentences", [])]
        enabled = _normalize_enabled_sentences(enabled, sections)
        _write_caption_file(str(entry), enabled, data.get("free_text", ""), sections)

    return {"ok": True, "sections": sections}


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


class AutoCaptionRequest(BaseModel):
    image_path: Optional[str] = None
    image_paths: Optional[list[str]] = None
    model: Optional[str] = None
    prompt_template: Optional[str] = None
    group_prompt_template: Optional[str] = None
    enable_free_text: Optional[bool] = None
    free_text_only: Optional[bool] = None
    target_section_index: Optional[int] = None
    target_group_index: Optional[int] = None
    free_text_prompt_template: Optional[str] = None
    timeout_seconds: Optional[int] = None


@app.post("/api/caption/save")
async def save_caption(data: SaveCaptionFull):
    """Full save of caption data for one image."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image not found")
    # Load sections config for proper formatting
    cfg = _load_config()
    folder = str(Path(data.image_path).parent)
    sections = _get_folder_sections(cfg, folder)
    enabled_sentences = _normalize_enabled_sentences(data.enabled_sentences, sections)
    _write_caption_file(data.image_path, enabled_sentences, data.free_text, sections)
    return {"ok": True}


@app.post("/api/auto-caption")
async def auto_caption(data: AutoCaptionRequest):
    """Ask a local Ollama vision model to verify each configured caption."""
    image_path = data.image_path or (data.image_paths[0] if data.image_paths else None)
    if not image_path or not os.path.isfile(image_path):
        raise HTTPException(status_code=404, detail="Image not found")

    cfg = _load_config()
    model = (data.model or _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL)).strip()
    host = _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS)

    if not model:
        raise HTTPException(status_code=400, detail="No Ollama model configured")

    folder = str(Path(image_path).parent)
    sections = _get_folder_sections(cfg, folder)
    all_sentences = _all_sentences_from_sections(sections)
    if not all_sentences and not free_text_only:
        raise HTTPException(status_code=400, detail="No captions configured for this folder")
    target_group = _get_group_target(sections, data.target_section_index, data.target_group_index)
    if data.target_section_index is not None or data.target_group_index is not None:
        if not target_group:
            raise HTTPException(status_code=400, detail="Invalid target group")

    headers = _all_headers_from_sections(sections)
    existing = _read_caption_file(image_path, all_sentences, headers)
    existing_free_text = existing.get("free_text", "")
    enabled = list(existing.get("enabled_sentences", []))
    results: list[dict] = []

    loop = asyncio.get_event_loop()
    if not free_text_only:
        try:
            if target_group:
                payload = {
                    "model": model,
                    "prompt": _ollama_prompt_for_group(target_group.get("group_name", ""), target_group["sentences"], group_prompt_template),
                    "images": [await loop.run_in_executor(executor, _encode_image_for_ollama, image_path)],
                    "stream": False,
                    "options": {"temperature": 0},
                }
                response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                raw_answer = str(response.get("response") or "").strip()
                selection_index = _parse_ollama_selection(raw_answer, target_group["sentences"])
                selected_sentence = target_group["sentences"][selection_index - 1] if selection_index else None
                enabled = [sentence for sentence in enabled if sentence not in target_group["sentences"]]
                if selected_sentence:
                    enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
                enabled = _normalize_enabled_sentences(enabled, sections)
                results = [{
                    "type": "group",
                    "section_index": target_group["section_index"],
                    "group_index": target_group["group_index"],
                    "group_name": target_group.get("group_name", ""),
                    "sentences": target_group["sentences"],
                    "selected_sentence": selected_sentence,
                    "selection_index": selection_index,
                    "answer": raw_answer,
                }]
            else:
                enabled, results = await loop.run_in_executor(
                    executor,
                    _auto_caption_sections,
                    host,
                    model,
                    image_path,
                    sections,
                    prompt_template,
                    group_prompt_template,
                    timeout_seconds,
                )
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption failed: {e}") from e

    free_text = existing_free_text if target_group else ""
    free_text_model_output = ""
    added_free_text_lines: list[str] = []
    if enable_free_text:
        caption_text = _build_caption_text(enabled, free_text, sections)
        try:
            free_text_model_output = await loop.run_in_executor(
                executor,
                _suggest_free_text,
                host,
                model,
                image_path,
                caption_text,
                free_text_prompt_template,
                timeout_seconds,
            )
            free_text, added_free_text_lines = _merge_free_text("", free_text_model_output, enabled)
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption free-text step failed: {e}") from e

    _write_caption_file(image_path, enabled, free_text, sections)
    return {
        "ok": True,
        "model": model,
        "host": host,
        "timeout_seconds": timeout_seconds,
        "prompt_template": prompt_template,
        "group_prompt_template": group_prompt_template,
        "enable_free_text": enable_free_text,
        "free_text_only": free_text_only,
        "free_text_prompt_template": free_text_prompt_template,
        "enabled_sentences": enabled,
        "free_text": free_text,
        "free_text_model_output": free_text_model_output,
        "added_free_text_lines": added_free_text_lines,
        "results": results,
    }


@app.post("/api/auto-caption/stream")
async def auto_caption_stream(data: AutoCaptionRequest, request: Request):
    """Stream real-time auto-caption progress for one or more images as NDJSON."""
    image_paths = data.image_paths or ([data.image_path] if data.image_path else [])
    if not image_paths:
        raise HTTPException(status_code=400, detail="No images provided")

    cfg = _load_config()
    model = (data.model or _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL)).strip()
    host = _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS)

    if not model:
        raise HTTPException(status_code=400, detail="No Ollama model configured")

    def _event_bytes(payload: dict) -> bytes:
        return (json.dumps(payload, ensure_ascii=False) + "\n").encode("utf-8")

    async def event_stream():
        loop = asyncio.get_event_loop()
        total_processed = 0
        total_errors = 0
        yield _event_bytes({
            "type": "start",
            "count": len(image_paths),
            "model": model,
            "host": host,
            "group_prompt_template": group_prompt_template,
            "enable_free_text": enable_free_text,
            "free_text_only": free_text_only,
            "timeout_seconds": timeout_seconds,
        })

        for image_path in image_paths:
            if await request.is_disconnected():
                break
            if not os.path.isfile(image_path):
                total_errors += 1
                yield _event_bytes({"type": "error", "path": image_path, "message": "Image not found"})
                continue

            folder = str(Path(image_path).parent)
            sections = _get_folder_sections(cfg, folder)
            all_sentences = _all_sentences_from_sections(sections)
            if not all_sentences and not free_text_only:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": image_path, "message": "No captions configured for this folder"})
                continue

            target_group = _get_group_target(sections, data.target_section_index, data.target_group_index)
            if data.target_section_index is not None or data.target_group_index is not None:
                if not target_group:
                    total_errors += 1
                    yield _event_bytes({"type": "error", "path": image_path, "message": "Invalid target group"})
                    continue

            headers = _all_headers_from_sections(sections)
            existing = _read_caption_file(image_path, all_sentences, headers)
            free_text = existing.get("free_text", "") if target_group else ""
            enabled = list(existing.get("enabled_sentences", []))
            results = []
            image_b64 = await loop.run_in_executor(executor, _encode_image_for_ollama, image_path)

            total_targets = 0 if free_text_only else (1 if target_group else sum(1 for _ in _iter_caption_targets_with_indices(sections)))
            yield _event_bytes({
                "type": "image-start",
                "path": image_path,
                "total_sentences": len(all_sentences),
                "total_targets": total_targets,
                "free_text_only": free_text_only,
                "target_group": {
                    "section_index": target_group["section_index"],
                    "group_index": target_group["group_index"],
                    "group_name": target_group.get("group_name", ""),
                } if target_group else None,
                "free_text": free_text,
            })

            try:
                if not free_text_only:
                    targets = [target_group] if target_group else list(_iter_caption_targets_with_indices(sections))
                    for index, target in enumerate(targets, start=1):
                        if await request.is_disconnected():
                            return
                        if target["type"] == "sentence":
                            sentence = target["sentence"]
                            payload = {
                                "model": model,
                                "prompt": _ollama_prompt_for_sentence(sentence, prompt_template),
                                "images": [image_b64],
                                "stream": False,
                                "options": {"temperature": 0},
                            }
                            response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                            raw_answer = str(response.get("response") or "").strip()
                            is_match = _parse_ollama_yes_no(raw_answer)
                            enabled = _apply_sentence_selection(enabled, sentence, sections, is_match)
                            result = {"type": "sentence", "sentence": sentence, "enabled": is_match, "answer": raw_answer}
                            results.append(result)
                            _write_caption_file(image_path, enabled, free_text, sections)
                            yield _event_bytes({
                                "type": "caption-check",
                                "path": image_path,
                                "index": index,
                                "total": total_targets,
                                "sentence": sentence,
                                "enabled": is_match,
                                "answer": raw_answer,
                            })
                            continue

                        group_sentences = target["sentences"]
                        payload = {
                            "model": model,
                            "prompt": _ollama_prompt_for_group(target.get("group_name", ""), group_sentences, group_prompt_template),
                            "images": [image_b64],
                            "stream": False,
                            "options": {"temperature": 0},
                        }
                        response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                        raw_answer = str(response.get("response") or "").strip()
                        selection_index = _parse_ollama_selection(raw_answer, group_sentences)
                        selected_sentence = group_sentences[selection_index - 1] if selection_index else None
                        selected_hidden = bool(selected_sentence and _is_hidden_group_sentence(sections, selected_sentence))
                        enabled = [sentence for sentence in enabled if sentence not in group_sentences]
                        if selected_sentence:
                            enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
                        result = {
                            "type": "group",
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "group_name": target.get("group_name", ""),
                            "sentences": group_sentences,
                            "selected_sentence": selected_sentence,
                            "selected_hidden": selected_hidden,
                            "selection_index": selection_index,
                            "answer": raw_answer,
                        }
                        results.append(result)
                        _write_caption_file(image_path, enabled, free_text, sections)
                        yield _event_bytes({
                            "type": "group-selection",
                            "path": image_path,
                            "index": index,
                            "total": total_targets,
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "group_name": target.get("group_name", ""),
                            "sentences": group_sentences,
                            "selected_sentence": selected_sentence,
                            "selected_hidden": selected_hidden,
                            "selection_index": selection_index,
                            "answer": raw_answer,
                        })

                free_text_model_output = ""
                added_free_text_lines: list[str] = []
                if enable_free_text:
                    if await request.is_disconnected():
                        return
                    caption_text = _build_caption_text(enabled, free_text, sections)
                    free_payload = {
                        "model": model,
                        "prompt": _ollama_prompt_for_free_text(caption_text, free_text_prompt_template),
                        "images": [image_b64],
                        "stream": False,
                        "options": {"temperature": 0},
                    }
                    response = await loop.run_in_executor(executor, _ollama_generate, host, free_payload, timeout_seconds)
                    free_text_model_output = str(response.get("response") or "").strip()
                    free_text, added_free_text_lines = _merge_free_text("", free_text_model_output, enabled)
                    _write_caption_file(image_path, enabled, free_text, sections)
                    yield _event_bytes({
                        "type": "free-text",
                        "path": image_path,
                        "answer": free_text_model_output,
                        "free_text": free_text,
                        "added_lines": added_free_text_lines,
                    })

                _write_caption_file(image_path, enabled, free_text, sections)
                total_processed += 1
                yield _event_bytes({
                    "type": "image-complete",
                    "path": image_path,
                    "free_text_only": free_text_only,
                    "enabled_sentences": enabled,
                    "free_text": free_text,
                    "results": results,
                })
            except RuntimeError as e:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": image_path, "message": f"Ollama error: {e}"})
            except Exception as e:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": image_path, "message": f"Auto caption failed: {e}"})

        yield _event_bytes({
            "type": "done",
            "processed": total_processed,
            "errors": total_errors,
            "count": len(image_paths),
        })

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


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
        headers = _all_headers_from_sections(sections)
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
    crop_aspect_ratios: Optional[list[str]] = None
    ollama_host: Optional[str] = None
    ollama_server: Optional[str] = None
    ollama_port: Optional[int] = None
    ollama_timeout_seconds: Optional[int] = None
    ollama_model: Optional[str] = None
    ollama_prompt_template: Optional[str] = None
    ollama_group_prompt_template: Optional[str] = None
    ollama_enable_free_text: Optional[bool] = None
    ollama_free_text_prompt_template: Optional[str] = None


@app.get("/api/settings")
async def get_settings(folder: Optional[str] = Query(default=None)):
    """Get full settings. If folder is specified, include sections for that folder."""
    cfg = _load_config()
    result = {
        "last_folder": cfg.get("last_folder", ""),
        "crop_aspect_ratios": _get_crop_aspect_ratios(cfg, DEFAULT_CROP_ASPECT_RATIOS),
        "ollama_host": _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_server": _get_ollama_server(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_port": _get_ollama_port(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_timeout_seconds": _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS),
        "ollama_model": _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL),
        "ollama_prompt_template": _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE),
        "ollama_group_prompt_template": _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE),
        "ollama_enable_free_text": _get_ollama_enable_free_text(cfg),
        "ollama_free_text_prompt_template": _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE),
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
    if data.crop_aspect_ratios is not None:
        cfg["crop_aspect_ratios"] = [str(r).strip() for r in data.crop_aspect_ratios if str(r).strip()] or list(DEFAULT_CROP_ASPECT_RATIOS)
    if data.ollama_host is not None:
        server_name, server_port = _split_ollama_host(data.ollama_host, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
        cfg["ollama_server"] = server_name
        cfg["ollama_port"] = server_port
        cfg["ollama_host"] = _compose_ollama_host(server_name, server_port, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    if data.ollama_server is not None:
        cfg["ollama_server"] = data.ollama_server.strip() or DEFAULT_OLLAMA_SERVER
    if data.ollama_port is not None:
        cfg["ollama_port"] = int(data.ollama_port)
    if data.ollama_timeout_seconds is not None:
        cfg["ollama_timeout_seconds"] = max(1, int(data.ollama_timeout_seconds))
    if data.ollama_model is not None:
        cfg["ollama_model"] = data.ollama_model.strip()
    if data.ollama_prompt_template is not None:
        cfg["ollama_prompt_template"] = data.ollama_prompt_template or DEFAULT_OLLAMA_PROMPT_TEMPLATE
    if data.ollama_group_prompt_template is not None:
        cfg["ollama_group_prompt_template"] = data.ollama_group_prompt_template or DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE
    if data.ollama_enable_free_text is not None:
        cfg["ollama_enable_free_text"] = bool(data.ollama_enable_free_text)
    if data.ollama_free_text_prompt_template is not None:
        cfg["ollama_free_text_prompt_template"] = data.ollama_free_text_prompt_template or DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE
    cfg["ollama_host"] = _compose_ollama_host(cfg.get("ollama_server"), cfg.get("ollama_port"), DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
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
