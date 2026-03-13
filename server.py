#!/usr/bin/env python3
"""
Image Captioning Tool - Backend Server
FastAPI-based server for browsing images and managing captions.
"""

import os
import sys
import json
import asyncio
import copy
import ipaddress
import subprocess
import platform
import shutil
import threading
import urllib.error
import urllib.request
import warnings
from pathlib import Path
from io import BytesIO
from functools import partial
from typing import Optional
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse, HTMLResponse, PlainTextResponse
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
    _auto_caption_sections,
    _auto_caption_sentences,
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
    _suggest_free_text,
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
    _remove_group_from_sections,
    _remove_section_from_sections,
    _remove_sentence_from_sections,
    _rename_section_in_sections,
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
DEFAULT_HTTPS_CERTFILE = ""
DEFAULT_HTTPS_KEYFILE = ""
DEFAULT_HTTPS_PORT = 8900
DEFAULT_REMOTE_HTTP_MODE = "redirect-to-https"
REMOTE_HTTP_MODES = {"allow", "redirect-to-https", "block"}
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


def _list_ollama_models(host: str, timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS) -> list[str]:
    """Return the list of available Ollama model names from /api/tags."""
    url = f"{host.rstrip('/')}/api/tags"
    request = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(detail or str(exc)) from exc
    except TimeoutError as exc:
        raise RuntimeError(f"request timed out after {timeout} seconds") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason or exc)) from exc

    models: list[str] = []
    seen: set[str] = set()
    for item in payload.get("models", []) or []:
        name = str((item or {}).get("name") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        models.append(name)
    return models


def _load_config() -> dict:
    """Load config from disk. Returns default structure if file missing."""
    default = {
        "last_folder": "",
        # Sentences shared across all folders when a folder has no own config yet
        "default_sentences": [],
        "thumb_size": 160,
        "crop_aspect_ratios": list(DEFAULT_CROP_ASPECT_RATIOS),
        "image_crops": {},
        "https_certfile": DEFAULT_HTTPS_CERTFILE,
        "https_keyfile": DEFAULT_HTTPS_KEYFILE,
        "https_port": DEFAULT_HTTPS_PORT,
        "remote_http_mode": DEFAULT_REMOTE_HTTP_MODE,
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


def _resolve_config_path(path_value: str | None) -> str:
    """Resolve a config path relative to the config file directory when needed."""
    raw_path = str(path_value or "").strip()
    if not raw_path:
        return ""
    if os.path.isabs(raw_path):
        return raw_path
    return os.path.abspath(os.path.join(os.path.dirname(CONFIG_PATH), raw_path))


def _has_https_config(cfg: dict) -> bool:
    """Return whether both HTTPS file paths are configured."""
    return bool(str(cfg.get("https_certfile") or "").strip() and str(cfg.get("https_keyfile") or "").strip())


def _get_https_port(cfg: dict, default_port: int = DEFAULT_HTTPS_PORT) -> int:
    """Return the configured HTTPS listener port."""
    try:
        port = int(cfg.get("https_port", default_port))
    except (TypeError, ValueError):
        port = default_port
    return max(1, min(65535, port))


def _get_remote_http_mode(cfg: dict, default_mode: str = DEFAULT_REMOTE_HTTP_MODE) -> str:
    """Return how remote HTTP requests should be handled when HTTPS is configured."""
    mode = str(cfg.get("remote_http_mode") or default_mode).strip().lower()
    if mode not in REMOTE_HTTP_MODES:
        return default_mode
    return mode


def _is_local_client_host(host: str | None) -> bool:
    """Return whether the request client is the local machine."""
    value = str(host or "").strip().lower()
    if not value:
        return False
    if value in {"localhost", "127.0.0.1", "::1", "testclient"}:
        return True
    try:
        return ipaddress.ip_address(value).is_loopback
    except ValueError:
        return False


def _build_https_redirect_url(request: Request, https_port: int) -> str:
    """Build the HTTPS redirect URL for an incoming HTTP request."""
    host = request.url.hostname or "localhost"
    if ":" in host and not host.startswith("["):
        host = f"[{host}]"
    netloc = host if https_port == 443 else f"{host}:{https_port}"
    return str(request.url.replace(scheme="https", netloc=netloc))


def _get_remote_http_response(request: Request, cfg: dict):
    """Return a response for remote HTTP requests when HTTPS is configured, else None."""
    if request.url.scheme == "https" or not _has_https_config(cfg):
        return None
    if _is_local_client_host(request.client.host if request.client else None):
        return None

    mode = _get_remote_http_mode(cfg)
    if mode == "allow":
        return None
    if mode == "block":
        return PlainTextResponse("Remote HTTP access is disabled. Use HTTPS instead.", status_code=403)
    return RedirectResponse(_build_https_redirect_url(request, _get_https_port(cfg)), status_code=307)


def _get_https_uvicorn_kwargs(cfg: dict) -> dict:
    """Return Uvicorn SSL kwargs for an optional local HTTPS certificate."""
    certfile = _resolve_config_path(cfg.get("https_certfile"))
    keyfile = _resolve_config_path(cfg.get("https_keyfile"))

    if not certfile and not keyfile:
        return {}
    if not certfile or not keyfile:
        raise RuntimeError("HTTPS requires both https_certfile and https_keyfile in config.json")
    if not os.path.isfile(certfile):
        raise RuntimeError(f"HTTPS certificate file not found: {certfile}")
    if not os.path.isfile(keyfile):
        raise RuntimeError(f"HTTPS key file not found: {keyfile}")

    return {
        "ssl_certfile": certfile,
        "ssl_keyfile": keyfile,
    }


@app.middleware("http")
async def enforce_remote_http_policy(request: Request, call_next):
    """Allow local HTTP while redirecting or blocking remote HTTP when HTTPS is configured."""
    response = _get_remote_http_response(request, _load_config())
    if response is not None:
        return response
    return await call_next(request)


def _read_live_caption_state(
    image_path: str,
    all_sentences: list[str],
    headers: list[str],
) -> tuple[list[str], str]:
    """Read the latest caption state from disk for merge-safe auto-caption writes."""
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")
    current = _read_caption_file(image_path, all_sentences, headers)
    return list(current.get("enabled_sentences", [])), current.get("free_text", "")


def _apply_sentence_result_to_live_caption(
    image_path: str,
    sentence: str,
    should_enable: bool,
    all_sentences: list[str],
    headers: list[str],
    sections: list[dict],
) -> tuple[list[str], str]:
    """Apply a single AI sentence verdict on top of the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
    enabled = _apply_sentence_selection(enabled, sentence, sections, should_enable)
    enabled = _normalize_enabled_sentences(enabled, sections)
    _write_caption_file(image_path, enabled, free_text, sections)
    return enabled, free_text


def _apply_group_result_to_live_caption(
    image_path: str,
    group_sentences: list[str],
    selected_sentence: str | None,
    all_sentences: list[str],
    headers: list[str],
    sections: list[dict],
) -> tuple[list[str], str]:
    """Apply a single AI group selection on top of the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
    enabled = [sentence for sentence in enabled if sentence not in group_sentences]
    if selected_sentence:
        enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
    enabled = _normalize_enabled_sentences(enabled, sections)
    _write_caption_file(image_path, enabled, free_text, sections)
    return enabled, free_text


def _apply_free_text_result_to_live_caption(
    image_path: str,
    model_output: str,
    all_sentences: list[str],
    headers: list[str],
    sections: list[dict],
    preserve_existing: bool = False,
) -> tuple[list[str], str, list[str]]:
    """Merge AI free-text output with the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
    base_free_text = free_text if preserve_existing else ""
    merged_free_text, added_lines = _merge_free_text(base_free_text, model_output, enabled)
    _write_caption_file(image_path, enabled, merged_free_text, sections)
    return enabled, merged_free_text, added_lines


@app.get("/api/list-images")
async def list_images(folder: str = Query(...)):
    """List all image files in the given folder."""
    folder_path = _resolve_folder_path(folder)

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


@app.post("/api/images/upload")
async def upload_images(
    folder: str = Form(...),
    files: Optional[list[UploadFile]] = File(default=None),
):
    """Upload image files into the currently loaded folder."""
    folder_path = _resolve_folder_path(folder)
    uploads = [upload for upload in (files or []) if upload is not None]
    if not uploads:
        raise HTTPException(status_code=400, detail="No files provided")

    uploaded: list[dict] = []
    skipped: list[dict] = []

    for upload in uploads:
        try:
            original_name = _sanitize_upload_filename(upload.filename or "")
            suffix = Path(original_name).suffix.lower()
            if suffix not in IMAGE_EXTENSIONS:
                skipped.append({
                    "name": original_name,
                    "reason": "Unsupported image file",
                })
                continue

            destination = _get_unique_upload_path(folder_path, original_name)
            destination.parent.mkdir(parents=True, exist_ok=True)
            upload.file.seek(0)
            with destination.open("wb") as handle:
                shutil.copyfileobj(upload.file, handle)

            _clear_thumbnail_cache_for_path(str(destination))
            stat = destination.stat()
            uploaded.append({
                "name": destination.name,
                "path": str(destination),
                "size": stat.st_size,
                "mtime": stat.st_mtime,
                "source_name": original_name,
                "renamed": destination.name != original_name,
            })
        except HTTPException:
            raise
        except Exception as exc:
            skipped.append({
                "name": str(upload.filename or "").strip() or "(unnamed)",
                "reason": str(exc),
            })
        finally:
            await upload.close()

    return {
        "ok": len(uploaded) > 0,
        "folder": str(folder_path),
        "uploaded": uploaded,
        "uploaded_count": len(uploaded),
        "skipped": skipped,
        "skipped_count": len(skipped),
        "renamed_count": sum(1 for item in uploaded if item["renamed"]),
    }


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


class RenameSectionUpdate(BaseModel):
    folder: str
    old_name: str
    new_name: str


class DeleteCaptionPresetUpdate(BaseModel):
    folder: str
    sentence: str


class DeleteGroupUpdate(BaseModel):
    folder: str
    section_index: int
    group_index: int


class DeleteSectionUpdate(BaseModel):
    folder: str
    section_index: int


class CropUpdate(BaseModel):
    image_path: str
    crop: Optional[dict] = None

class RotateUpdate(BaseModel):
    image_path: str
    direction: str


class BatchDeleteImagesRequest(BaseModel):
    image_paths: list[str]


class CloneFolderRequest(BaseModel):
    source_folder: str
    new_folder_name: str
    image_paths: list[str] = []


def _event_bytes(event: dict) -> bytes:
    return (json.dumps(event, ensure_ascii=False) + "\n").encode("utf-8")


def _normalize_folder_clone_name(name: str) -> str:
    normalized = str(name or "").strip()
    if not normalized:
        raise HTTPException(status_code=400, detail="New folder name is required")
    if normalized in {".", ".."}:
        raise HTTPException(status_code=400, detail="Invalid folder name")
    if os.path.basename(normalized) != normalized:
        raise HTTPException(status_code=400, detail="Folder name must not include path separators")
    if any(char in normalized for char in '<>:"/\\|?*'):
        raise HTTPException(status_code=400, detail="Folder name contains invalid characters")
    return normalized


def _resolve_folder_path(folder: str, *, detail: str = "Not a valid directory") -> Path:
    folder_path = Path(str(folder or "")).resolve()
    if not folder_path.is_dir():
        raise HTTPException(status_code=400, detail=detail)
    return folder_path


def _sanitize_upload_filename(name: str) -> str:
    raw_name = str(name or "").replace("\\", "/")
    base_name = os.path.basename(raw_name).strip()
    if not base_name or base_name in {".", ".."}:
        raise HTTPException(status_code=400, detail="Upload contains an invalid filename")
    return base_name


def _get_unique_upload_path(folder_path: Path, file_name: str) -> Path:
    candidate = folder_path / file_name
    if not candidate.exists():
        return candidate

    stem = candidate.stem or "image"
    suffix = candidate.suffix
    counter = 1
    while True:
        renamed = folder_path / f"{stem} ({counter}){suffix}"
        if not renamed.exists():
            return renamed
        counter += 1


def _prepare_clone_plan(source_folder: str, requested_name: str, image_paths: list[str]) -> dict:
    source_path = Path(source_folder).resolve()
    if not source_path.is_dir():
        raise HTTPException(status_code=400, detail="Source folder is not a valid directory")

    target_name = _normalize_folder_clone_name(requested_name)
    target_path = source_path.parent / target_name
    if target_path.exists():
        raise HTTPException(status_code=400, detail="Target folder already exists")

    normalized_source = os.path.normcase(str(source_path))
    normalized_selected: list[Path] = []
    for raw_path in image_paths or []:
        candidate = Path(str(raw_path or "")).resolve()
        if not candidate.is_file():
            raise HTTPException(status_code=400, detail=f"Image not found: {candidate}")
        if os.path.normcase(str(candidate.parent)) != normalized_source:
            raise HTTPException(status_code=400, detail="Selected images must belong to the current folder")
        if candidate.suffix.lower() not in IMAGE_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"Unsupported image file: {candidate.name}")
        if candidate not in normalized_selected:
            normalized_selected.append(candidate)

    selected_mode = len(normalized_selected) > 1
    copy_items: list[tuple[Path, Path]] = []
    copied_image_count = 0

    if selected_mode:
        for image_path in normalized_selected:
            dest_image = target_path / image_path.name
            copy_items.append((image_path, dest_image))
            copied_image_count += 1
            caption_path = _get_caption_path(str(image_path))
            if caption_path.exists():
                copy_items.append((caption_path, target_path / caption_path.name))
    else:
        for entry in sorted(source_path.iterdir(), key=lambda item: item.name.lower()):
            dest_entry = target_path / entry.name
            copy_items.append((entry, dest_entry))
            if entry.is_file() and entry.suffix.lower() in IMAGE_EXTENSIONS:
                copied_image_count += 1

    if copied_image_count == 0:
        raise HTTPException(status_code=400, detail="No images available to clone")

    return {
        "source_path": source_path,
        "target_path": target_path,
        "selected_mode": selected_mode,
        "copy_items": copy_items,
        "copied_image_count": copied_image_count,
    }


def _copy_folder_config_for_clone(source_folder: str, target_folder: str):
    cfg = _load_config()
    folders_cfg = cfg.setdefault("folders", {})
    source_key = os.path.normpath(source_folder)
    target_key = os.path.normpath(target_folder)
    if source_key in folders_cfg:
        folders_cfg[target_key] = copy.deepcopy(folders_cfg[source_key])
    else:
        sections = _get_folder_sections(cfg, source_folder)
        _set_folder_sections(cfg, target_folder, sections)
    _save_config(cfg)


def _iter_folder_image_entries(folder: str):
    """Yield image files in a folder."""
    folder_path = Path(folder)
    for entry in folder_path.iterdir():
        if entry.is_file() and entry.suffix.lower() in IMAGE_EXTENSIONS:
            yield entry


def _remove_deleted_caption_lines(free_text: str, removed_sentences: set[str]) -> str:
    """Drop free-text lines that exactly match deleted captions."""
    if not removed_sentences:
        return str(free_text or "")
    filtered_lines = [
        line
        for line in str(free_text or "").splitlines()
        if line.strip() not in removed_sentences
    ]
    return "\n".join(filtered_lines)


def _remove_sentences_from_caption_files(folder: str, sections_before: list[dict], sections_after: list[dict], removed_sentences: list[str]):
    """Rewrite caption files in a folder after configured captions are deleted."""
    removed_set = {sentence for sentence in removed_sentences if sentence}
    if not removed_set:
        return

    all_sentences_before = _all_sentences_from_sections(sections_before)
    headers_before = _all_headers_from_sections(sections_before)

    for entry in _iter_folder_image_entries(folder):
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_sentences_before, headers_before)
        enabled = [
            sentence for sentence in data.get("enabled_sentences", [])
            if sentence not in removed_set
        ]
        enabled = _normalize_enabled_sentences(enabled, sections_after)
        free_text = _remove_deleted_caption_lines(str(data.get("free_text", "") or ""), removed_set)
        _write_caption_file(str(entry), enabled, free_text, sections_after)


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


@app.post("/api/images/delete")
async def delete_images(data: BatchDeleteImagesRequest):
    """Delete image files and their local sidecar artifacts."""
    image_paths = [str(path or "").strip() for path in data.image_paths]
    image_paths = [path for path in image_paths if path]
    if not image_paths:
        raise HTTPException(status_code=400, detail="No image paths provided")

    cfg = _load_config()
    image_crops = cfg.get("image_crops") if isinstance(cfg.get("image_crops"), dict) else None
    deleted_paths: list[str] = []
    errors: list[dict] = []
    config_changed = False

    for image_path in image_paths:
        normalized_path = os.path.abspath(os.path.normpath(image_path))
        path_obj = Path(normalized_path)

        if path_obj.suffix.lower() not in IMAGE_EXTENSIONS:
            errors.append({"path": normalized_path, "error": "Unsupported image file"})
            continue
        if not path_obj.is_file():
            errors.append({"path": normalized_path, "error": "File not found"})
            continue

        caption_path = _get_caption_path(normalized_path)
        backup_path = _get_crop_backup_path(normalized_path)

        try:
            path_obj.unlink()
            if caption_path.exists():
                caption_path.unlink()
            if os.path.isfile(backup_path):
                os.remove(backup_path)
            _clear_thumbnail_cache_for_path(normalized_path)
            if image_crops is not None and image_crops.pop(_normalize_image_key(normalized_path), None) is not None:
                config_changed = True
            deleted_paths.append(normalized_path)
        except PermissionError:
            errors.append({"path": normalized_path, "error": "Permission denied"})
        except OSError as exc:
            errors.append({"path": normalized_path, "error": str(exc)})

    if config_changed:
        _save_config(cfg)

    return {
        "ok": len(errors) == 0,
        "deleted_paths": deleted_paths,
        "deleted_count": len(deleted_paths),
        "errors": errors,
    }


@app.post("/api/folder/clone/stream")
async def clone_folder_stream(data: CloneFolderRequest):
    """Clone the current folder or a multi-selection into a sibling folder with progress events."""
    plan = _prepare_clone_plan(data.source_folder, data.new_folder_name, data.image_paths or [])
    source_path: Path = plan["source_path"]
    target_path: Path = plan["target_path"]
    copy_items: list[tuple[Path, Path]] = plan["copy_items"]
    selected_mode = bool(plan["selected_mode"])
    copied_image_count = int(plan["copied_image_count"])

    async def event_stream():
        try:
            target_path.mkdir(parents=True, exist_ok=False)
            yield _event_bytes({
                "type": "start",
                "mode": "selected" if selected_mode else "folder",
                "source_folder": str(source_path),
                "target_folder": str(target_path),
                "total": len(copy_items),
                "image_count": copied_image_count,
            })

            copied_count = 0
            for index, (src, dest) in enumerate(copy_items, start=1):
                if src.is_dir():
                    shutil.copytree(src, dest)
                else:
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, dest)
                copied_count += 1
                yield _event_bytes({
                    "type": "progress",
                    "index": index,
                    "total": len(copy_items),
                    "path": str(dest),
                    "name": dest.name,
                    "copied": copied_count,
                })

            _copy_folder_config_for_clone(str(source_path), str(target_path))
            yield _event_bytes({
                "type": "config-copied",
                "target_folder": str(target_path),
            })
            yield _event_bytes({
                "type": "done",
                "target_folder": str(target_path),
                "copied": copied_count,
                "total": len(copy_items),
                "image_count": copied_image_count,
                "mode": "selected" if selected_mode else "folder",
            })
        except HTTPException:
            raise
        except Exception as exc:
            shutil.rmtree(target_path, ignore_errors=True)
            yield _event_bytes({"type": "error", "message": str(exc)})

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


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
        free_text = str(data.get("free_text", "") or "").replace(old_sentence, new_sentence)
        _write_caption_file(str(entry), enabled, free_text, sections)

    return {"ok": True, "sections": sections}


@app.post("/api/section/rename")
async def rename_section(update: RenameSectionUpdate):
    """Rename a configured section and migrate existing caption files."""
    folder = os.path.normpath(update.folder)
    old_name = str(update.old_name or "").strip()
    new_name = str(update.new_name or "").strip()
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")
    if old_name == new_name:
        cfg = _load_config()
        return {"ok": True, "sections": _get_folder_sections(cfg, folder)}

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    section_names_before = [str(section.get("name") or "") for section in sections_before]
    if old_name not in section_names_before:
        raise HTTPException(status_code=404, detail="Section not found")
    if new_name and new_name in section_names_before:
        raise HTTPException(status_code=400, detail="A section with that name already exists")

    all_sentences = _all_sentences_from_sections(sections_before)
    headers_before = _all_headers_from_sections(sections_before)
    sections = _get_folder_sections(cfg, folder)
    renamed = _rename_section_in_sections(sections, old_name, new_name)
    if not renamed:
        raise HTTPException(status_code=404, detail="Section not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)

    folder_path = Path(folder)
    for entry in folder_path.iterdir():
        if not entry.is_file() or entry.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_sentences, headers_before)
        free_text = str(data.get("free_text", "") or "").replace(old_name, new_name)
        _write_caption_file(str(entry), list(data.get("enabled_sentences", [])), free_text, sections)

    return {"ok": True, "sections": sections}


@app.post("/api/caption/delete-preset")
async def delete_caption_preset(update: DeleteCaptionPresetUpdate):
    """Delete a configured caption and remove it from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    sentence = str(update.sentence or "").strip()
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")
    if not sentence:
        raise HTTPException(status_code=400, detail="Caption text is required")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    if sentence not in _all_sentences_from_sections(sections_before):
        raise HTTPException(status_code=404, detail="Caption not found")

    sections = _get_folder_sections(cfg, folder)
    removed = _remove_sentence_from_sections(sections, sentence)
    if not removed:
        raise HTTPException(status_code=404, detail="Caption not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)
    _remove_sentences_from_caption_files(folder, sections_before, sections, [sentence])
    return {"ok": True, "sections": sections, "removed_sentences": [sentence]}


@app.post("/api/group/delete")
async def delete_group(update: DeleteGroupUpdate):
    """Delete a configured group and remove its captions from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    sections = _get_folder_sections(cfg, folder)
    removed_sentences = _remove_group_from_sections(sections, update.section_index, update.group_index)
    if removed_sentences is None:
        raise HTTPException(status_code=404, detail="Group not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)
    _remove_sentences_from_caption_files(folder, sections_before, sections, removed_sentences)
    return {"ok": True, "sections": sections, "removed_sentences": removed_sentences}


@app.post("/api/section/delete")
async def delete_section(update: DeleteSectionUpdate):
    """Delete a configured section and remove its captions from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    sections = _get_folder_sections(cfg, folder)
    updated_sections, removed_sentences = _remove_section_from_sections(sections, update.section_index)
    if updated_sections is None or removed_sentences is None:
        raise HTTPException(status_code=404, detail="Section not found")

    _set_folder_sections(cfg, folder, updated_sections)
    _save_config(cfg)
    _remove_sentences_from_caption_files(folder, sections_before, updated_sections, removed_sentences)
    return {"ok": True, "sections": updated_sections, "removed_sentences": removed_sentences}


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
    target_sentence: Optional[str] = None
    free_text_prompt_template: Optional[str] = None
    timeout_seconds: Optional[int] = None


def _resolve_caption_targets(
    sections: list[dict],
    target_section_index: int | None = None,
    target_group_index: int | None = None,
    target_sentence: str | None = None,
) -> tuple[list[dict] | None, dict | None]:
    """Resolve full or scoped caption targets for auto captioning."""
    sentence = str(target_sentence or "").strip()
    if sentence:
        for target in _iter_caption_targets_with_indices(sections):
            if target.get("type") == "sentence" and target.get("sentence") == sentence:
                return [target], {
                    "type": "sentence",
                    "sentence": sentence,
                    "section_index": target.get("section_index"),
                    "group_index": target.get("group_index"),
                    "section_name": target.get("section_name", ""),
                }
        return None, None

    if target_group_index is not None:
        target_group = _get_group_target(sections, target_section_index, target_group_index)
        if not target_group:
            return None, None
        return [target_group], {
            "type": "group",
            "section_index": target_group.get("section_index"),
            "group_index": target_group.get("group_index"),
            "section_name": target_group.get("section_name", ""),
            "group_name": target_group.get("group_name", ""),
        }

    if target_section_index is not None:
        if target_section_index < 0 or target_section_index >= len(sections):
            return None, None
        scoped_targets = [
            target
            for target in _iter_caption_targets_with_indices(sections)
            if target.get("section_index") == target_section_index
        ]
        if not scoped_targets:
            return None, None
        return scoped_targets, {
            "type": "section",
            "section_index": target_section_index,
            "section_name": sections[target_section_index].get("name", ""),
        }

    return list(_iter_caption_targets_with_indices(sections)), {"type": "full"}


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
    targets, target_scope = _resolve_caption_targets(
        sections,
        data.target_section_index,
        data.target_group_index,
        data.target_sentence,
    )
    if not free_text_only and targets is None:
        if data.target_sentence:
            raise HTTPException(status_code=400, detail="Invalid target sentence")
        if data.target_group_index is not None:
            raise HTTPException(status_code=400, detail="Invalid target group")
        if data.target_section_index is not None:
            raise HTTPException(status_code=400, detail="Invalid target section")
        raise HTTPException(status_code=400, detail="Invalid target")

    headers = _all_headers_from_sections(sections)
    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
    results: list[dict] = []
    is_scoped = bool(target_scope and target_scope.get("type") in {"group", "section", "sentence"})

    loop = asyncio.get_event_loop()
    if not free_text_only:
        try:
            if target_scope and target_scope.get("type") == "full":
                auto_caption_call = partial(
                    _auto_caption_sections,
                    host,
                    model,
                    image_path,
                    sections,
                    encode_image_func=_encode_image_for_ollama,
                    generate_func=_ollama_generate,
                    prompt_template=prompt_template,
                    group_prompt_template=group_prompt_template,
                    timeout=timeout_seconds,
                )
                enabled, results = await loop.run_in_executor(
                    executor,
                    auto_caption_call,
                )
                free_text = free_text if is_scoped else ""
                if not os.path.isfile(image_path):
                    raise FileNotFoundError(f"Image not found: {image_path}")
                _write_caption_file(image_path, enabled, free_text, sections)
            else:
                image_b64 = await loop.run_in_executor(executor, _encode_image_for_ollama, image_path)
                for target in targets or []:
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
                        enabled, free_text = _apply_sentence_result_to_live_caption(
                            image_path,
                            sentence,
                            is_match,
                            all_sentences,
                            headers,
                            sections,
                        )
                        results.append({
                            "type": "sentence",
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "section_name": target.get("section_name", ""),
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
                    enabled, free_text = _apply_group_result_to_live_caption(
                        image_path,
                        group_sentences,
                        selected_sentence,
                        all_sentences,
                        headers,
                        sections,
                    )
                    results.append({
                        "type": "group",
                        "section_index": target.get("section_index"),
                        "group_index": target.get("group_index"),
                        "section_name": target.get("section_name", ""),
                        "group_name": target.get("group_name", ""),
                        "sentences": group_sentences,
                        "selected_sentence": selected_sentence,
                        "selected_hidden": selected_hidden,
                        "selection_index": selection_index,
                        "answer": raw_answer,
                    })
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except FileNotFoundError:
            raise HTTPException(status_code=409, detail="Image was deleted during auto caption") from None
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption failed: {e}") from e

    free_text_model_output = ""
    added_free_text_lines: list[str] = []
    if enable_free_text:
        enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
        caption_text = _build_caption_text(enabled, free_text, sections)
        try:
            suggest_free_text_call = partial(
                _suggest_free_text,
                host,
                model,
                image_path,
                caption_text,
                encode_image_func=_encode_image_for_ollama,
                generate_func=_ollama_generate,
                prompt_template=free_text_prompt_template,
                timeout=timeout_seconds,
            )
            free_text_model_output = await loop.run_in_executor(
                executor,
                suggest_free_text_call,
            )
            enabled, free_text, added_free_text_lines = _apply_free_text_result_to_live_caption(
                image_path,
                free_text_model_output,
                all_sentences,
                headers,
                sections,
                preserve_existing=is_scoped,
            )
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except FileNotFoundError:
            raise HTTPException(status_code=409, detail="Image was deleted during auto caption") from None
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption free-text step failed: {e}") from e

    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
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

            targets, target_scope = _resolve_caption_targets(
                sections,
                data.target_section_index,
                data.target_group_index,
                data.target_sentence,
            )
            if not free_text_only and targets is None:
                total_errors += 1
                message = "Invalid target"
                if data.target_sentence:
                    message = "Invalid target sentence"
                elif data.target_group_index is not None:
                    message = "Invalid target group"
                elif data.target_section_index is not None:
                    message = "Invalid target section"
                yield _event_bytes({"type": "error", "path": image_path, "message": message})
                continue

            headers = _all_headers_from_sections(sections)
            enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
            results = []
            image_b64 = await loop.run_in_executor(executor, _encode_image_for_ollama, image_path)

            total_targets = 0 if free_text_only else len(targets or [])
            yield _event_bytes({
                "type": "image-start",
                "path": image_path,
                "total_sentences": len(all_sentences),
                "total_targets": total_targets,
                "free_text_only": free_text_only,
                "target_scope": target_scope,
                "enabled_sentences": enabled,
                "free_text": free_text,
            })

            try:
                if not free_text_only:
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
                            enabled, free_text = _apply_sentence_result_to_live_caption(
                                image_path,
                                sentence,
                                is_match,
                                all_sentences,
                                headers,
                                sections,
                            )
                            result = {"type": "sentence", "sentence": sentence, "enabled": is_match, "answer": raw_answer}
                            results.append(result)
                            yield _event_bytes({
                                "type": "caption-check",
                                "path": image_path,
                                "index": index,
                                "total": total_targets,
                                "sentence": sentence,
                                "enabled": is_match,
                                "enabled_sentences": enabled,
                                "free_text": free_text,
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
                        enabled, free_text = _apply_group_result_to_live_caption(
                            image_path,
                            group_sentences,
                            selected_sentence,
                            all_sentences,
                            headers,
                            sections,
                        )
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
                            "enabled_sentences": enabled,
                            "free_text": free_text,
                            "answer": raw_answer,
                        })

                free_text_model_output = ""
                added_free_text_lines: list[str] = []
                if enable_free_text:
                    if await request.is_disconnected():
                        return
                    is_scoped = bool(target_scope and target_scope.get("type") in {"group", "section", "sentence"})
                    enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
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
                    enabled, free_text, added_free_text_lines = _apply_free_text_result_to_live_caption(
                        image_path,
                        free_text_model_output,
                        all_sentences,
                        headers,
                        sections,
                        preserve_existing=is_scoped,
                    )
                    yield _event_bytes({
                        "type": "free-text",
                        "path": image_path,
                        "answer": free_text_model_output,
                        "enabled_sentences": enabled,
                        "free_text": free_text,
                        "added_lines": added_free_text_lines,
                    })

                enabled, free_text = _read_live_caption_state(image_path, all_sentences, headers)
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
            except FileNotFoundError:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": image_path, "message": "Image was deleted during auto caption"})
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

    return _get_bulk_caption_results(image_paths, predefined)


def _get_bulk_caption_results(image_paths: list[str], predefined: list[str]) -> dict:
    """Load caption data for multiple images."""

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


class BulkCaptionsRequest(BaseModel):
    paths: list[str]
    sentences: list[str] = []


@app.post("/api/captions/bulk")
async def post_captions_bulk(payload: BulkCaptionsRequest):
    """Get caption status for multiple images at once from a JSON body."""
    return _get_bulk_caption_results(payload.paths or [], payload.sentences or [])


# ===== SETTINGS API =====

class SettingsUpdate(BaseModel):
    last_folder: Optional[str] = None
    sections: Optional[list[dict]] = None
    folder: Optional[str] = None  # which folder these sections belong to
    thumb_size: Optional[int] = None
    crop_aspect_ratios: Optional[list[str]] = None
    https_certfile: Optional[str] = None
    https_keyfile: Optional[str] = None
    https_port: Optional[int] = None
    remote_http_mode: Optional[str] = None
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
        "thumb_size": int(cfg.get("thumb_size", 160) or 160),
        "crop_aspect_ratios": _get_crop_aspect_ratios(cfg, DEFAULT_CROP_ASPECT_RATIOS),
        "https_certfile": str(cfg.get("https_certfile") or ""),
        "https_keyfile": str(cfg.get("https_keyfile") or ""),
        "https_port": _get_https_port(cfg, DEFAULT_HTTPS_PORT),
        "remote_http_mode": _get_remote_http_mode(cfg, DEFAULT_REMOTE_HTTP_MODE),
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


@app.get("/api/ollama/models")
async def get_ollama_models(
    server: Optional[str] = Query(default=None),
    port: Optional[int] = Query(default=None),
):
    """List available model names from the configured Ollama instance."""
    cfg = _load_config()
    host = _compose_ollama_host(
        server if server is not None else cfg.get("ollama_server"),
        port if port is not None else cfg.get("ollama_port"),
        DEFAULT_OLLAMA_SERVER,
        DEFAULT_OLLAMA_PORT,
    )
    timeout = _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS)
    try:
        models = _list_ollama_models(host, timeout)
    except RuntimeError as exc:
        raise HTTPException(status_code=502, detail=f"Ollama error: {exc}") from exc
    return {"host": host, "models": models}


@app.post("/api/settings")
async def update_settings(data: SettingsUpdate):
    """Update settings. Saves last_folder and/or per-folder sections."""
    cfg = _load_config()
    if data.last_folder is not None:
        cfg["last_folder"] = data.last_folder
    if data.thumb_size is not None:
        cfg["thumb_size"] = max(60, min(400, int(data.thumb_size)))
    if data.crop_aspect_ratios is not None:
        cfg["crop_aspect_ratios"] = [str(r).strip() for r in data.crop_aspect_ratios if str(r).strip()] or list(DEFAULT_CROP_ASPECT_RATIOS)
    if data.https_certfile is not None:
        cfg["https_certfile"] = str(data.https_certfile or "").strip()
    if data.https_keyfile is not None:
        cfg["https_keyfile"] = str(data.https_keyfile or "").strip()
    if data.https_port is not None:
        cfg["https_port"] = max(1, min(65535, int(data.https_port)))
    if data.remote_http_mode is not None:
        cfg["remote_http_mode"] = _get_remote_http_mode({"remote_http_mode": data.remote_http_mode}, DEFAULT_REMOTE_HTTP_MODE)
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


def _open_file_manager_selection(path: str):
    """Open a file manager and select the requested file when possible."""
    normalized_path = os.path.normpath(path)
    system = platform.system()

    if system == "Windows":
        subprocess.Popen(["explorer", "/select,", normalized_path])
        return

    if system == "Darwin":
        subprocess.Popen(["open", "-R", normalized_path])
        return

    directory = os.path.dirname(normalized_path)
    file_uri = Path(normalized_path).resolve().as_uri()
    commands: list[list[str]] = []

    if shutil.which("dbus-send"):
        commands.append([
            "dbus-send",
            "--session",
            "--dest=org.freedesktop.FileManager1",
            "--type=method_call",
            "/org/freedesktop/FileManager1",
            "org.freedesktop.FileManager1.ShowItems",
            f"array:string:{file_uri}",
            "string:",
        ])
    if shutil.which("nautilus"):
        commands.append(["nautilus", "--select", normalized_path])
    if shutil.which("dolphin"):
        commands.append(["dolphin", "--select", normalized_path])
    if shutil.which("thunar"):
        commands.append(["thunar", "--select", normalized_path])
    if shutil.which("nemo"):
        commands.append(["nemo", normalized_path])
    commands.append(["xdg-open", directory])

    last_error: Exception | None = None
    for command in commands:
        try:
            subprocess.Popen(command)
            return
        except Exception as exc:
            last_error = exc

    if last_error is not None:
        raise last_error
    raise RuntimeError("No file manager command available")


def _open_file_direct(path: str):
    """Open a file directly with the OS default handler."""
    normalized_path = os.path.normpath(path)
    system = platform.system()

    if system == "Windows":
        if not hasattr(os, "startfile"):
            raise RuntimeError("os.startfile is not available")
        os.startfile(normalized_path)
        return

    if system == "Darwin":
        subprocess.Popen(["open", normalized_path])
        return

    subprocess.Popen(["xdg-open", normalized_path])


def _run_uvicorn_instance(host: str, port: int, ssl_kwargs: dict | None = None):
    """Run a single Uvicorn instance."""
    import uvicorn

    uvicorn.run(app, host=host, port=port, **(ssl_kwargs or {}))


@app.get("/api/open-in-explorer")
async def open_in_explorer(path: str = Query(...)):
    """Open the OS file explorer with the given file selected."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    try:
        _open_file_manager_selection(path)
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/open-file")
async def open_file(path: str = Query(...)):
    """Open a file directly with the default OS application."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    try:
        _open_file_direct(path)
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
    cfg = _load_config()
    http_port = int(sys.argv[1]) if len(sys.argv) > 1 else 8899
    try:
        https_kwargs = _get_https_uvicorn_kwargs(cfg)
    except RuntimeError as exc:
        raise SystemExit(str(exc)) from exc
    if not https_kwargs:
        _run_uvicorn_instance("0.0.0.0", http_port)
        raise SystemExit(0)

    https_port = _get_https_port(cfg, DEFAULT_HTTPS_PORT)
    if https_port == http_port:
        raise SystemExit("HTTPS port must differ from the HTTP port when local HTTP is enabled")

    http_thread = threading.Thread(
        target=_run_uvicorn_instance,
        args=("0.0.0.0", http_port),
        kwargs={"ssl_kwargs": None},
        daemon=True,
    )
    http_thread.start()
    _run_uvicorn_instance("0.0.0.0", https_port, https_kwargs)
