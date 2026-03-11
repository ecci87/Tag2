#!/usr/bin/env python3
"""
Image Captioning Tool - Backend Server
FastAPI-based server for browsing images and managing captions.
"""

import os
import sys
import json
import re
import asyncio
import atexit
import base64
import hashlib
import shutil
import subprocess
import platform
import tempfile
import warnings
import urllib.error
import urllib.request
from urllib.parse import urlparse
from pathlib import Path
from io import BytesIO
from typing import Optional
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Query, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from PIL import Image, ImageOps

try:
    import piexif
except ImportError:  # pragma: no cover - optional during import
    piexif = None

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

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"}

# In-memory thumbnail cache: (path, mtime) -> bytes
thumbnail_cache: dict[tuple[str, float], bytes] = {}
THUMBNAIL_SIZES = [64, 128, 256, 400]
PREVIEW_MAX_SIZE = 2048  # Max edge for preview images

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


RUNTIME_CROP_BACKUP_DIR = tempfile.mkdtemp(prefix="tag2-crop-")


@atexit.register
def _cleanup_runtime_crop_backups():
    """Delete temporary crop backups when the server process exits."""
    shutil.rmtree(RUNTIME_CROP_BACKUP_DIR, ignore_errors=True)


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
                    str(data.get("ollama_host") or default["ollama_host"])
                )
                data.setdefault("ollama_server", parsed_server)
                data.setdefault("ollama_port", parsed_port)
            data["ollama_host"] = _compose_ollama_host(
                data.get("ollama_server"),
                data.get("ollama_port"),
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


def _get_folder_sections(cfg: dict, folder: str) -> list[dict]:
    """Get sections for a folder, falling back to default_sections.
    Returns list of dicts: [{"name": "SectionName", "sentences": [...], "groups": [...]}, ...]
    The unnamed section (name='') holds sentences without a section header.
    """
    def _normalize_sentences(sentences) -> list[str]:
        cleaned: list[str] = []
        seen: set[str] = set()
        for raw in sentences or []:
            text = str(raw).strip()
            if not text or text in seen:
                continue
            seen.add(text)
            cleaned.append(text)
        return cleaned

    def _normalize_group(group: dict | None) -> dict:
        group = dict(group or {})
        return {
            "name": str(group.get("name") or "").strip(),
            "sentences": _normalize_sentences(group.get("sentences", [])),
        }

    def _normalize_section(section: dict | None) -> dict:
        section = dict(section or {})
        groups = []
        for group in section.get("groups", []) or []:
            normalized_group = _normalize_group(group)
            if normalized_group["name"] or normalized_group["sentences"]:
                groups.append(normalized_group)
        return {
            "name": str(section.get("name") or "").strip(),
            "sentences": _normalize_sentences(section.get("sentences", [])),
            "groups": groups,
        }

    folder_key = os.path.normpath(folder)
    folder_cfg = cfg.get("folders", {}).get(folder_key, None)
    if folder_cfg is not None:
        if "sections" in folder_cfg:
            return [_normalize_section(s) for s in folder_cfg["sections"]]
        # Migration: old flat sentences list -> single unnamed section
        if "sentences" in folder_cfg:
            return [_normalize_section({"name": "", "sentences": list(folder_cfg["sentences"]), "groups": []})]
    # Try default_sections
    default_sections = cfg.get("default_sections", [])
    if default_sections:
        return [_normalize_section(s) for s in default_sections]
    # Migration from old default_sentences
    default_sentences = cfg.get("default_sentences", [])
    if default_sentences:
        return [_normalize_section({"name": "", "sentences": list(default_sentences), "groups": []})]
    return [_normalize_section({"name": "", "sentences": [], "groups": []})]


def _set_folder_sections(cfg: dict, folder: str, sections: list[dict]):
    """Set sections for a specific folder."""
    folder_key = os.path.normpath(folder)
    if "folders" not in cfg:
        cfg["folders"] = {}
    if folder_key not in cfg["folders"]:
        cfg["folders"][folder_key] = {}
    cfg["folders"][folder_key]["sections"] = _get_folder_sections({
        "folders": {folder_key: {"sections": sections}}
    }, folder_key)
    # Remove legacy key if present
    cfg["folders"][folder_key].pop("sentences", None)


def _all_sentences_from_sections(sections: list[dict]) -> list[str]:
    """Flatten sections into a single list of all predefined sentences."""
    result = []
    for sec in sections:
        result.extend(sec.get("sentences", []))
        for group in sec.get("groups", []) or []:
            result.extend(group.get("sentences", []))
    return result


def _all_headers_from_sections(sections: list[dict]) -> list[str]:
    """Collect all structured header lines from sections and groups."""
    headers: list[str] = []
    for section in sections:
        if section.get("name"):
            headers.append(section["name"])
        for group in section.get("groups", []) or []:
            if group.get("name"):
                headers.append(group["name"])
    return headers


def _iter_caption_targets(sections: list[dict]):
    """Yield caption targets in display order."""
    for section in sections:
        for sentence in section.get("sentences", []):
            yield {
                "type": "sentence",
                "section_name": section.get("name", ""),
                "sentence": sentence,
            }
        for group in section.get("groups", []) or []:
            group_sentences = list(group.get("sentences", []))
            if not group_sentences:
                continue
            yield {
                "type": "group",
                "section_name": section.get("name", ""),
                "group_name": group.get("name", ""),
                "sentences": group_sentences,
            }


def _find_group_for_sentence(sections: list[dict], sentence: str) -> dict | None:
    """Return the group containing a sentence, if any."""
    for section in sections:
        for group in section.get("groups", []) or []:
            if sentence in group.get("sentences", []):
                return group
    return None


def _apply_sentence_selection(enabled_sentences: list[str], sentence: str,
                              sections: list[dict], should_enable: bool) -> list[str]:
    """Apply a sentence toggle while enforcing group exclusivity."""
    updated = [item for item in enabled_sentences if item != sentence]
    if should_enable:
        group = _find_group_for_sentence(sections, sentence)
        if group:
            group_sentences = set(group.get("sentences", []))
            updated = [item for item in updated if item not in group_sentences]
        updated.append(sentence)
    return updated


def _normalize_enabled_sentences(enabled_sentences: list[str], sections: list[dict]) -> list[str]:
    """Normalize enabled sentences against known captions and group exclusivity."""
    known = set(_all_sentences_from_sections(sections))
    normalized: list[str] = []
    for sentence in enabled_sentences:
        if sentence not in known:
            continue
        normalized = _apply_sentence_selection(normalized, sentence, sections, True)
    return normalized


def _rename_sentence_in_sections(sections: list[dict], old_sentence: str, new_sentence: str) -> bool:
    """Rename a configured sentence inside sections/groups."""
    renamed = False
    for section in sections:
        section_sentences = list(section.get("sentences", []))
        if old_sentence in section_sentences:
            renamed = True
        section["sentences"] = [new_sentence if sentence == old_sentence else sentence for sentence in section_sentences]
        for group in section.get("groups", []) or []:
            group_sentences = list(group.get("sentences", []))
            if old_sentence in group_sentences:
                renamed = True
            group["sentences"] = [new_sentence if sentence == old_sentence else sentence for sentence in group_sentences]
    return renamed


def _get_crop_aspect_ratios(cfg: dict) -> list[str]:
    """Get the configured list of allowed crop aspect ratios."""
    ratios = cfg.get("crop_aspect_ratios", DEFAULT_CROP_ASPECT_RATIOS)
    if not isinstance(ratios, list):
        return list(DEFAULT_CROP_ASPECT_RATIOS)
    cleaned = [str(r).strip() for r in ratios if str(r).strip()]
    return cleaned or list(DEFAULT_CROP_ASPECT_RATIOS)


def _normalize_image_key(image_path: str) -> str:
    """Normalize an image path for use as a config key."""
    return os.path.normcase(os.path.abspath(os.path.normpath(image_path)))


def _get_crop_backup_dir() -> str:
    """Return the folder used for reversible crop backups."""
    return RUNTIME_CROP_BACKUP_DIR


def _get_crop_backup_path(image_path: str) -> str:
    """Build a unique backup file path for a cropped image."""
    normalized = _normalize_image_key(image_path)
    digest = hashlib.sha256(normalized.encode("utf-8", errors="surrogatepass")).hexdigest()[:16]
    image = Path(image_path)
    stem = image.stem or "image"
    suffix = image.suffix or ".img"
    return os.path.join(_get_crop_backup_dir(), f"{stem}.{digest}{suffix}")


def _clear_thumbnail_cache_for_path(image_path: str):
    """Drop cached thumbnails/previews for a specific image."""
    normalized = _normalize_image_key(image_path)
    for key in [k for k in thumbnail_cache if _normalize_image_key(k[0]) == normalized]:
        thumbnail_cache.pop(key, None)


def _normalize_crop_rect(crop: dict, image_size: tuple[int, int]) -> dict:
    """Clamp and normalize a crop rectangle to fit the source image."""
    img_w, img_h = image_size
    x = int(round(float(crop.get("x", 0))))
    y = int(round(float(crop.get("y", 0))))
    w = int(round(float(crop.get("w", 0))))
    h = int(round(float(crop.get("h", 0))))
    ratio = str(crop.get("ratio", "")).strip()

    x = max(0, min(x, max(0, img_w - 1)))
    y = max(0, min(y, max(0, img_h - 1)))
    w = max(1, min(w, img_w - x))
    h = max(1, min(h, img_h - y))

    result = {"x": x, "y": y, "w": w, "h": h}
    if ratio:
        result["ratio"] = ratio
    return result


def _normalize_exif_bytes(exif_bytes: bytes | None) -> bytes | None:
    """Normalize EXIF orientation so served images are not rotated twice."""
    if not exif_bytes:
        return None
    if piexif is None:
        return None
    try:
        exif_dict = piexif.load(exif_bytes)
        exif_dict.setdefault("0th", {})
        exif_dict["0th"][piexif.ImageIFD.Orientation] = 1
        return piexif.dump(exif_dict)
    except Exception:
        return None


def _load_oriented_image(filepath: str) -> tuple[Image.Image, str, bytes | None, bytes | None]:
    """Load an image with EXIF orientation applied."""
    with Image.open(filepath) as source:
        original_format = (source.format or Path(filepath).suffix.lstrip(".") or "PNG").upper()
        exif_bytes = _normalize_exif_bytes(source.info.get("exif"))
        icc_profile = source.info.get("icc_profile")
        image = ImageOps.exif_transpose(source)
        image.load()
    return image, original_format, exif_bytes, icc_profile


def _get_display_image_size(filepath: str) -> tuple[int, int]:
    """Return the image size as displayed after EXIF transforms are applied."""
    with Image.open(filepath) as source:
        return ImageOps.exif_transpose(source).size


def _save_image_file(filepath: str, image: Image.Image, original_format: str,
                     exif_bytes: bytes | None = None, icc_profile: bytes | None = None):
    """Save an image while preserving useful metadata where possible."""
    fmt = "JPEG" if original_format in {"JPG", "JPEG"} else ("PNG" if original_format == "PNG" else original_format)
    save_kwargs: dict = {}
    if exif_bytes:
        save_kwargs["exif"] = exif_bytes
    if icc_profile:
        save_kwargs["icc_profile"] = icc_profile

    output = image
    if fmt == "JPEG":
        if output.mode not in ("RGB", "L"):
            output = output.convert("RGB")
        output.save(filepath, format=fmt, quality=95, optimize=True, **save_kwargs)
        return

    output.save(filepath, format=fmt, **save_kwargs)


def _get_image_crop(image_path: str) -> dict:
    """Return current/original display dimensions and crop state for an image."""
    backup_path = _get_crop_backup_path(image_path)
    try:
        current_w, current_h = _get_display_image_size(image_path)
    except Exception:
        return {"applied": os.path.isfile(backup_path)}

    original_w = current_w
    original_h = current_h
    applied = False
    if os.path.isfile(backup_path):
        applied = True
        try:
            original_w, original_h = _get_display_image_size(backup_path)
        except Exception:
            original_w, original_h = current_w, current_h

    return {
        "applied": applied,
        "current_width": current_w,
        "current_height": current_h,
        "original_width": original_w,
        "original_height": original_h,
    }


def _apply_real_crop(image_path: str, crop: dict) -> dict:
    """Crop the actual image file while keeping a reversible backup."""
    os.makedirs(_get_crop_backup_dir(), exist_ok=True)
    backup_path = _get_crop_backup_path(image_path)
    if not os.path.isfile(backup_path):
        shutil.copy2(image_path, backup_path)

    image, original_format, exif_bytes, icc_profile = _load_oriented_image(image_path)
    try:
        normalized = _normalize_crop_rect(crop, image.size)
        cropped = image.crop((
            normalized["x"],
            normalized["y"],
            normalized["x"] + normalized["w"],
            normalized["y"] + normalized["h"],
        ))
        _save_image_file(image_path, cropped, original_format, exif_bytes, icc_profile)
    finally:
        image.close()

    _clear_thumbnail_cache_for_path(image_path)
    return normalized


def _remove_real_crop(image_path: str) -> bool:
    """Restore the original image file if a crop backup exists."""
    backup_path = _get_crop_backup_path(image_path)
    if not os.path.isfile(backup_path):
        return False
    shutil.copy2(backup_path, image_path)
    os.remove(backup_path)
    _clear_thumbnail_cache_for_path(image_path)
    return True

def _rotate_image_file(filepath: str, clockwise: bool):
    """Rotate an image file by 90 degrees while preserving metadata."""
    image, original_format, exif_bytes, icc_profile = _load_oriented_image(filepath)
    try:
        rotated = image.transpose(Image.Transpose.ROTATE_270 if clockwise else Image.Transpose.ROTATE_90)
        _save_image_file(filepath, rotated, original_format, exif_bytes, icc_profile)
    finally:
        image.close()

def _rotate_image(image_path: str, direction: str) -> dict:
    """Rotate the actual image file and any active crop backup by 90 degrees."""
    normalized_direction = str(direction or "").strip().lower()
    if normalized_direction not in {"left", "right"}:
        raise ValueError("direction must be 'left' or 'right'")

    clockwise = normalized_direction == "right"
    _rotate_image_file(image_path, clockwise)

    backup_path = _get_crop_backup_path(image_path)
    if os.path.isfile(backup_path):
        _rotate_image_file(backup_path, clockwise)

    _clear_thumbnail_cache_for_path(image_path)
    return _get_image_crop(image_path)


def _get_ollama_host(cfg: dict) -> str:
    """Get the configured Ollama host."""
    server = cfg.get("ollama_server")
    port = cfg.get("ollama_port")
    if server is not None or port is not None:
        return _compose_ollama_host(server, port)
    return str(cfg.get("ollama_host") or f"http://{DEFAULT_OLLAMA_SERVER}:{DEFAULT_OLLAMA_PORT}").rstrip("/")


def _split_ollama_host(host: str) -> tuple[str, int]:
    """Split a full Ollama host URL into server and port."""
    raw = (host or "").strip()
    if not raw:
        return DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT
    if "://" not in raw:
        raw = f"http://{raw}"
    parsed = urlparse(raw)
    server = parsed.hostname or DEFAULT_OLLAMA_SERVER
    port = parsed.port or DEFAULT_OLLAMA_PORT
    return server, port


def _compose_ollama_host(server: str | None, port: int | str | None) -> str:
    """Compose a normalized Ollama host URL from server and port."""
    host_server = str(server or DEFAULT_OLLAMA_SERVER).strip() or DEFAULT_OLLAMA_SERVER
    try:
        host_port = int(port if port is not None else DEFAULT_OLLAMA_PORT)
    except (TypeError, ValueError):
        host_port = DEFAULT_OLLAMA_PORT
    return f"http://{host_server}:{host_port}"


def _get_ollama_server(cfg: dict) -> str:
    """Get configured Ollama server hostname."""
    server = str(cfg.get("ollama_server") or "").strip()
    if server:
        return server
    return _split_ollama_host(_get_ollama_host(cfg))[0]


def _get_ollama_port(cfg: dict) -> int:
    """Get configured Ollama server port."""
    port = cfg.get("ollama_port")
    try:
        if port is not None:
            return int(port)
    except (TypeError, ValueError):
        pass
    return _split_ollama_host(_get_ollama_host(cfg))[1]


def _get_ollama_model(cfg: dict) -> str:
    """Get the configured Ollama model name."""
    return str(cfg.get("ollama_model") or DEFAULT_OLLAMA_MODEL).strip()


def _get_ollama_timeout_seconds(cfg: dict) -> int:
    """Get the configured Ollama request timeout."""
    try:
        timeout = int(cfg.get("ollama_timeout_seconds", DEFAULT_OLLAMA_TIMEOUT_SECONDS))
    except (TypeError, ValueError):
        timeout = DEFAULT_OLLAMA_TIMEOUT_SECONDS
    return max(1, timeout)


def _get_ollama_prompt_template(cfg: dict) -> str:
    """Get the configured Ollama prompt template."""
    return str(cfg.get("ollama_prompt_template") or DEFAULT_OLLAMA_PROMPT_TEMPLATE)


def _get_ollama_group_prompt_template(cfg: dict) -> str:
    """Get the configured grouped-caption prompt template."""
    return str(cfg.get("ollama_group_prompt_template") or DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE)


def _get_ollama_enable_free_text(cfg: dict) -> bool:
    """Get whether the free-text enhancement step is enabled."""
    return bool(cfg.get("ollama_enable_free_text", True))


def _get_ollama_free_text_prompt_template(cfg: dict) -> str:
    """Get the configured free-text prompt template."""
    return str(cfg.get("ollama_free_text_prompt_template") or DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE)


def _generate_thumbnail(filepath: str, size: int, crop: dict | None = None) -> bytes:
    """Generate a JPEG thumbnail of the given size (longest edge)."""
    try:
        img, _, _, _ = _load_oriented_image(filepath)
        try:
            if crop:
                img = img.crop((crop["x"], crop["y"], crop["x"] + crop["w"], crop["y"] + crop["h"]))
            img.thumbnail((size, size), Image.LANCZOS)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            buf = BytesIO()
            quality = 85 if size >= 1024 else 80
            img.save(buf, format="JPEG", quality=quality, optimize=True)
            return buf.getvalue()
        finally:
            img.close()
    except Exception:
        return b""


def _get_thumbnail(filepath: str, size: int, crop: dict | None = None) -> bytes:
    """Get thumbnail from cache or generate it."""
    mtime = os.path.getmtime(filepath)
    crop_key = None if not crop else (crop.get("x"), crop.get("y"), crop.get("w"), crop.get("h"), crop.get("ratio"))
    key = (filepath, mtime, size, crop_key)
    if key not in thumbnail_cache:
        thumbnail_cache[key] = _generate_thumbnail(filepath, size, crop)
    return thumbnail_cache[key]


def _render_image_bytes(filepath: str, crop: dict | None = None,
                        max_size: int | None = None, force_jpeg: bool = False) -> tuple[bytes, str]:
    """Render an image variant for serving through the API."""
    img, original_format, exif_bytes, icc_profile = _load_oriented_image(filepath)
    try:
        save_kwargs: dict = {}

        if crop:
            img = img.crop((crop["x"], crop["y"], crop["x"] + crop["w"], crop["y"] + crop["h"]))
        if max_size:
            img.thumbnail((max_size, max_size), Image.LANCZOS)

        if force_jpeg or original_format in {"JPG", "JPEG"}:
            media_type = "image/jpeg"
            fmt = "JPEG"
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            if exif_bytes:
                save_kwargs["exif"] = exif_bytes
        else:
            fmt = "PNG" if original_format == "PNG" else original_format
            media_type = Image.MIME.get(fmt, "application/octet-stream")
            if exif_bytes:
                save_kwargs["exif"] = exif_bytes

        if icc_profile:
            save_kwargs["icc_profile"] = icc_profile

        buf = BytesIO()
        if fmt == "JPEG":
            img.save(buf, format=fmt, quality=90, optimize=True, **save_kwargs)
        else:
            img.save(buf, format=fmt, **save_kwargs)
        return buf.getvalue(), media_type
    finally:
        img.close()


def _encode_image_for_ollama(filepath: str) -> str:
    """Encode an image as base64 JPEG for Ollama vision models."""
    data = _get_thumbnail(filepath, PREVIEW_MAX_SIZE)
    if not data:
        data = Path(filepath).read_bytes()
    return base64.b64encode(data).decode("ascii")


def _ollama_prompt_for_sentence(sentence: str, template: str | None = None) -> str:
    """Build a strict yes/no prompt for one candidate caption."""
    prompt_template = template or DEFAULT_OLLAMA_PROMPT_TEMPLATE
    return prompt_template.replace("{caption}", sentence).replace("{sentence}", sentence)


def _ollama_prompt_for_group(group_name: str, sentences: list[str], template: str | None = None) -> str:
    """Build a numbered-choice prompt for a mutually-exclusive caption group."""
    prompt_template = template or DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE
    options = "\n".join(f"{index}. {sentence}" for index, sentence in enumerate(sentences, start=1))
    resolved_group_name = (group_name or "Caption group").strip() or "Caption group"
    return (
        prompt_template
        .replace("{group_name}", resolved_group_name)
        .replace("{group}", resolved_group_name)
        .replace("{options}", options)
        .replace("{captions}", options)
        .replace("{choices}", options)
        .replace("{count}", str(len(sentences)))
    )


def _ollama_prompt_for_free_text(caption_text: str, template: str | None = None) -> str:
    """Build a prompt asking for additional important image details."""
    prompt_template = template or DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE
    return (
        prompt_template
        .replace("{caption_text}", caption_text)
        .replace("{current_caption}", caption_text)
        .replace("{caption}", caption_text)
    )


def _ollama_generate(host: str, payload: dict, timeout: int = 120) -> dict:
    """Call the local Ollama generate API."""
    url = f"{host.rstrip('/')}/api/generate"
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(detail or str(e)) from e
    except TimeoutError as e:
        raise RuntimeError(f"request timed out after {timeout} seconds") from e
    except urllib.error.URLError as e:
        raise RuntimeError(str(e.reason or e)) from e


def _parse_ollama_yes_no(response_text: str) -> bool:
    """Parse a yes/no response from Ollama conservatively."""
    text = (response_text or "").strip().upper()
    if text.startswith("YES"):
        return True
    if text.startswith("NO"):
        return False
    tokens = [t.strip(" .,!?:;()[]{}\"'") for t in text.split()]
    if "YES" in tokens:
        return True
    if "NO" in tokens:
        return False
    return False


def _parse_ollama_selection(response_text: str, sentences: list[str]) -> int | None:
    """Parse a 1-based choice from an Ollama response."""
    text = (response_text or "").strip()
    if not text:
        return None

    for match in re.findall(r"\d+", text):
        value = int(match)
        if 1 <= value <= len(sentences):
            return value

    normalized = _normalize_caption_line(text)
    for index, sentence in enumerate(sentences, start=1):
        sentence_normalized = _normalize_caption_line(sentence)
        if normalized == sentence_normalized:
            return index
    for index, sentence in enumerate(sentences, start=1):
        sentence_normalized = _normalize_caption_line(sentence)
        if sentence_normalized and sentence_normalized in normalized:
            return index
    return None


def _normalize_caption_line(text: str) -> str:
    """Normalize a caption/free-text line for duplicate detection."""
    stripped = (text or "").strip()
    stripped = re.sub(r"^[\-•*\d\.)\s]+", "", stripped)
    stripped = re.sub(r"\s+", " ", stripped)
    return stripped.casefold()


def _extract_free_text_lines(response_text: str) -> list[str]:
    """Extract meaningful suggestion lines from an Ollama free-text response."""
    text = (response_text or "").strip()
    if not text or text.upper() == "NONE":
        return []

    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.upper() == "NONE":
            continue
        line = re.sub(r"^[\-•*\d\.)\s]+", "", line).strip()
        if line:
            lines.append(line)
    return lines


def _merge_free_text(existing_free_text: str, suggested_text: str,
                     enabled_sentences: list[str]) -> tuple[str, list[str]]:
    """Merge new free-text suggestions while avoiding duplicates."""
    existing_text = existing_free_text or ""
    existing_lines = [line.rstrip() for line in existing_text.splitlines() if line.strip()]
    known = {_normalize_caption_line(line) for line in existing_lines}
    known.update(_normalize_caption_line(sentence) for sentence in enabled_sentences)

    added_lines = []
    for line in _extract_free_text_lines(suggested_text):
        normalized = _normalize_caption_line(line)
        if not normalized or normalized in known:
            continue
        known.add(normalized)
        added_lines.append(line)

    merged_lines = list(existing_lines)
    merged_lines.extend(added_lines)
    return "\n".join(merged_lines), added_lines


def _build_caption_text(enabled_sentences: list[str], free_text: str,
                        sections: list[dict] | None = None) -> str:
    """Build the final caption text in the same format as the caption file."""
    if sections:
        enabled_sentences = _normalize_enabled_sentences(enabled_sentences, sections)
        enabled_set = set(enabled_sentences)
        blocks = []
        for section in sections:
            sec_name = section.get("name", "")
            sec_sentences = [s for s in section.get("sentences", []) if s in enabled_set]
            group_blocks = []
            for group in section.get("groups", []) or []:
                group_sentences = [s for s in group.get("sentences", []) if s in enabled_set]
                if not group_sentences:
                    continue
                group_lines = []
                if group.get("name"):
                    group_lines.append(group["name"])
                for sentence in group_sentences:
                    group_lines.append(f"- {sentence}")
                group_blocks.append("\n".join(group_lines))

            if not sec_sentences and not group_blocks:
                continue

            lines = []
            if sec_name:
                lines.append(sec_name)
            for sentence in sec_sentences:
                lines.append(f"- {sentence}")
            if group_blocks:
                if lines:
                    lines.append("")
                lines.append("\n\n".join(group_blocks))
            blocks.append("\n".join(lines))

        parts = []
        if blocks:
            parts.append("\n\n".join(blocks))
        if free_text and free_text.strip():
            parts.append(free_text.strip())
        return "\n\n".join(parts)

    parts = []
    if enabled_sentences:
        parts.append("\n".join(f"- {s}" for s in enabled_sentences))
    if free_text and free_text.strip():
        parts.append(free_text.strip())
    return "\n\n".join(parts)


def _auto_caption_sentences(host: str, model: str, image_path: str,
                            sentences: list[str], prompt_template: str | None = None,
                            timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS) -> tuple[list[str], list[dict]]:
    """Ask Ollama about each caption candidate and return enabled sentences."""
    image_b64 = _encode_image_for_ollama(image_path)
    enabled = []
    results = []

    for sentence in sentences:
        payload = {
            "model": model,
            "prompt": _ollama_prompt_for_sentence(sentence, prompt_template),
            "images": [image_b64],
            "stream": False,
            "options": {
                "temperature": 0,
            },
        }
        response = _ollama_generate(host, payload, timeout=timeout)
        raw_answer = str(response.get("response") or "").strip()
        is_match = _parse_ollama_yes_no(raw_answer)
        results.append({
            "sentence": sentence,
            "enabled": is_match,
            "answer": raw_answer,
        })
        if is_match:
            enabled.append(sentence)

    return enabled, results


def _auto_caption_sections(host: str, model: str, image_path: str,
                           sections: list[dict], prompt_template: str | None = None,
                           group_prompt_template: str | None = None,
                           timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS) -> tuple[list[str], list[dict]]:
    """Ask Ollama about configured captions, including exclusive groups."""
    image_b64 = _encode_image_for_ollama(image_path)
    enabled: list[str] = []
    results: list[dict] = []

    for target in _iter_caption_targets(sections):
        if target["type"] == "sentence":
            sentence = target["sentence"]
            payload = {
                "model": model,
                "prompt": _ollama_prompt_for_sentence(sentence, prompt_template),
                "images": [image_b64],
                "stream": False,
                "options": {"temperature": 0},
            }
            response = _ollama_generate(host, payload, timeout=timeout)
            raw_answer = str(response.get("response") or "").strip()
            is_match = _parse_ollama_yes_no(raw_answer)
            enabled = _apply_sentence_selection(enabled, sentence, sections, is_match)
            results.append({
                "type": "sentence",
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
        response = _ollama_generate(host, payload, timeout=timeout)
        raw_answer = str(response.get("response") or "").strip()
        selection_index = _parse_ollama_selection(raw_answer, group_sentences)
        selected_sentence = group_sentences[selection_index - 1] if selection_index else None
        enabled = [sentence for sentence in enabled if sentence not in group_sentences]
        if selected_sentence:
            enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
        results.append({
            "type": "group",
            "group_name": target.get("group_name", ""),
            "sentences": group_sentences,
            "selected_sentence": selected_sentence,
            "selection_index": selection_index,
            "answer": raw_answer,
        })

    return enabled, results


def _suggest_free_text(host: str, model: str, image_path: str, caption_text: str,
                       prompt_template: str | None = None,
                       timeout: int = DEFAULT_OLLAMA_TIMEOUT_SECONDS) -> str:
    """Ask Ollama for additional free-text image details."""
    image_b64 = _encode_image_for_ollama(image_path)
    payload = {
        "model": model,
        "prompt": _ollama_prompt_for_free_text(caption_text, prompt_template),
        "images": [image_b64],
        "stream": False,
        "options": {
            "temperature": 0,
        },
    }
    response = _ollama_generate(host, payload, timeout=timeout)
    return str(response.get("response") or "").strip()


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


def _get_caption_path(image_path: str) -> Path:
    """Get the corresponding .txt path for an image."""
    p = Path(image_path)
    return p.with_suffix(".txt")


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
    content = _build_caption_text(enabled_sentences, free_text, sections)

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
    model = (data.model or _get_ollama_model(cfg)).strip()
    host = _get_ollama_host(cfg)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg)

    if not model:
        raise HTTPException(status_code=400, detail="No Ollama model configured")

    folder = str(Path(image_path).parent)
    sections = _get_folder_sections(cfg, folder)
    all_sentences = _all_sentences_from_sections(sections)
    if not all_sentences and not free_text_only:
        raise HTTPException(status_code=400, detail="No captions configured for this folder")

    headers = _all_headers_from_sections(sections)
    existing = _read_caption_file(image_path, all_sentences, headers)
    enabled = list(existing.get("enabled_sentences", []))
    results: list[dict] = []

    loop = asyncio.get_event_loop()
    if not free_text_only:
        try:
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

    free_text = existing.get("free_text", "")
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
            free_text, added_free_text_lines = _merge_free_text(free_text, free_text_model_output, enabled)
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
    model = (data.model or _get_ollama_model(cfg)).strip()
    host = _get_ollama_host(cfg)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg)

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

            headers = _all_headers_from_sections(sections)
            existing = _read_caption_file(image_path, all_sentences, headers)
            free_text = existing.get("free_text", "")
            enabled = list(existing.get("enabled_sentences", []))
            results = []
            image_b64 = await loop.run_in_executor(executor, _encode_image_for_ollama, image_path)

            total_targets = 0 if free_text_only else sum(1 for _ in _iter_caption_targets(sections))
            yield _event_bytes({
                "type": "image-start",
                "path": image_path,
                "total_sentences": len(all_sentences),
                "total_targets": total_targets,
                "free_text_only": free_text_only,
                "free_text": free_text,
            })

            try:
                if not free_text_only:
                    for index, target in enumerate(_iter_caption_targets(sections), start=1):
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
                        enabled = [sentence for sentence in enabled if sentence not in group_sentences]
                        if selected_sentence:
                            enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
                        result = {
                            "type": "group",
                            "group_name": target.get("group_name", ""),
                            "sentences": group_sentences,
                            "selected_sentence": selected_sentence,
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
                            "group_name": target.get("group_name", ""),
                            "sentences": group_sentences,
                            "selected_sentence": selected_sentence,
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
                    free_text, added_free_text_lines = _merge_free_text(free_text, free_text_model_output, enabled)
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
        "crop_aspect_ratios": _get_crop_aspect_ratios(cfg),
        "ollama_host": _get_ollama_host(cfg),
        "ollama_server": _get_ollama_server(cfg),
        "ollama_port": _get_ollama_port(cfg),
        "ollama_timeout_seconds": _get_ollama_timeout_seconds(cfg),
        "ollama_model": _get_ollama_model(cfg),
        "ollama_prompt_template": _get_ollama_prompt_template(cfg),
        "ollama_group_prompt_template": _get_ollama_group_prompt_template(cfg),
        "ollama_enable_free_text": _get_ollama_enable_free_text(cfg),
        "ollama_free_text_prompt_template": _get_ollama_free_text_prompt_template(cfg),
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
        server_name, server_port = _split_ollama_host(data.ollama_host)
        cfg["ollama_server"] = server_name
        cfg["ollama_port"] = server_port
        cfg["ollama_host"] = _compose_ollama_host(server_name, server_port)
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
    cfg["ollama_host"] = _compose_ollama_host(cfg.get("ollama_server"), cfg.get("ollama_port"))
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
