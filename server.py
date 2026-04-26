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
import mimetypes
import subprocess
import platform
import shutil
import threading
import time
import re
import urllib.error
import urllib.parse
import urllib.request
import warnings
from collections import deque
from pathlib import Path
from io import BytesIO
from functools import partial
from typing import Optional
from concurrent.futures import ThreadPoolExecutor
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse, HTMLResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from PIL import Image

from tag2_captions import (
    _build_caption_text,
    _get_caption_path,
    _read_caption_file,
    _write_caption_file,
)
from tag2_metadata import (
    _apply_metadata_changes,
    _get_metadata_path,
    _read_metadata_file,
    _write_metadata_file,
)
from tag2_images import (
    IMAGE_EXTENSIONS,
    MASK_SIDECAR_SUFFIX,
    MEDIA_EXTENSIONS,
    PREVIEW_MAX_SIZE,
    THUMBNAIL_SIZES,
    VIDEO_EXTENSIONS,
    _apply_real_crop,
    _build_clip_output_path,
    _clear_thumbnail_cache_for_path,
    _clip_video_file,
    _crop_video_file,
    _encode_media_for_ollama,
    _extract_video_frame,
    _flip_image,
    _generate_thumbnail,
    _get_media_type_for_path,
    _get_crop_backup_dir,
    _get_crop_backup_path,
    _get_crop_mask_backup_path,
    _get_display_image_size,
    _get_image_crop,
    _get_image_mask_info,
    _get_image_mask_path,
    _get_video_mask_frame_info,
    _get_video_mask_info,
    _get_thumbnail,
    _convert_gif_to_mp4_file,
    _is_image_path,
    _is_mask_sidecar_path,
    _is_video_path,
    _list_video_mask_paths,
    _load_oriented_image,
    _normalize_crop_rect,
    _normalize_exif_bytes,
    _normalize_image_key,
    _probe_video_info,
    _remove_real_crop,
    _render_image_bytes,
    _rotate_image,
    _rotate_image_file,
    _save_edited_image,
    _save_image_mask,
    _save_video_mask,
    _save_image_file,
    _suspend_thumbnail_processing_for_path,
    _resume_thumbnail_processing_for_path,
    _write_default_video_mask,
    _write_default_image_mask,
)
from tag2_ollama import (
    _apply_media_prompt_context,
    _auto_caption_sections,
    _auto_caption_captions,
    _build_ollama_generate_payload,
    _compose_ollama_host,
    _extract_free_text_lines,
    _get_ollama_enable_free_text,
    _get_ollama_free_text_prompt_template,
    _get_ollama_group_prompt_template,
    _get_ollama_host,
    _get_ollama_max_output_tokens,
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
    _ollama_prompt_for_caption,
    _ollama_response_meta,
    _parse_ollama_selection,
    _parse_ollama_yes_no,
    _split_ollama_host,
    _suggest_free_text,
)
from tag2_sections import (
    _caption_values,
    _all_headers_from_sections,
    _all_captions_from_sections,
    _apply_caption_selection,
    _find_group_for_caption,
    _get_crop_aspect_ratios,
    _get_folder_sections,
    _get_group_target,
    _group_hidden_captions,
    _is_general_section_name,
    _is_hidden_group_caption,
    _iter_caption_targets,
    _iter_caption_targets_with_indices,
    _normalize_enabled_captions,
    _ordered_sections_for_output,
    _remove_group_from_sections,
    _remove_section_from_sections,
    _remove_caption_from_sections,
    _rename_section_in_sections,
    _rename_caption_in_sections,
    _set_folder_sections,
)

_all_sentences_from_sections = _all_captions_from_sections
_apply_sentence_selection = _apply_caption_selection
_find_group_for_sentence = _find_group_for_caption
_group_hidden_sentences = _group_hidden_captions
_is_hidden_group_sentence = _is_hidden_group_caption
_normalize_enabled_sentences = _normalize_enabled_captions
_remove_sentence_from_sections = _remove_caption_from_sections
_rename_sentence_in_sections = _rename_caption_in_sections
_auto_caption_sentences = _auto_caption_captions
_ollama_prompt_for_sentence = _ollama_prompt_for_caption

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

# Processing defaults aim to keep the UI responsive while using most available cores.
HOST_CPU_COUNT = max(1, os.cpu_count() or 1)
INITIAL_RESERVED_CORES = max(1, min(4, HOST_CPU_COUNT // 4 or 1))
DEFAULT_EXECUTOR_WORKERS = max(1, min(16, HOST_CPU_COUNT - INITIAL_RESERVED_CORES))

# Thread pool for thumbnail / preview generation.
executor = ThreadPoolExecutor(max_workers=DEFAULT_EXECUTOR_WORKERS)
video_job_lock = threading.Lock()
video_job_pending: deque[str] = deque()
video_job_registry: dict[str, dict] = {}
video_job_recent: list[str] = []
video_job_active_id: str | None = None
video_job_wakeup = threading.Event()
MAX_VIDEO_JOB_HISTORY = 24
comfyui_job_lock = threading.Lock()
comfyui_job_registry: dict[str, dict] = {}
MAX_COMFYUI_JOB_HISTORY = 96
media_stream_lock = threading.Lock()
media_stream_active_counts: dict[str, int] = {}
media_stream_cancelled_paths: set[str] = set()
MEDIA_STREAM_CHUNK_SIZE = 256 * 1024
MEDIA_RANGE_HEADER_PATTERN = re.compile(r"bytes=(\d*)-(\d*)$", re.IGNORECASE)

# In-memory thumbnail cache lives in `tag2_images`.

# ===== CONFIG FILE =====
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
DEFAULT_OLLAMA_SERVER = "127.0.0.1"
DEFAULT_OLLAMA_PORT = 11434
DEFAULT_OLLAMA_TIMEOUT_SECONDS = 20
DEFAULT_OLLAMA_MAX_OUTPUT_TOKENS = 64
DEFAULT_OLLAMA_MODEL = "llava"
DEFAULT_COMFYUI_SERVER = "127.0.0.1"
DEFAULT_COMFYUI_PORT = 8188
DEFAULT_COMFYUI_WORKFLOW_PATH = ""
DEFAULT_COMFYUI_OUTPUT_FOLDER = ""
DEFAULT_COMFYUI_AUTO_PREVIEW = False
DEFAULT_CROP_ASPECT_RATIOS = ["4:3", "16:9", "3:4", "1:1", "9:16", "2:3", "3:2"]
DEFAULT_MASK_LATENT_BASE_WIDTH_PRESETS = [512, 768, 1024, 1280]
DEFAULT_HTTPS_CERTFILE = ""
DEFAULT_HTTPS_KEYFILE = ""
DEFAULT_HTTPS_PORT = 8900
DEFAULT_REMOTE_HTTP_MODE = "redirect-to-https"
DEFAULT_FFMPEG_PATH = ""
DEFAULT_FFMPEG_THREADS = 0
DEFAULT_FFMPEG_HWACCEL = "auto"
DEFAULT_PROCESSING_RESERVED_CORES = INITIAL_RESERVED_CORES
DEFAULT_FOLDER_SUGGESTION_LIMIT = 12
REMOTE_HTTP_MODES = {"allow", "redirect-to-https", "block"}
DEFAULT_VIDEO_TRAINING_PRESET_DEFINITIONS = [
    {
        "key": "wan-40f-16fps",
        "label": "Wan 40f @ 16 fps",
        "target_family": "wan",
        "num_frames": 40,
        "fps": 16,
        "shrink_video_to_frames": True,
        "short_clip_factor": 0.75,
        "long_clip_factor": 1.5,
        "preferred_extensions": [".mp4", ".mov", ".webm"],
        "description": "Good baseline for short Wan character or identity clips.",
    },
    {
        "key": "wan-81f-16fps",
        "label": "Wan 81f @ 16 fps",
        "target_family": "wan",
        "num_frames": 81,
        "fps": 16,
        "shrink_video_to_frames": True,
        "short_clip_factor": 0.8,
        "long_clip_factor": 1.35,
        "preferred_extensions": [".mp4", ".mov", ".webm"],
        "description": "Longer Wan motion clips with more temporal context.",
    },
    {
        "key": "ltx-97f-24fps",
        "label": "LTX 97f @ 24 fps",
        "target_family": "ltx",
        "num_frames": 97,
        "fps": 24,
        "shrink_video_to_frames": True,
        "short_clip_factor": 0.8,
        "long_clip_factor": 1.3,
        "preferred_extensions": [".mp4", ".mov"],
        "description": "Default LTX profile with a tighter duration window.",
    },
]
DEFAULT_OLLAMA_PROMPT_TEMPLATE = (
    "You are verifying a caption for one media item. "
    "Reply with exactly one word: YES or NO. "
    "Reply YES only if the caption is clearly correct for the media. "
    "Reply NO if it is wrong, uncertain, too specific, or not clearly visible.\n\n"
    "Caption: {caption}\n"
    "Answer:"
)
DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE = (
    "You are selecting the single best caption for one media item from a numbered list. "
    "Reply with exactly one number from 1 to {count}. "
    "Pick the most likely correct caption for the media.\n\n"
    "Group: {group_name}\n"
    "{options}\n\n"
    "Answer:"
)
DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE = (
    "You are improving a media caption file. The caption text below already covers known details and must not be repeated. "
    "Look at the media and return only notable, important visual details that are still missing. "
    "Return either NONE or one short line per missing detail, with no bullets or numbering.\n\n"
    "Current caption text:\n{caption_text}\n\n"
    "Answer:"
)
COMFYUI_PROMPT_PLACEHOLDER = "{{TAG2_PROMPT}}"
COMFYUI_FILENAME_PREFIX_PLACEHOLDER = "{{TAG2_FILENAME_PREFIX}}"
COMFYUI_SEED_PLACEHOLDER = "{{TAG2_SEED}}"
COMFYUI_RANDOM_SEED_PLACEHOLDER = "{{TAG2_RANDOM_SEED}}"
COMFYUI_RANDOMIZED_SEED_INPUT_NAMES = {"seed", "noise_seed", "seed_num", "sampler_seed"}


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


def _compose_comfyui_host(server: str | None, port: int | None, default_server: str, default_port: int) -> str:
    """Return the normalized HTTP base URL for the configured ComfyUI instance."""
    server_name = str(server or default_server).strip() or default_server
    try:
        port_number = int(port or default_port)
    except (TypeError, ValueError):
        port_number = default_port
    port_number = max(1, min(65535, port_number))
    if server_name.startswith("http://") or server_name.startswith("https://"):
        return server_name.rstrip("/")
    return f"http://{server_name}:{port_number}"


def _get_comfyui_server(cfg: dict, default_server: str = DEFAULT_COMFYUI_SERVER) -> str:
    """Return the configured ComfyUI host name without protocol."""
    return str(cfg.get("comfyui_server") or default_server).strip() or default_server


def _get_comfyui_port(cfg: dict, default_port: int = DEFAULT_COMFYUI_PORT) -> int:
    """Return the configured ComfyUI port."""
    try:
        port = int(cfg.get("comfyui_port", default_port))
    except (TypeError, ValueError):
        port = default_port
    return max(1, min(65535, port))


def _get_comfyui_host(cfg: dict, default_server: str = DEFAULT_COMFYUI_SERVER, default_port: int = DEFAULT_COMFYUI_PORT) -> str:
    """Return the normalized HTTP base URL for the configured ComfyUI instance."""
    return _compose_comfyui_host(
        cfg.get("comfyui_server"),
        cfg.get("comfyui_port"),
        default_server,
        default_port,
    )


def _get_comfyui_workflow_path(cfg: dict, default_path: str = DEFAULT_COMFYUI_WORKFLOW_PATH) -> str:
    """Return the configured ComfyUI workflow API JSON path."""
    return _resolve_config_path(cfg.get("comfyui_workflow_path") or default_path)


def _get_comfyui_output_folder(cfg: dict, default_path: str = DEFAULT_COMFYUI_OUTPUT_FOLDER) -> str:
    """Return the configured ComfyUI output directory."""
    return _resolve_config_path(cfg.get("comfyui_output_folder") or default_path)


def _get_comfyui_auto_preview(cfg: dict, default_value: bool = DEFAULT_COMFYUI_AUTO_PREVIEW) -> bool:
    """Return whether auto-preview should queue ComfyUI jobs after auto-caption completes."""
    return bool(cfg.get("comfyui_auto_preview", default_value))


def _fetch_comfyui_json(host: str, route: str, timeout: int = 20) -> dict:
    """Fetch one JSON payload from the ComfyUI HTTP API."""
    url = f"{host.rstrip('/')}/{route.lstrip('/')}"
    request = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(detail or str(exc)) from exc
    except TimeoutError as exc:
        raise RuntimeError(f"request timed out after {timeout} seconds") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason or exc)) from exc


def _post_comfyui_json(host: str, route: str, payload: dict, timeout: int = 20) -> dict:
    """Post one JSON payload to the ComfyUI HTTP API."""
    url = f"{host.rstrip('/')}/{route.lstrip('/')}"
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        method="POST",
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(detail or str(exc)) from exc
    except TimeoutError as exc:
        raise RuntimeError(f"request timed out after {timeout} seconds") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason or exc)) from exc


def _load_comfyui_workflow_template(workflow_path: str) -> dict:
    """Read and validate the configured ComfyUI API prompt-graph JSON from disk."""
    normalized_path = os.path.abspath(os.path.normpath(str(workflow_path or "").strip()))
    if not normalized_path:
        raise RuntimeError("ComfyUI workflow path is not configured")
    if not os.path.isfile(normalized_path):
        raise RuntimeError(f"ComfyUI workflow file not found: {normalized_path}")
    try:
        with open(normalized_path, "r", encoding="utf-8") as f:
            workflow = json.load(f)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Invalid ComfyUI workflow JSON: {exc}") from exc
    if not isinstance(workflow, dict) or not workflow:
        raise RuntimeError("ComfyUI workflow JSON must be a non-empty object")

    prompt_graph = workflow.get("prompt") if isinstance(workflow.get("prompt"), dict) and workflow.get("prompt") else workflow
    if not isinstance(prompt_graph, dict) or not prompt_graph:
        raise RuntimeError("ComfyUI API prompt graph must be a non-empty object")

    if isinstance(prompt_graph.get("nodes"), list):
        raise RuntimeError(
            "This looks like a normal ComfyUI workflow, not an API prompt graph. "
            "Export a second file with Save (API format). The API JSON is for Tag2 and will not reopen in the ComfyUI canvas."
        )

    invalid_top_level_keys = [key for key, node in prompt_graph.items() if not isinstance(node, dict)]
    if invalid_top_level_keys:
        raise RuntimeError(
            f"ComfyUI API prompt graph is invalid at top-level key '{invalid_top_level_keys[0]}'. "
            "Export with Save (API format), not the normal workflow JSON."
        )

    invalid_nodes = [
        key for key, node in prompt_graph.items()
        if "class_type" not in node or "inputs" not in node
    ]
    if invalid_nodes:
        raise RuntimeError(
            f"ComfyUI API prompt graph is invalid at node '{invalid_nodes[0]}': missing class_type or inputs. "
            "Export with Save (API format)."
        )

    return prompt_graph


def _replace_comfyui_workflow_placeholders(value, replacements: dict[str, str]):
    """Recursively replace placeholder tokens inside a workflow JSON payload."""
    if isinstance(value, str):
        result = value
        for placeholder, replacement in replacements.items():
            result = result.replace(placeholder, replacement)
        return result
    if isinstance(value, list):
        return [_replace_comfyui_workflow_placeholders(item, replacements) for item in value]
    if isinstance(value, dict):
        return {key: _replace_comfyui_workflow_placeholders(item, replacements) for key, item in value.items()}
    return value


def _find_comfyui_placeholder_usage(value, placeholder: str) -> bool:
    """Return whether a workflow JSON payload references a specific placeholder."""
    if isinstance(value, str):
        return placeholder in value
    if isinstance(value, list):
        return any(_find_comfyui_placeholder_usage(item, placeholder) for item in value)
    if isinstance(value, dict):
        return any(_find_comfyui_placeholder_usage(item, placeholder) for item in value.values())
    return False


def _generate_comfyui_seed() -> int:
    """Return one preview seed suitable for ComfyUI sampler nodes."""
    return uuid4().int & 0xFFFFFFFF


def _coerce_comfyui_seed_input(value, seed: int):
    """Replace one supported ComfyUI seed input with the generated preview seed."""
    if isinstance(value, str):
        normalized = value.strip()
        if normalized in {COMFYUI_SEED_PLACEHOLDER, COMFYUI_RANDOM_SEED_PLACEHOLDER}:
            return seed
        if re.fullmatch(r"-?\d+", normalized):
            return seed
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return seed
    return value


def _inject_comfyui_seed_inputs(prompt_payload: dict, seed: int) -> None:
    """Randomize known seed-style inputs in one ComfyUI API prompt graph."""
    if not isinstance(prompt_payload, dict):
        return
    for node in prompt_payload.values():
        if not isinstance(node, dict):
            continue
        inputs = node.get("inputs")
        if not isinstance(inputs, dict):
            continue
        for input_name, input_value in list(inputs.items()):
            if str(input_name or "").strip().casefold() not in COMFYUI_RANDOMIZED_SEED_INPUT_NAMES:
                continue
            inputs[input_name] = _coerce_comfyui_seed_input(input_value, seed)


def _build_comfyui_filename_prefix(source_path: str) -> str:
    """Return a sortable filename prefix that stays grouped under the source stem."""
    stem = Path(str(source_path or "")).stem.strip() or "tag2-preview"
    return f"{stem}__tag2__{time.time_ns()}"


def _build_comfyui_prompt_payload(workflow_template: dict, prompt_text: str, filename_prefix: str) -> dict:
    """Build one ComfyUI prompt payload by replacing supported placeholder tokens."""
    if not _find_comfyui_placeholder_usage(workflow_template, COMFYUI_PROMPT_PLACEHOLDER):
        raise RuntimeError(f"ComfyUI workflow must reference {COMFYUI_PROMPT_PLACEHOLDER}")
    if not _find_comfyui_placeholder_usage(workflow_template, COMFYUI_FILENAME_PREFIX_PLACEHOLDER):
        raise RuntimeError(f"ComfyUI workflow must reference {COMFYUI_FILENAME_PREFIX_PLACEHOLDER}")
    preview_seed = _generate_comfyui_seed()
    replacements = {
        COMFYUI_PROMPT_PLACEHOLDER: str(prompt_text or ""),
        COMFYUI_FILENAME_PREFIX_PLACEHOLDER: str(filename_prefix or ""),
        COMFYUI_SEED_PLACEHOLDER: str(preview_seed),
        COMFYUI_RANDOM_SEED_PLACEHOLDER: str(preview_seed),
    }
    prompt_payload = _replace_comfyui_workflow_placeholders(copy.deepcopy(workflow_template), replacements)
    _inject_comfyui_seed_inputs(prompt_payload, preview_seed)
    return prompt_payload


def _queue_comfyui_prompt(host: str, prompt: dict) -> dict:
    """Submit one prompt graph to ComfyUI and return the queue response."""
    if not isinstance(prompt, dict) or not prompt:
        raise RuntimeError("ComfyUI prompt payload is empty")
    return _post_comfyui_json(host, "/prompt", {"prompt": prompt})


def _matches_comfyui_preview_filename(entry_stem: str, filename_prefix: str) -> bool:
    """Return whether a ComfyUI output filename belongs to the selected source stem."""
    normalized_stem = str(entry_stem or "").casefold()
    normalized_prefix = str(filename_prefix or "").casefold()
    if not normalized_prefix:
        return False
    if normalized_stem == normalized_prefix:
        return True
    if re.fullmatch(rf"{re.escape(normalized_prefix)}_\d+(?:_.*)?", normalized_stem):
        return True
    return bool(
        re.fullmatch(
            rf"{re.escape(normalized_prefix)}__tag2__\d+(?:_\d+(?:_.*)?)?",
            normalized_stem,
        )
    )


def _scan_comfyui_preview_files(output_folder: str, source_path: str) -> list[str]:
    """Return generated preview images that share the selected source filename stem."""
    normalized_folder = os.path.abspath(os.path.normpath(str(output_folder or "").strip()))
    if not normalized_folder:
        raise RuntimeError("ComfyUI output folder is not configured")
    if not os.path.isdir(normalized_folder):
        raise RuntimeError(f"ComfyUI output folder not found: {normalized_folder}")

    prefix = Path(str(source_path or "")).stem
    if not prefix:
        return []

    def sort_key(path_obj: Path):
        tokens = re.split(r"(\d+)", path_obj.stem.lower())
        normalized_tokens = tuple((1, int(token)) if token.isdigit() else (0, token) for token in tokens)
        return normalized_tokens, path_obj.name.lower()

    matches: list[str] = []
    for entry in Path(normalized_folder).iterdir():
        if not entry.is_file():
            continue
        if entry.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        if not _matches_comfyui_preview_filename(entry.stem, prefix):
            continue
        matches.append(str(entry))
    matches.sort(key=lambda item: sort_key(Path(item)))
    return matches


def _select_latest_comfyui_preview_file(files: list[str]) -> str:
    """Return the newest preview file path from a matched preview file list."""
    existing_files = [str(path) for path in files or [] if os.path.isfile(path)]
    if not existing_files:
        return ""
    return max(
        existing_files,
        key=lambda item: (float(os.path.getmtime(item)), Path(item).name.lower()),
    )


def _get_comfyui_queue_prompt_ids(queue_payload: dict) -> tuple[set[str], set[str]]:
    """Extract running and pending prompt ids from the ComfyUI queue response."""
    running: set[str] = set()
    pending: set[str] = set()

    def collect_ids(items, target: set[str]):
        for item in items or []:
            if isinstance(item, dict):
                prompt_id = str(item.get("prompt_id") or item.get("id") or "").strip()
                if prompt_id:
                    target.add(prompt_id)
                continue
            if isinstance(item, (list, tuple)):
                for value in item:
                    if isinstance(value, str) and value:
                        target.add(value)
                        break

    collect_ids(queue_payload.get("queue_running"), running)
    collect_ids(queue_payload.get("queue_pending"), pending)
    return running, pending


def _prune_comfyui_job_registry_locked() -> None:
    """Keep the in-memory ComfyUI job registry from growing without bound."""
    if len(comfyui_job_registry) <= MAX_COMFYUI_JOB_HISTORY:
        return
    removable = sorted(
        comfyui_job_registry.items(),
        key=lambda item: float(item[1].get("created_at") or 0.0),
    )
    for prompt_id, job in removable:
        if len(comfyui_job_registry) <= MAX_COMFYUI_JOB_HISTORY:
            break
        if str(job.get("status") or "") in {"queued", "running"}:
            continue
        comfyui_job_registry.pop(prompt_id, None)


def _upsert_comfyui_job_record(prompt_id: str, **patch) -> dict | None:
    """Create or update one tracked ComfyUI job record."""
    normalized_prompt_id = str(prompt_id or "").strip()
    if not normalized_prompt_id:
        return None
    with comfyui_job_lock:
        job = comfyui_job_registry.get(normalized_prompt_id)
        now = time.time()
        if not job:
            job = {
                "prompt_id": normalized_prompt_id,
                "status": "queued",
                "created_at": now,
                "updated_at": now,
                "message": "Queued",
                "error": "",
                "image_path": "",
                "folder": "",
                "filename_prefix": "",
                "caption_text": "",
            }
            comfyui_job_registry[normalized_prompt_id] = job
        job.update(patch)
        job["updated_at"] = now
        _prune_comfyui_job_registry_locked()
        return dict(job)


def _get_comfyui_jobs_for_image(image_path: str) -> list[dict]:
    """Return tracked ComfyUI jobs for one source image sorted by creation time."""
    normalized_path = os.path.abspath(os.path.normpath(str(image_path or "").strip()))
    with comfyui_job_lock:
        jobs = [
            dict(job)
            for job in comfyui_job_registry.values()
            if os.path.abspath(os.path.normpath(str(job.get("image_path") or "").strip())) == normalized_path
        ]
    jobs.sort(key=lambda item: float(item.get("created_at") or 0.0))
    return jobs


def _refresh_comfyui_jobs_for_image(cfg: dict, image_path: str) -> dict:
    """Refresh tracked ComfyUI jobs for one source image against the live ComfyUI queue/history."""
    normalized_path = os.path.abspath(os.path.normpath(str(image_path or "").strip()))
    jobs = _get_comfyui_jobs_for_image(normalized_path)
    output_files = _scan_comfyui_preview_files(_get_comfyui_output_folder(cfg), normalized_path)
    latest_output_path = _select_latest_comfyui_preview_file(output_files)
    if not jobs:
        return {
            "jobs": [],
            "summary": {
                "total": 0,
                "spawned": 0,
                "queued": 0,
                "running": 0,
                "completed": 0,
                "failed": 0,
                "latest_prompt_id": "",
                "latest_output_path": latest_output_path,
            },
            "files": output_files,
        }

    host = _get_comfyui_host(cfg, DEFAULT_COMFYUI_SERVER, DEFAULT_COMFYUI_PORT)
    queue_payload = _fetch_comfyui_json(host, "/queue")
    running_ids, pending_ids = _get_comfyui_queue_prompt_ids(queue_payload)
    refreshed_jobs: list[dict] = []
    for job in jobs:
        prompt_id = str(job.get("prompt_id") or "").strip()
        if not prompt_id:
            continue
        next_status = str(job.get("status") or "queued")
        message = str(job.get("message") or "")
        error = str(job.get("error") or "")
        if prompt_id in running_ids:
            next_status = "running"
            message = "Running in ComfyUI"
        elif prompt_id in pending_ids:
            next_status = "queued"
            message = "Queued in ComfyUI"
        else:
            history_payload = _fetch_comfyui_json(host, f"/history/{urllib.parse.quote(prompt_id)}")
            history_entry = history_payload.get(prompt_id) if isinstance(history_payload, dict) else None
            if history_entry:
                status_info = history_entry.get("status") if isinstance(history_entry, dict) else None
                completed = bool((status_info or {}).get("completed")) if isinstance(status_info, dict) else False
                history_error = (status_info or {}).get("messages") if isinstance(status_info, dict) else None
                if completed:
                    next_status = "completed"
                    message = "Completed"
                elif history_error:
                    next_status = "error"
                    error = json.dumps(history_error, ensure_ascii=False)
                    message = "ComfyUI execution failed"
                else:
                    next_status = "completed"
                    message = "Completed"
            elif next_status not in {"completed", "error"}:
                next_status = "completed" if output_files else "unknown"
                message = "Finished" if output_files else "Prompt not found in live queue"
        refreshed = _upsert_comfyui_job_record(prompt_id, status=next_status, message=message, error=error) or job
        refreshed_jobs.append(refreshed)

    refreshed_jobs.sort(key=lambda item: float(item.get("created_at") or 0.0))
    summary = {
        "total": len(refreshed_jobs),
        "spawned": len(refreshed_jobs),
        "queued": sum(1 for job in refreshed_jobs if job.get("status") == "queued"),
        "running": sum(1 for job in refreshed_jobs if job.get("status") == "running"),
        "completed": sum(1 for job in refreshed_jobs if job.get("status") == "completed"),
        "failed": sum(1 for job in refreshed_jobs if job.get("status") == "error"),
        "latest_prompt_id": str(refreshed_jobs[-1].get("prompt_id") or "") if refreshed_jobs else "",
        "latest_output_path": latest_output_path,
    }
    return {"jobs": refreshed_jobs, "summary": summary, "files": output_files}


def _load_config() -> dict:
    """Load config from disk. Returns default structure if file missing."""
    default = {
        "last_folder": "",
        # Captions shared across all folders when a folder has no own config yet.
        "default_captions": [],
        "thumb_size": 160,
        "crop_aspect_ratios": list(DEFAULT_CROP_ASPECT_RATIOS),
        "mask_latent_base_width_presets": list(DEFAULT_MASK_LATENT_BASE_WIDTH_PRESETS),
        "image_crops": {},
        "https_certfile": DEFAULT_HTTPS_CERTFILE,
        "https_keyfile": DEFAULT_HTTPS_KEYFILE,
        "https_port": DEFAULT_HTTPS_PORT,
        "remote_http_mode": DEFAULT_REMOTE_HTTP_MODE,
        "ffmpeg_path": DEFAULT_FFMPEG_PATH,
        "ffmpeg_threads": DEFAULT_FFMPEG_THREADS,
        "ffmpeg_hwaccel": DEFAULT_FFMPEG_HWACCEL,
        "processing_reserved_cores": DEFAULT_PROCESSING_RESERVED_CORES,
        "video_training_presets": copy.deepcopy(DEFAULT_VIDEO_TRAINING_PRESET_DEFINITIONS),
        # Local Ollama server settings
        "ollama_server": DEFAULT_OLLAMA_SERVER,
        "ollama_port": DEFAULT_OLLAMA_PORT,
        "ollama_timeout_seconds": DEFAULT_OLLAMA_TIMEOUT_SECONDS,
        "ollama_max_output_tokens": DEFAULT_OLLAMA_MAX_OUTPUT_TOKENS,
        "ollama_host": f"http://{DEFAULT_OLLAMA_SERVER}:{DEFAULT_OLLAMA_PORT}",
        "ollama_model": DEFAULT_OLLAMA_MODEL,
        "ollama_prompt_template": DEFAULT_OLLAMA_PROMPT_TEMPLATE,
        "ollama_group_prompt_template": DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE,
        "ollama_enable_free_text": True,
        "ollama_free_text_prompt_template": DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE,
        # Local ComfyUI prompt-preview settings
        "comfyui_server": DEFAULT_COMFYUI_SERVER,
        "comfyui_port": DEFAULT_COMFYUI_PORT,
        "comfyui_workflow_path": DEFAULT_COMFYUI_WORKFLOW_PATH,
        "comfyui_output_folder": DEFAULT_COMFYUI_OUTPUT_FOLDER,
        "comfyui_auto_preview": DEFAULT_COMFYUI_AUTO_PREVIEW,
        # Per-folder caption lists. Key = absolute folder path.
        # To copy captions to a new folder, duplicate a block here.
        "folders": {}
    }
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "default_captions" not in data and "default_sentences" in data:
                data["default_captions"] = list(data.get("default_sentences") or [])
            data.pop("default_sentences", None)
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
            data.setdefault("comfyui_server", DEFAULT_COMFYUI_SERVER)
            data.setdefault("comfyui_port", DEFAULT_COMFYUI_PORT)
            data.setdefault("comfyui_workflow_path", DEFAULT_COMFYUI_WORKFLOW_PATH)
            data.setdefault("comfyui_output_folder", DEFAULT_COMFYUI_OUTPUT_FOLDER)
            data.setdefault("comfyui_auto_preview", DEFAULT_COMFYUI_AUTO_PREVIEW)
            # Merge with defaults so new keys are always present
            for k, v in default.items():
                data.setdefault(k, v)
            data["mask_latent_base_width_presets"] = _normalize_mask_latent_base_width_presets(
                data.get("mask_latent_base_width_presets")
            )
            return data
        except Exception:
            pass
    return default


def _normalize_mask_latent_base_width_presets(raw_presets) -> list[int]:
    """Normalize mask latent base-width presets into a sorted unique list."""
    normalized: list[int] = []
    seen: set[int] = set()
    for raw_value in raw_presets if isinstance(raw_presets, list) else []:
        try:
            value = int(raw_value)
        except (TypeError, ValueError):
            continue
        value = max(64, min(2048, value))
        if value in seen:
            continue
        seen.add(value)
        normalized.append(value)
    normalized.sort()
    return normalized or list(DEFAULT_MASK_LATENT_BASE_WIDTH_PRESETS)


def _save_config(cfg: dict):
    """Persist config to disk with pretty formatting for easy hand-editing."""
    cfg = copy.deepcopy(cfg)
    if "default_captions" not in cfg and "default_sentences" in cfg:
        cfg["default_captions"] = list(cfg.get("default_sentences") or [])
    cfg.pop("default_sentences", None)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def _coalesce_caption_query(captions: str | None, sentences: str | None = None) -> str:
    if captions is not None:
        return captions
    if sentences is not None:
        return sentences
    return "[]"


def _read_enabled_captions(data: dict | None) -> list[str]:
    payload = dict(data or {})
    enabled_captions = payload.get("enabled_captions")
    if isinstance(enabled_captions, list):
        return list(enabled_captions)
    enabled_sentences = payload.get("enabled_sentences")
    if isinstance(enabled_sentences, list):
        return list(enabled_sentences)
    return []


def _read_caption_text(value: str | None, legacy_value: str | None = None) -> str:
    return str(value or legacy_value or "").strip()


def _slugify_video_training_preset_key(value: str, fallback: str) -> str:
    """Convert a preset label or key into a config-safe identifier."""
    raw = str(value or fallback).strip().lower()
    pieces: list[str] = []
    previous_dash = False
    for char in raw:
        if char.isalnum():
            pieces.append(char)
            previous_dash = False
            continue
        if previous_dash:
            continue
        pieces.append("-")
        previous_dash = True
    normalized = "".join(pieces).strip("-")
    return normalized or fallback


def _normalize_video_training_preset(raw_preset: dict, index: int, used_keys: set[str]) -> dict | None:
    """Validate and normalize one editable video training preset."""
    if not isinstance(raw_preset, dict):
        return None
    fallback_key = f"profile-{index + 1}"
    base_key = _slugify_video_training_preset_key(
        raw_preset.get("key") or raw_preset.get("label") or raw_preset.get("name"),
        fallback_key,
    )
    key = base_key
    counter = 2
    while key in used_keys:
        key = f"{base_key}-{counter}"
        counter += 1
    used_keys.add(key)

    label = str(raw_preset.get("label") or raw_preset.get("name") or key).strip() or key
    target_family = str(raw_preset.get("target_family") or "custom").strip().lower() or "custom"
    if target_family not in {"wan", "ltx", "hunyuan", "custom"}:
        target_family = "custom"
    try:
        num_frames = max(1, int(raw_preset.get("num_frames") or 1))
    except (TypeError, ValueError):
        num_frames = 1
    try:
        fps = max(1, int(raw_preset.get("fps") or 16))
    except (TypeError, ValueError):
        fps = 16
    try:
        short_clip_factor = float(raw_preset.get("short_clip_factor") or 0.8)
    except (TypeError, ValueError):
        short_clip_factor = 0.8
    try:
        long_clip_factor = float(raw_preset.get("long_clip_factor") or 1.5)
    except (TypeError, ValueError):
        long_clip_factor = 1.5
    short_clip_factor = max(0.25, min(1.0, short_clip_factor))
    long_clip_factor = max(short_clip_factor, min(4.0, long_clip_factor))
    preferred_extensions: list[str] = []
    for extension in raw_preset.get("preferred_extensions") or []:
        normalized_extension = str(extension or "").strip().lower()
        if not normalized_extension:
            continue
        if not normalized_extension.startswith("."):
            normalized_extension = f".{normalized_extension}"
        if normalized_extension not in VIDEO_EXTENSIONS or normalized_extension in preferred_extensions:
            continue
        preferred_extensions.append(normalized_extension)
    ideal_clip_seconds = round(num_frames / max(1, fps), 3)
    return {
        "key": key,
        "label": label,
        "target_family": target_family,
        "num_frames": num_frames,
        "fps": fps,
        "shrink_video_to_frames": bool(raw_preset.get("shrink_video_to_frames", True)),
        "short_clip_factor": short_clip_factor,
        "long_clip_factor": long_clip_factor,
        "ideal_clip_seconds": ideal_clip_seconds,
        "min_clip_seconds": round(ideal_clip_seconds * short_clip_factor, 3),
        "max_clip_seconds": round(ideal_clip_seconds * long_clip_factor, 3),
        "preferred_extensions": preferred_extensions,
        "description": str(raw_preset.get("description") or "").strip(),
    }


def _normalize_video_training_presets(raw_presets) -> list[dict]:
    """Return normalized editable video training presets."""
    preset_source = raw_presets if isinstance(raw_presets, list) else copy.deepcopy(DEFAULT_VIDEO_TRAINING_PRESET_DEFINITIONS)
    normalized: list[dict] = []
    used_keys: set[str] = set()
    for index, raw_preset in enumerate(preset_source):
        preset = _normalize_video_training_preset(raw_preset, index, used_keys)
        if preset is not None:
            normalized.append(preset)
    if normalized:
        return normalized
    fallback_used_keys: set[str] = set()
    fallback_presets: list[dict] = []
    for index, raw_preset in enumerate(DEFAULT_VIDEO_TRAINING_PRESET_DEFINITIONS):
        preset = _normalize_video_training_preset(raw_preset, index, fallback_used_keys)
        if preset is not None:
            fallback_presets.append(preset)
    return fallback_presets


def _get_video_training_presets(cfg: dict) -> list[dict]:
    """Return the normalized video training preset library."""
    return _normalize_video_training_presets(cfg.get("video_training_presets"))


def _get_video_training_profile_by_key(presets: list[dict], key: str | None) -> dict | None:
    """Return one normalized preset by key."""
    normalized_key = str(key or "").strip()
    if not normalized_key:
        return presets[0] if presets else None
    for preset in presets:
        if preset.get("key") == normalized_key:
            return preset
    return presets[0] if presets else None


def _get_folder_video_training_profile_key(cfg: dict, folder: str, presets: list[dict] | None = None) -> str:
    """Return the selected video training preset key for a folder."""
    normalized_folder = os.path.normpath(str(folder or ""))
    presets = presets if presets is not None else _get_video_training_presets(cfg)
    default_key = presets[0].get("key") if presets else ""
    folder_cfg = cfg.get("folders", {}).get(normalized_folder, {}) if normalized_folder else {}
    selected_key = str(folder_cfg.get("video_training_profile_key") or "").strip()
    if _get_video_training_profile_by_key(presets, selected_key):
        return selected_key or default_key
    return default_key


def _set_folder_video_training_profile_key(cfg: dict, folder: str, profile_key: str, presets: list[dict] | None = None) -> str:
    """Persist the selected video training preset key for a folder."""
    normalized_folder = os.path.normpath(str(folder or ""))
    if not normalized_folder:
        return ""
    presets = presets if presets is not None else _get_video_training_presets(cfg)
    selected_preset = _get_video_training_profile_by_key(presets, profile_key)
    selected_key = selected_preset.get("key") if selected_preset else ""
    folders_cfg = cfg.setdefault("folders", {})
    folder_cfg = folders_cfg.setdefault(normalized_folder, {})
    folder_cfg["video_training_profile_key"] = selected_key
    return selected_key


def _resolve_config_path(path_value: str | None) -> str:
    """Resolve a config path relative to the config file directory when needed."""
    raw_path = str(path_value or "").strip()
    if not raw_path:
        return ""
    if os.path.isabs(raw_path):
        return raw_path
    return os.path.abspath(os.path.join(os.path.dirname(CONFIG_PATH), raw_path))


def _get_ffmpeg_path(cfg: dict) -> str:
    """Return the configured ffmpeg path override, if any."""
    return str(cfg.get("ffmpeg_path") or "").strip()


def _get_processing_reserved_cores(cfg: dict) -> int:
    """Return how many CPU cores to keep free for responsiveness."""
    try:
        configured = int(cfg.get("processing_reserved_cores", DEFAULT_PROCESSING_RESERVED_CORES) or DEFAULT_PROCESSING_RESERVED_CORES)
    except (TypeError, ValueError):
        configured = DEFAULT_PROCESSING_RESERVED_CORES
    return max(0, min(HOST_CPU_COUNT - 1, configured)) if HOST_CPU_COUNT > 1 else 0


def _get_ffmpeg_threads(cfg: dict) -> int:
    """Return the ffmpeg thread budget for one processing job."""
    try:
        configured = int(cfg.get("ffmpeg_threads", DEFAULT_FFMPEG_THREADS) or 0)
    except (TypeError, ValueError):
        configured = 0
    if configured > 0:
        return max(1, configured)
    return max(1, HOST_CPU_COUNT - _get_processing_reserved_cores(cfg))


def _get_ffmpeg_hwaccel(cfg: dict) -> str:
    """Return the ffmpeg hardware acceleration mode."""
    mode = str(cfg.get("ffmpeg_hwaccel") or DEFAULT_FFMPEG_HWACCEL).strip().lower()
    return mode if mode in {"auto", "off"} else DEFAULT_FFMPEG_HWACCEL


def _get_tool_binary_name(base_name: str) -> str:
    """Return the platform-appropriate binary file name."""
    if os.name == "nt" and not base_name.lower().endswith(".exe"):
        return f"{base_name}.exe"
    return base_name


def _resolve_ffmpeg_binaries(cfg: dict) -> tuple[str, str]:
    """Resolve ffmpeg and ffprobe executable paths from config or PATH."""
    configured = _get_ffmpeg_path(cfg)
    ffmpeg_name = _get_tool_binary_name("ffmpeg")
    ffprobe_name = _get_tool_binary_name("ffprobe")
    if not configured:
        return ffmpeg_name, ffprobe_name

    resolved = _resolve_config_path(configured)
    if os.path.isdir(resolved):
        return os.path.join(resolved, ffmpeg_name), os.path.join(resolved, ffprobe_name)

    ffmpeg_binary = resolved
    ffprobe_binary = os.path.join(os.path.dirname(resolved), ffprobe_name) if os.path.dirname(resolved) else ffprobe_name
    return ffmpeg_binary, ffprobe_binary


def _build_media_entry(entry: Path, prompt_preview_path: str = "", prompt_preview_mtime: float = 0.0) -> dict:
    """Build a list response payload for a supported media file."""
    stat = entry.stat()
    caption_file = entry.with_suffix(".txt")
    media_type = _get_media_type_for_path(str(entry))
    mask_count = 0
    if media_type == "image":
        mask_count = 1 if _get_image_mask_path(str(entry)).exists() else 0
    elif media_type == "video":
        mask_count = int(_get_video_mask_info(str(entry)).get("count") or 0)
    return {
        "name": entry.name,
        "path": str(entry),
        "size": stat.st_size,
        "mtime": stat.st_mtime,
        "has_caption": caption_file.exists(),
        "has_mask": mask_count > 0,
        "mask_count": mask_count,
        "media_type": media_type,
        "prompt_preview_path": str(prompt_preview_path or ""),
        "prompt_preview_mtime": float(prompt_preview_mtime or 0.0),
    }


def _dump_model(model: BaseModel, **kwargs):
    """Compatibility wrapper for Pydantic v1/v2 model serialization."""
    if hasattr(model, "model_dump"):
        return model.model_dump(**kwargs)
    return model.dict(**kwargs)


def _normalize_existing_media_path(path: str) -> str:
    """Normalize and validate a media file path from API input."""
    normalized_path = os.path.abspath(os.path.normpath(str(path or "").strip()))
    if not os.path.isfile(normalized_path):
        raise HTTPException(status_code=404, detail="Media not found")
    if Path(normalized_path).suffix.lower() not in MEDIA_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Unsupported media file")
    return normalized_path


def _serialize_video_job(job: dict) -> dict:
    """Return a client-safe video job payload."""
    return {
        "id": job.get("id"),
        "type": job.get("type"),
        "video_path": job.get("video_path"),
        "folder": job.get("folder"),
        "status": job.get("status"),
        "progress": float(job.get("progress") or 0.0),
        "message": job.get("message") or "",
        "error": job.get("error") or "",
        "output_path": job.get("output_path") or "",
        "created_at": float(job.get("created_at") or 0.0),
        "updated_at": float(job.get("updated_at") or 0.0),
        "start_seconds": job.get("start_seconds"),
        "end_seconds": job.get("end_seconds"),
    }


def _update_video_job(job_id: str, **patch) -> None:
    """Update a queued video job in the registry."""
    with video_job_lock:
        job = video_job_registry.get(job_id)
        if not job:
            return
        job.update(patch)
        job["updated_at"] = time.time()


def _record_recent_video_job_locked(job_id: str) -> None:
    """Track a recently finished job and prune old history."""
    if job_id in video_job_recent:
        video_job_recent.remove(job_id)
    video_job_recent.insert(0, job_id)
    while len(video_job_recent) > MAX_VIDEO_JOB_HISTORY:
        stale_id = video_job_recent.pop()
        if stale_id == video_job_active_id or stale_id in video_job_pending:
            continue
        video_job_registry.pop(stale_id, None)


def _snapshot_video_jobs() -> dict:
    """Return a consistent snapshot of queued video job state."""
    with video_job_lock:
        active_job = video_job_registry.get(video_job_active_id) if video_job_active_id else None
        queued_jobs = [
            _serialize_video_job(video_job_registry[job_id])
            for job_id in list(video_job_pending)
            if job_id in video_job_registry
        ]
        recent_jobs = [
            _serialize_video_job(video_job_registry[job_id])
            for job_id in video_job_recent
            if job_id in video_job_registry
        ]
        completed_count = sum(1 for job in video_job_registry.values() if job.get("status") == "completed")
        failed_count = sum(1 for job in video_job_registry.values() if job.get("status") == "error")
        total_count = completed_count + failed_count + len(queued_jobs) + (1 if active_job else 0)
        return {
            "active_job": _serialize_video_job(active_job) if active_job else None,
            "queued_jobs": queued_jobs,
            "recent_jobs": recent_jobs,
            "summary": {
                "total": total_count,
                "completed": completed_count,
                "failed": failed_count,
                "queued": len(queued_jobs),
                "running": 1 if active_job else 0,
            },
        }


def _run_video_job(job_id: str) -> None:
    """Execute a queued video crop, clip, or conversion job."""
    job = video_job_registry.get(job_id)
    if not job:
        return

    cfg = _load_config()
    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)

    def on_progress(progress_value: float, message: str):
        _update_video_job(job_id, progress=max(0.0, min(1.0, progress_value)), message=message)

    if job.get("type") == "crop":
        result = _crop_video_file(
            job["video_path"],
            job["crop"],
            job.get("output_path"),
            ffmpeg_path,
            ffprobe_path,
            ffmpeg_threads,
            ffmpeg_hwaccel,
            on_progress=on_progress,
        )
        _update_video_job(
            job_id,
            status="completed",
            progress=1.0,
            message="Crop complete",
            output_path=result.get("output_path") or job.get("output_path") or "",
            result=result,
        )
        return

    if job.get("type") == "clip":
        result = _clip_video_file(
            job["video_path"],
            job["start_seconds"],
            job["end_seconds"],
            job.get("crop"),
            job.get("output_path"),
            ffmpeg_path,
            ffprobe_path,
            ffmpeg_threads,
            ffmpeg_hwaccel,
            on_progress=on_progress,
        )
        _update_video_job(
            job_id,
            status="completed",
            progress=1.0,
            message="Clip complete",
            output_path=result.get("output_path") or job.get("output_path") or "",
            result=result,
        )
        return

    if job.get("type") == "gif_to_mp4":
        result = _convert_gif_to_mp4_file(
            job["video_path"],
            job.get("output_path"),
            ffmpeg_path,
            ffprobe_path,
            ffmpeg_threads,
            ffmpeg_hwaccel,
            on_progress=on_progress,
        )
        source_caption = _get_caption_path(job["video_path"])
        output_caption = _get_caption_path(result.get("output_path") or job.get("output_path") or "")
        if source_caption.exists() and str(output_caption) and not output_caption.exists():
            shutil.copy2(source_caption, output_caption)
        _update_video_job(
            job_id,
            status="completed",
            progress=1.0,
            message="GIF conversion complete",
            output_path=result.get("output_path") or job.get("output_path") or "",
            result=result,
        )
        return

    raise RuntimeError("Unsupported video job type")


def _video_job_worker_loop() -> None:
    """Process queued video jobs sequentially in the background."""
    global video_job_active_id
    while True:
        video_job_wakeup.wait()
        while True:
            with video_job_lock:
                if not video_job_pending:
                    video_job_active_id = None
                    video_job_wakeup.clear()
                    break
                job_id = video_job_pending.popleft()
                job = video_job_registry.get(job_id)
                if not job:
                    continue
                video_job_active_id = job_id
                job["status"] = "running"
                job["progress"] = 0.0
                job["message"] = "Queued job started"
                job["error"] = ""
                job["updated_at"] = time.time()
            try:
                _run_video_job(job_id)
            except Exception as exc:
                _update_video_job(job_id, status="error", error=str(exc), message="Job failed")
            finally:
                with video_job_lock:
                    _record_recent_video_job_locked(job_id)
                    video_job_active_id = None


video_job_worker = threading.Thread(target=_video_job_worker_loop, name="tag2-video-jobs", daemon=True)
video_job_worker.start()


def _enqueue_video_job(job: dict) -> dict:
    """Queue a video job for background processing."""
    with video_job_lock:
        job["id"] = str(uuid4())
        now = time.time()
        job["status"] = "queued"
        job["progress"] = 0.0
        job["message"] = "Queued"
        job["error"] = ""
        job["created_at"] = now
        job["updated_at"] = now
        video_job_registry[job["id"]] = job
        video_job_pending.append(job["id"])
        video_job_wakeup.set()
        return _serialize_video_job(job)


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
    response = await call_next(request)
    path = str(request.url.path or "")
    if path == "/" or path.endswith(".html") or path.endswith(".js") or path.endswith(".css"):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


def _read_live_caption_state(
    image_path: str,
    all_captions: list[str],
    headers: list[str],
) -> tuple[list[str], str]:
    """Read the latest caption state from disk for merge-safe auto-caption writes."""
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")
    current = _read_caption_file(image_path, all_captions, headers)
    return _read_enabled_captions(current), current.get("free_text", "")


def _apply_caption_result_to_live_caption(
    image_path: str,
    caption: str,
    should_enable: bool,
    all_captions: list[str],
    headers: list[str],
    sections: list[dict],
) -> tuple[list[str], str]:
    """Apply a single AI caption verdict on top of the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_captions, headers)
    enabled = _apply_caption_selection(enabled, caption, sections, should_enable)
    enabled = _normalize_enabled_captions(enabled, sections)
    _write_caption_file(image_path, enabled, free_text, sections)
    return enabled, free_text


def _apply_group_result_to_live_caption(
    image_path: str,
    group_captions: list[str],
    selected_caption: str | None,
    all_captions: list[str],
    headers: list[str],
    sections: list[dict],
) -> tuple[list[str], str]:
    """Apply a single AI group selection on top of the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_captions, headers)
    enabled = [caption for caption in enabled if caption not in group_captions]
    if selected_caption:
        enabled = _apply_caption_selection(enabled, selected_caption, sections, True)
    enabled = _normalize_enabled_captions(enabled, sections)
    _write_caption_file(image_path, enabled, free_text, sections)
    return enabled, free_text


def _apply_free_text_result_to_live_caption(
    image_path: str,
    model_output: str,
    all_captions: list[str],
    headers: list[str],
    sections: list[dict],
    preserve_existing: bool = False,
) -> tuple[list[str], str, list[str]]:
    """Merge AI free-text output with the latest on-disk caption state."""
    enabled, free_text = _read_live_caption_state(image_path, all_captions, headers)
    base_free_text = free_text if preserve_existing else ""
    merged_free_text, added_lines = _merge_free_text(base_free_text, model_output, enabled)
    _write_caption_file(image_path, enabled, merged_free_text, sections)
    return enabled, merged_free_text, added_lines


def _build_ollama_answer_meta(response: dict) -> dict:
    """Normalize Ollama completion metadata for API responses."""
    return _ollama_response_meta(response)


def _build_skipped_caption_result(target: dict, enabled: list[str]) -> dict:
    """Return a consistent payload for a skipped caption check."""
    caption = str(target.get("caption") or target.get("sentence") or "").strip()
    return {
        "type": "sentence",
        "section_index": target.get("section_index"),
        "group_index": target.get("group_index"),
        "section_name": target.get("section_name", ""),
        "caption": caption,
        "sentence": caption,
        "enabled": caption in enabled,
        "skipped": True,
        "skip_reason": str(target.get("skip_reason") or ""),
        "answer": "",
    }


def _build_skipped_group_result(target: dict, enabled: list[str], sections: list[dict]) -> dict:
    """Return a consistent payload for a skipped exclusive-group check."""
    group_captions = [caption for caption in (target.get("captions") or target.get("sentences") or []) if caption]
    selected_caption = next((caption for caption in group_captions if caption in enabled), None)
    selected_hidden = bool(selected_caption and _is_hidden_group_caption(sections, selected_caption))
    selection_index = group_captions.index(selected_caption) + 1 if selected_caption in group_captions else None
    skipped_captions = list(target.get("skip_captions") or target.get("skip_sentences") or [])
    return {
        "type": "group",
        "section_index": target.get("section_index"),
        "group_index": target.get("group_index"),
        "section_name": target.get("section_name", ""),
        "group_name": target.get("group_name", ""),
        "captions": group_captions,
        "sentences": list(group_captions),
        "selected_caption": selected_caption,
        "selected_sentence": selected_caption,
        "selected_hidden": selected_hidden,
        "selection_index": selection_index,
        "skipped": True,
        "skip_reason": str(target.get("skip_reason") or ""),
        "skip_captions": skipped_captions,
        "skip_sentences": list(skipped_captions),
        "answer": "",
    }


@app.get("/api/list-images")
async def list_images(folder: str = Query(...)):
    """List all supported media files in the given folder."""
    folder_path = _resolve_folder_path(folder)
    cfg = _load_config()
    preview_output_folder = ""
    try:
        configured_output_folder = _get_comfyui_output_folder(cfg, DEFAULT_COMFYUI_OUTPUT_FOLDER)
        if configured_output_folder and os.path.isdir(configured_output_folder):
            preview_output_folder = configured_output_folder
    except RuntimeError:
        preview_output_folder = ""

    images = []
    try:
        for entry in sorted(folder_path.iterdir()):
            if entry.is_file() and entry.suffix.lower() in MEDIA_EXTENSIONS and not _is_mask_sidecar_path(str(entry)):
                prompt_preview_path = ""
                prompt_preview_mtime = 0.0
                if preview_output_folder and _is_image_path(str(entry)):
                    preview_files = _scan_comfyui_preview_files(preview_output_folder, str(entry))
                    prompt_preview_path = _select_latest_comfyui_preview_file(preview_files)
                    if prompt_preview_path:
                        prompt_preview_mtime = float(os.path.getmtime(prompt_preview_path))
                images.append(_build_media_entry(entry, prompt_preview_path, prompt_preview_mtime))
    except PermissionError:
        raise HTTPException(status_code=403, detail="Permission denied")

    return {"images": images, "folder": str(folder_path.resolve())}


@app.get("/api/folders/suggest")
async def suggest_folders(
    query: str = Query(default=""),
    limit: int = Query(default=DEFAULT_FOLDER_SUGGESTION_LIMIT, ge=1, le=50),
):
    """Return matching directory suggestions for the current folder input query."""
    return {
        "query": query,
        "suggestions": _suggest_folder_paths(query, limit),
    }


@app.post("/api/images/upload")
async def upload_images(
    folder: str = Form(...),
    files: Optional[list[UploadFile]] = File(default=None),
):
    """Upload supported media files into the currently loaded folder."""
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
            if suffix not in MEDIA_EXTENSIONS or original_name.lower().endswith(MASK_SIDECAR_SUFFIX):
                skipped.append({
                    "name": original_name,
                    "reason": "Unsupported media file",
                })
                continue

            destination = _get_unique_upload_path(folder_path, original_name)
            destination.parent.mkdir(parents=True, exist_ok=True)
            upload.file.seek(0)
            with destination.open("wb") as handle:
                shutil.copyfileobj(upload.file, handle)

            _clear_thumbnail_cache_for_path(str(destination))
            media_entry = _build_media_entry(destination)
            media_entry.update({
                "source_name": original_name,
                "renamed": destination.name != original_name,
            })
            uploaded.append(media_entry)
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
    """Return a thumbnail for the given media path."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    if _get_media_type_for_path(path) not in {"image", "video"}:
        raise HTTPException(status_code=400, detail="Unsupported media file")

    # Clamp size to nearest available
    actual_size = min(THUMBNAIL_SIZES, key=lambda s: abs(s - size))
    loop = asyncio.get_event_loop()
    cfg = _load_config()
    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    try:
        data = await loop.run_in_executor(executor, _get_thumbnail, path, actual_size, None, ffmpeg_path, ffprobe_path, ffmpeg_threads, ffmpeg_hwaccel)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    if not data:
        raise HTTPException(status_code=500, detail="Failed to generate thumbnail")

    return StreamingResponse(BytesIO(data), media_type="image/jpeg")


def _clear_media_stream_cancellation_if_idle(normalized_path: str) -> None:
    """Drop a delete-time media stream cancellation flag once no streams remain."""
    if media_stream_active_counts.get(normalized_path, 0) > 0:
        return
    media_stream_cancelled_paths.discard(normalized_path)


def _start_media_stream(path_value: str) -> str:
    """Track an active streamed media response for one path."""
    normalized_path = _normalize_image_key(path_value)
    with media_stream_lock:
        if normalized_path in media_stream_cancelled_paths:
            raise RuntimeError(f"Media stream cancelled for {Path(path_value).name}")
        media_stream_active_counts[normalized_path] = media_stream_active_counts.get(normalized_path, 0) + 1
    return normalized_path


def _finish_media_stream(normalized_path: str) -> None:
    """Release active streamed media bookkeeping for one path."""
    with media_stream_lock:
        remaining = media_stream_active_counts.get(normalized_path, 0) - 1
        if remaining > 0:
            media_stream_active_counts[normalized_path] = remaining
        else:
            media_stream_active_counts.pop(normalized_path, None)
        _clear_media_stream_cancellation_if_idle(normalized_path)


def _is_media_stream_cancelled(normalized_path: str) -> bool:
    """Return whether delete has cancelled streamed media for a path."""
    with media_stream_lock:
        return normalized_path in media_stream_cancelled_paths


def _suspend_media_stream_for_path(path_value: str, timeout_seconds: float = 1.5) -> None:
    """Block new media streams for a path and wait briefly for active ones to stop."""
    normalized_path = _normalize_image_key(path_value)
    deadline = time.monotonic() + max(0.0, float(timeout_seconds or 0.0))
    with media_stream_lock:
        media_stream_cancelled_paths.add(normalized_path)
    while True:
        with media_stream_lock:
            active_count = media_stream_active_counts.get(normalized_path, 0)
        if active_count <= 0:
            return
        if time.monotonic() >= deadline:
            return
        time.sleep(0.05)


def _resume_media_stream_for_path(path_value: str) -> None:
    """Allow future media streams again once delete-time draining is complete."""
    normalized_path = _normalize_image_key(path_value)
    with media_stream_lock:
        _clear_media_stream_cancellation_if_idle(normalized_path)


def _parse_media_range_header(range_header: str, file_size: int) -> tuple[int, int]:
    """Parse a single HTTP byte-range request."""
    match = MEDIA_RANGE_HEADER_PATTERN.fullmatch(str(range_header or "").strip())
    if not match:
        raise HTTPException(status_code=416, detail="Invalid byte range")

    start_text, end_text = match.groups()
    if not start_text and not end_text:
        raise HTTPException(status_code=416, detail="Invalid byte range")

    try:
        if start_text:
            start = max(0, int(start_text))
            end = int(end_text) if end_text else max(0, file_size - 1)
        else:
            suffix_length = int(end_text)
            if suffix_length <= 0:
                raise ValueError("Suffix length must be positive")
            end = max(0, file_size - 1)
            start = max(0, file_size - suffix_length)
    except ValueError as exc:
        raise HTTPException(status_code=416, detail="Invalid byte range") from exc

    if file_size <= 0:
        return 0, -1
    end = min(max(0, end), file_size - 1)
    if start >= file_size or end < start:
        raise HTTPException(status_code=416, detail="Invalid byte range")
    return start, end


async def _iter_media_file_chunks(request: Request, path_value: str, start: int, end: int):
    """Stream a media file in chunks while honoring delete-time cancellation."""
    normalized_path = _start_media_stream(path_value)
    remaining = max(0, end - start + 1)
    try:
        with open(path_value, "rb") as media_file:
            if start > 0:
                media_file.seek(start)
            while remaining > 0:
                if _is_media_stream_cancelled(normalized_path):
                    break
                if await request.is_disconnected():
                    break
                chunk = media_file.read(min(MEDIA_STREAM_CHUNK_SIZE, remaining))
                if not chunk:
                    break
                remaining -= len(chunk)
                yield chunk
                await asyncio.sleep(0)
    finally:
        _finish_media_stream(normalized_path)


@app.get("/api/media")
async def get_media(request: Request, path: str = Query(...)):
    """Serve the original media file with an inferred content type."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")
    if _get_media_type_for_path(path) not in {"image", "video"}:
        raise HTTPException(status_code=400, detail="Unsupported media file")

    media_type = mimetypes.guess_type(path)[0] or "application/octet-stream"
    if _is_video_path(path):
        file_size = os.path.getsize(path)
        range_header = request.headers.get("range")
        start = 0
        end = max(0, file_size - 1)
        status_code = 200
        headers = {
            "Accept-Ranges": "bytes",
            "Cache-Control": "no-store",
        }
        if range_header:
            start, end = _parse_media_range_header(range_header, file_size)
            status_code = 206
            headers["Content-Range"] = f"bytes {start}-{end}/{file_size}"
        headers["Content-Length"] = str(max(0, end - start + 1))
        return StreamingResponse(
            _iter_media_file_chunks(request, path, start, end),
            media_type=media_type,
            status_code=status_code,
            headers=headers,
        )
    return FileResponse(path, media_type=media_type)


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
    """Serve a preview-quality image for a supported media file."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="File not found")

    media_type = _get_media_type_for_path(path)
    if media_type == "other":
        raise HTTPException(status_code=400, detail="Unsupported media file")

    loop = asyncio.get_event_loop()
    cfg = _load_config()
    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    try:
        data = await loop.run_in_executor(executor, _get_thumbnail, path, PREVIEW_MAX_SIZE, None, ffmpeg_path, ffprobe_path, ffmpeg_threads, ffmpeg_hwaccel)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    if not data:
        if media_type == "video":
            raise HTTPException(status_code=500, detail="Failed to generate video preview")
        rendered, media_type = await loop.run_in_executor(executor, _render_image_bytes, path, None, PREVIEW_MAX_SIZE, True)
        return StreamingResponse(BytesIO(rendered), media_type=media_type)

    return StreamingResponse(
        BytesIO(data),
        media_type="image/jpeg",
        headers={"Cache-Control": "public, max-age=3600"},
    )


@app.get("/api/video/meta")
async def get_video_meta(path: str = Query(...)):
    """Return basic metadata for a video file."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Video file not found")
    if not _is_video_path(path):
        raise HTTPException(status_code=400, detail="Unsupported video file")

    cfg = _load_config()
    _, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    try:
        metadata = _probe_video_info(path, ffprobe_path, cancel_path=path)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    return {
        "path": path,
        "name": Path(path).name,
        "width": int(metadata.get("width") or 0),
        "height": int(metadata.get("height") or 0),
        "duration": float(metadata.get("duration") or 0.0),
        "fps": float(metadata.get("fps") or 0.0),
        "mask_keyframes": [
            int(keyframe.get("frame_index") or 0)
            for keyframe in (_get_video_mask_info(path).get("keyframes") or [])
            if int(keyframe.get("frame_index") or 0) >= 0
        ],
    }


@app.get("/api/video/frame")
async def get_video_frame(
    path: str = Query(...),
    time_seconds: float = Query(default=0.0),
    width: int = Query(default=160),
    height: int = Query(default=90),
):
    """Return a single JPEG frame from a video at a selected timestamp."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Video file not found")
    if not _is_video_path(path):
        raise HTTPException(status_code=400, detail="Unsupported video file")

    cfg = _load_config()
    ffmpeg_path, _ = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    loop = asyncio.get_event_loop()
    try:
        data = await loop.run_in_executor(
            executor,
            _extract_video_frame,
            path,
            max(0.0, float(time_seconds or 0.0)),
            max(32, min(640, int(width or 160))),
            max(18, min(360, int(height or 90))),
            ffmpeg_path,
            ffmpeg_threads,
            ffmpeg_hwaccel,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    return StreamingResponse(
        BytesIO(data),
        media_type="image/jpeg",
        headers={"Cache-Control": "public, max-age=3600"},
    )


class VideoCropJobRequest(BaseModel):
    video_path: str
    crop: dict


class VideoFrameExtractRequest(BaseModel):
    video_path: str
    time_seconds: float = 0.0
    frame_index: Optional[int] = None


@app.post("/api/video/extract-frame")
async def extract_video_frame_to_image(data: VideoFrameExtractRequest):
    """Extract the current video frame into a new JPG file beside the source video."""
    video_path = os.path.abspath(os.path.normpath(str(data.video_path or "").strip()))
    if not os.path.isfile(video_path):
        raise HTTPException(status_code=404, detail="Video file not found")
    if not _is_video_path(video_path):
        raise HTTPException(status_code=400, detail="Unsupported video file")

    cfg = _load_config()
    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    loop = asyncio.get_event_loop()
    try:
        video_info = await loop.run_in_executor(executor, _probe_video_info, video_path, ffprobe_path)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    image_width = int(video_info.get("width") or 0)
    image_height = int(video_info.get("height") or 0)
    if image_width <= 0 or image_height <= 0:
        raise HTTPException(status_code=500, detail="Video metadata does not include valid frame dimensions")

    duration_seconds = max(0.0, float(video_info.get("duration") or 0.0))
    time_seconds = max(0.0, float(data.time_seconds or 0.0))
    if duration_seconds > 0:
        time_seconds = min(time_seconds, duration_seconds)

    if data.frame_index is None:
        fps = max(0.0, float(video_info.get("fps") or 0.0))
        requested_frame_index = max(0, int((time_seconds * fps) + 1e-6)) if fps > 0 else 0
    else:
        requested_frame_index = max(0, int(data.frame_index or 0))

    output_path = _get_next_extracted_video_frame_path(video_path)
    created_paths: list[Path] = []
    try:
        frame_bytes = await loop.run_in_executor(
            executor,
            _extract_video_frame,
            video_path,
            time_seconds,
            image_width,
            image_height,
            ffmpeg_path,
            ffmpeg_threads,
            ffmpeg_hwaccel,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    if not frame_bytes:
        raise HTTPException(status_code=500, detail="ffmpeg returned an empty frame")

    try:
        output_path.write_bytes(frame_bytes)
        created_paths.append(output_path)
        sidecar_info = _copy_video_frame_sidecars(
            video_path,
            str(output_path),
            requested_frame_index,
            image_width,
            image_height,
        )
        if sidecar_info.get("caption_copied"):
            created_paths.append(_get_caption_path(str(output_path)))
        if sidecar_info.get("metadata_copied"):
            created_paths.append(_get_metadata_path(str(output_path)))
        if sidecar_info.get("mask_copied"):
            created_paths.append(_get_image_mask_path(str(output_path)))
    except OSError as exc:
        for created_path in reversed(created_paths):
            created_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"Failed to save extracted frame: {exc}") from exc

    _clear_thumbnail_cache_for_path(str(output_path))
    return {
        "ok": True,
        "video_path": video_path,
        "image_path": str(output_path),
        "name": output_path.name,
        "time_seconds": time_seconds,
        "requested_frame_index": requested_frame_index,
        **sidecar_info,
    }


class VideoClipJobRequest(BaseModel):
    video_path: str
    start_seconds: float
    end_seconds: float
    crop: Optional[dict] = None


class GifConvertJobRequest(BaseModel):
    media_path: str


@app.get("/api/video/jobs/status")
async def get_video_job_status():
    """Return the current queued video job state."""
    return _snapshot_video_jobs()


@app.post("/api/video/jobs/crop")
async def enqueue_video_crop_job(data: VideoCropJobRequest):
    """Queue a video crop job."""
    video_path = os.path.abspath(os.path.normpath(str(data.video_path or "").strip()))
    if not os.path.isfile(video_path):
        raise HTTPException(status_code=404, detail="Video file not found")
    if not _is_video_path(video_path):
        raise HTTPException(status_code=400, detail="Unsupported video file")
    if not isinstance(data.crop, dict):
        raise HTTPException(status_code=400, detail="Crop payload is required")

    queued = _enqueue_video_job({
        "type": "crop",
        "video_path": video_path,
        "folder": str(Path(video_path).parent),
        "crop": data.crop,
        "output_path": str(_get_unique_upload_path(
            Path(video_path).parent,
            Path(video_path).with_name(
                f"{Path(video_path).stem}__crop_{max(1, int(round(float(data.crop.get('w') or 1))))}x{max(1, int(round(float(data.crop.get('h') or 1))))}_{max(0, int(round(float(data.crop.get('x') or 0))))}-{max(0, int(round(float(data.crop.get('y') or 0))))}_{str(data.crop.get('ratio') or 'crop').replace(':', '-').replace('/', '-')}{Path(video_path).suffix or '.mp4'}"
            ).name,
        )),
    })
    return {"ok": True, "job": queued}


@app.post("/api/video/jobs/clip")
async def enqueue_video_clip_job(data: VideoClipJobRequest):
    """Queue a video clip job."""
    video_path = os.path.abspath(os.path.normpath(str(data.video_path or "").strip()))
    if not os.path.isfile(video_path):
        raise HTTPException(status_code=404, detail="Video file not found")
    if not _is_video_path(video_path):
        raise HTTPException(status_code=400, detail="Unsupported video file")

    try:
        start_seconds = max(0.0, float(data.start_seconds or 0.0))
        end_seconds = max(0.0, float(data.end_seconds or 0.0))
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="Invalid clip time range") from exc

    if end_seconds <= start_seconds:
        raise HTTPException(status_code=400, detail="Clip end time must be greater than start time")

    crop = data.crop if isinstance(data.crop, dict) else None
    folder_path = Path(video_path).parent
    clip_name = Path(_build_clip_output_path(video_path, start_seconds, end_seconds, crop)).name
    output_path = str(_get_unique_upload_path(folder_path, clip_name))
    queued = _enqueue_video_job({
        "type": "clip",
        "video_path": video_path,
        "folder": str(folder_path),
        "start_seconds": start_seconds,
        "end_seconds": end_seconds,
        "crop": crop,
        "output_path": output_path,
    })
    return {"ok": True, "job": queued}


@app.post("/api/media/jobs/convert-gif")
async def enqueue_gif_convert_job(data: GifConvertJobRequest):
    """Queue a GIF-to-MP4 conversion job in the current folder."""
    media_path = os.path.abspath(os.path.normpath(str(data.media_path or "").strip()))
    if not os.path.isfile(media_path):
        raise HTTPException(status_code=404, detail="GIF file not found")
    if Path(media_path).suffix.lower() != ".gif":
        raise HTTPException(status_code=400, detail="Only GIF files can be converted with this action")

    folder_path = Path(media_path).parent
    output_path = str(_get_unique_upload_path(folder_path, f"{Path(media_path).stem}.mp4"))
    queued = _enqueue_video_job({
        "type": "gif_to_mp4",
        "video_path": media_path,
        "folder": str(folder_path),
        "output_path": output_path,
    })
    return {"ok": True, "job": queued}


class BatchCaptionUpdate(BaseModel):
    image_paths: list[str]
    caption: str | None = None
    sentence: str | None = None
    enabled: bool

    def resolved_caption(self) -> str:
        return _read_caption_text(self.caption, self.sentence)


class BatchFreeTextUpdate(BaseModel):
    image_path: str
    free_text: str


class MediaMetadataFields(BaseModel):
    seed: int | None = None
    min_t: int | None = None
    max_t: int | None = None
    sampling_frequency: float | None = None
    caption_dropout_enabled: bool | None = None
    caption_dropout_caption: str | None = None


class SaveMediaMetadataRequest(BaseModel):
    path: str
    metadata: MediaMetadataFields = Field(default_factory=MediaMetadataFields)


class ApplyMediaMetadataRequest(BaseModel):
    paths: list[str]
    changes: MediaMetadataFields = Field(default_factory=MediaMetadataFields)


class BulkMediaMetadataRequest(BaseModel):
    paths: list[str]


class RenameCaptionPresetUpdate(BaseModel):
    folder: str
    old_caption: str | None = None
    new_caption: str | None = None
    old_sentence: str | None = None
    new_sentence: str | None = None

    def resolved_old_caption(self) -> str:
        return _read_caption_text(self.old_caption, self.old_sentence)

    def resolved_new_caption(self) -> str:
        return _read_caption_text(self.new_caption, self.new_sentence)


class RenameSectionUpdate(BaseModel):
    folder: str
    old_name: str
    new_name: str


class DeleteCaptionPresetUpdate(BaseModel):
    folder: str
    caption: str | None = None
    sentence: str | None = None

    def resolved_caption(self) -> str:
        return _read_caption_text(self.caption, self.sentence)


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


class FlipUpdate(BaseModel):
    image_path: str
    axis: str


class BatchDeleteImagesRequest(BaseModel):
    image_paths: list[str]


class DuplicateImageRequest(BaseModel):
    image_path: str
    new_name: str


class CloneFolderRequest(BaseModel):
    source_folder: str
    new_folder_name: str
    image_paths: list[str] = []


class MoveSelectedMediaRequest(BaseModel):
    source_folder: str
    target_folder: str
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


def _build_folder_suggestion_entry(path: Path) -> dict:
    resolved_path = path.resolve()
    normalized_path = os.path.normpath(str(resolved_path))
    parent_path = os.path.normpath(str(resolved_path.parent))
    return {
        "path": normalized_path,
        "name": resolved_path.name or normalized_path,
        "parent": "" if parent_path == normalized_path else parent_path,
    }


def _list_windows_drive_roots(prefix: str = "") -> list[Path]:
    normalized_prefix = str(prefix or "").strip().upper()
    roots: list[Path] = []
    for code in range(ord("A"), ord("Z") + 1):
        drive_letter = chr(code)
        if normalized_prefix and not drive_letter.startswith(normalized_prefix):
            continue
        drive_root = Path(f"{drive_letter}:\\")
        if drive_root.is_dir():
            roots.append(drive_root)
    return roots


def _suggest_folder_paths(query: str, limit: int = DEFAULT_FOLDER_SUGGESTION_LIMIT) -> list[dict]:
    expanded_query = os.path.expandvars(os.path.expanduser(str(query or "").strip()))
    if not expanded_query or limit <= 0:
        return []

    query_value = expanded_query.replace("/", "\\") if os.name == "nt" else expanded_query

    if os.name == "nt":
        if len(query_value) == 1 and query_value.isalpha():
            return [
                _build_folder_suggestion_entry(path)
                for path in _list_windows_drive_roots(query_value)[:limit]
            ]
        if len(query_value) == 2 and query_value[0].isalpha() and query_value[1] == ":":
            drive_root = Path(f"{query_value[0].upper()}:\\")
            return [_build_folder_suggestion_entry(drive_root)] if drive_root.is_dir() else []

    has_trailing_separator = query_value[-1] in {"/", "\\"}
    if has_trailing_separator:
        base_query = query_value
        prefix = ""
    else:
        base_query = os.path.dirname(query_value)
        prefix = os.path.basename(query_value)

    if os.name == "nt" and len(base_query) == 2 and base_query[0].isalpha() and base_query[1] == ":":
        base_query = f"{base_query}\\"

    try:
        base_path = Path(base_query or ".").resolve()
    except OSError:
        return []

    if not base_path.is_dir():
        return []

    prefix_casefold = prefix.casefold()
    suggestions: list[dict] = []
    try:
        for entry in sorted(base_path.iterdir(), key=lambda item: item.name.lower()):
            if not entry.is_dir():
                continue
            if prefix_casefold and not entry.name.casefold().startswith(prefix_casefold):
                continue
            suggestions.append(_build_folder_suggestion_entry(entry))
            if len(suggestions) >= limit:
                break
    except (OSError, PermissionError):
        return []

    return suggestions


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


def _get_next_extracted_video_frame_path(video_path: str, output_suffix: str = ".jpg") -> Path:
    """Return the next available extracted-frame image path beside a source video."""
    video_file = Path(video_path)
    folder_path = video_file.parent
    stem = video_file.stem or "video"
    normalized_suffix = str(output_suffix or ".jpg").strip() or ".jpg"
    if not normalized_suffix.startswith("."):
        normalized_suffix = f".{normalized_suffix}"

    pattern = re.compile(
        rf"^{re.escape(stem)}-frame_(?P<index>\d{{4}}){re.escape(normalized_suffix)}$",
        re.IGNORECASE,
    )
    highest_index = 0
    for candidate in folder_path.glob(f"{stem}-frame_*{normalized_suffix}"):
        if not candidate.is_file():
            continue
        match = pattern.match(candidate.name)
        if not match:
            continue
        try:
            highest_index = max(highest_index, int(match.group("index") or 0))
        except (TypeError, ValueError):
            continue

    next_index = highest_index + 1
    while True:
        candidate = folder_path / f"{stem}-frame_{next_index:04d}{normalized_suffix}"
        if not candidate.exists():
            return candidate
        next_index += 1


def _copy_video_frame_sidecars(
    video_path: str,
    image_path: str,
    requested_frame_index: int | None,
    image_width: int,
    image_height: int,
) -> dict:
    """Copy caption, metadata, and the resolved key-frame mask onto an extracted frame image."""
    source_caption_path = _get_caption_path(video_path)
    target_caption_path = _get_caption_path(image_path)
    caption_copied = False
    if source_caption_path.exists():
        shutil.copy2(source_caption_path, target_caption_path)
        caption_copied = True

    source_metadata_path = _get_metadata_path(video_path)
    target_metadata_path = _get_metadata_path(image_path)
    metadata_copied = False
    if source_metadata_path.exists():
        shutil.copy2(source_metadata_path, target_metadata_path)
        metadata_copied = True

    mask_copied = False
    mask_source_frame_index = None
    if requested_frame_index is not None:
        mask_info = _get_video_mask_frame_info(
            video_path,
            requested_frame_index,
            image_width,
            image_height,
            create_new=False,
        )
        source_mask_path = Path(str(mask_info.get("path") or ""))
        if mask_info.get("exists") and source_mask_path.is_file():
            shutil.copy2(source_mask_path, _get_image_mask_path(image_path))
            mask_copied = True
            mask_source_frame_index = int(mask_info.get("frame_index") or 0)

    return {
        "caption_copied": caption_copied,
        "metadata_copied": metadata_copied,
        "mask_copied": mask_copied,
        "mask_source_frame_index": mask_source_frame_index,
    }


def _normalize_duplicate_image_name(image_path: str, requested_name: str) -> str:
    normalized_name = _sanitize_upload_filename(requested_name)
    source_suffix = Path(image_path).suffix
    requested_suffix = Path(normalized_name).suffix
    if not requested_suffix:
        return f"{normalized_name}{source_suffix}"
    if requested_suffix.lower() != source_suffix.lower():
        raise HTTPException(status_code=400, detail=f"Duplicate name must keep the {source_suffix} extension")
    return normalized_name


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
        if candidate.suffix.lower() not in MEDIA_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"Unsupported media file: {candidate.name}")
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
            metadata_path = _get_metadata_path(str(image_path))
            if metadata_path.exists():
                copy_items.append((metadata_path, target_path / metadata_path.name))
            if _is_image_path(str(image_path)):
                mask_path = _get_image_mask_path(str(image_path))
                if mask_path.exists():
                    copy_items.append((mask_path, target_path / mask_path.name))
            elif _is_video_path(str(image_path)):
                for mask_path in _list_video_mask_paths(str(image_path)):
                    copy_items.append((mask_path, target_path / mask_path.name))
    else:
        for entry in sorted(source_path.iterdir(), key=lambda item: item.name.lower()):
            dest_entry = target_path / entry.name
            copy_items.append((entry, dest_entry))
            if entry.is_file() and entry.suffix.lower() in MEDIA_EXTENSIONS and not _is_mask_sidecar_path(str(entry)):
                copied_image_count += 1

    if copied_image_count == 0:
        raise HTTPException(status_code=400, detail="No media files available to clone")

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


def _normalize_section_merge_name(value: str | None) -> str:
    return str(value or "").strip().casefold()


def _section_names_match(left: str | None, right: str | None) -> bool:
    if _is_general_section_name(left) and _is_general_section_name(right):
        return True
    return _normalize_section_merge_name(left) == _normalize_section_merge_name(right)


def _make_empty_section(name: str | None = "") -> dict:
    return {
        "name": str(name or "").strip(),
        "captions": [],
        "sentences": [],
        "groups": [],
        "item_order": [],
    }


def _normalize_section_snapshot(sections: list[dict]) -> list[dict]:
    temp_folder = os.path.normpath("__move_merge__")
    return _get_folder_sections({"folders": {temp_folder: {"sections": sections}}}, temp_folder)


def _find_section_index_for_merge(sections: list[dict], section_name: str | None) -> int | None:
    for index, section in enumerate(sections):
        if _section_names_match(section.get("name", ""), section_name):
            return index
    return None


def _append_caption_to_section(section: dict, caption: str):
    normalized_caption = str(caption or "").strip()
    if not normalized_caption:
        return

    captions = list(section.get("captions") or section.get("sentences") or [])
    if normalized_caption not in captions:
        captions.append(normalized_caption)
    section["captions"] = captions
    section["sentences"] = list(captions)

    item_order = [item for item in (section.get("item_order") or []) if isinstance(item, dict)]
    exists_in_order = any(
        str(item.get("type") or "").strip().lower() in {"sentence", "caption"}
        and str(item.get("sentence") or item.get("caption") or "").strip() == normalized_caption
        for item in item_order
    )
    if not exists_in_order:
        item_order.append({"type": "sentence", "sentence": normalized_caption})
    section["item_order"] = item_order


def _remove_simple_caption_from_sections(sections: list[dict], caption_to_remove: str) -> bool:
    normalized_caption = str(caption_to_remove or "").strip()
    if not normalized_caption:
        return False

    removed = False
    for section in sections:
        section_captions = list(section.get("captions") or section.get("sentences") or [])
        if normalized_caption in section_captions:
            removed = True
        filtered_captions = [caption for caption in section_captions if caption != normalized_caption]
        section["captions"] = filtered_captions
        section["sentences"] = list(filtered_captions)
        section["item_order"] = [
            {
                "type": "sentence",
                "sentence": str(item.get("sentence") or item.get("caption") or "").strip(),
            }
            if str(item.get("type") or "").strip().lower() in {"sentence", "caption"}
            else {
                "type": "group",
                "group_id": str(item.get("group_id") or "").strip(),
            }
            for item in (section.get("item_order") or [])
            if isinstance(item, dict)
            and not (
                str(item.get("type") or "").strip().lower() in {"sentence", "caption"}
                and str(item.get("sentence") or item.get("caption") or "").strip() == normalized_caption
            )
        ]
    return removed


def _iter_group_locations(sections: list[dict]):
    for section_index, section in enumerate(sections):
        for group_index, group in enumerate(section.get("groups", []) or []):
            yield {
                "section_index": section_index,
                "group_index": group_index,
                "section_name": str(section.get("name") or "").strip(),
                "group_name": str(group.get("name") or "").strip(),
                "captions": _caption_values(group),
            }


def _collect_caption_locations(sections: list[dict]) -> dict[str, list[dict]]:
    locations: dict[str, list[dict]] = {}
    for section_index, section in enumerate(sections):
        for caption in _caption_values(section):
            locations.setdefault(caption, []).append({
                "kind": "simple",
                "section_index": section_index,
                "section_name": str(section.get("name") or "").strip(),
            })
        for group_location in _iter_group_locations([section]):
            actual_group_index = group_location["group_index"]
            captions = group_location["captions"]
            for caption in captions:
                locations.setdefault(caption, []).append({
                    "kind": "group",
                    "section_index": section_index,
                    "group_index": actual_group_index,
                    "section_name": str(section.get("name") or "").strip(),
                    "group_name": group_location["group_name"],
                })
    return locations


def _resolve_target_group_merge_location(
    target_sections: list[dict],
    source_section_name: str | None,
    source_group_name: str | None,
    source_group_captions: list[str],
) -> dict | None:
    source_caption_set = {caption for caption in source_group_captions if caption}
    group_name_key = _normalize_section_merge_name(source_group_name)
    all_groups = list(_iter_group_locations(target_sections))
    groups_by_key = {
        (item["section_index"], item["group_index"]): item
        for item in all_groups
    }

    overlap_keys: list[tuple[int, int]] = []
    exact_name_keys: list[tuple[int, int]] = []
    same_name_keys: list[tuple[int, int]] = []
    for item in all_groups:
        key = (item["section_index"], item["group_index"])
        if source_caption_set.intersection(item["captions"]):
            overlap_keys.append(key)
        if group_name_key and _normalize_section_merge_name(item["group_name"]) == group_name_key:
            same_name_keys.append(key)
            if _section_names_match(item["section_name"], source_section_name):
                exact_name_keys.append(key)

    overlap_keys = list(dict.fromkeys(overlap_keys))
    exact_name_keys = list(dict.fromkeys(exact_name_keys))
    same_name_keys = list(dict.fromkeys(same_name_keys))

    if len(overlap_keys) > 1:
        raise HTTPException(
            status_code=400,
            detail=f'Cannot merge source group "{source_group_name or "(unnamed group)"}" because its captions already belong to multiple groups in the target folder',
        )

    if overlap_keys:
        overlap_key = overlap_keys[0]
        if exact_name_keys and exact_name_keys[0] != overlap_key:
            raise HTTPException(
                status_code=400,
                detail=f'Cannot merge source group "{source_group_name or "(unnamed group)"}" because the target folder has conflicting group matches',
            )
        return groups_by_key[overlap_key]

    if len(exact_name_keys) > 1:
        raise HTTPException(
            status_code=400,
            detail=f'Cannot merge source group "{source_group_name or "(unnamed group)"}" because the target folder has multiple matching groups with the same name',
        )
    if exact_name_keys:
        return groups_by_key[exact_name_keys[0]]

    if len(same_name_keys) > 1:
        raise HTTPException(
            status_code=400,
            detail=f'Cannot merge source group "{source_group_name or "(unnamed group)"}" because the target folder has multiple groups with that name',
        )
    if same_name_keys:
        return groups_by_key[same_name_keys[0]]

    return None


def _merge_source_caption_into_target_sections(target_sections: list[dict], source_section_name: str | None, caption: str):
    normalized_caption = str(caption or "").strip()
    if not normalized_caption:
        return

    locations = _collect_caption_locations(target_sections).get(normalized_caption, [])
    group_keys = list(dict.fromkeys(
        (item["section_index"], item["group_index"])
        for item in locations
        if item.get("kind") == "group"
    ))
    simple_sections = list(dict.fromkeys(
        item["section_index"]
        for item in locations
        if item.get("kind") == "simple"
    ))

    if len(group_keys) > 1 or len(simple_sections) > 1:
        raise HTTPException(
            status_code=400,
            detail=f'Cannot move captions because "{normalized_caption}" already exists in multiple places in the target folder',
        )

    if group_keys:
        if simple_sections:
            _remove_simple_caption_from_sections(target_sections, normalized_caption)
        return
    if simple_sections:
        return

    section_index = _find_section_index_for_merge(target_sections, source_section_name)
    if section_index is None:
        target_sections.append(_make_empty_section(source_section_name))
        section_index = len(target_sections) - 1
    _append_caption_to_section(target_sections[section_index], normalized_caption)


def _merge_source_group_into_target_sections(target_sections: list[dict], source_section_name: str | None, source_group: dict):
    source_group_name = str(source_group.get("name") or "").strip()
    source_group_captions = _caption_values(source_group)
    if not source_group_captions:
        return

    target_location = _resolve_target_group_merge_location(
        target_sections,
        source_section_name,
        source_group_name,
        source_group_captions,
    )

    for caption in source_group_captions:
        _remove_simple_caption_from_sections(target_sections, caption)

    if target_location is None:
        section_index = _find_section_index_for_merge(target_sections, source_section_name)
        if section_index is None:
            target_sections.append(_make_empty_section(source_section_name))
            section_index = len(target_sections) - 1
        target_sections[section_index].setdefault("groups", []).append({
            "id": str(source_group.get("id") or "") or f"move-group-{uuid4().hex[:8]}",
            "name": source_group_name,
            "captions": list(source_group_captions),
            "sentences": list(source_group_captions),
            "hidden_captions": [caption for caption in _group_hidden_captions(source_group) if caption in source_group_captions],
            "hidden_sentences": [caption for caption in _group_hidden_captions(source_group) if caption in source_group_captions],
        })
        return

    group = target_sections[target_location["section_index"]].get("groups", [])[target_location["group_index"]]
    existing_captions = _caption_values(group)
    existing_caption_set = set(existing_captions)
    newly_added: list[str] = []
    for caption in source_group_captions:
        if caption in existing_caption_set:
            continue
        existing_captions.append(caption)
        existing_caption_set.add(caption)
        newly_added.append(caption)
    group["captions"] = list(existing_captions)
    group["sentences"] = list(existing_captions)
    if not str(group.get("name") or "").strip() and source_group_name:
        group["name"] = source_group_name

    existing_hidden = list(group.get("hidden_captions") or group.get("hidden_sentences") or [])
    for caption in _group_hidden_captions(source_group):
        if caption in newly_added and caption not in existing_hidden and caption in existing_caption_set:
            existing_hidden.append(caption)
    group["hidden_captions"] = [caption for caption in existing_hidden if caption in existing_caption_set]
    group["hidden_sentences"] = list(group["hidden_captions"])


def _merge_move_captions_into_target_sections(
    source_sections: list[dict],
    target_sections: list[dict],
    used_captions: set[str],
) -> list[dict]:
    updated_sections = copy.deepcopy(target_sections)
    processed_captions: set[str] = set()
    processed_groups: set[tuple[int | None, int | None]] = set()

    for target in _iter_caption_targets_with_indices(source_sections):
        if target["type"] == "group":
            section_index = target.get("section_index")
            group_index = target.get("group_index")
            group_key = (section_index, group_index)
            group_captions = [caption for caption in (target.get("captions") or []) if caption]
            if not group_captions or not any(caption in used_captions for caption in group_captions):
                continue
            if group_key in processed_groups or section_index is None or group_index is None:
                continue
            source_group = source_sections[section_index].get("groups", [])[group_index]
            _merge_source_group_into_target_sections(updated_sections, target.get("section_name", ""), source_group)
            processed_groups.add(group_key)
            processed_captions.update(group_captions)
            continue

        caption = str(target.get("caption") or target.get("sentence") or "").strip()
        if not caption or caption not in used_captions or caption in processed_captions:
            continue
        _merge_source_caption_into_target_sections(updated_sections, target.get("section_name", ""), caption)
        processed_captions.add(caption)

    return _normalize_section_snapshot(updated_sections)


def _rewrite_caption_files_for_section_change(
    folder: str,
    sections_before: list[dict],
    sections_after: list[dict],
    *,
    exclude_media_paths: set[str] | None = None,
) -> list[tuple[Path, str | None]]:
    excluded = {
        os.path.normcase(os.path.abspath(os.path.normpath(path)))
        for path in (exclude_media_paths or set())
        if path
    }
    backups: list[tuple[Path, str | None]] = []
    all_captions_before = _all_captions_from_sections(sections_before)
    headers_before = _all_headers_from_sections(sections_before)

    for entry in _iter_folder_image_entries(folder):
        normalized_entry = os.path.normcase(os.path.abspath(os.path.normpath(str(entry))))
        if normalized_entry in excluded:
            continue
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        backups.append((caption_path, caption_path.read_text(encoding="utf-8")))
        data = _read_caption_file(str(entry), all_captions_before, headers_before)
        enabled = _normalize_enabled_captions(_read_enabled_captions(data), sections_after)
        free_text = str(data.get("free_text", "") or "")
        _write_caption_file(str(entry), enabled, free_text, sections_after)

    return backups


def _restore_caption_file_backups(backups: list[tuple[Path, str | None]]):
    for path_obj, content in reversed(backups):
        try:
            if content is None:
                path_obj.unlink(missing_ok=True)
            else:
                path_obj.parent.mkdir(parents=True, exist_ok=True)
                path_obj.write_text(content, encoding="utf-8")
        except OSError:
            continue


def _build_move_artifacts_for_media(source_path: Path, target_path: Path, image_crops: dict | None) -> tuple[list[dict], dict | None]:
    artifacts: list[dict] = [{
        "source": source_path,
        "target": target_path,
        "kind": "media",
        "allow_replace": False,
    }]

    caption_path = _get_caption_path(str(source_path))
    if caption_path.exists():
        artifacts.append({
            "source": caption_path,
            "target": target_path.with_suffix(".txt"),
            "kind": "caption",
            "allow_replace": False,
        })

    metadata_path = _get_metadata_path(str(source_path))
    if metadata_path.exists():
        artifacts.append({
            "source": metadata_path,
            "target": _get_metadata_path(str(target_path)),
            "kind": "metadata",
            "allow_replace": False,
        })

    crop_state = None
    if _is_image_path(str(source_path)):
        mask_path = _get_image_mask_path(str(source_path))
        if mask_path.exists():
            artifacts.append({
                "source": mask_path,
                "target": _get_image_mask_path(str(target_path)),
                "kind": "mask",
                "allow_replace": False,
            })

        crop_backup_path = Path(_get_crop_backup_path(str(source_path)))
        if crop_backup_path.exists():
            artifacts.append({
                "source": crop_backup_path,
                "target": Path(_get_crop_backup_path(str(target_path))),
                "kind": "crop-backup",
                "allow_replace": True,
            })

        crop_mask_backup_path = Path(_get_crop_mask_backup_path(str(source_path)))
        if crop_mask_backup_path.exists():
            artifacts.append({
                "source": crop_mask_backup_path,
                "target": Path(_get_crop_mask_backup_path(str(target_path))),
                "kind": "crop-mask-backup",
                "allow_replace": True,
            })

        if isinstance(image_crops, dict):
            crop_state = copy.deepcopy(image_crops.get(_normalize_image_key(str(source_path))))
    elif _is_video_path(str(source_path)):
        for mask_path in _list_video_mask_paths(str(source_path)):
            artifacts.append({
                "source": mask_path,
                "target": target_path.parent / mask_path.name.replace(source_path.name, target_path.name, 1),
                "kind": "video-mask",
                "allow_replace": False,
            })

    return artifacts, crop_state


def _prepare_move_plan(source_folder: str, target_folder: str, image_paths: list[str], cfg: dict) -> dict:
    source_path = _resolve_folder_path(source_folder, detail="Source folder is not a valid directory")
    target_path = _resolve_folder_path(target_folder, detail="Target folder is not a valid directory")
    if os.path.normcase(str(source_path)) == os.path.normcase(str(target_path)):
        raise HTTPException(status_code=400, detail="Target folder must be different from the current folder")

    normalized_source = os.path.normcase(str(source_path))
    normalized_selected: list[Path] = []
    for raw_path in image_paths or []:
        candidate = Path(str(raw_path or "")).resolve()
        if not candidate.is_file():
            raise HTTPException(status_code=400, detail=f"Image not found: {candidate}")
        if os.path.normcase(str(candidate.parent)) != normalized_source:
            raise HTTPException(status_code=400, detail="Selected images must belong to the current folder")
        if candidate.suffix.lower() not in MEDIA_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"Unsupported media file: {candidate.name}")
        if candidate not in normalized_selected:
            normalized_selected.append(candidate)

    if not normalized_selected:
        raise HTTPException(status_code=400, detail="Select at least one media file to copy")

    image_crops = cfg.get("image_crops") if isinstance(cfg.get("image_crops"), dict) else {}
    source_sections = _get_folder_sections(cfg, str(source_path))
    source_all_captions = _all_captions_from_sections(source_sections)
    source_headers = _all_headers_from_sections(source_sections)

    planned_target_paths: set[str] = set()
    conflicts: list[str] = []
    media_entries: list[dict] = []
    used_captions: set[str] = set()

    for media_path in normalized_selected:
        destination_path = target_path / media_path.name
        artifacts, crop_state = _build_move_artifacts_for_media(media_path, destination_path, image_crops)
        for artifact in artifacts:
            target_artifact = artifact["target"]
            target_key = os.path.normcase(os.path.abspath(os.path.normpath(str(target_artifact))))
            if not artifact.get("allow_replace"):
                if target_key in planned_target_paths:
                    conflicts.append(f"Multiple selected files would overwrite {target_artifact.name} in the target folder")
                elif target_artifact.exists():
                    conflicts.append(f"Target already contains {target_artifact.name}")
                planned_target_paths.add(target_key)

        caption_path = _get_caption_path(str(media_path))
        if caption_path.exists():
            original_caption_text = caption_path.read_text(encoding="utf-8")
            caption_data = _read_caption_file(str(media_path), source_all_captions, source_headers)
        else:
            original_caption_text = None
            caption_data = {"enabled_captions": [], "free_text": ""}

        enabled_captions = _read_enabled_captions(caption_data)
        used_captions.update(enabled_captions)
        media_entries.append({
            "source_path": media_path,
            "target_path": destination_path,
            "artifacts": artifacts,
            "crop_state": crop_state,
            "caption_snapshot": {
                "had_caption_file": caption_path.exists(),
                "enabled_captions": enabled_captions,
                "free_text": str(caption_data.get("free_text", "") or ""),
                "original_text": original_caption_text,
            },
        })

    if conflicts:
        conflict_preview = "\n".join(f"- {message}" for message in conflicts[:8])
        if len(conflicts) > 8:
            conflict_preview += f"\n- ...and {len(conflicts) - 8} more conflict(s)"
        raise HTTPException(status_code=400, detail=f"Copy blocked by existing target files:\n{conflict_preview}")

    target_sections_before = _get_folder_sections(cfg, str(target_path))
    target_sections_after = _merge_move_captions_into_target_sections(source_sections, target_sections_before, used_captions)

    return {
        "source_path": source_path,
        "target_path": target_path,
        "source_sections": source_sections,
        "target_sections_before": target_sections_before,
        "target_sections_after": target_sections_after,
        "sections_changed": target_sections_after != target_sections_before,
        "media_entries": media_entries,
        "used_captions": used_captions,
    }


def _iter_folder_image_entries(folder: str):
    """Yield supported media files in a folder."""
    folder_path = Path(folder)
    for entry in folder_path.iterdir():
        if entry.is_file() and entry.suffix.lower() in MEDIA_EXTENSIONS and not _is_mask_sidecar_path(str(entry)):
            yield entry


def _remove_deleted_caption_lines(free_text: str, removed_captions: set[str]) -> str:
    """Drop free-text lines that exactly match deleted captions."""
    if not removed_captions:
        return str(free_text or "")
    filtered_lines = [
        line
        for line in str(free_text or "").splitlines()
        if line.strip() not in removed_captions
    ]
    return "\n".join(filtered_lines)


def _remove_captions_from_caption_files(folder: str, sections_before: list[dict], sections_after: list[dict], removed_captions: list[str]):
    """Rewrite caption files in a folder after configured captions are deleted."""
    removed_set = {caption for caption in removed_captions if caption}
    if not removed_set:
        return 0

    all_captions_before = _all_captions_from_sections(sections_before)
    headers_before = _all_headers_from_sections(sections_before)
    touched_caption_files = 0

    for entry in _iter_folder_image_entries(folder):
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_captions_before, headers_before)
        enabled = [
            caption for caption in _read_enabled_captions(data)
            if caption not in removed_set
        ]
        enabled = _normalize_enabled_captions(enabled, sections_after)
        free_text = _remove_deleted_caption_lines(str(data.get("free_text", "") or ""), removed_set)
        _write_caption_file(str(entry), enabled, free_text, sections_after)
        touched_caption_files += 1

    return touched_caption_files


@app.get("/api/crop")
async def get_crop(path: str = Query(...)):
    """Get whether an image currently has a reversible real crop applied."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Image file not found")
    if not _is_image_path(path):
        raise HTTPException(status_code=400, detail="Crop is only available for image files")
    return {
        "path": path,
        "crop": _get_image_crop(path),
    }


@app.get("/api/mask")
async def get_mask(
    path: str = Query(...),
    ensure: bool = Query(default=False),
    frame_index: Optional[int] = Query(default=None),
    create_new: bool = Query(default=False),
):
    """Return mask sidecar metadata for an image or video key frame."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Media file not found")

    loop = asyncio.get_event_loop()
    if _is_image_path(path):
        info = await loop.run_in_executor(executor, _write_default_image_mask if ensure else _get_image_mask_info, path)
        if not ensure:
            image_width, image_height = await loop.run_in_executor(executor, _get_display_image_size, path)
            info["created"] = False
            info["image_width"] = image_width
            info["image_height"] = image_height
        info["media_type"] = "image"
        return info

    if not _is_video_path(path):
        raise HTTPException(status_code=400, detail="Mask editing is only available for image and video files")
    if frame_index is None:
        raise HTTPException(status_code=400, detail="Frame index is required for video mask editing")

    cfg = _load_config()
    _, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    video_info = await loop.run_in_executor(executor, _probe_video_info, path, ffprobe_path)
    image_width = int(video_info.get("width") or 0)
    image_height = int(video_info.get("height") or 0)
    if image_width <= 0 or image_height <= 0:
        raise HTTPException(status_code=503, detail="Could not determine video dimensions")

    info = await loop.run_in_executor(
        executor,
        _write_default_video_mask if ensure else _get_video_mask_frame_info,
        path,
        int(frame_index),
        image_width,
        image_height,
        bool(create_new),
    )
    info["media_type"] = "video"
    info["video_width"] = image_width
    info["video_height"] = image_height
    info["video_fps"] = float(video_info.get("fps") or 0.0)
    return info


@app.get("/api/mask/image")
async def get_mask_image(
    path: str = Query(...),
    ensure: bool = Query(default=False),
    frame_index: Optional[int] = Query(default=None),
    create_new: bool = Query(default=False),
):
    """Serve the grayscale mask sidecar for an image or video key frame."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Media file not found")

    loop = asyncio.get_event_loop()
    if _is_image_path(path):
        if ensure:
            await loop.run_in_executor(executor, _write_default_image_mask, path)

        mask_path = _get_image_mask_path(path)
        if not mask_path.is_file():
            raise HTTPException(status_code=404, detail="Mask file not found")
        return FileResponse(mask_path, media_type="image/png")

    if not _is_video_path(path):
        raise HTTPException(status_code=400, detail="Mask editing is only available for image and video files")
    if frame_index is None:
        raise HTTPException(status_code=400, detail="Frame index is required for video mask editing")

    cfg = _load_config()
    _, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    video_info = await loop.run_in_executor(executor, _probe_video_info, path, ffprobe_path)
    image_width = int(video_info.get("width") or 0)
    image_height = int(video_info.get("height") or 0)
    if image_width <= 0 or image_height <= 0:
        raise HTTPException(status_code=503, detail="Could not determine video dimensions")

    info = await loop.run_in_executor(
        executor,
        _write_default_video_mask if ensure else _get_video_mask_frame_info,
        path,
        int(frame_index),
        image_width,
        image_height,
        bool(create_new),
    )
    mask_path = Path(str(info.get("path") or ""))
    if not mask_path.is_file():
        raise HTTPException(status_code=404, detail="Mask file not found")

    return FileResponse(mask_path, media_type="image/png")


@app.post("/api/mask")
async def save_mask(
    image_path: Optional[str] = Form(default=None),
    media_path: Optional[str] = Form(default=None),
    frame_index: Optional[int] = Form(default=None),
    mask: UploadFile = File(...),
):
    """Save a grayscale PNG mask sidecar for an image or video key frame."""
    target_path = str(media_path or image_path or "").strip()
    if not target_path or not os.path.isfile(target_path):
        raise HTTPException(status_code=404, detail="Media file not found")

    try:
        mask_bytes = await mask.read()
        if not mask_bytes:
            raise HTTPException(status_code=400, detail="Mask upload is empty")
        loop = asyncio.get_event_loop()
        if _is_image_path(target_path):
            info = await loop.run_in_executor(executor, _save_image_mask, target_path, mask_bytes)
            info["media_type"] = "image"
            return {"ok": True, **info}

        if not _is_video_path(target_path):
            raise HTTPException(status_code=400, detail="Mask editing is only available for image and video files")
        if frame_index is None:
            raise HTTPException(status_code=400, detail="Frame index is required for video mask editing")

        cfg = _load_config()
        _, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
        video_info = await loop.run_in_executor(executor, _probe_video_info, target_path, ffprobe_path)
        image_width = int(video_info.get("width") or 0)
        image_height = int(video_info.get("height") or 0)
        if image_width <= 0 or image_height <= 0:
            raise HTTPException(status_code=503, detail="Could not determine video dimensions")
        info = await loop.run_in_executor(executor, _save_video_mask, target_path, int(frame_index), mask_bytes, image_width, image_height)
        info["media_type"] = "video"
        return {"ok": True, **info}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid mask image: {exc}") from exc
    finally:
        await mask.close()


@app.post("/api/image/edit")
async def save_image_edit(
    image_path: str = Form(...),
    image: UploadFile = File(...),
):
    """Save an edited image back onto the source file while preserving metadata."""
    if not image_path or not os.path.isfile(image_path):
        raise HTTPException(status_code=404, detail="Image file not found")
    if not _is_image_path(image_path):
        raise HTTPException(status_code=400, detail="Image editing is only available for image files")

    try:
        image_bytes = await image.read()
        if not image_bytes:
            raise HTTPException(status_code=400, detail="Image upload is empty")
        loop = asyncio.get_event_loop()
        info = await loop.run_in_executor(executor, _save_edited_image, image_path, image_bytes)
        return {"ok": True, **info}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid edited image: {exc}") from exc
    finally:
        await image.close()


@app.post("/api/crop")
async def save_crop(data: CropUpdate):
    """Apply or remove a real crop on the image file."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image file not found")
    if not _is_image_path(data.image_path):
        raise HTTPException(status_code=400, detail="Crop is only available for image files")
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
    if not _is_image_path(data.image_path):
        raise HTTPException(status_code=400, detail="Rotate is only available for image files")
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


@app.post("/api/flip")
async def flip_image(data: FlipUpdate):
    """Flip an image horizontally or vertically."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Image file not found")
    if not _is_image_path(data.image_path):
        raise HTTPException(status_code=400, detail="Flip is only available for image files")
    try:
        crop_state = _flip_image(data.image_path, data.axis)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid flip: {e}") from e
    return {
        "ok": True,
        "crop": crop_state,
    }


@app.post("/api/images/delete")
async def delete_images(data: BatchDeleteImagesRequest):
    """Delete supported media files and their local sidecar artifacts."""
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

        if path_obj.suffix.lower() not in MEDIA_EXTENSIONS:
            errors.append({"path": normalized_path, "error": "Unsupported media file"})
            continue
        if not path_obj.is_file():
            errors.append({"path": normalized_path, "error": "File not found"})
            continue

        caption_path = _get_caption_path(normalized_path)
        metadata_path = _get_metadata_path(normalized_path)
        mask_path = _get_image_mask_path(normalized_path) if _is_image_path(normalized_path) else None
        video_mask_paths = _list_video_mask_paths(normalized_path) if _is_video_path(normalized_path) else []
        backup_path = _get_crop_backup_path(normalized_path) if _is_image_path(normalized_path) else ""

        try:
            _suspend_media_stream_for_path(normalized_path)
            _suspend_thumbnail_processing_for_path(normalized_path)
            last_permission_error: PermissionError | None = None
            for attempt in range(6):
                try:
                    path_obj.unlink()
                    last_permission_error = None
                    break
                except PermissionError as exc:
                    last_permission_error = exc
                    if attempt >= 5:
                        raise
                    time.sleep(0.15)
            if last_permission_error is not None:
                raise last_permission_error
            if caption_path.exists():
                caption_path.unlink()
            if metadata_path.exists():
                metadata_path.unlink()
            if mask_path and mask_path.exists():
                mask_path.unlink()
            for video_mask_path in video_mask_paths:
                if video_mask_path.exists():
                    video_mask_path.unlink()
            if backup_path and os.path.isfile(backup_path):
                os.remove(backup_path)
            _clear_thumbnail_cache_for_path(normalized_path)
            if image_crops is not None and image_crops.pop(_normalize_image_key(normalized_path), None) is not None:
                config_changed = True
            deleted_paths.append(normalized_path)
        except PermissionError:
            errors.append({"path": normalized_path, "error": "Permission denied"})
        except OSError as exc:
            errors.append({"path": normalized_path, "error": str(exc)})
        finally:
            _resume_thumbnail_processing_for_path(normalized_path)
            _resume_media_stream_for_path(normalized_path)

    if config_changed:
        _save_config(cfg)

    return {
        "ok": len(errors) == 0,
        "deleted_paths": deleted_paths,
        "deleted_count": len(deleted_paths),
        "errors": errors,
    }


@app.post("/api/image/duplicate")
async def duplicate_image(data: DuplicateImageRequest):
    """Duplicate a single image file together with its caption and mask sidecars."""
    source_path = Path(str(data.image_path or "")).resolve()
    if not source_path.is_file():
        raise HTTPException(status_code=404, detail="Image file not found")
    if not _is_image_path(str(source_path)):
        raise HTTPException(status_code=400, detail="Image duplication is only available for image files")

    target_name = _normalize_duplicate_image_name(str(source_path), data.new_name)
    target_path = source_path.parent / target_name
    if target_path.exists():
        raise HTTPException(status_code=400, detail="Target image already exists")

    caption_source = _get_caption_path(str(source_path))
    caption_target = target_path.with_suffix(".txt")
    metadata_source = _get_metadata_path(str(source_path))
    metadata_target = _get_metadata_path(str(target_path))
    mask_source = _get_image_mask_path(str(source_path))
    mask_target = _get_image_mask_path(str(target_path))

    try:
        shutil.copy2(source_path, target_path)
        if caption_source.exists():
            shutil.copy2(caption_source, caption_target)
        if metadata_source.exists():
            shutil.copy2(metadata_source, metadata_target)
        if mask_source.exists():
            shutil.copy2(mask_source, mask_target)
    except Exception as exc:
        target_path.unlink(missing_ok=True)
        caption_target.unlink(missing_ok=True)
        metadata_target.unlink(missing_ok=True)
        mask_target.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=f"Could not duplicate image: {exc}") from exc

    return {
        "ok": True,
        "image_path": str(target_path),
        "caption_path": str(caption_target) if caption_target.exists() else "",
        "mask_path": str(mask_target) if mask_target.exists() else "",
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


@app.post("/api/media/copy/stream")
@app.post("/api/media/move/stream")
async def move_selected_media_stream(data: MoveSelectedMediaRequest):
    """Copy selected media into another folder while merging caption presets and sidecars."""
    cfg = _load_config()
    cfg_before = copy.deepcopy(cfg)
    plan = _prepare_move_plan(data.source_folder, data.target_folder, data.image_paths or [], cfg)
    source_path: Path = plan["source_path"]
    target_path: Path = plan["target_path"]
    source_sections: list[dict] = plan["source_sections"]
    target_sections_before: list[dict] = plan["target_sections_before"]
    target_sections_after: list[dict] = plan["target_sections_after"]
    sections_changed = bool(plan["sections_changed"])
    media_entries: list[dict] = plan["media_entries"]

    async def event_stream():
        copied_artifacts: list[Path] = []
        caption_backups: list[tuple[Path, str | None]] = []
        try:
            yield _event_bytes({
                "type": "start",
                "mode": "selected",
                "source_folder": str(source_path),
                "target_folder": str(target_path),
                "total": len(media_entries),
                "sections_changed": sections_changed,
            })

            for index, media_entry in enumerate(media_entries, start=1):
                source_media_path: Path = media_entry["source_path"]
                target_media_path: Path = media_entry["target_path"]
                for artifact in media_entry["artifacts"]:
                    source_artifact: Path = artifact["source"]
                    target_artifact: Path = artifact["target"]
                    if artifact.get("allow_replace") and target_artifact.exists():
                        target_artifact.unlink(missing_ok=True)
                    target_artifact.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source_artifact, target_artifact)
                    copied_artifacts.append(target_artifact)

                _clear_thumbnail_cache_for_path(str(source_media_path))
                _clear_thumbnail_cache_for_path(str(target_media_path))
                yield _event_bytes({
                    "type": "progress",
                    "index": index,
                    "total": len(media_entries),
                    "source_path": str(source_media_path),
                    "target_path": str(target_media_path),
                    "name": target_media_path.name,
                })

            folders_cfg = cfg.setdefault("folders", {})
            image_crops = cfg.setdefault("image_crops", {}) if isinstance(cfg.get("image_crops"), dict) else {}
            if sections_changed or os.path.normpath(str(target_path)) in folders_cfg:
                _set_folder_sections(cfg, str(target_path), target_sections_after)

            moved_target_paths = {str(item["target_path"]) for item in media_entries}
            if sections_changed:
                caption_backups.extend(
                    _rewrite_caption_files_for_section_change(
                        str(target_path),
                        target_sections_before,
                        target_sections_after,
                        exclude_media_paths=moved_target_paths,
                    )
                )

            for media_entry in media_entries:
                source_media_path: Path = media_entry["source_path"]
                target_media_path: Path = media_entry["target_path"]
                if media_entry.get("crop_state") is not None:
                    image_crops[_normalize_image_key(str(target_media_path))] = media_entry["crop_state"]

                caption_snapshot = media_entry["caption_snapshot"]
                if not caption_snapshot.get("had_caption_file"):
                    continue
                target_caption_path = _get_caption_path(str(target_media_path))
                caption_backups.append((target_caption_path, caption_snapshot.get("original_text")))
                _write_caption_file(
                    str(target_media_path),
                    list(caption_snapshot.get("enabled_captions") or []),
                    str(caption_snapshot.get("free_text") or ""),
                    target_sections_after,
                )

            _save_config(cfg)
            yield _event_bytes({
                "type": "config-updated",
                "target_folder": str(target_path),
                "sections_changed": sections_changed,
            })
            yield _event_bytes({
                "type": "done",
                "source_folder": str(source_path),
                "target_folder": str(target_path),
                "total": len(media_entries),
                "copied": len(media_entries),
                "moved": len(media_entries),
            })
        except HTTPException:
            raise
        except Exception as exc:
            _restore_caption_file_backups(caption_backups)
            for target_artifact in reversed(copied_artifacts):
                try:
                    if not target_artifact.exists():
                        continue
                    target_artifact.unlink()
                except OSError:
                    continue
            _save_config(cfg_before)
            yield _event_bytes({"type": "error", "message": str(exc)})

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


@app.get("/api/caption")
async def get_caption(
    path: str = Query(...),
    captions: str | None = Query(default=None),
    sentences: str | None = Query(default=None),
):
    """Read caption data for an image. captions is a JSON array of predefined captions."""
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Image file not found")
    try:
        predefined = json.loads(_coalesce_caption_query(captions, sentences))
    except json.JSONDecodeError:
        predefined = []

    # Load section headers for proper parsing
    cfg = _load_config()
    folder = str(Path(path).parent)
    sections = _get_folder_sections(cfg, folder)
    headers = _all_headers_from_sections(sections)

    data = _read_caption_file(path, predefined, headers)
    return {"enabled_captions": _read_enabled_captions(data), **data}


@app.get("/api/media/meta")
async def get_media_metadata(path: str = Query(...)):
    """Read metadata for a supported media file."""
    normalized_path = _normalize_existing_media_path(path)
    return _read_metadata_file(normalized_path)


@app.post("/api/media/meta/bulk")
async def post_media_metadata_bulk(payload: BulkMediaMetadataRequest):
    """Read metadata for multiple supported media files."""
    results = {}
    for raw_path in payload.paths or []:
        normalized_path = os.path.abspath(os.path.normpath(str(raw_path or "").strip()))
        if os.path.isfile(normalized_path) and Path(normalized_path).suffix.lower() in MEDIA_EXTENSIONS:
            results[normalized_path] = _read_metadata_file(normalized_path)
        else:
            results[normalized_path] = {}
    return results


@app.post("/api/media/meta/save")
async def save_media_metadata(payload: SaveMediaMetadataRequest):
    """Replace metadata for one media file."""
    normalized_path = _normalize_existing_media_path(payload.path)
    try:
        saved = _write_metadata_file(normalized_path, _dump_model(payload.metadata, exclude_none=False))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"ok": True, "metadata": saved}


@app.post("/api/media/meta/apply")
async def apply_media_metadata(payload: ApplyMediaMetadataRequest):
    """Apply sparse metadata changes to multiple media files."""
    if not payload.paths:
        raise HTTPException(status_code=400, detail="No media paths provided")

    raw_changes = _dump_model(payload.changes, exclude_unset=True)
    if not raw_changes:
        raise HTTPException(status_code=400, detail="No metadata changes provided")

    try:
        _apply_metadata_changes({}, raw_changes)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    results = []
    errors = 0
    for raw_path in payload.paths:
        normalized_path = os.path.abspath(os.path.normpath(str(raw_path or "").strip()))
        if not os.path.isfile(normalized_path):
            errors += 1
            results.append({"path": normalized_path, "error": "Media not found"})
            continue
        if Path(normalized_path).suffix.lower() not in MEDIA_EXTENSIONS:
            errors += 1
            results.append({"path": normalized_path, "error": "Unsupported media file"})
            continue

        try:
            merged_metadata = _apply_metadata_changes(_read_metadata_file(normalized_path), raw_changes)
            saved_metadata = _write_metadata_file(normalized_path, merged_metadata)
            results.append({"path": normalized_path, "ok": True, "metadata": saved_metadata})
        except ValueError as exc:
            errors += 1
            results.append({"path": normalized_path, "error": str(exc)})

    return {"ok": errors == 0, "results": results}


@app.post("/api/caption/batch-toggle")
async def batch_toggle_caption(update: BatchCaptionUpdate):
    """Toggle a predefined caption on or off for multiple images."""
    results = []
    caption = update.resolved_caption()
    if not caption:
        raise HTTPException(status_code=400, detail="Caption text is required")

    # Load sections config from the folder of the first image
    cfg = _load_config()
    folder = str(Path(update.image_paths[0]).parent) if update.image_paths else ""
    sections = _get_folder_sections(cfg, folder)
    all_captions = _all_captions_from_sections(sections)

    for img_path in update.image_paths:
        if not os.path.isfile(img_path):
            results.append({"path": img_path, "error": "File not found"})
            continue

        # Read existing caption data
        headers = _all_headers_from_sections(sections)
        data = _read_caption_file(img_path, all_captions, headers)
        enabled = _read_enabled_captions(data)
        free_text = data["free_text"]

        if update.enabled:
            if caption not in enabled:
                enabled.append(caption)
        else:
            enabled = [item for item in enabled if item != caption]

        _write_caption_file(img_path, enabled, free_text, sections)
        results.append({"path": img_path, "ok": True})

    return {"results": results}


@app.post("/api/caption/rename-preset")
async def rename_caption_preset(update: RenameCaptionPresetUpdate):
    """Rename a configured caption preset and migrate existing caption files."""
    folder = os.path.normpath(update.folder)
    old_caption = update.resolved_old_caption()
    new_caption = update.resolved_new_caption()
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")
    if not old_caption or not new_caption:
        raise HTTPException(status_code=400, detail="Both old and new caption text are required")
    if old_caption == new_caption:
        cfg = _load_config()
        return {"ok": True, "sections": _get_folder_sections(cfg, folder), "touched_caption_files": 0}

    cfg = _load_config()
    sections = _get_folder_sections(cfg, folder)
    all_captions_before = _all_captions_from_sections(sections)
    if old_caption not in all_captions_before:
        raise HTTPException(status_code=404, detail="Caption not found")
    if new_caption in all_captions_before:
        raise HTTPException(status_code=400, detail="A caption with that text already exists")

    renamed = _rename_caption_in_sections(sections, old_caption, new_caption)
    if not renamed:
        raise HTTPException(status_code=404, detail="Caption not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)

    headers = _all_headers_from_sections(sections)
    folder_path = Path(folder)
    touched_caption_files = 0
    for entry in folder_path.iterdir():
        if not entry.is_file() or entry.suffix.lower() not in MEDIA_EXTENSIONS:
            continue
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_captions_before, headers)
        enabled_before = _read_enabled_captions(data)
        free_text_before = str(data.get("free_text", "") or "")
        if old_caption not in enabled_before and old_caption not in free_text_before:
            continue
        enabled = [new_caption if caption == old_caption else caption for caption in enabled_before]
        enabled = _normalize_enabled_captions(enabled, sections)
        free_text = free_text_before.replace(old_caption, new_caption)
        _write_caption_file(str(entry), enabled, free_text, sections)
        touched_caption_files += 1

    return {"ok": True, "sections": sections, "touched_caption_files": touched_caption_files}


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
        return {"ok": True, "sections": _get_folder_sections(cfg, folder), "touched_caption_files": 0}

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    section_names_before = [str(section.get("name") or "") for section in sections_before]
    if old_name not in section_names_before:
        raise HTTPException(status_code=404, detail="Section not found")
    if new_name and new_name in section_names_before:
        raise HTTPException(status_code=400, detail="A section with that name already exists")

    all_captions = _all_captions_from_sections(sections_before)
    headers_before = _all_headers_from_sections(sections_before)
    sections = _get_folder_sections(cfg, folder)
    renamed = _rename_section_in_sections(sections, old_name, new_name)
    if not renamed:
        raise HTTPException(status_code=404, detail="Section not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)

    folder_path = Path(folder)
    touched_caption_files = 0
    for entry in folder_path.iterdir():
        if not entry.is_file() or entry.suffix.lower() not in MEDIA_EXTENSIONS:
            continue
        caption_path = entry.with_suffix(".txt")
        if not caption_path.exists():
            continue
        data = _read_caption_file(str(entry), all_captions, headers_before)
        free_text = str(data.get("free_text", "") or "").replace(old_name, new_name)
        _write_caption_file(str(entry), _read_enabled_captions(data), free_text, sections)
        touched_caption_files += 1

    return {"ok": True, "sections": sections, "touched_caption_files": touched_caption_files}


@app.post("/api/caption/delete-preset")
async def delete_caption_preset(update: DeleteCaptionPresetUpdate):
    """Delete a configured caption and remove it from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    caption = update.resolved_caption()
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")
    if not caption:
        raise HTTPException(status_code=400, detail="Caption text is required")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    if caption not in _all_captions_from_sections(sections_before):
        raise HTTPException(status_code=404, detail="Caption not found")

    sections = _get_folder_sections(cfg, folder)
    removed = _remove_caption_from_sections(sections, caption)
    if not removed:
        raise HTTPException(status_code=404, detail="Caption not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)
    touched_caption_files = _remove_captions_from_caption_files(folder, sections_before, sections, [caption])
    return {"ok": True, "sections": sections, "removed_captions": [caption], "removed_sentences": [caption], "touched_caption_files": touched_caption_files}


@app.post("/api/group/delete")
async def delete_group(update: DeleteGroupUpdate):
    """Delete a configured group and remove its captions from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    sections = _get_folder_sections(cfg, folder)
    removed_captions = _remove_group_from_sections(sections, update.section_index, update.group_index)
    if removed_captions is None:
        raise HTTPException(status_code=404, detail="Group not found")

    _set_folder_sections(cfg, folder, sections)
    _save_config(cfg)
    touched_caption_files = _remove_captions_from_caption_files(folder, sections_before, sections, removed_captions)
    return {"ok": True, "sections": sections, "removed_captions": removed_captions, "removed_sentences": removed_captions, "touched_caption_files": touched_caption_files}


@app.post("/api/section/delete")
async def delete_section(update: DeleteSectionUpdate):
    """Delete a configured section and remove its captions from caption files in the folder."""
    folder = os.path.normpath(update.folder)
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="Folder not found")

    cfg = _load_config()
    sections_before = _get_folder_sections(cfg, folder)
    sections = _get_folder_sections(cfg, folder)
    updated_sections, removed_captions = _remove_section_from_sections(sections, update.section_index)
    if updated_sections is None or removed_captions is None:
        raise HTTPException(status_code=404, detail="Section not found")

    _set_folder_sections(cfg, folder, updated_sections)
    _save_config(cfg)
    touched_caption_files = _remove_captions_from_caption_files(folder, sections_before, updated_sections, removed_captions)
    return {"ok": True, "sections": updated_sections, "removed_captions": removed_captions, "removed_sentences": removed_captions, "touched_caption_files": touched_caption_files}


@app.post("/api/caption/save-free-text")
async def save_free_text(data: BatchFreeTextUpdate):
    """Save free text for a single media file, preserving predefined captions."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Media not found")

    # We need to know which lines are predefined to preserve them
    # Client will send the captions query param.
    # Actually let's accept captions in the body too.
    return {"ok": True}


class SaveCaptionFull(BaseModel):
    image_path: str
    enabled_captions: list[str] = Field(default_factory=list)
    enabled_sentences: list[str] = Field(default_factory=list)
    free_text: str

    def resolved_enabled_captions(self) -> list[str]:
        return list(self.enabled_captions or self.enabled_sentences)


class AutoCaptionRequest(BaseModel):
    media_path: Optional[str] = None
    media_paths: Optional[list[str]] = None
    image_path: Optional[str] = None
    image_paths: Optional[list[str]] = None
    model: Optional[str] = None
    prompt_template: Optional[str] = None
    group_prompt_template: Optional[str] = None
    enable_free_text: Optional[bool] = None
    free_text_only: Optional[bool] = None
    target_section_index: Optional[int] = None
    target_group_index: Optional[int] = None
    target_caption: Optional[str] = None
    target_sentence: Optional[str] = None
    free_text_prompt_template: Optional[str] = None
    timeout_seconds: Optional[int] = None
    max_output_tokens: Optional[int] = None

    def resolved_target_caption(self) -> str | None:
        caption = _read_caption_text(self.target_caption, self.target_sentence)
        return caption or None


def _resolve_caption_targets(
    sections: list[dict],
    target_section_index: int | None = None,
    target_group_index: int | None = None,
    target_caption: str | None = None,
) -> tuple[list[dict] | None, dict | None]:
    """Resolve full or scoped caption targets for auto captioning."""
    caption = str(target_caption or "").strip()
    if caption:
        for target in _iter_caption_targets_with_indices(sections):
            if target.get("type") == "caption" and target.get("caption") == caption:
                return [target], {
                    "type": "caption",
                    "caption": caption,
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
    """Full save of caption data for one media file."""
    if not os.path.isfile(data.image_path):
        raise HTTPException(status_code=404, detail="Media not found")
    # Load sections config for proper formatting
    cfg = _load_config()
    folder = str(Path(data.image_path).parent)
    sections = _get_folder_sections(cfg, folder)
    enabled_captions = _normalize_enabled_captions(data.resolved_enabled_captions(), sections)
    _write_caption_file(data.image_path, enabled_captions, data.free_text, sections)
    return {"ok": True}


@app.post("/api/auto-caption")
async def auto_caption(data: AutoCaptionRequest):
    """Ask a local Ollama vision model to verify each configured caption for one media file."""
    media_path = data.media_path or (data.media_paths[0] if data.media_paths else None) or data.image_path or (data.image_paths[0] if data.image_paths else None)
    if not media_path or not os.path.isfile(media_path):
        raise HTTPException(status_code=404, detail="Media not found")
    if _get_media_type_for_path(media_path) not in {"image", "video"}:
        raise HTTPException(status_code=400, detail="Unsupported media file for auto caption")

    cfg = _load_config()
    model = (data.model or _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL)).strip()
    host = _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS)
    max_output_tokens = data.max_output_tokens or _get_ollama_max_output_tokens(cfg, DEFAULT_OLLAMA_MAX_OUTPUT_TOKENS)

    if not model:
        raise HTTPException(status_code=400, detail="No Ollama model configured")

    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    encode_media_call = partial(
        _encode_media_for_ollama,
        ffmpeg_path=ffmpeg_path,
        ffprobe_path=ffprobe_path,
        thread_count=ffmpeg_threads,
        hwaccel_mode=ffmpeg_hwaccel,
    )

    folder = str(Path(media_path).parent)
    sections = _get_folder_sections(cfg, folder)
    all_captions = _all_captions_from_sections(sections)
    if not all_captions and not free_text_only:
        raise HTTPException(status_code=400, detail="No captions configured for this folder")
    targets, target_scope = _resolve_caption_targets(
        sections,
        data.target_section_index,
        data.target_group_index,
        data.resolved_target_caption(),
    )
    if not free_text_only and targets is None:
        if data.resolved_target_caption():
            raise HTTPException(status_code=400, detail="Invalid target caption")
        if data.target_group_index is not None:
            raise HTTPException(status_code=400, detail="Invalid target group")
        if data.target_section_index is not None:
            raise HTTPException(status_code=400, detail="Invalid target section")
        raise HTTPException(status_code=400, detail="Invalid target")

    headers = _all_headers_from_sections(sections)
    enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
    results: list[dict] = []
    is_scoped = bool(target_scope and target_scope.get("type") in {"group", "section", "caption"})

    loop = asyncio.get_event_loop()
    if not free_text_only:
        try:
            if target_scope and target_scope.get("type") == "full":
                auto_caption_call = partial(
                    _auto_caption_sections,
                    host,
                    model,
                    media_path,
                    sections,
                    initial_enabled_captions=enabled,
                    encode_image_func=encode_media_call,
                    generate_func=_ollama_generate,
                    prompt_template=prompt_template,
                    group_prompt_template=group_prompt_template,
                    timeout=timeout_seconds,
                    max_output_tokens=max_output_tokens,
                )
                enabled, results = await loop.run_in_executor(
                    executor,
                    auto_caption_call,
                )
                free_text = free_text if is_scoped else ""
                if not os.path.isfile(media_path):
                    raise FileNotFoundError(f"Media not found: {media_path}")
                _write_caption_file(media_path, enabled, free_text, sections)
            else:
                media_payload = await loop.run_in_executor(executor, encode_media_call, media_path)
                for target in targets or []:
                    if target["type"] == "caption":
                        if target.get("skip_auto_caption"):
                            results.append(_build_skipped_caption_result(target, enabled))
                            continue
                        caption = target["caption"]
                        payload = _build_ollama_generate_payload(
                            model,
                            _apply_media_prompt_context(_ollama_prompt_for_caption(caption, prompt_template), media_path),
                            media_payload,
                            max_output_tokens=max_output_tokens,
                        )
                        response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                        raw_answer = str(response.get("response") or "").strip()
                        is_match = _parse_ollama_yes_no(raw_answer)
                        enabled, free_text = _apply_caption_result_to_live_caption(
                            media_path,
                            caption,
                            is_match,
                            all_captions,
                            headers,
                            sections,
                        )
                        results.append({
                            "type": "sentence",
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "section_name": target.get("section_name", ""),
                            "caption": caption,
                            "sentence": caption,
                            "enabled": is_match,
                            "answer": raw_answer,
                            **_build_ollama_answer_meta(response),
                        })
                        continue

                    group_captions = target["captions"]
                    if target.get("skip_auto_caption"):
                        results.append(_build_skipped_group_result(target, enabled, sections))
                        continue
                    payload = _build_ollama_generate_payload(
                        model,
                        _apply_media_prompt_context(_ollama_prompt_for_group(target.get("group_name", ""), group_captions, group_prompt_template), media_path),
                        media_payload,
                        max_output_tokens=max_output_tokens,
                    )
                    response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                    raw_answer = str(response.get("response") or "").strip()
                    selection_index = _parse_ollama_selection(raw_answer, group_captions)
                    selected_caption = group_captions[selection_index - 1] if selection_index else None
                    selected_hidden = bool(selected_caption and _is_hidden_group_caption(sections, selected_caption))
                    enabled, free_text = _apply_group_result_to_live_caption(
                        media_path,
                        group_captions,
                        selected_caption,
                        all_captions,
                        headers,
                        sections,
                    )
                    results.append({
                        "type": "group",
                        "section_index": target.get("section_index"),
                        "group_index": target.get("group_index"),
                        "section_name": target.get("section_name", ""),
                        "group_name": target.get("group_name", ""),
                        "captions": group_captions,
                        "sentences": group_captions,
                        "selected_caption": selected_caption,
                        "selected_sentence": selected_caption,
                        "selected_hidden": selected_hidden,
                        "selection_index": selection_index,
                        "answer": raw_answer,
                        **_build_ollama_answer_meta(response),
                    })
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except FileNotFoundError:
            raise HTTPException(status_code=409, detail="Media was deleted during auto caption") from None
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption failed: {e}") from e

    free_text_model_output = ""
    added_free_text_lines: list[str] = []
    if enable_free_text:
        enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
        caption_text = _build_caption_text(enabled, free_text, sections)
        try:
            suggest_free_text_call = partial(
                _suggest_free_text,
                host,
                model,
                media_path,
                caption_text,
                encode_image_func=encode_media_call,
                generate_func=_ollama_generate,
                prompt_template=free_text_prompt_template,
                timeout=timeout_seconds,
                max_output_tokens=max_output_tokens,
            )
            free_text_model_output = await loop.run_in_executor(
                executor,
                suggest_free_text_call,
            )
            enabled, free_text, added_free_text_lines = _apply_free_text_result_to_live_caption(
                media_path,
                free_text_model_output,
                all_captions,
                headers,
                sections,
                preserve_existing=is_scoped,
            )
        except RuntimeError as e:
            raise HTTPException(status_code=502, detail=f"Ollama error: {e}") from e
        except FileNotFoundError:
            raise HTTPException(status_code=409, detail="Media was deleted during auto caption") from None
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Auto caption free-text step failed: {e}") from e

    enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
    return {
        "ok": True,
        "model": model,
        "host": host,
        "timeout_seconds": timeout_seconds,
        "max_output_tokens": max_output_tokens,
        "prompt_template": prompt_template,
        "group_prompt_template": group_prompt_template,
        "enable_free_text": enable_free_text,
        "free_text_only": free_text_only,
        "free_text_prompt_template": free_text_prompt_template,
        "enabled_captions": enabled,
        "enabled_sentences": enabled,
        "free_text": free_text,
        "free_text_model_output": free_text_model_output,
        "added_free_text_lines": added_free_text_lines,
        "results": results,
    }


@app.post("/api/auto-caption/stream")
async def auto_caption_stream(data: AutoCaptionRequest, request: Request):
    """Stream real-time auto-caption progress for one or more media files as NDJSON."""
    media_paths = data.media_paths or data.image_paths or ([data.media_path] if data.media_path else []) or ([data.image_path] if data.image_path else [])
    if not media_paths:
        raise HTTPException(status_code=400, detail="No media provided")

    cfg = _load_config()
    model = (data.model or _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL)).strip()
    host = _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    prompt_template = data.prompt_template or _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE)
    group_prompt_template = data.group_prompt_template or _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE)
    enable_free_text = _get_ollama_enable_free_text(cfg) if data.enable_free_text is None else data.enable_free_text
    free_text_only = bool(data.free_text_only)
    free_text_prompt_template = data.free_text_prompt_template or _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE)
    timeout_seconds = data.timeout_seconds or _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS)
    max_output_tokens = data.max_output_tokens or _get_ollama_max_output_tokens(cfg, DEFAULT_OLLAMA_MAX_OUTPUT_TOKENS)

    if not model:
        raise HTTPException(status_code=400, detail="No Ollama model configured")

    ffmpeg_path, ffprobe_path = _resolve_ffmpeg_binaries(cfg)
    ffmpeg_threads = _get_ffmpeg_threads(cfg)
    ffmpeg_hwaccel = _get_ffmpeg_hwaccel(cfg)
    encode_media_call = partial(
        _encode_media_for_ollama,
        ffmpeg_path=ffmpeg_path,
        ffprobe_path=ffprobe_path,
        thread_count=ffmpeg_threads,
        hwaccel_mode=ffmpeg_hwaccel,
    )

    def _event_bytes(payload: dict) -> bytes:
        return (json.dumps(payload, ensure_ascii=False) + "\n").encode("utf-8")

    async def event_stream():
        loop = asyncio.get_event_loop()
        total_processed = 0
        total_errors = 0
        yield _event_bytes({
            "type": "start",
            "count": len(media_paths),
            "model": model,
            "host": host,
            "group_prompt_template": group_prompt_template,
            "enable_free_text": enable_free_text,
            "free_text_only": free_text_only,
            "timeout_seconds": timeout_seconds,
            "max_output_tokens": max_output_tokens,
        })

        for media_path in media_paths:
            if await request.is_disconnected():
                break
            if not os.path.isfile(media_path):
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": "Media not found"})
                continue
            if _get_media_type_for_path(media_path) not in {"image", "video"}:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": "Unsupported media file for auto caption"})
                continue

            folder = str(Path(media_path).parent)
            sections = _get_folder_sections(cfg, folder)
            all_captions = _all_captions_from_sections(sections)
            if not all_captions and not free_text_only:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": "No captions configured for this folder"})
                continue

            targets, target_scope = _resolve_caption_targets(
                sections,
                data.target_section_index,
                data.target_group_index,
                data.resolved_target_caption(),
            )
            if not free_text_only and targets is None:
                total_errors += 1
                message = "Invalid target"
                if data.resolved_target_caption():
                    message = "Invalid target caption"
                elif data.target_group_index is not None:
                    message = "Invalid target group"
                elif data.target_section_index is not None:
                    message = "Invalid target section"
                yield _event_bytes({"type": "error", "path": media_path, "message": message})
                continue

            headers = _all_headers_from_sections(sections)
            enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
            results = []
            media_payload = await loop.run_in_executor(executor, encode_media_call, media_path)

            total_targets = 0 if free_text_only else len(targets or [])
            yield _event_bytes({
                "type": "image-start",
                "path": media_path,
                "total_captions": len(all_captions),
                "total_sentences": len(all_captions),
                "total_targets": total_targets,
                "free_text_only": free_text_only,
                "target_scope": target_scope,
                "enabled_captions": enabled,
                "enabled_sentences": enabled,
                "free_text": free_text,
            })

            try:
                if not free_text_only:
                    for index, target in enumerate(targets, start=1):
                        if await request.is_disconnected():
                            return
                        if target["type"] == "caption":
                            if target.get("skip_auto_caption"):
                                skipped_result = _build_skipped_caption_result(target, enabled)
                                results.append(skipped_result)
                                skipped_event = {key: value for key, value in skipped_result.items() if key != "type"}
                                yield _event_bytes({
                                    "type": "caption-check",
                                    "path": media_path,
                                    "index": index,
                                    "total": total_targets,
                                    "enabled_captions": enabled,
                                    "enabled_sentences": enabled,
                                    "free_text": free_text,
                                    **skipped_event,
                                })
                                continue
                            caption = target["caption"]
                            payload = _build_ollama_generate_payload(
                                model,
                                _apply_media_prompt_context(_ollama_prompt_for_caption(caption, prompt_template), media_path),
                                media_payload,
                                max_output_tokens=max_output_tokens,
                            )
                            response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                            raw_answer = str(response.get("response") or "").strip()
                            is_match = _parse_ollama_yes_no(raw_answer)
                            response_meta = _build_ollama_answer_meta(response)
                            enabled, free_text = _apply_caption_result_to_live_caption(
                                media_path,
                                caption,
                                is_match,
                                all_captions,
                                headers,
                                sections,
                            )
                            result = {
                                "type": "caption",
                                "caption": caption,
                                "sentence": caption,
                                "enabled": is_match,
                                "answer": raw_answer,
                                **response_meta,
                            }
                            results.append(result)
                            yield _event_bytes({
                                "type": "caption-check",
                                "path": media_path,
                                "index": index,
                                "total": total_targets,
                                "caption": caption,
                                "sentence": caption,
                                "enabled": is_match,
                                "enabled_captions": enabled,
                                "enabled_sentences": enabled,
                                "free_text": free_text,
                                "answer": raw_answer,
                                **response_meta,
                            })
                            continue

                        group_captions = target["captions"]
                        if target.get("skip_auto_caption"):
                            skipped_result = _build_skipped_group_result(target, enabled, sections)
                            results.append(skipped_result)
                            skipped_event = {key: value for key, value in skipped_result.items() if key != "type"}
                            yield _event_bytes({
                                "type": "group-selection",
                                "path": media_path,
                                "index": index,
                                "total": total_targets,
                                "enabled_captions": enabled,
                                "enabled_sentences": enabled,
                                "free_text": free_text,
                                **skipped_event,
                            })
                            continue
                        payload = _build_ollama_generate_payload(
                            model,
                            _apply_media_prompt_context(_ollama_prompt_for_group(target.get("group_name", ""), group_captions, group_prompt_template), media_path),
                            media_payload,
                            max_output_tokens=max_output_tokens,
                        )
                        response = await loop.run_in_executor(executor, _ollama_generate, host, payload, timeout_seconds)
                        raw_answer = str(response.get("response") or "").strip()
                        response_meta = _build_ollama_answer_meta(response)
                        selection_index = _parse_ollama_selection(raw_answer, group_captions)
                        selected_caption = group_captions[selection_index - 1] if selection_index else None
                        selected_hidden = bool(selected_caption and _is_hidden_group_caption(sections, selected_caption))
                        enabled, free_text = _apply_group_result_to_live_caption(
                            media_path,
                            group_captions,
                            selected_caption,
                            all_captions,
                            headers,
                            sections,
                        )
                        result = {
                            "type": "group",
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "group_name": target.get("group_name", ""),
                            "captions": group_captions,
                            "sentences": group_captions,
                            "selected_caption": selected_caption,
                            "selected_sentence": selected_caption,
                            "selected_hidden": selected_hidden,
                            "selection_index": selection_index,
                            "answer": raw_answer,
                            **response_meta,
                        }
                        results.append(result)
                        yield _event_bytes({
                            "type": "group-selection",
                            "path": media_path,
                            "index": index,
                            "total": total_targets,
                            "section_index": target.get("section_index"),
                            "group_index": target.get("group_index"),
                            "group_name": target.get("group_name", ""),
                            "captions": group_captions,
                            "sentences": group_captions,
                            "selected_caption": selected_caption,
                            "selected_sentence": selected_caption,
                            "selected_hidden": selected_hidden,
                            "selection_index": selection_index,
                            "enabled_captions": enabled,
                            "enabled_sentences": enabled,
                            "free_text": free_text,
                            "answer": raw_answer,
                            **response_meta,
                        })

                free_text_model_output = ""
                added_free_text_lines: list[str] = []
                if enable_free_text:
                    if await request.is_disconnected():
                        return
                    is_scoped = bool(target_scope and target_scope.get("type") in {"group", "section", "caption"})
                    enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
                    caption_text = _build_caption_text(enabled, free_text, sections)
                    free_payload = _build_ollama_generate_payload(
                        model,
                        _apply_media_prompt_context(_ollama_prompt_for_free_text(caption_text, free_text_prompt_template), media_path),
                        media_payload,
                        max_output_tokens=max_output_tokens,
                    )
                    response = await loop.run_in_executor(executor, _ollama_generate, host, free_payload, timeout_seconds)
                    free_text_model_output = str(response.get("response") or "").strip()
                    response_meta = _build_ollama_answer_meta(response)
                    enabled, free_text, added_free_text_lines = _apply_free_text_result_to_live_caption(
                        media_path,
                        free_text_model_output,
                        all_captions,
                        headers,
                        sections,
                        preserve_existing=is_scoped,
                    )
                    yield _event_bytes({
                        "type": "free-text",
                        "path": media_path,
                        "answer": free_text_model_output,
                        "enabled_captions": enabled,
                        "enabled_sentences": enabled,
                        "free_text": free_text,
                        "added_lines": added_free_text_lines,
                        **response_meta,
                    })

                enabled, free_text = _read_live_caption_state(media_path, all_captions, headers)
                total_processed += 1
                yield _event_bytes({
                    "type": "image-complete",
                    "path": media_path,
                    "free_text_only": free_text_only,
                    "enabled_captions": enabled,
                    "enabled_sentences": enabled,
                    "free_text": free_text,
                    "results": results,
                })
            except RuntimeError as e:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": f"Ollama error: {e}"})
            except FileNotFoundError:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": "Media was deleted during auto caption"})
            except Exception as e:
                total_errors += 1
                yield _event_bytes({"type": "error", "path": media_path, "message": f"Auto caption failed: {e}"})

        yield _event_bytes({
            "type": "done",
            "processed": total_processed,
            "errors": total_errors,
            "count": len(media_paths),
        })

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


@app.get("/api/captions/bulk")
async def get_captions_bulk(
    paths: str = Query(...),
    captions: str | None = Query(default=None),
    sentences: str | None = Query(default=None),
):
    """Get caption status for multiple images at once.
    paths: JSON array of image paths
    captions: JSON array of predefined captions
    """
    try:
        image_paths = json.loads(paths)
        predefined = json.loads(_coalesce_caption_query(captions, sentences))
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid JSON")

    return _get_bulk_caption_results(image_paths, predefined)


def _get_bulk_caption_results(image_paths: list[str], predefined_captions: list[str]) -> dict:
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
            data = _read_caption_file(img_path, predefined_captions, headers)
            results[img_path] = {"enabled_captions": _read_enabled_captions(data), **data}
        else:
            results[img_path] = {"enabled_captions": [], "enabled_sentences": [], "free_text": ""}

    return results


class BulkCaptionsRequest(BaseModel):
    paths: list[str]
    captions: list[str] = Field(default_factory=list)
    sentences: list[str] = Field(default_factory=list)

    def resolved_captions(self) -> list[str]:
        return list(self.captions or self.sentences)


@app.post("/api/captions/bulk")
async def post_captions_bulk(payload: BulkCaptionsRequest):
    """Get caption status for multiple images at once from a JSON body."""
    return _get_bulk_caption_results(payload.paths or [], payload.resolved_captions())


# ===== SETTINGS API =====

class SettingsUpdate(BaseModel):
    last_folder: Optional[str] = None
    sections: Optional[list[dict]] = None
    folder: Optional[str] = None  # which folder these sections belong to
    video_training_presets: Optional[list[dict]] = None
    video_training_profile_key: Optional[str] = None
    thumb_size: Optional[int] = None
    crop_aspect_ratios: Optional[list[str]] = None
    mask_latent_base_width_presets: Optional[list[int]] = None
    https_certfile: Optional[str] = None
    https_keyfile: Optional[str] = None
    https_port: Optional[int] = None
    remote_http_mode: Optional[str] = None
    ffmpeg_path: Optional[str] = None
    ffmpeg_threads: Optional[int] = None
    ffmpeg_hwaccel: Optional[str] = None
    processing_reserved_cores: Optional[int] = None
    ollama_host: Optional[str] = None
    ollama_server: Optional[str] = None
    ollama_port: Optional[int] = None
    ollama_timeout_seconds: Optional[int] = None
    ollama_max_output_tokens: Optional[int] = None
    ollama_model: Optional[str] = None
    ollama_prompt_template: Optional[str] = None
    ollama_group_prompt_template: Optional[str] = None
    ollama_enable_free_text: Optional[bool] = None
    ollama_free_text_prompt_template: Optional[str] = None
    comfyui_server: Optional[str] = None
    comfyui_port: Optional[int] = None
    comfyui_workflow_path: Optional[str] = None
    comfyui_output_folder: Optional[str] = None
    comfyui_auto_preview: Optional[bool] = None


class ComfyUiPromptPreviewRequest(BaseModel):
    image_path: str
    enabled_captions: list[str] = Field(default_factory=list)
    enabled_sentences: list[str] = Field(default_factory=list)
    free_text: str = ""

    def resolved_enabled_captions(self) -> list[str]:
        return list(self.enabled_captions or self.enabled_sentences)


@app.get("/api/settings")
async def get_settings(folder: Optional[str] = Query(default=None)):
    """Get full settings. If folder is specified, include sections for that folder."""
    cfg = _load_config()
    video_training_presets = _get_video_training_presets(cfg)
    result = {
        "last_folder": cfg.get("last_folder", ""),
        "thumb_size": int(cfg.get("thumb_size", 160) or 160),
        "crop_aspect_ratios": _get_crop_aspect_ratios(cfg, DEFAULT_CROP_ASPECT_RATIOS),
        "mask_latent_base_width_presets": _normalize_mask_latent_base_width_presets(cfg.get("mask_latent_base_width_presets")),
        "video_training_presets": video_training_presets,
        "https_certfile": str(cfg.get("https_certfile") or ""),
        "https_keyfile": str(cfg.get("https_keyfile") or ""),
        "https_port": _get_https_port(cfg, DEFAULT_HTTPS_PORT),
        "remote_http_mode": _get_remote_http_mode(cfg, DEFAULT_REMOTE_HTTP_MODE),
        "ffmpeg_path": _get_ffmpeg_path(cfg),
        "ffmpeg_threads": _get_ffmpeg_threads(cfg),
        "ffmpeg_hwaccel": _get_ffmpeg_hwaccel(cfg),
        "processing_reserved_cores": _get_processing_reserved_cores(cfg),
        "ollama_host": _get_ollama_host(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_server": _get_ollama_server(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_port": _get_ollama_port(cfg, DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT),
        "ollama_timeout_seconds": _get_ollama_timeout_seconds(cfg, DEFAULT_OLLAMA_TIMEOUT_SECONDS),
        "ollama_max_output_tokens": _get_ollama_max_output_tokens(cfg, DEFAULT_OLLAMA_MAX_OUTPUT_TOKENS),
        "ollama_model": _get_ollama_model(cfg, DEFAULT_OLLAMA_MODEL),
        "ollama_prompt_template": _get_ollama_prompt_template(cfg, DEFAULT_OLLAMA_PROMPT_TEMPLATE),
        "ollama_group_prompt_template": _get_ollama_group_prompt_template(cfg, DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE),
        "ollama_enable_free_text": _get_ollama_enable_free_text(cfg),
        "ollama_free_text_prompt_template": _get_ollama_free_text_prompt_template(cfg, DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE),
        "comfyui_server": _get_comfyui_server(cfg, DEFAULT_COMFYUI_SERVER),
        "comfyui_port": _get_comfyui_port(cfg, DEFAULT_COMFYUI_PORT),
        "comfyui_workflow_path": str(cfg.get("comfyui_workflow_path") or ""),
        "comfyui_output_folder": str(cfg.get("comfyui_output_folder") or ""),
        "comfyui_auto_preview": _get_comfyui_auto_preview(cfg),
    }
    if folder:
        result["sections"] = _get_folder_sections(cfg, folder)
        result["folder"] = os.path.normpath(folder)
        profile_key = _get_folder_video_training_profile_key(cfg, folder, video_training_presets)
        result["video_training_profile_key"] = profile_key
        result["video_training_profile"] = _get_video_training_profile_by_key(video_training_presets, profile_key)
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
    sections_before = None
    sections_after = None
    touched_caption_files = None
    if data.last_folder is not None:
        cfg["last_folder"] = data.last_folder
    if data.video_training_presets is not None:
        cfg["video_training_presets"] = _normalize_video_training_presets(data.video_training_presets)
    if data.thumb_size is not None:
        cfg["thumb_size"] = max(60, min(400, int(data.thumb_size)))
    if data.crop_aspect_ratios is not None:
        cfg["crop_aspect_ratios"] = [str(r).strip() for r in data.crop_aspect_ratios if str(r).strip()] or list(DEFAULT_CROP_ASPECT_RATIOS)
    if data.mask_latent_base_width_presets is not None:
        cfg["mask_latent_base_width_presets"] = _normalize_mask_latent_base_width_presets(data.mask_latent_base_width_presets)
    if data.https_certfile is not None:
        cfg["https_certfile"] = str(data.https_certfile or "").strip()
    if data.https_keyfile is not None:
        cfg["https_keyfile"] = str(data.https_keyfile or "").strip()
    if data.https_port is not None:
        cfg["https_port"] = max(1, min(65535, int(data.https_port)))
    if data.remote_http_mode is not None:
        cfg["remote_http_mode"] = _get_remote_http_mode({"remote_http_mode": data.remote_http_mode}, DEFAULT_REMOTE_HTTP_MODE)
    if data.ffmpeg_path is not None:
        cfg["ffmpeg_path"] = str(data.ffmpeg_path or "").strip()
    if data.ffmpeg_threads is not None:
        cfg["ffmpeg_threads"] = max(0, int(data.ffmpeg_threads))
    if data.ffmpeg_hwaccel is not None:
        cfg["ffmpeg_hwaccel"] = _get_ffmpeg_hwaccel({"ffmpeg_hwaccel": data.ffmpeg_hwaccel})
    if data.processing_reserved_cores is not None:
        cfg["processing_reserved_cores"] = max(0, min(HOST_CPU_COUNT - 1, int(data.processing_reserved_cores))) if HOST_CPU_COUNT > 1 else 0
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
    if data.ollama_max_output_tokens is not None:
        cfg["ollama_max_output_tokens"] = max(1, int(data.ollama_max_output_tokens))
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
    if data.comfyui_server is not None:
        cfg["comfyui_server"] = str(data.comfyui_server or "").strip() or DEFAULT_COMFYUI_SERVER
    if data.comfyui_port is not None:
        cfg["comfyui_port"] = max(1, min(65535, int(data.comfyui_port)))
    if data.comfyui_workflow_path is not None:
        cfg["comfyui_workflow_path"] = str(data.comfyui_workflow_path or "").strip()
    if data.comfyui_output_folder is not None:
        cfg["comfyui_output_folder"] = str(data.comfyui_output_folder or "").strip()
    if data.comfyui_auto_preview is not None:
        cfg["comfyui_auto_preview"] = bool(data.comfyui_auto_preview)
    cfg["ollama_host"] = _compose_ollama_host(cfg.get("ollama_server"), cfg.get("ollama_port"), DEFAULT_OLLAMA_SERVER, DEFAULT_OLLAMA_PORT)
    if data.video_training_profile_key is not None and data.folder:
        presets = _get_video_training_presets(cfg)
        _set_folder_video_training_profile_key(cfg, data.folder, data.video_training_profile_key, presets)
    if data.sections is not None and data.folder:
        sections_before = _get_folder_sections(cfg, data.folder)
        _set_folder_sections(cfg, data.folder, data.sections)
        sections_after = _get_folder_sections(cfg, data.folder)
    _save_config(cfg)
    if sections_before is not None and sections_after is not None:
        touched_caption_files = len(_rewrite_caption_files_for_section_change(data.folder, sections_before, sections_after))
    response = {"ok": True}
    if sections_after is not None and data.folder:
        response["folder"] = os.path.normpath(data.folder)
        response["sections"] = sections_after
        response["touched_caption_files"] = touched_caption_files if touched_caption_files is not None else 0
    return response


@app.get("/api/comfyui/prompt-preview/status")
async def get_comfyui_prompt_preview_status(image_path: str = Query(...)):
    """Return tracked ComfyUI prompt-preview jobs and summary information for one source image."""
    normalized_path = _normalize_existing_media_path(image_path)
    if not _is_image_path(normalized_path):
        raise HTTPException(status_code=400, detail="Prompt preview currently supports images only")
    cfg = _load_config()
    try:
        snapshot = _refresh_comfyui_jobs_for_image(cfg, normalized_path)
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "image_path": normalized_path,
        "jobs": snapshot["jobs"],
        "summary": snapshot["summary"],
        "files": snapshot["files"],
    }


@app.get("/api/comfyui/prompt-preview/files")
async def get_comfyui_prompt_preview_files(image_path: str = Query(...)):
    """Return generated preview files for one source image based on filename prefix matching."""
    normalized_path = _normalize_existing_media_path(image_path)
    if not _is_image_path(normalized_path):
        raise HTTPException(status_code=400, detail="Prompt preview currently supports images only")
    cfg = _load_config()
    try:
        files = _scan_comfyui_preview_files(_get_comfyui_output_folder(cfg), normalized_path)
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "image_path": normalized_path,
        "filename_prefix": Path(normalized_path).stem,
        "files": files,
    }


@app.post("/api/comfyui/prompt-preview")
async def post_comfyui_prompt_preview(data: ComfyUiPromptPreviewRequest):
    """Submit the selected image caption text to the configured ComfyUI workflow."""
    normalized_path = _normalize_existing_media_path(data.image_path)
    if not _is_image_path(normalized_path):
        raise HTTPException(status_code=400, detail="Prompt preview currently supports images only")

    cfg = _load_config()
    host = _get_comfyui_host(cfg, DEFAULT_COMFYUI_SERVER, DEFAULT_COMFYUI_PORT)
    workflow_path = _get_comfyui_workflow_path(cfg, DEFAULT_COMFYUI_WORKFLOW_PATH)
    output_folder = _get_comfyui_output_folder(cfg, DEFAULT_COMFYUI_OUTPUT_FOLDER)
    folder = str(Path(normalized_path).parent)
    sections = _get_folder_sections(cfg, folder)
    enabled_captions = _normalize_enabled_captions(data.resolved_enabled_captions(), sections)
    free_text = str(data.free_text or "")
    caption_text = _build_caption_text(enabled_captions, free_text, sections)
    if not caption_text.strip():
        raise HTTPException(status_code=400, detail="Prompt preview requires caption text")

    try:
        workflow_template = _load_comfyui_workflow_template(workflow_path)
        if not os.path.isdir(output_folder):
            raise RuntimeError(f"ComfyUI output folder not found: {output_folder}")
        filename_prefix = _build_comfyui_filename_prefix(normalized_path)
        prompt_payload = _build_comfyui_prompt_payload(workflow_template, caption_text, filename_prefix)
        response = _queue_comfyui_prompt(host, prompt_payload)
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    prompt_id = str(response.get("prompt_id") or "").strip()
    if not prompt_id:
        raise HTTPException(status_code=502, detail="ComfyUI did not return a prompt_id")

    job = _upsert_comfyui_job_record(
        prompt_id,
        status="queued",
        message="Queued in ComfyUI",
        error="",
        image_path=normalized_path,
        folder=folder,
        filename_prefix=filename_prefix,
        caption_text=caption_text,
        workflow_path=workflow_path,
        output_folder=output_folder,
        queue_number=response.get("number"),
    ) or {}

    try:
        snapshot = _refresh_comfyui_jobs_for_image(cfg, normalized_path)
    except RuntimeError:
        snapshot = {
            "jobs": [job],
            "summary": {
                "total": 1,
                "spawned": 1,
                "queued": 1,
                "running": 0,
                "completed": 0,
                "failed": 0,
                "latest_prompt_id": prompt_id,
                "latest_output_path": "",
            },
            "files": [],
        }

    return {
        "ok": True,
        "job": job,
        "jobs": snapshot["jobs"],
        "summary": snapshot["summary"],
        "files": snapshot["files"],
    }


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
