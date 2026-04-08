"""Per-media metadata sidecar helpers."""

from __future__ import annotations

import json
import math
from pathlib import Path
from uuid import uuid4


METADATA_FIELD_NAMES = ("seed", "min_t", "max_t", "sampling_frequency")


def _get_metadata_path(media_path: str) -> Path:
    """Return the sidecar path for a media metadata JSON file."""
    media = Path(media_path)
    return media.with_name(f"{media.name}.meta.json")


def _parse_int_field(field_name: str, value: object) -> int:
    """Parse an integer metadata field from JSON-compatible input."""
    if isinstance(value, bool):
        raise ValueError(f"{field_name} must be an integer")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise ValueError(f"{field_name} must be an integer")
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} must be an integer")
        try:
            parsed = float(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be an integer") from exc
        if not math.isfinite(parsed) or not parsed.is_integer():
            raise ValueError(f"{field_name} must be an integer")
        return int(parsed)
    raise ValueError(f"{field_name} must be an integer")


def _parse_float_field(field_name: str, value: object) -> float:
    """Parse a numeric metadata field from JSON-compatible input."""
    if isinstance(value, bool):
        raise ValueError(f"{field_name} must be a number")
    if isinstance(value, (int, float)):
        parsed = float(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} must be a number")
        try:
            parsed = float(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a number") from exc
    else:
        raise ValueError(f"{field_name} must be a number")

    if not math.isfinite(parsed):
        raise ValueError(f"{field_name} must be a finite number")
    return parsed


def _normalize_metadata_dict(metadata: object) -> dict:
    """Normalize sparse metadata fields and validate cross-field constraints."""
    if metadata is None:
        return {}
    if not isinstance(metadata, dict):
        raise ValueError("Metadata must be a JSON object")

    normalized: dict[str, int | float] = {}
    for field_name in METADATA_FIELD_NAMES:
        if field_name not in metadata:
            continue
        value = metadata.get(field_name)
        if value is None:
            continue
        if field_name == "sampling_frequency":
            parsed = _parse_float_field(field_name, value)
            if parsed < 0:
                raise ValueError("sampling_frequency must be greater than or equal to 0")
            normalized[field_name] = parsed
        else:
            normalized[field_name] = _parse_int_field(field_name, value)

    min_t = normalized.get("min_t")
    max_t = normalized.get("max_t")
    if min_t is not None and max_t is not None and min_t > max_t:
        raise ValueError("min_t must be less than or equal to max_t")
    return normalized


def _read_metadata_file(media_path: str) -> dict:
    """Read and normalize a per-media metadata sidecar."""
    metadata_path = _get_metadata_path(media_path)
    if not metadata_path.exists():
        return {}

    try:
        raw = json.loads(metadata_path.read_text(encoding="utf-8"))
    except Exception:
        return {}

    try:
        return _normalize_metadata_dict(raw)
    except ValueError:
        return {}


def _write_metadata_file(media_path: str, metadata: object) -> dict:
    """Write sparse metadata JSON or remove the sidecar when empty."""
    normalized = _normalize_metadata_dict(metadata)
    metadata_path = _get_metadata_path(media_path)

    if not normalized:
        metadata_path.unlink(missing_ok=True)
        return {}

    temp_path = metadata_path.with_name(f"{metadata_path.name}.{uuid4().hex}.tmp")
    temp_path.write_text(json.dumps(normalized, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(metadata_path)
    return normalized


def _apply_metadata_changes(metadata: object, changes: object) -> dict:
    """Merge sparse metadata changes into an existing metadata record."""
    current = _normalize_metadata_dict(metadata)
    if not isinstance(changes, dict):
        raise ValueError("Metadata changes must be a JSON object")

    merged = dict(current)
    for field_name in METADATA_FIELD_NAMES:
        if field_name not in changes:
            continue
        value = changes.get(field_name)
        if value is None:
            merged.pop(field_name, None)
        else:
            merged[field_name] = value
    return _normalize_metadata_dict(merged)