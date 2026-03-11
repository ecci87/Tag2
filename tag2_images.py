"""Image loading, preview, crop, and thumbnail helpers."""

from __future__ import annotations

import atexit
import base64
import hashlib
import os
import shutil
import tempfile
import warnings
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageOps

try:
    import piexif
except ImportError:  # pragma: no cover - optional during import
    piexif = None


Image.MAX_IMAGE_PIXELS = None
warnings.simplefilter("ignore", Image.DecompressionBombWarning)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"}
THUMBNAIL_SIZES = [64, 128, 256, 400]
PREVIEW_MAX_SIZE = 2048
thumbnail_cache: dict[tuple[str, float, int, tuple | None], bytes] = {}
RUNTIME_CROP_BACKUP_DIR = tempfile.mkdtemp(prefix="tag2-crop-")


@atexit.register
def _cleanup_runtime_crop_backups():
    """Delete temporary crop backups when the server process exits."""
    shutil.rmtree(RUNTIME_CROP_BACKUP_DIR, ignore_errors=True)


def _normalize_image_key(image_path: str) -> str:
    """Normalize an image path for use as a cache or config key."""
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
    """Drop cached thumbnails and previews for a specific image."""
    normalized = _normalize_image_key(image_path)
    for key in [item for item in thumbnail_cache if _normalize_image_key(item[0]) == normalized]:
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
    if not exif_bytes or piexif is None:
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


def _save_image_file(
    filepath: str,
    image: Image.Image,
    original_format: str,
    exif_bytes: bytes | None = None,
    icc_profile: bytes | None = None,
):
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
    """Return current and original display dimensions and crop state."""
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


def _generate_thumbnail(filepath: str, size: int, crop: dict | None = None) -> bytes:
    """Generate a JPEG thumbnail of the given size."""
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
    """Get a thumbnail from cache or generate it."""
    mtime = os.path.getmtime(filepath)
    crop_key = None if not crop else (crop.get("x"), crop.get("y"), crop.get("w"), crop.get("h"), crop.get("ratio"))
    key = (filepath, mtime, size, crop_key)
    if key not in thumbnail_cache:
        thumbnail_cache[key] = _generate_thumbnail(filepath, size, crop)
    return thumbnail_cache[key]


def _render_image_bytes(
    filepath: str,
    crop: dict | None = None,
    max_size: int | None = None,
    force_jpeg: bool = False,
) -> tuple[bytes, str]:
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
