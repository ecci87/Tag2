"""Image and video loading, preview, crop, and thumbnail helpers."""

from __future__ import annotations

import atexit
import base64
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import warnings
from functools import lru_cache
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
VIDEO_EXTENSIONS = {".mp4", ".m4v", ".mov", ".webm", ".mkv", ".avi"}
MEDIA_EXTENSIONS = IMAGE_EXTENSIONS | VIDEO_EXTENSIONS
MASK_SIDECAR_SUFFIX = ".mask.png"
THUMBNAIL_SIZES = [64, 128, 256, 400]
PREVIEW_MAX_SIZE = 2048
thumbnail_cache: dict[tuple[str, float, int, tuple | None], bytes] = {}
RUNTIME_CROP_BACKUP_DIR = tempfile.mkdtemp(prefix="tag2-crop-")
GPU_ENCODER_CANDIDATES = (
    "h264_v4l2m2m",
    "h264_nvenc",
    "h264_videotoolbox",
    "h264_amf",
)
GPU_ENCODER_EXTENSIONS = {".mp4", ".m4v", ".mov", ".mkv"}
OLLAMA_VIDEO_FRAME_COUNT = 4
OLLAMA_VIDEO_FRAME_SIZE = 896


@atexit.register
def _cleanup_runtime_crop_backups():
    """Delete temporary crop backups when the server process exits."""
    shutil.rmtree(RUNTIME_CROP_BACKUP_DIR, ignore_errors=True)


def _normalize_image_key(image_path: str) -> str:
    """Normalize an image path for use as a cache or config key."""
    return os.path.normcase(os.path.abspath(os.path.normpath(image_path)))


def _get_media_type_for_path(path_value: str) -> str:
    """Return the supported media type for a path."""
    if _is_mask_sidecar_path(path_value):
        return "other"
    suffix = Path(str(path_value or "")).suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        return "image"
    if suffix in VIDEO_EXTENSIONS:
        return "video"
    return "other"


def _is_image_path(path_value: str) -> bool:
    """Return whether a path refers to a supported image file."""
    return _get_media_type_for_path(path_value) == "image"


def _is_video_path(path_value: str) -> bool:
    """Return whether a path refers to a supported video file."""
    return _get_media_type_for_path(path_value) == "video"


def _is_mask_sidecar_path(path_value: str) -> bool:
    """Return whether a path is one of this app's image mask sidecars."""
    file_name = Path(str(path_value or "")).name.lower()
    return file_name.endswith(MASK_SIDECAR_SUFFIX)


def _get_image_mask_path(image_path: str) -> Path:
    """Build the mask sidecar path for an image file."""
    return Path(f"{image_path}{MASK_SIDECAR_SUFFIX}")


def _get_image_mask_info(image_path: str) -> dict:
    """Return mask sidecar metadata for an image file."""
    mask_path = _get_image_mask_path(image_path)
    exists = mask_path.is_file()
    info = {
        "path": str(mask_path),
        "exists": exists,
        "mtime": mask_path.stat().st_mtime if exists else 0.0,
    }
    if exists:
        with Image.open(mask_path) as mask_image:
            info["width"], info["height"] = mask_image.size
    return info


def _write_default_image_mask(image_path: str, default_value: int = 128) -> dict:
    """Create or normalize a grayscale mask sidecar for an image."""
    image_width, image_height = _get_display_image_size(image_path)
    mask_path = _get_image_mask_path(image_path)
    default_fill = max(0, min(255, int(round(default_value))))
    created = False

    if mask_path.is_file():
        with Image.open(mask_path) as existing_mask:
            normalized_mask = ImageOps.exif_transpose(existing_mask).convert("L")
            if normalized_mask.size != (image_width, image_height):
                normalized_mask = normalized_mask.resize((image_width, image_height), Image.Resampling.LANCZOS)
            normalized_mask.save(mask_path, format="PNG")
    else:
        Image.new("L", (image_width, image_height), color=default_fill).save(mask_path, format="PNG")
        created = True

    info = _get_image_mask_info(image_path)
    info["created"] = created
    info["image_width"] = image_width
    info["image_height"] = image_height
    return info


def _save_image_mask(image_path: str, mask_bytes: bytes) -> dict:
    """Persist an uploaded mask image as a normalized grayscale PNG sidecar."""
    image_width, image_height = _get_display_image_size(image_path)
    mask_path = _get_image_mask_path(image_path)
    with Image.open(BytesIO(mask_bytes)) as uploaded_mask:
        normalized_mask = ImageOps.exif_transpose(uploaded_mask).convert("L")
        if normalized_mask.size != (image_width, image_height):
            normalized_mask = normalized_mask.resize((image_width, image_height), Image.Resampling.LANCZOS)
        normalized_mask.save(mask_path, format="PNG")

    info = _get_image_mask_info(image_path)
    info["created"] = False
    info["image_width"] = image_width
    info["image_height"] = image_height
    return info


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


@lru_cache(maxsize=8)
def _list_ffmpeg_encoders(ffmpeg_path: str) -> frozenset[str]:
    """Return the set of available ffmpeg encoder names."""
    try:
        result = subprocess.run(
            [ffmpeg_path, "-hide_banner", "-encoders"],
            capture_output=True,
            check=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception:
        return frozenset()

    encoders: set[str] = set()
    for line in result.stdout.splitlines():
        stripped = line.strip()
        if not stripped or not stripped.startswith(("V", ".V")):
            continue
        parts = stripped.split()
        if len(parts) >= 2:
            encoders.add(parts[1].strip())
    return frozenset(encoders)


def _build_ffmpeg_decode_args(hwaccel_mode: str | None = "auto") -> list[str]:
    """Build optional decode-side hardware acceleration args."""
    mode = str(hwaccel_mode or "auto").strip().lower()
    if mode in {"", "off", "false", "none", "cpu"}:
        return []
    return ["-hwaccel", "auto"]


def _choose_gpu_video_encoder(ffmpeg_path: str, output_path: str) -> str | None:
    """Pick a supported hardware video encoder for compatible containers."""
    if Path(output_path).suffix.lower() not in GPU_ENCODER_EXTENSIONS:
        return None
    encoders = _list_ffmpeg_encoders(ffmpeg_path)
    for encoder in GPU_ENCODER_CANDIDATES:
        if encoder in encoders:
            return encoder
    return None


def _build_ffmpeg_processing_args(
    ffmpeg_path: str,
    output_path: str,
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
    allow_gpu_encode: bool = True,
) -> tuple[list[str], list[str], list[str]]:
    """Build decode args and preferred/fallback encode args for ffmpeg jobs."""
    threads = max(0, int(thread_count or 0))
    decode_args = _build_ffmpeg_decode_args(hwaccel_mode)
    cpu_args = ["-threads", str(threads)] if threads > 0 else []
    gpu_encoder = _choose_gpu_video_encoder(ffmpeg_path, output_path) if allow_gpu_encode else None
    gpu_args = []
    if gpu_encoder:
        gpu_args = ["-c:v", gpu_encoder]
        if threads > 0:
            gpu_args.extend(["-threads", str(threads)])
    return decode_args, gpu_args, cpu_args


def _run_ffmpeg_with_optional_fallback(
    primary_command: list[str],
    duration_seconds: float,
    on_progress=None,
    fallback_command: list[str] | None = None,
) -> None:
    """Run ffmpeg and retry once with fallback args when hardware acceleration fails."""
    try:
        _run_ffmpeg_with_progress(primary_command, duration_seconds, on_progress)
        return
    except RuntimeError:
        if not fallback_command:
            raise
        if on_progress:
            on_progress(0.0, "Retrying on CPU")
        _run_ffmpeg_with_progress(fallback_command, duration_seconds, on_progress)


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


def _probe_video_info(filepath: str, ffprobe_path: str = "ffprobe") -> dict:
    """Probe a video file for primary stream dimensions and duration."""
    command = [
        ffprobe_path,
        "-v",
        "error",
        "-select_streams",
        "v:0",
        "-show_entries",
        "stream=width,height:format=duration",
        "-of",
        "json",
        filepath,
    ]
    try:
        result = subprocess.run(command, capture_output=True, check=True)
    except FileNotFoundError as exc:
        raise RuntimeError(f"ffprobe not found: {ffprobe_path}") from exc
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(stderr or f"ffprobe failed for {Path(filepath).name}") from exc

    try:
        payload = json.loads(result.stdout.decode("utf-8", errors="replace") or "{}")
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"ffprobe returned invalid JSON for {Path(filepath).name}") from exc

    stream = ((payload.get("streams") or [{}])[0]) if isinstance(payload, dict) else {}
    format_data = payload.get("format") or {} if isinstance(payload, dict) else {}
    duration_value = stream.get("duration") or format_data.get("duration") or 0

    try:
        duration = max(0.0, float(duration_value or 0))
    except (TypeError, ValueError):
        duration = 0.0

    return {
        "width": int(stream.get("width") or 0),
        "height": int(stream.get("height") or 0),
        "duration": duration,
    }


def _choose_video_preview_time(video_info: dict) -> float:
    """Pick a stable preview frame timestamp away from black first frames when possible."""
    duration = float(video_info.get("duration") or 0.0)
    if duration <= 0:
        return 0.0
    if duration <= 1.0:
        return max(0.0, duration * 0.25)
    return min(1.0, duration * 0.1)


def _choose_video_ollama_timestamps(video_info: dict, frame_count: int = OLLAMA_VIDEO_FRAME_COUNT) -> list[float]:
    """Pick representative timestamps for sending a video as sampled frames to a vision model."""
    duration = max(0.0, float(video_info.get("duration") or 0.0))
    if duration <= 0.0 or frame_count <= 1:
        return [round(_choose_video_preview_time(video_info), 3)]

    start_time = min(max(duration * 0.12, 0.0), max(0.0, duration - 0.15))
    end_time = max(start_time, duration * 0.88)
    if end_time - start_time < 0.25:
        return [round(_choose_video_preview_time(video_info), 3)]

    timestamps: list[float] = []
    for index in range(frame_count):
        ratio = index / max(1, frame_count - 1)
        timestamps.append(round(start_time + ((end_time - start_time) * ratio), 3))
    return list(dict.fromkeys(timestamps))


def _format_ffmpeg_seconds(seconds: float) -> str:
    """Format seconds for ffmpeg CLI arguments."""
    return f"{max(0.0, float(seconds or 0.0)):.3f}"


def _make_scale_pad_filter(width: int, height: int) -> str:
    """Build a scale+pad filter that preserves aspect ratio."""
    return (
        f"scale={width}:{height}:force_original_aspect_ratio=decrease,"
        f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2:black"
    )


def _generate_video_thumbnail(
    filepath: str,
    size: int,
    ffmpeg_path: str = "ffmpeg",
    ffprobe_path: str = "ffprobe",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
) -> bytes:
    """Generate a JPEG thumbnail for a video by extracting a representative frame."""
    video_info = _probe_video_info(filepath, ffprobe_path)
    timestamp = _choose_video_preview_time(video_info)
    vf = f"scale={size}:{size}:force_original_aspect_ratio=decrease"
    command = [
        ffmpeg_path,
        "-hide_banner",
        "-loglevel",
        "error",
        *_build_ffmpeg_decode_args(hwaccel_mode),
        *(["-threads", str(max(1, int(thread_count or 0)))] if int(thread_count or 0) > 0 else []),
        "-ss",
        f"{timestamp:.3f}",
        "-i",
        filepath,
        "-frames:v",
        "1",
        "-vf",
        vf,
        "-f",
        "image2pipe",
        "-vcodec",
        "mjpeg",
        "pipe:1",
    ]
    try:
        result = subprocess.run(command, capture_output=True, check=True)
    except FileNotFoundError as exc:
        raise RuntimeError(f"ffmpeg not found: {ffmpeg_path}") from exc
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(stderr or f"ffmpeg failed for {Path(filepath).name}") from exc
    return result.stdout


def _extract_video_frame(
    filepath: str,
    time_seconds: float,
    width: int,
    height: int,
    ffmpeg_path: str = "ffmpeg",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
) -> bytes:
    """Extract a single JPEG video frame at the given timestamp."""
    command = [
        ffmpeg_path,
        "-hide_banner",
        "-loglevel",
        "error",
        *_build_ffmpeg_decode_args(hwaccel_mode),
        *(["-threads", str(max(1, int(thread_count or 0)))] if int(thread_count or 0) > 0 else []),
        "-ss",
        _format_ffmpeg_seconds(time_seconds),
        "-i",
        filepath,
        "-frames:v",
        "1",
        "-vf",
        _make_scale_pad_filter(width, height),
        "-f",
        "image2pipe",
        "-vcodec",
        "mjpeg",
        "pipe:1",
    ]
    try:
        result = subprocess.run(command, capture_output=True, check=True)
    except FileNotFoundError as exc:
        raise RuntimeError(f"ffmpeg not found: {ffmpeg_path}") from exc
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(stderr or f"ffmpeg failed for {Path(filepath).name}") from exc
    return result.stdout


def _run_ffmpeg_with_progress(command: list[str], duration_seconds: float, on_progress=None) -> None:
    """Run an ffmpeg command and emit normalized progress updates when available."""
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    try:
        if on_progress:
            on_progress(0.0, "Starting")
        if process.stdout is not None:
            for raw_line in process.stdout:
                line = raw_line.strip()
                if not line or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                if key in {"out_time_ms", "out_time_us"} and duration_seconds > 0:
                    try:
                        current_seconds = float(value) / (1000000.0 if key == "out_time_us" else 1000000.0)
                    except (TypeError, ValueError):
                        continue
                    progress = max(0.0, min(1.0, current_seconds / max(duration_seconds, 0.001)))
                    if on_progress:
                        on_progress(progress, "Processing")
                elif key == "progress" and value == "end":
                    if on_progress:
                        on_progress(1.0, "Finalizing")
        stderr_output = process.stderr.read() if process.stderr is not None else ""
        return_code = process.wait()
        if return_code != 0:
            raise RuntimeError(stderr_output.strip() or "ffmpeg command failed")
        if on_progress:
            on_progress(1.0, "Done")
    finally:
        if process.stdout is not None:
            process.stdout.close()
        if process.stderr is not None:
            process.stderr.close()


def _replace_file_from_temp(temp_output: str, final_output: str) -> None:
    """Atomically replace an output file with a rendered temporary file."""
    os.replace(temp_output, final_output)
    _clear_thumbnail_cache_for_path(final_output)


def _crop_video_file(
    filepath: str,
    crop: dict,
    output_path: str | None = None,
    ffmpeg_path: str = "ffmpeg",
    ffprobe_path: str = "ffprobe",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
    on_progress=None,
) -> dict:
    """Create a cropped copy of a video file using ffmpeg and return updated metadata."""
    video_info = _probe_video_info(filepath, ffprobe_path)
    width = int(video_info.get("width") or 0)
    height = int(video_info.get("height") or 0)
    if width <= 0 or height <= 0:
        raise RuntimeError("Could not determine video dimensions")

    normalized = _normalize_crop_rect(crop, (width, height))
    input_path = str(filepath)
    source = Path(input_path)
    ratio_label = str(normalized.get("ratio") or "crop").replace(":", "-").replace("/", "-")
    final_output = str(output_path or source.with_name(
        f"{source.stem}__crop_{normalized['w']}x{normalized['h']}_{normalized['x']}-{normalized['y']}_{ratio_label}{source.suffix or '.mp4'}"
    ))
    temp_output = Path(final_output).with_name(f"{Path(final_output).stem}.tag2-crop-temp{Path(final_output).suffix}")
    filter_expr = f"crop={normalized['w']}:{normalized['h']}:{normalized['x']}:{normalized['y']}"
    decode_args, gpu_processing_args, cpu_processing_args = _build_ffmpeg_processing_args(
        ffmpeg_path,
        str(final_output),
        thread_count=thread_count,
        hwaccel_mode=hwaccel_mode,
        allow_gpu_encode=True,
    )
    command = [
        ffmpeg_path,
        "-hide_banner",
        "-loglevel",
        "error",
        "-y",
        *decode_args,
        "-i",
        input_path,
        "-map",
        "0:v:0",
        "-map",
        "0:a?",
        "-vf",
        filter_expr,
        *(gpu_processing_args or cpu_processing_args),
        "-c:a",
        "copy",
        "-progress",
        "pipe:1",
        "-nostats",
        str(temp_output),
    ]
    fallback_command = None
    if gpu_processing_args:
        fallback_command = [
            ffmpeg_path,
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-i",
            input_path,
            "-map",
            "0:v:0",
            "-map",
            "0:a?",
            "-vf",
            filter_expr,
            *cpu_processing_args,
            "-c:a",
            "copy",
            "-progress",
            "pipe:1",
            "-nostats",
            str(temp_output),
        ]
    try:
        _run_ffmpeg_with_optional_fallback(command, float(video_info.get("duration") or 0.0), on_progress, fallback_command)
        _replace_file_from_temp(str(temp_output), final_output)
    finally:
        if temp_output.exists():
            temp_output.unlink(missing_ok=True)

    refreshed = _probe_video_info(final_output, ffprobe_path)
    return {
        "source_path": input_path,
        "output_path": final_output,
        "path": final_output,
        "width": int(refreshed.get("width") or 0),
        "height": int(refreshed.get("height") or 0),
        "duration": float(refreshed.get("duration") or 0.0),
        "crop": normalized,
    }


def _format_timestamp_label(seconds: float) -> str:
    """Format a filename-safe timestamp label."""
    total_ms = max(0, int(round(float(seconds or 0.0) * 1000)))
    hours, remainder = divmod(total_ms, 3600000)
    minutes, remainder = divmod(remainder, 60000)
    secs, millis = divmod(remainder, 1000)
    if hours > 0:
        return f"{hours:02d}-{minutes:02d}-{secs:02d}-{millis:03d}"
    return f"{minutes:02d}-{secs:02d}-{millis:03d}"


def _build_clip_output_path(filepath: str, start_seconds: float, end_seconds: float, crop: dict | None = None) -> str:
    """Build a default output path for a clipped video."""
    source = Path(filepath)
    start_label = _format_timestamp_label(start_seconds)
    end_label = _format_timestamp_label(end_seconds)
    suffix = source.suffix or ".mp4"
    crop_suffix = ""
    if isinstance(crop, dict):
        crop_suffix = (
            f"__crop_{max(1, int(round(float(crop.get('w') or 1))))}x{max(1, int(round(float(crop.get('h') or 1))))}"
            f"_{max(0, int(round(float(crop.get('x') or 0))))}-{max(0, int(round(float(crop.get('y') or 0))))}"
        )
    return str(source.with_name(f"{source.stem}__{start_label}__{end_label}{crop_suffix}{suffix}"))


def _clip_video_file(
    filepath: str,
    start_seconds: float,
    end_seconds: float,
    crop: dict | None = None,
    output_path: str | None = None,
    ffmpeg_path: str = "ffmpeg",
    ffprobe_path: str = "ffprobe",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
    on_progress=None,
) -> dict:
    """Create a clipped copy of a video file using ffmpeg."""
    video_info = _probe_video_info(filepath, ffprobe_path)
    duration = float(video_info.get("duration") or 0.0)
    start_seconds = max(0.0, float(start_seconds or 0.0))
    if duration > 0:
        start_seconds = min(start_seconds, duration)
    end_seconds = max(start_seconds, float(end_seconds or 0.0))
    if duration > 0:
        end_seconds = min(end_seconds, duration)
    clip_duration = max(0.0, end_seconds - start_seconds)
    if clip_duration <= 0:
        raise RuntimeError("Clip duration must be greater than zero")

    normalized_crop = _normalize_crop_rect(crop, (int(video_info.get("width") or 0), int(video_info.get("height") or 0))) if crop else None
    final_output = str(output_path or _build_clip_output_path(filepath, start_seconds, end_seconds, normalized_crop))
    temp_output = str(Path(final_output).with_name(f"{Path(final_output).stem}.tag2-clip-temp{Path(final_output).suffix}"))
    decode_args, gpu_processing_args, cpu_processing_args = _build_ffmpeg_processing_args(
        ffmpeg_path,
        final_output,
        thread_count=thread_count,
        hwaccel_mode=hwaccel_mode,
        allow_gpu_encode=True,
    )
    command = [
        ffmpeg_path,
        "-hide_banner",
        "-loglevel",
        "error",
        "-y",
        *decode_args,
        "-ss",
        _format_ffmpeg_seconds(start_seconds),
        "-i",
        filepath,
        "-t",
        _format_ffmpeg_seconds(clip_duration),
        "-map",
        "0:v:0",
        "-map",
        "0:a?",
    ]
    if normalized_crop:
        command.extend([
            "-vf",
            f"crop={normalized_crop['w']}:{normalized_crop['h']}:{normalized_crop['x']}:{normalized_crop['y']}",
        ])
    command.extend([
        *(gpu_processing_args or cpu_processing_args),
        "-c:a",
        "copy",
        "-progress",
        "pipe:1",
        "-nostats",
        temp_output,
    ])
    fallback_command = None
    if gpu_processing_args:
        fallback_command = [
            ffmpeg_path,
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-ss",
            _format_ffmpeg_seconds(start_seconds),
            "-i",
            filepath,
            "-t",
            _format_ffmpeg_seconds(clip_duration),
            "-map",
            "0:v:0",
            "-map",
            "0:a?",
        ]
        if normalized_crop:
            fallback_command.extend([
                "-vf",
                f"crop={normalized_crop['w']}:{normalized_crop['h']}:{normalized_crop['x']}:{normalized_crop['y']}",
            ])
        fallback_command.extend([
            *cpu_processing_args,
            "-c:a",
            "copy",
            "-progress",
            "pipe:1",
            "-nostats",
            temp_output,
        ])
    try:
        _run_ffmpeg_with_optional_fallback(command, clip_duration, on_progress, fallback_command)
        _replace_file_from_temp(temp_output, final_output)
    finally:
        temp_candidate = Path(temp_output)
        if temp_candidate.exists():
            temp_candidate.unlink(missing_ok=True)

    return {
        "source_path": filepath,
        "output_path": final_output,
        "start_seconds": start_seconds,
        "end_seconds": end_seconds,
        "duration": clip_duration,
        "crop": normalized_crop,
    }


def _get_thumbnail(
    filepath: str,
    size: int,
    crop: dict | None = None,
    ffmpeg_path: str = "ffmpeg",
    ffprobe_path: str = "ffprobe",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
) -> bytes:
    """Get a thumbnail from cache or generate it."""
    mtime = os.path.getmtime(filepath)
    crop_key = None if not crop else (crop.get("x"), crop.get("y"), crop.get("w"), crop.get("h"), crop.get("ratio"))
    key = (filepath, mtime, size, crop_key)
    if key not in thumbnail_cache:
        if _is_video_path(filepath):
            thumbnail_cache[key] = _generate_video_thumbnail(filepath, size, ffmpeg_path, ffprobe_path, thread_count, hwaccel_mode)
        else:
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
    if not _is_image_path(filepath):
        raise ValueError("Only image files can be encoded for Ollama vision models")
    data = _get_thumbnail(filepath, PREVIEW_MAX_SIZE)
    if not data:
        data = Path(filepath).read_bytes()
    return base64.b64encode(data).decode("ascii")


def _encode_media_for_ollama(
    filepath: str,
    ffmpeg_path: str = "ffmpeg",
    ffprobe_path: str = "ffprobe",
    thread_count: int = 0,
    hwaccel_mode: str | None = "auto",
) -> list[str]:
    """Encode supported media as one or more base64 JPEG inputs for Ollama vision models."""
    if _is_image_path(filepath):
        return [_encode_image_for_ollama(filepath)]
    if not _is_video_path(filepath):
        raise ValueError("Only image and video files can be encoded for Ollama vision models")

    video_info = _probe_video_info(filepath, ffprobe_path)
    timestamps = _choose_video_ollama_timestamps(video_info)
    encoded_frames: list[str] = []
    for time_seconds in timestamps:
        frame_bytes = _extract_video_frame(
            filepath,
            time_seconds,
            OLLAMA_VIDEO_FRAME_SIZE,
            OLLAMA_VIDEO_FRAME_SIZE,
            ffmpeg_path,
            thread_count,
            hwaccel_mode,
        )
        encoded_frames.append(base64.b64encode(frame_bytes).decode("ascii"))
    if not encoded_frames:
        raise RuntimeError(f"Could not extract representative frames for {Path(filepath).name}")
    return encoded_frames
