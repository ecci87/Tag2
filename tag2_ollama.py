"""Ollama-related helpers for prompts, parsing, and caption generation."""

from __future__ import annotations

import json
import re
import socket
import urllib.error
import urllib.request
from urllib.parse import urlparse

from tag2_images import _is_video_path

from tag2_sections import (
    _apply_caption_selection,
    _is_hidden_group_caption,
    _iter_caption_targets,
)


def _split_ollama_host(host: str, default_server: str, default_port: int) -> tuple[str, int]:
    """Split a full Ollama host URL into server and port."""
    raw = (host or "").strip()
    if not raw:
        return default_server, default_port
    if "://" not in raw:
        raw = f"http://{raw}"
    parsed = urlparse(raw)
    server = parsed.hostname or default_server
    port = parsed.port or default_port
    return server, port


def _compose_ollama_host(
    server: str | None,
    port: int | str | None,
    default_server: str,
    default_port: int,
) -> str:
    """Compose a normalized Ollama host URL from server and port."""
    host_server = str(server or default_server).strip() or default_server
    try:
        host_port = int(port if port is not None else default_port)
    except (TypeError, ValueError):
        host_port = default_port
    return f"http://{host_server}:{host_port}"


def _get_ollama_host(cfg: dict, default_server: str, default_port: int) -> str:
    """Get the configured Ollama host."""
    server = cfg.get("ollama_server")
    port = cfg.get("ollama_port")
    if server is not None or port is not None:
        return _compose_ollama_host(server, port, default_server, default_port)
    return str(cfg.get("ollama_host") or f"http://{default_server}:{default_port}").rstrip("/")


def _get_ollama_server(cfg: dict, default_server: str, default_port: int) -> str:
    """Get configured Ollama server hostname."""
    server = str(cfg.get("ollama_server") or "").strip()
    if server:
        return server
    return _split_ollama_host(_get_ollama_host(cfg, default_server, default_port), default_server, default_port)[0]


def _get_ollama_port(cfg: dict, default_server: str, default_port: int) -> int:
    """Get configured Ollama server port."""
    port = cfg.get("ollama_port")
    try:
        if port is not None:
            return int(port)
    except (TypeError, ValueError):
        pass
    return _split_ollama_host(_get_ollama_host(cfg, default_server, default_port), default_server, default_port)[1]


def _get_ollama_model(cfg: dict, default_model: str) -> str:
    """Get the configured Ollama model name."""
    return str(cfg.get("ollama_model") or default_model).strip()


def _get_ollama_timeout_seconds(cfg: dict, default_timeout: int) -> int:
    """Get the configured Ollama request timeout."""
    try:
        timeout = int(cfg.get("ollama_timeout_seconds", default_timeout))
    except (TypeError, ValueError):
        timeout = default_timeout
    return max(1, timeout)


def _get_ollama_max_output_tokens(cfg: dict, default_max_output_tokens: int) -> int:
    """Get the configured Ollama output token cap."""
    try:
        max_output_tokens = int(cfg.get("ollama_max_output_tokens", default_max_output_tokens))
    except (TypeError, ValueError):
        max_output_tokens = default_max_output_tokens
    return max(1, max_output_tokens)


def _get_ollama_prompt_template(cfg: dict, default_template: str) -> str:
    """Get the configured Ollama prompt template."""
    return str(cfg.get("ollama_prompt_template") or default_template)


def _get_ollama_group_prompt_template(cfg: dict, default_template: str) -> str:
    """Get the configured grouped-caption prompt template."""
    return str(cfg.get("ollama_group_prompt_template") or default_template)


def _get_ollama_enable_free_text(cfg: dict) -> bool:
    """Get whether the free-text enhancement step is enabled."""
    return bool(cfg.get("ollama_enable_free_text", True))


def _get_ollama_free_text_prompt_template(cfg: dict, default_template: str) -> str:
    """Get the configured free-text prompt template."""
    return str(cfg.get("ollama_free_text_prompt_template") or default_template)


def _get_ollama_region_system_prompt_template(cfg: dict, default_template: str) -> str:
    """Get the configured selected-region system prompt template."""
    return str(cfg.get("ollama_region_system_prompt_template") or default_template)


def _ollama_prompt_for_caption(caption: str, template: str) -> str:
    """Build a strict yes/no prompt for one candidate caption."""
    return template.replace("{caption}", caption).replace("{sentence}", caption)


def _ollama_prompt_for_group(group_name: str, captions: list[str], template: str) -> str:
    """Build a numbered-choice prompt for a mutually-exclusive caption group."""
    options = "\n".join(f"{index}. {caption}" for index, caption in enumerate(captions, start=1))
    resolved_group_name = (group_name or "Caption group").strip() or "Caption group"
    return (
        template
        .replace("{group_name}", resolved_group_name)
        .replace("{group}", resolved_group_name)
        .replace("{options}", options)
        .replace("{captions}", options)
        .replace("{choices}", options)
        .replace("{count}", str(len(captions)))
    )


def _ollama_prompt_for_free_text(caption_text: str, template: str) -> str:
    """Build a prompt asking for additional important image details."""
    return (
        template
        .replace("{caption_text}", caption_text)
        .replace("{current_caption}", caption_text)
        .replace("{caption}", caption_text)
    )


def _normalize_ollama_images(encoded_media: str | list[str] | tuple[str, ...]) -> list[str]:
    """Normalize one or more encoded vision inputs into the Ollama images payload shape."""
    if isinstance(encoded_media, str):
        return [encoded_media] if encoded_media else []
    return [item for item in list(encoded_media or []) if item]


def _apply_media_prompt_context(prompt: str, media_path: str) -> str:
    """Inject extra guidance when a prompt is evaluating sampled video frames."""
    if not _is_video_path(media_path):
        return prompt
    prefix = (
        "The attached visual input contains multiple representative frames sampled from one video clip. "
        "Base your answer on the overall clip content across all provided frames.\n\n"
    )
    return prefix + prompt


def _build_ollama_generate_payload(
    model: str,
    prompt: str,
    images: list[str],
    *,
    system: str | None = None,
    max_output_tokens: int | None = None,
    temperature: float = 0,
) -> dict:
    """Build a generate payload with conservative decoding defaults."""
    options = {"temperature": temperature}
    try:
        if max_output_tokens is not None:
            options["num_predict"] = max(1, int(max_output_tokens))
    except (TypeError, ValueError):
        pass
    payload = {
        "model": model,
        "prompt": prompt,
        "images": images,
        "stream": True,
        "options": options,
    }
    if system:
        payload["system"] = str(system)
    return payload


def _ollama_response_meta(response: dict) -> dict:
    """Normalize completion metadata for downstream logging and UI state."""
    incomplete = bool(response.get("incomplete") or response.get("done") is False)
    done_reason = str(response.get("done_reason") or "").strip()
    if incomplete and not done_reason:
        done_reason = "timeout"
    return {
        "answer_incomplete": incomplete,
        "answer_done_reason": done_reason,
    }


def _is_timeout_error(reason: object) -> bool:
    """Return whether an exception reason represents a network timeout."""
    if isinstance(reason, (TimeoutError, socket.timeout)):
        return True
    return "timed out" in str(reason or "").lower()


def _ollama_generate(host: str, payload: dict, timeout: int = 120) -> dict:
    """Call the local Ollama generate API and keep partial streamed output on timeout."""
    url = f"{host.rstrip('/')}/api/generate"
    request_payload = dict(payload or {})
    request_payload["stream"] = True
    request = urllib.request.Request(
        url,
        data=json.dumps(request_payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    response_chunks: list[str] = []
    last_event: dict = {}

    def _build_response(*, incomplete: bool = False, done_reason: str | None = None) -> dict:
        merged = dict(last_event)
        merged["response"] = "".join(response_chunks).strip()
        if incomplete:
            merged["done"] = False
            merged["incomplete"] = True
            merged["done_reason"] = done_reason or merged.get("done_reason") or "timeout"
        elif "done" not in merged:
            merged["done"] = True
        return merged

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            for raw_line in response:
                line = raw_line.decode("utf-8", errors="replace").strip()
                if not line:
                    continue
                try:
                    event = json.loads(line)
                except json.JSONDecodeError as exc:
                    raise RuntimeError("invalid response from Ollama") from exc
                if event.get("error"):
                    raise RuntimeError(str(event.get("error")))
                last_event = event
                chunk = event.get("response")
                if chunk:
                    response_chunks.append(str(chunk))
            return _build_response()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(detail or str(exc)) from exc
    except (TimeoutError, socket.timeout) as exc:
        partial = "".join(response_chunks).strip()
        if partial:
            return _build_response(incomplete=True, done_reason="timeout")
        raise RuntimeError(f"request timed out after {timeout} seconds") from exc
    except urllib.error.URLError as exc:
        if _is_timeout_error(exc.reason or exc):
            partial = "".join(response_chunks).strip()
            if partial:
                return _build_response(incomplete=True, done_reason="timeout")
        raise RuntimeError(str(exc.reason or exc)) from exc


def _parse_ollama_yes_no(response_text: str) -> bool:
    """Parse a yes/no response from Ollama conservatively."""
    text = (response_text or "").strip().upper()
    if text.startswith("YES"):
        return True
    if text.startswith("NO"):
        return False
    tokens = [token.strip(" .,!?:;()[]{}\"'") for token in text.split()]
    if "YES" in tokens:
        return True
    if "NO" in tokens:
        return False
    return False


def _normalize_caption_line(text: str) -> str:
    """Normalize a caption or free-text line for duplicate detection."""
    stripped = (text or "").strip()
    stripped = re.sub(r"^[\-•*\d\.)\s]+", "", stripped)
    stripped = re.sub(r"\s+", " ", stripped)
    return stripped.casefold()


def _parse_ollama_selection(response_text: str, captions: list[str]) -> int | None:
    """Parse a 1-based choice from an Ollama response."""
    text = (response_text or "").strip()
    if not text:
        return None

    for match in re.findall(r"\d+", text):
        value = int(match)
        if 1 <= value <= len(captions):
            return value

    normalized = _normalize_caption_line(text)
    for index, caption in enumerate(captions, start=1):
        caption_normalized = _normalize_caption_line(caption)
        if normalized == caption_normalized:
            return index
    for index, caption in enumerate(captions, start=1):
        caption_normalized = _normalize_caption_line(caption)
        if caption_normalized and caption_normalized in normalized:
            return index
    return None


def _extract_free_text_lines(response_text: str) -> list[str]:
    """Extract meaningful suggestion lines from an Ollama free-text response."""
    text = (response_text or "").strip()
    if not text or text.upper() == "NONE":
        return []

    lines: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.upper() == "NONE":
            continue
        line = re.sub(r"^[\-•*\d\.)\s]+", "", line).strip()
        if line:
            lines.append(line)
    return lines


def _merge_free_text(
    existing_free_text: str,
    suggested_text: str,
    enabled_captions: list[str],
) -> tuple[str, list[str]]:
    """Merge new free-text suggestions while avoiding duplicates."""
    existing_text = existing_free_text or ""
    existing_lines = [line.rstrip() for line in existing_text.splitlines() if line.strip()]
    known = {_normalize_caption_line(line) for line in existing_lines}
    known.update(_normalize_caption_line(caption) for caption in enabled_captions)

    added_lines: list[str] = []
    for line in _extract_free_text_lines(suggested_text):
        normalized = _normalize_caption_line(line)
        if not normalized or normalized in known:
            continue
        known.add(normalized)
        added_lines.append(line)

    merged_lines = list(existing_lines)
    merged_lines.extend(added_lines)
    return "\n".join(merged_lines), added_lines


def _auto_caption_captions(
    host: str,
    model: str,
    image_path: str,
    captions: list[str],
    *,
    encode_image_func,
    generate_func,
    prompt_template: str,
    timeout: int,
    max_output_tokens: int | None = None,
) -> tuple[list[str], list[dict]]:
    """Ask Ollama about each caption candidate and return enabled captions."""
    image_payload = _normalize_ollama_images(encode_image_func(image_path))
    enabled: list[str] = []
    results: list[dict] = []

    for caption in captions:
        payload = _build_ollama_generate_payload(
            model,
            _apply_media_prompt_context(_ollama_prompt_for_caption(caption, prompt_template), image_path),
            image_payload,
            max_output_tokens=max_output_tokens,
        )
        response = generate_func(host, payload, timeout=timeout)
        raw_answer = str(response.get("response") or "").strip()
        is_match = _parse_ollama_yes_no(raw_answer)
        results.append({
            "caption": caption,
            "sentence": caption,
            "enabled": is_match,
            "answer": raw_answer,
            **_ollama_response_meta(response),
        })
        if is_match:
            enabled.append(caption)

    return enabled, results


def _auto_caption_sections(
    host: str,
    model: str,
    image_path: str,
    sections: list[dict],
    *,
    initial_enabled_captions: list[str] | None = None,
    encode_image_func,
    generate_func,
    prompt_template: str,
    group_prompt_template: str,
    timeout: int,
    max_output_tokens: int | None = None,
) -> tuple[list[str], list[dict]]:
    """Ask Ollama about configured captions, including exclusive groups."""
    image_payload = _normalize_ollama_images(encode_image_func(image_path))
    enabled: list[str] = list(initial_enabled_captions or [])
    results: list[dict] = []

    for target in _iter_caption_targets(sections):
        if target["type"] == "caption":
            caption = target["caption"]
            if target.get("skip_auto_caption"):
                results.append({
                    "type": "sentence",
                    "caption": caption,
                    "sentence": caption,
                    "enabled": caption in enabled,
                    "skipped": True,
                    "skip_reason": target.get("skip_reason", ""),
                    "answer": "",
                })
                continue
            payload = _build_ollama_generate_payload(
                model,
                _apply_media_prompt_context(_ollama_prompt_for_caption(caption, prompt_template), image_path),
                image_payload,
                max_output_tokens=max_output_tokens,
            )
            response = generate_func(host, payload, timeout=timeout)
            raw_answer = str(response.get("response") or "").strip()
            is_match = _parse_ollama_yes_no(raw_answer)
            enabled = _apply_caption_selection(enabled, caption, sections, is_match)
            results.append({
                "type": "sentence",
                "caption": caption,
                "sentence": caption,
                "enabled": is_match,
                "answer": raw_answer,
                **_ollama_response_meta(response),
            })
            continue

        group_captions = target["captions"]
        if target.get("skip_auto_caption"):
            selected_caption = next((caption for caption in group_captions if caption in enabled), None)
            selected_hidden = bool(selected_caption and _is_hidden_group_caption(sections, selected_caption))
            results.append({
                "type": "group",
                "group_name": target.get("group_name", ""),
                "captions": group_captions,
                "sentences": list(group_captions),
                "selected_caption": selected_caption,
                "selected_sentence": selected_caption,
                "selected_hidden": selected_hidden,
                "selection_index": group_captions.index(selected_caption) + 1 if selected_caption in group_captions else None,
                "skipped": True,
                "skip_reason": target.get("skip_reason", ""),
                "skip_captions": list(target.get("skip_captions") or target.get("skip_sentences") or []),
                "answer": "",
            })
            continue
        payload = _build_ollama_generate_payload(
            model,
            _apply_media_prompt_context(
                _ollama_prompt_for_group(
                    target.get("group_name", ""),
                    group_captions,
                    group_prompt_template,
                ),
                image_path,
            ),
            image_payload,
            max_output_tokens=max_output_tokens,
        )
        response = generate_func(host, payload, timeout=timeout)
        raw_answer = str(response.get("response") or "").strip()
        selection_index = _parse_ollama_selection(raw_answer, group_captions)
        selected_caption = group_captions[selection_index - 1] if selection_index else None
        selected_hidden = bool(selected_caption and _is_hidden_group_caption(sections, selected_caption))
        enabled = [caption for caption in enabled if caption not in group_captions]
        if selected_caption:
            enabled = _apply_caption_selection(enabled, selected_caption, sections, True)
        results.append({
            "type": "group",
            "group_name": target.get("group_name", ""),
            "captions": group_captions,
            "sentences": list(group_captions),
            "selected_caption": selected_caption,
            "selected_sentence": selected_caption,
            "selected_hidden": selected_hidden,
            "selection_index": selection_index,
            "answer": raw_answer,
            **_ollama_response_meta(response),
        })

    return enabled, results


def _suggest_free_text(
    host: str,
    model: str,
    image_path: str,
    caption_text: str,
    *,
    encode_image_func,
    generate_func,
    prompt_template: str,
    system_prompt: str | None = None,
    timeout: int,
    max_output_tokens: int | None = None,
) -> str:
    """Ask Ollama for additional free-text image details."""
    image_payload = _normalize_ollama_images(encode_image_func(image_path))
    payload = _build_ollama_generate_payload(
        model,
        _apply_media_prompt_context(_ollama_prompt_for_free_text(caption_text, prompt_template), image_path),
        image_payload,
        system=system_prompt,
        max_output_tokens=max_output_tokens,
    )
    response = generate_func(host, payload, timeout=timeout)
    return str(response.get("response") or "").strip()
