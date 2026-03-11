"""Ollama-related helpers for prompts, parsing, and caption generation."""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from urllib.parse import urlparse

from tag2_sections import (
    _apply_sentence_selection,
    _is_hidden_group_sentence,
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


def _ollama_prompt_for_sentence(sentence: str, template: str) -> str:
    """Build a strict yes/no prompt for one candidate caption."""
    return template.replace("{caption}", sentence).replace("{sentence}", sentence)


def _ollama_prompt_for_group(group_name: str, sentences: list[str], template: str) -> str:
    """Build a numbered-choice prompt for a mutually-exclusive caption group."""
    options = "\n".join(f"{index}. {sentence}" for index, sentence in enumerate(sentences, start=1))
    resolved_group_name = (group_name or "Caption group").strip() or "Caption group"
    return (
        template
        .replace("{group_name}", resolved_group_name)
        .replace("{group}", resolved_group_name)
        .replace("{options}", options)
        .replace("{captions}", options)
        .replace("{choices}", options)
        .replace("{count}", str(len(sentences)))
    )


def _ollama_prompt_for_free_text(caption_text: str, template: str) -> str:
    """Build a prompt asking for additional important image details."""
    return (
        template
        .replace("{caption_text}", caption_text)
        .replace("{current_caption}", caption_text)
        .replace("{caption}", caption_text)
    )


def _ollama_generate(host: str, payload: dict, timeout: int = 120) -> dict:
    """Call the local Ollama generate API."""
    url = f"{host.rstrip('/')}/api/generate"
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
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
    enabled_sentences: list[str],
) -> tuple[str, list[str]]:
    """Merge new free-text suggestions while avoiding duplicates."""
    existing_text = existing_free_text or ""
    existing_lines = [line.rstrip() for line in existing_text.splitlines() if line.strip()]
    known = {_normalize_caption_line(line) for line in existing_lines}
    known.update(_normalize_caption_line(sentence) for sentence in enabled_sentences)

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


def _auto_caption_sentences(
    host: str,
    model: str,
    image_path: str,
    sentences: list[str],
    *,
    encode_image_func,
    generate_func,
    prompt_template: str,
    timeout: int,
) -> tuple[list[str], list[dict]]:
    """Ask Ollama about each caption candidate and return enabled sentences."""
    image_b64 = encode_image_func(image_path)
    enabled: list[str] = []
    results: list[dict] = []

    for sentence in sentences:
        payload = {
            "model": model,
            "prompt": _ollama_prompt_for_sentence(sentence, prompt_template),
            "images": [image_b64],
            "stream": False,
            "options": {"temperature": 0},
        }
        response = generate_func(host, payload, timeout=timeout)
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


def _auto_caption_sections(
    host: str,
    model: str,
    image_path: str,
    sections: list[dict],
    *,
    encode_image_func,
    generate_func,
    prompt_template: str,
    group_prompt_template: str,
    timeout: int,
) -> tuple[list[str], list[dict]]:
    """Ask Ollama about configured captions, including exclusive groups."""
    image_b64 = encode_image_func(image_path)
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
            response = generate_func(host, payload, timeout=timeout)
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
            "prompt": _ollama_prompt_for_group(
                target.get("group_name", ""),
                group_sentences,
                group_prompt_template,
            ),
            "images": [image_b64],
            "stream": False,
            "options": {"temperature": 0},
        }
        response = generate_func(host, payload, timeout=timeout)
        raw_answer = str(response.get("response") or "").strip()
        selection_index = _parse_ollama_selection(raw_answer, group_sentences)
        selected_sentence = group_sentences[selection_index - 1] if selection_index else None
        selected_hidden = bool(selected_sentence and _is_hidden_group_sentence(sections, selected_sentence))
        enabled = [sentence for sentence in enabled if sentence not in group_sentences]
        if selected_sentence:
            enabled = _apply_sentence_selection(enabled, selected_sentence, sections, True)
        results.append({
            "type": "group",
            "group_name": target.get("group_name", ""),
            "sentences": group_sentences,
            "selected_sentence": selected_sentence,
            "selected_hidden": selected_hidden,
            "selection_index": selection_index,
            "answer": raw_answer,
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
    timeout: int,
) -> str:
    """Ask Ollama for additional free-text image details."""
    image_b64 = encode_image_func(image_path)
    payload = {
        "model": model,
        "prompt": _ollama_prompt_for_free_text(caption_text, prompt_template),
        "images": [image_b64],
        "stream": False,
        "options": {"temperature": 0},
    }
    response = generate_func(host, payload, timeout=timeout)
    return str(response.get("response") or "").strip()
