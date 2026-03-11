"""Caption file read/write helpers."""

from __future__ import annotations

from pathlib import Path

from tag2_sections import (
    _is_hidden_group_sentence,
    _iter_section_items,
    _normalize_enabled_sentences,
    _ordered_sections_for_output,
)


def _get_caption_path(image_path: str) -> Path:
    """Get the corresponding .txt path for an image."""
    return Path(image_path).with_suffix(".txt")


def _read_caption_file(
    image_path: str,
    predefined_sentences: list[str],
    section_headers: list[str] | None = None,
) -> dict:
    """Read caption file and separate predefined sentences from free text."""
    caption_path = _get_caption_path(image_path)
    enabled_sentences: list[str] = []
    free_lines: list[str] = []
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

                if in_structured and stripped in header_set:
                    continue

                if in_structured and stripped.startswith("- "):
                    sentence_text = stripped[2:]
                    if sentence_text in pred_set:
                        enabled_sentences.append(sentence_text)
                        continue

                if in_structured and stripped in pred_set:
                    enabled_sentences.append(stripped)
                    continue

                in_structured = False
                free_lines.append(line)
        except Exception:
            pass

    return {
        "enabled_sentences": enabled_sentences,
        "free_text": "\n".join(free_lines),
    }


def _build_caption_text(
    enabled_sentences: list[str],
    free_text: str,
    sections: list[dict] | None = None,
) -> str:
    """Build the final caption text in the same format as the caption file."""
    if sections:
        enabled_sentences = _normalize_enabled_sentences(enabled_sentences, sections)
        enabled_set = {
            sentence
            for sentence in enabled_sentences
            if not _is_hidden_group_sentence(sections, sentence)
        }
        blocks: list[str] = []
        for section in _ordered_sections_for_output(sections):
            section_name = section.get("name", "")
            visible_sentences: list[str] = []
            for item in _iter_section_items(section):
                if item["type"] == "sentence":
                    sentence = item["sentence"]
                    if sentence in enabled_set:
                        visible_sentences.append(sentence)
                    continue

                group = item["group"]
                visible_sentences.extend(
                    sentence for sentence in group.get("sentences", []) if sentence in enabled_set
                )
            if not visible_sentences:
                continue

            lines: list[str] = []
            if section_name:
                lines.append(section_name)
            lines.extend(visible_sentences)
            blocks.append("\n".join(lines))

        parts: list[str] = []
        if blocks:
            parts.append("\n\n".join(blocks))
        if free_text and free_text.strip():
            parts.append(free_text.strip())
        return "\n\n".join(parts)

    parts: list[str] = []
    if enabled_sentences:
        parts.append("\n".join(enabled_sentences))
    if free_text and free_text.strip():
        parts.append(free_text.strip())
    return "\n\n".join(parts)


def _write_caption_file(
    image_path: str,
    enabled_sentences: list[str],
    free_text: str,
    sections: list[dict] | None = None,
):
    """Write caption file with sectioned format."""
    caption_path = _get_caption_path(image_path)
    content = _build_caption_text(enabled_sentences, free_text, sections)

    if content:
        caption_path.write_text(content + "\n", encoding="utf-8")
    elif caption_path.exists():
        caption_path.write_text("", encoding="utf-8")
