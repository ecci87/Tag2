"""Caption file read/write helpers."""

from __future__ import annotations

from pathlib import Path

from tag2_sections import (
    _caption_values,
    _is_hidden_group_caption,
    _iter_section_items,
    _normalize_enabled_captions,
    _ordered_sections_for_output,
)


def _get_caption_path(image_path: str) -> Path:
    """Get the corresponding .txt path for an image."""
    return Path(image_path).with_suffix(".txt")


def _read_caption_file(
    image_path: str,
    predefined_captions: list[str],
    section_headers: list[str] | None = None,
) -> dict:
    """Read caption file and separate predefined captions from free text."""
    caption_path = _get_caption_path(image_path)
    enabled_captions: list[str] = []
    free_lines: list[str] = []
    header_set = set(section_headers or [])

    if caption_path.exists():
        try:
            content = caption_path.read_text(encoding="utf-8")
            pred_set = set(predefined_captions)
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
                    caption_text = stripped[2:]
                    if caption_text in pred_set:
                        enabled_captions.append(caption_text)
                        continue

                if in_structured and stripped in pred_set:
                    enabled_captions.append(stripped)
                    continue

                in_structured = False
                free_lines.append(line)
        except Exception:
            pass

    return {
        "enabled_sentences": list(enabled_captions),
        "free_text": "\n".join(free_lines),
    }


def _build_caption_text(
    enabled_captions: list[str],
    free_text: str,
    sections: list[dict] | None = None,
) -> str:
    """Build the final caption text in the same format as the caption file."""
    if sections:
        enabled_captions = _normalize_enabled_captions(enabled_captions, sections)
        enabled_set = {
            caption
            for caption in enabled_captions
            if not _is_hidden_group_caption(sections, caption)
        }
        blocks: list[str] = []
        for section in _ordered_sections_for_output(sections):
            section_name = section.get("name", "")
            visible_captions: list[str] = []
            for item in _iter_section_items(section):
                if item["type"] == "caption":
                    caption = item["caption"]
                    if caption in enabled_set:
                        visible_captions.append(caption)
                    continue

                group = item["group"]
                visible_captions.extend(
                    caption for caption in _caption_values(group) if caption in enabled_set
                )
            if not visible_captions:
                continue

            lines: list[str] = []
            if section_name:
                lines.append(section_name)
            lines.extend(visible_captions)
            blocks.append("\n".join(lines))

        parts: list[str] = []
        if blocks:
            parts.append("\n\n".join(blocks))
        if free_text and free_text.strip():
            parts.append(free_text.strip())
        return "\n\n".join(parts)

    parts: list[str] = []
    if enabled_captions:
        parts.append("\n".join(enabled_captions))
    if free_text and free_text.strip():
        parts.append(free_text.strip())
    return "\n\n".join(parts)


def _write_caption_file(
    image_path: str,
    enabled_captions: list[str],
    free_text: str,
    sections: list[dict] | None = None,
):
    """Write caption file with sectioned format."""
    caption_path = _get_caption_path(image_path)
    content = _build_caption_text(enabled_captions, free_text, sections)

    if content:
        caption_path.write_text(content + "\n", encoding="utf-8")
    elif caption_path.exists():
        caption_path.write_text("", encoding="utf-8")
