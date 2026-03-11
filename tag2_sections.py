"""Helpers for folder section and caption option management."""

from __future__ import annotations

import os


def _get_folder_sections(cfg: dict, folder: str) -> list[dict]:
    """Get sections for a folder, falling back to default sections."""

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
        sentences = _normalize_sentences(group.get("sentences", []))
        hidden_sentences = [
            sentence
            for sentence in _normalize_sentences(group.get("hidden_sentences", []))
            if sentence in sentences
        ]
        return {
            "id": str(group.get("id") or "").strip(),
            "name": str(group.get("name") or "").strip(),
            "sentences": sentences,
            "hidden_sentences": hidden_sentences,
        }

    def _normalize_section(section: dict | None) -> dict:
        section = dict(section or {})
        groups = []
        used_group_ids: set[str] = set()
        for index, group in enumerate(section.get("groups", []) or []):
            normalized_group = _normalize_group(group)
            if normalized_group["name"] or normalized_group["sentences"]:
                base_group_id = normalized_group.get("id") or f"group-{index + 1}"
                group_id = base_group_id
                suffix = 2
                while group_id in used_group_ids:
                    group_id = f"{base_group_id}-{suffix}"
                    suffix += 1
                normalized_group["id"] = group_id
                used_group_ids.add(group_id)
                groups.append(normalized_group)
        sentences = _normalize_sentences(section.get("sentences", []))
        sentence_set = set(sentences)
        group_ids = {group["id"] for group in groups}
        item_order: list[dict] = []
        seen_sentences: set[str] = set()
        seen_groups: set[str] = set()
        for raw_item in section.get("item_order", []) or []:
            if not isinstance(raw_item, dict):
                continue
            item_type = str(raw_item.get("type") or "").strip().lower()
            if item_type == "sentence":
                sentence = str(raw_item.get("sentence") or "").strip()
                if sentence and sentence in sentence_set and sentence not in seen_sentences:
                    item_order.append({"type": "sentence", "sentence": sentence})
                    seen_sentences.add(sentence)
            elif item_type == "group":
                group_id = str(raw_item.get("group_id") or "").strip()
                if group_id and group_id in group_ids and group_id not in seen_groups:
                    item_order.append({"type": "group", "group_id": group_id})
                    seen_groups.add(group_id)
        for sentence in sentences:
            if sentence not in seen_sentences:
                item_order.append({"type": "sentence", "sentence": sentence})
        for group in groups:
            group_id = group.get("id", "")
            if group_id not in seen_groups:
                item_order.append({"type": "group", "group_id": group_id})
        return {
            "name": str(section.get("name") or "").strip(),
            "sentences": sentences,
            "groups": groups,
            "item_order": item_order,
        }

    folder_key = os.path.normpath(folder)
    folder_cfg = cfg.get("folders", {}).get(folder_key, None)
    if folder_cfg is not None:
        if "sections" in folder_cfg:
            return [_normalize_section(section) for section in folder_cfg["sections"]]
        if "sentences" in folder_cfg:
            return [_normalize_section({"name": "", "sentences": list(folder_cfg["sentences"]), "groups": []})]

    default_sections = cfg.get("default_sections", [])
    if default_sections:
        return [_normalize_section(section) for section in default_sections]

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
    cfg["folders"][folder_key]["sections"] = _get_folder_sections(
        {"folders": {folder_key: {"sections": sections}}},
        folder_key,
    )
    cfg["folders"][folder_key].pop("sentences", None)


def _all_sentences_from_sections(sections: list[dict]) -> list[str]:
    """Flatten sections into a single list of all predefined sentences."""
    result: list[str] = []
    for section in sections:
        result.extend(section.get("sentences", []))
        for group in section.get("groups", []) or []:
            result.extend(group.get("sentences", []))
    return result


def _group_hidden_sentences(group: dict | None) -> list[str]:
    return [sentence for sentence in (group or {}).get("hidden_sentences", []) or [] if sentence]


def _is_hidden_group_sentence(sections: list[dict], sentence: str) -> bool:
    for section in sections:
        for group in section.get("groups", []) or []:
            if sentence in _group_hidden_sentences(group):
                return True
    return False


def _is_general_section_name(name: str | None) -> bool:
    normalized = str(name or "").strip()
    if not normalized:
        return True
    normalized = normalized.lstrip("#").strip().lower()
    return normalized == "general"


def _ordered_sections_for_output(sections: list[dict]) -> list[dict]:
    """Keep section order stable while pinning the general section to the top."""
    indexed_sections = list(enumerate(sections))
    indexed_sections.sort(
        key=lambda item: (0 if _is_general_section_name(item[1].get("name")) else 1, item[0])
    )
    return [section for _, section in indexed_sections]


def _all_headers_from_sections(sections: list[dict]) -> list[str]:
    """Collect structured section header lines used in caption files."""
    headers: list[str] = []
    seen: set[str] = set()
    for section in _ordered_sections_for_output(sections):
        section_name = str(section.get("name") or "").strip()
        aliases = [section_name] if section_name else []
        if _is_general_section_name(section_name):
            aliases.extend(["(General)", "## General"])
        for alias in aliases:
            if not alias or alias in seen:
                continue
            seen.add(alias)
            headers.append(alias)
    return headers


def _iter_section_items(section: dict):
    """Yield mixed section items in configured display order."""
    groups = list(section.get("groups", []) or [])
    groups_by_id = {
        str(group.get("id") or "").strip(): group
        for group in groups
        if str(group.get("id") or "").strip()
    }
    yielded_sentences: set[str] = set()
    yielded_groups: set[str] = set()

    for raw_item in section.get("item_order", []) or []:
        if not isinstance(raw_item, dict):
            continue
        item_type = str(raw_item.get("type") or "").strip().lower()
        if item_type == "sentence":
            sentence = str(raw_item.get("sentence") or "").strip()
            if sentence and sentence in (section.get("sentences", []) or []) and sentence not in yielded_sentences:
                yielded_sentences.add(sentence)
                yield {"type": "sentence", "sentence": sentence}
        elif item_type == "group":
            group_id = str(raw_item.get("group_id") or "").strip()
            group = groups_by_id.get(group_id)
            if group and group_id not in yielded_groups:
                yielded_groups.add(group_id)
                yield {"type": "group", "group": group}

    for sentence in section.get("sentences", []) or []:
        if sentence not in yielded_sentences:
            yield {"type": "sentence", "sentence": sentence}
    for group in groups:
        group_id = str(group.get("id") or "").strip()
        if group_id and group_id not in yielded_groups:
            yield {"type": "group", "group": group}


def _iter_caption_targets(sections: list[dict]):
    """Yield caption targets in display order."""
    for section in sections:
        for item in _iter_section_items(section):
            if item["type"] == "sentence":
                yield {
                    "type": "sentence",
                    "section_name": section.get("name", ""),
                    "sentence": item["sentence"],
                }
                continue

            group = item["group"]
            group_sentences = list(group.get("sentences", []))
            if not group_sentences:
                continue
            yield {
                "type": "group",
                "section_index": None,
                "group_index": None,
                "section_name": section.get("name", ""),
                "group_name": group.get("name", ""),
                "sentences": group_sentences,
            }


def _iter_caption_targets_with_indices(sections: list[dict]):
    """Yield caption targets in display order together with stable indices."""
    for section_index, section in enumerate(sections):
        group_index_by_id = {
            str(group.get("id") or "").strip(): group_index
            for group_index, group in enumerate(section.get("groups", []) or [])
        }
        for item in _iter_section_items(section):
            if item["type"] == "sentence":
                yield {
                    "type": "sentence",
                    "section_index": section_index,
                    "group_index": None,
                    "section_name": section.get("name", ""),
                    "sentence": item["sentence"],
                }
                continue

            group = item["group"]
            group_sentences = list(group.get("sentences", []))
            group_id = str(group.get("id") or "").strip()
            group_index = group_index_by_id.get(group_id)
            if not group_sentences or group_index is None:
                continue
            yield {
                "type": "group",
                "section_index": section_index,
                "group_index": group_index,
                "section_name": section.get("name", ""),
                "group_name": group.get("name", ""),
                "sentences": group_sentences,
            }


def _get_group_target(sections: list[dict], section_index: int | None, group_index: int | None) -> dict | None:
    """Return a specific group target by section and group indices."""
    if section_index is None or group_index is None:
        return None
    if section_index < 0 or section_index >= len(sections):
        return None
    groups = sections[section_index].get("groups", []) or []
    if group_index < 0 or group_index >= len(groups):
        return None
    group = groups[group_index]
    group_sentences = list(group.get("sentences", []))
    if not group_sentences:
        return None
    return {
        "type": "group",
        "section_index": section_index,
        "group_index": group_index,
        "section_name": sections[section_index].get("name", ""),
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


def _apply_sentence_selection(
    enabled_sentences: list[str],
    sentence: str,
    sections: list[dict],
    should_enable: bool,
) -> list[str]:
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
    """Rename a configured sentence inside sections and groups."""
    renamed = False
    for section in sections:
        section_item_order = []
        for item in section.get("item_order", []) or []:
            if not isinstance(item, dict):
                continue
            if str(item.get("type") or "").strip().lower() == "sentence":
                sentence = str(item.get("sentence") or "").strip()
                section_item_order.append({
                    "type": "sentence",
                    "sentence": new_sentence if sentence == old_sentence else sentence,
                })
            elif str(item.get("type") or "").strip().lower() == "group":
                section_item_order.append({
                    "type": "group",
                    "group_id": str(item.get("group_id") or "").strip(),
                })
        if section_item_order:
            section["item_order"] = section_item_order
        section_sentences = list(section.get("sentences", []))
        if old_sentence in section_sentences:
            renamed = True
        section["sentences"] = [
            new_sentence if sentence == old_sentence else sentence
            for sentence in section_sentences
        ]
        for group in section.get("groups", []) or []:
            group_sentences = list(group.get("sentences", []))
            if old_sentence in group_sentences:
                renamed = True
            group["sentences"] = [
                new_sentence if sentence == old_sentence else sentence
                for sentence in group_sentences
            ]
            group_hidden_sentences = _group_hidden_sentences(group)
            group["hidden_sentences"] = [
                new_sentence if sentence == old_sentence else sentence
                for sentence in group_hidden_sentences
                if (new_sentence if sentence == old_sentence else sentence) in group["sentences"]
            ]
    return renamed


def _get_crop_aspect_ratios(cfg: dict, default_crop_aspect_ratios: list[str]) -> list[str]:
    """Get the configured list of allowed crop aspect ratios."""
    ratios = cfg.get("crop_aspect_ratios", default_crop_aspect_ratios)
    if not isinstance(ratios, list):
        return list(default_crop_aspect_ratios)
    cleaned = [str(ratio).strip() for ratio in ratios if str(ratio).strip()]
    return cleaned or list(default_crop_aspect_ratios)
