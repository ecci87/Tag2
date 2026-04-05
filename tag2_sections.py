"""Helpers for folder section and caption option management."""

from __future__ import annotations

import os


def _normalize_caption_entries(values) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in values or []:
        text = str(raw).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        cleaned.append(text)
    return cleaned


def _caption_values(container: dict | None) -> list[str]:
    container = dict(container or {})
    return _normalize_caption_entries(container.get("captions", container.get("sentences", [])))


def _hidden_caption_values(group: dict | None) -> list[str]:
    group = dict(group or {})
    return _normalize_caption_entries(
        group.get("hidden_captions", group.get("hidden_sentences", []))
    )


def _normalize_item_type(raw_item: dict | None) -> str:
    item_type = str((raw_item or {}).get("type") or "").strip().lower()
    return "caption" if item_type == "sentence" else item_type


def _item_caption_value(raw_item: dict | None) -> str:
    raw_item = dict(raw_item or {})
    return str(raw_item.get("caption") or raw_item.get("sentence") or "").strip()


def _get_folder_sections(cfg: dict, folder: str) -> list[dict]:
    """Get sections for a folder, falling back to default sections."""

    def _normalize_group(group: dict | None) -> dict:
        group = dict(group or {})
        captions = _caption_values(group)
        hidden_captions = [
            caption
            for caption in _hidden_caption_values(group)
            if caption in captions
        ]
        return {
            "id": str(group.get("id") or "").strip(),
            "name": str(group.get("name") or "").strip(),
            "captions": captions,
            "sentences": list(captions),
            "hidden_captions": hidden_captions,
            "hidden_sentences": list(hidden_captions),
        }

    def _normalize_section(section: dict | None) -> dict:
        section = dict(section or {})
        groups: list[dict] = []
        used_group_ids: set[str] = set()
        for index, group in enumerate(section.get("groups", []) or []):
            normalized_group = _normalize_group(group)
            if normalized_group["name"] or normalized_group["captions"]:
                base_group_id = normalized_group.get("id") or f"group-{index + 1}"
                group_id = base_group_id
                suffix = 2
                while group_id in used_group_ids:
                    group_id = f"{base_group_id}-{suffix}"
                    suffix += 1
                normalized_group["id"] = group_id
                used_group_ids.add(group_id)
                groups.append(normalized_group)

        captions = _caption_values(section)
        caption_set = set(captions)
        group_ids = {group["id"] for group in groups}
        item_order: list[dict] = []
        seen_captions: set[str] = set()
        seen_groups: set[str] = set()
        for raw_item in section.get("item_order", []) or []:
            if not isinstance(raw_item, dict):
                continue
            item_type = _normalize_item_type(raw_item)
            if item_type == "caption":
                caption = _item_caption_value(raw_item)
                if caption and caption in caption_set and caption not in seen_captions:
                    item_order.append({"type": "sentence", "sentence": caption})
                    seen_captions.add(caption)
            elif item_type == "group":
                group_id = str(raw_item.get("group_id") or "").strip()
                if group_id and group_id in group_ids and group_id not in seen_groups:
                    item_order.append({"type": "group", "group_id": group_id})
                    seen_groups.add(group_id)
        for caption in captions:
            if caption not in seen_captions:
                item_order.append({"type": "sentence", "sentence": caption})
        for group in groups:
            group_id = group.get("id", "")
            if group_id not in seen_groups:
                item_order.append({"type": "group", "group_id": group_id})
        return {
            "name": str(section.get("name") or "").strip(),
            "captions": captions,
            "sentences": list(captions),
            "groups": groups,
            "item_order": item_order,
        }

    folder_key = os.path.normpath(folder)
    folder_cfg = cfg.get("folders", {}).get(folder_key, None)
    if folder_cfg is not None:
        if "sections" in folder_cfg:
            return [_normalize_section(section) for section in folder_cfg["sections"]]
        if "captions" in folder_cfg:
            return [_normalize_section({"name": "", "captions": list(folder_cfg["captions"]), "groups": []})]
        if "sentences" in folder_cfg:
            return [_normalize_section({"name": "", "captions": list(folder_cfg["sentences"]), "groups": []})]

    default_sections = cfg.get("default_sections", [])
    if default_sections:
        return [_normalize_section(section) for section in default_sections]

    default_captions = cfg.get("default_captions", cfg.get("default_sentences", []))
    if default_captions:
        return [_normalize_section({"name": "", "captions": list(default_captions), "groups": []})]

    return [_normalize_section({"name": "", "captions": [], "groups": []})]


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
    cfg["folders"][folder_key].pop("captions", None)
    cfg["folders"][folder_key].pop("sentences", None)


def _all_captions_from_sections(sections: list[dict]) -> list[str]:
    """Flatten sections into a single list of all predefined captions."""
    result: list[str] = []
    for section in sections:
        result.extend(_caption_values(section))
        for group in section.get("groups", []) or []:
            result.extend(_caption_values(group))
    return result


def _group_hidden_captions(group: dict | None) -> list[str]:
    return [caption for caption in _hidden_caption_values(group) if caption]


def _is_hidden_group_caption(sections: list[dict], caption: str) -> bool:
    for section in sections:
        for group in section.get("groups", []) or []:
            if caption in _group_hidden_captions(group):
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
    group_keys: list[str] = []
    for index, group in enumerate(groups):
        group_id = str(group.get("id") or "").strip()
        group_keys.append(group_id or f"__index__:{index}")
    groups_by_id = {
        key: group
        for key, group in zip(group_keys, groups)
    }
    section_captions = _caption_values(section)
    yielded_captions: set[str] = set()
    yielded_groups: set[str] = set()

    for raw_item in section.get("item_order", []) or []:
        if not isinstance(raw_item, dict):
            continue
        item_type = _normalize_item_type(raw_item)
        if item_type == "caption":
            caption = _item_caption_value(raw_item)
            if caption and caption in section_captions and caption not in yielded_captions:
                yielded_captions.add(caption)
                yield {"type": "caption", "caption": caption, "sentence": caption}
        elif item_type == "group":
            group_id = str(raw_item.get("group_id") or "").strip()
            group = groups_by_id.get(group_id)
            if group and group_id not in yielded_groups:
                yielded_groups.add(group_id)
                yield {"type": "group", "group": group}

    for caption in section_captions:
        if caption not in yielded_captions:
            yield {"type": "caption", "caption": caption, "sentence": caption}
    for group_key, group in zip(group_keys, groups):
        if group_key not in yielded_groups:
            yield {"type": "group", "group": group}


def _iter_caption_targets(sections: list[dict]):
    """Yield caption targets in display order."""
    for section in sections:
        for item in _iter_section_items(section):
            if item["type"] == "caption":
                yield {
                    "type": "caption",
                    "section_name": section.get("name", ""),
                    "caption": item["caption"],
                    "sentence": item["caption"],
                }
                continue

            group = item["group"]
            group_captions = _caption_values(group)
            if not group_captions:
                continue
            yield {
                "type": "group",
                "section_index": None,
                "group_index": None,
                "section_name": section.get("name", ""),
                "group_name": group.get("name", ""),
                "captions": group_captions,
                "sentences": list(group_captions),
            }


def _ordered_captions_from_sections(sections: list[dict]) -> list[str]:
    """Flatten all configured captions in their display order."""
    ordered: list[str] = []
    for section in sections:
        for item in _iter_section_items(section):
            if item["type"] == "caption":
                ordered.append(item["caption"])
                continue
            ordered.extend(_caption_values(item["group"]))
    return ordered


def _iter_caption_targets_with_indices(sections: list[dict]):
    """Yield caption targets in display order together with stable indices."""
    for section_index, section in enumerate(sections):
        group_index_by_id = {
            str(group.get("id") or "").strip(): group_index
            for group_index, group in enumerate(section.get("groups", []) or [])
        }
        for item in _iter_section_items(section):
            if item["type"] == "caption":
                yield {
                    "type": "caption",
                    "section_index": section_index,
                    "group_index": None,
                    "section_name": section.get("name", ""),
                    "caption": item["caption"],
                    "sentence": item["caption"],
                }
                continue

            group = item["group"]
            group_captions = _caption_values(group)
            group_id = str(group.get("id") or "").strip()
            group_index = group_index_by_id.get(group_id)
            if not group_captions or group_index is None:
                continue
            yield {
                "type": "group",
                "section_index": section_index,
                "group_index": group_index,
                "section_name": section.get("name", ""),
                "group_name": group.get("name", ""),
                "captions": group_captions,
                "sentences": list(group_captions),
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
    group_captions = _caption_values(group)
    if not group_captions:
        return None
    return {
        "type": "group",
        "section_index": section_index,
        "group_index": group_index,
        "section_name": sections[section_index].get("name", ""),
        "group_name": group.get("name", ""),
        "captions": group_captions,
        "sentences": list(group_captions),
    }


def _find_group_for_caption(sections: list[dict], caption: str) -> dict | None:
    """Return the group containing a caption, if any."""
    for section in sections:
        for group in section.get("groups", []) or []:
            if caption in _caption_values(group):
                return group
    return None


def _apply_caption_selection(
    enabled_captions: list[str],
    caption: str,
    sections: list[dict],
    should_enable: bool,
) -> list[str]:
    """Apply a caption toggle while enforcing group exclusivity."""
    updated = [item for item in enabled_captions if item != caption]
    if should_enable:
        group = _find_group_for_caption(sections, caption)
        if group:
            group_captions = set(_caption_values(group))
            updated = [item for item in updated if item not in group_captions]
        updated.append(caption)
    order_map = {value: index for index, value in enumerate(_ordered_captions_from_sections(sections))}
    return sorted(updated, key=lambda value: order_map.get(value, len(order_map)))


def _normalize_enabled_captions(enabled_captions: list[str], sections: list[dict]) -> list[str]:
    """Normalize enabled captions against known captions and group exclusivity."""
    known = set(_all_captions_from_sections(sections))
    normalized: list[str] = []
    for caption in enabled_captions:
        if caption not in known:
            continue
        normalized = _apply_caption_selection(normalized, caption, sections, True)
    return normalized


def _rename_caption_in_sections(sections: list[dict], old_caption: str, new_caption: str) -> bool:
    """Rename a configured caption inside sections and groups."""
    renamed = False
    for section in sections:
        section_item_order = []
        for item in section.get("item_order", []) or []:
            if not isinstance(item, dict):
                continue
            item_type = _normalize_item_type(item)
            if item_type == "caption":
                caption = _item_caption_value(item)
                section_item_order.append({
                    "type": "sentence",
                    "sentence": new_caption if caption == old_caption else caption,
                })
            elif item_type == "group":
                section_item_order.append({
                    "type": "group",
                    "group_id": str(item.get("group_id") or "").strip(),
                })
        if section_item_order or "item_order" in section:
            section["item_order"] = section_item_order

        section_captions = _caption_values(section)
        if old_caption in section_captions:
            renamed = True
        section["captions"] = [
            new_caption if caption == old_caption else caption
            for caption in section_captions
        ]
        section["sentences"] = list(section["captions"])

        for group in section.get("groups", []) or []:
            group_hidden_captions = _hidden_caption_values(group)
            group_captions = _caption_values(group)
            if old_caption in group_captions:
                renamed = True
            group["captions"] = [
                new_caption if caption == old_caption else caption
                for caption in group_captions
            ]
            group["sentences"] = list(group["captions"])
            group["hidden_captions"] = [
                new_caption if caption == old_caption else caption
                for caption in group_hidden_captions
                if (new_caption if caption == old_caption else caption) in group["captions"]
            ]
            group["hidden_sentences"] = list(group["hidden_captions"])
    return renamed


def _remove_caption_from_sections(sections: list[dict], caption_to_remove: str) -> bool:
    """Remove a configured caption from sections and groups."""
    removed = False
    for section in sections:
        section_captions = _caption_values(section)
        if caption_to_remove in section_captions:
            removed = True
        section["captions"] = [caption for caption in section_captions if caption != caption_to_remove]
        section["sentences"] = list(section["captions"])
        section["item_order"] = [
            {
                "type": "sentence",
                "sentence": _item_caption_value(item),
            }
            if _normalize_item_type(item) == "caption"
            else {
                "type": "group",
                "group_id": str(item.get("group_id") or "").strip(),
            }
            for item in (section.get("item_order", []) or [])
            if isinstance(item, dict)
            and not (
                _normalize_item_type(item) == "caption"
                and _item_caption_value(item) == caption_to_remove
            )
        ]
        for group in section.get("groups", []) or []:
            group_hidden_captions = _hidden_caption_values(group)
            group_captions = _caption_values(group)
            if caption_to_remove in group_captions:
                removed = True
            group["captions"] = [caption for caption in group_captions if caption != caption_to_remove]
            group["sentences"] = list(group["captions"])
            group["hidden_captions"] = [
                caption
                for caption in group_hidden_captions
                if caption != caption_to_remove and caption in group["captions"]
            ]
            group["hidden_sentences"] = list(group["hidden_captions"])
    return removed


def _remove_group_from_sections(sections: list[dict], section_index: int, group_index: int) -> list[str] | None:
    """Remove a configured group and return the removed group captions."""
    if section_index < 0 or section_index >= len(sections):
        return None
    groups = sections[section_index].get("groups", []) or []
    if group_index < 0 or group_index >= len(groups):
        return None
    removed_group = groups.pop(group_index)
    removed_group_id = str(removed_group.get("id") or "").strip()
    sections[section_index]["groups"] = groups
    sections[section_index]["item_order"] = [
        {
            "type": "sentence",
            "sentence": _item_caption_value(item),
        }
        if _normalize_item_type(item) == "caption"
        else {
            "type": "group",
            "group_id": str(item.get("group_id") or "").strip(),
        }
        for item in (sections[section_index].get("item_order", []) or [])
        if isinstance(item, dict)
        and not (
            _normalize_item_type(item) == "group"
            and str(item.get("group_id") or "").strip() == removed_group_id
        )
    ]
    return list(_caption_values(removed_group))


def _remove_section_from_sections(sections: list[dict], section_index: int) -> tuple[list[dict], list[str]] | tuple[None, None]:
    """Remove a configured section and return the updated sections with removed captions."""
    if section_index < 0 or section_index >= len(sections):
        return None, None
    removed_section = sections.pop(section_index)
    removed_captions = list(_caption_values(removed_section))
    for group in removed_section.get("groups", []) or []:
        removed_captions.extend(_caption_values(group))
    if not sections:
        sections.append({"name": "", "captions": [], "sentences": [], "groups": [], "item_order": []})
    return sections, removed_captions


def _rename_section_in_sections(sections: list[dict], old_name: str, new_name: str) -> bool:
    """Rename a configured section header."""
    renamed = False
    for section in sections:
        if str(section.get("name") or "") == old_name:
            section["name"] = new_name
            renamed = True
    return renamed


def _get_crop_aspect_ratios(cfg: dict, default_crop_aspect_ratios: list[str]) -> list[str]:
    """Get the configured list of allowed crop aspect ratios."""
    ratios = cfg.get("crop_aspect_ratios", default_crop_aspect_ratios)
    if not isinstance(ratios, list):
        return list(default_crop_aspect_ratios)
    cleaned = [str(ratio).strip() for ratio in ratios if str(ratio).strip()]
    return cleaned or list(default_crop_aspect_ratios)


_all_sentences_from_sections = _all_captions_from_sections
_group_hidden_sentences = _group_hidden_captions
_is_hidden_group_sentence = _is_hidden_group_caption
_ordered_sentences_from_sections = _ordered_captions_from_sections
_find_group_for_sentence = _find_group_for_caption
_apply_sentence_selection = _apply_caption_selection
_normalize_enabled_sentences = _normalize_enabled_captions
_rename_sentence_in_sections = _rename_caption_in_sections
_remove_sentence_from_sections = _remove_caption_from_sections