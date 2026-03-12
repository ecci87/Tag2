"""Unit tests for the Image Captioning Tool backend."""

import json
import os
import shutil
import tempfile
from io import BytesIO
from pathlib import Path

import pytest
import piexif
from fastapi.testclient import TestClient
from PIL import Image

import server


# ===== FIXTURES =====

@pytest.fixture(autouse=True)
def _reset_config(tmp_path, monkeypatch):
    """Use a temporary config file for every test."""
    config_path = str(tmp_path / "config.json")
    monkeypatch.setattr(server, "CONFIG_PATH", config_path)
    yield


@pytest.fixture
def client():
    return TestClient(server.app)


@pytest.fixture
def img_dir(tmp_path):
    """Create a temp directory with a few test images."""
    d = tmp_path / "images"
    d.mkdir()
    for name in ["photo1.jpg", "photo2.png", "photo3.jpg"]:
        img = Image.new("RGB", (100, 100), color="red")
        img.save(str(d / name))
    return d


@pytest.fixture
def single_image(img_dir):
    """Return the path to a single test image."""
    return str(img_dir / "photo1.jpg")


# ===== HELPER FUNCTIONS (UNIT TESTS) =====

class TestGetFolderSections:
    def test_default_empty(self):
        cfg = {"folders": {}}
        result = server._get_folder_sections(cfg, "/some/folder")
        assert len(result) == 1
        assert result[0]["name"] == ""
        assert result[0]["sentences"] == []
        assert result[0]["groups"] == []

    def test_folder_with_sections(self):
        sections = [
            {"name": "## Lighting", "sentences": ["bright", "dark"], "groups": []},
            {"name": "", "sentences": ["generic"], "groups": [{"name": "Car", "sentences": ["red", "blue"], "hidden_sentences": ["blue"]}]},
        ]
        cfg = {"folders": {os.path.normpath("/pics"): {"sections": sections}}}
        result = server._get_folder_sections(cfg, "/pics")
        assert len(result) == 2
        assert result[0]["name"] == "## Lighting"
        assert result[0]["sentences"] == ["bright", "dark"]
        assert result[1]["groups"][0]["sentences"] == ["red", "blue"]
        assert result[1]["groups"][0]["hidden_sentences"] == ["blue"]

    def test_migration_from_flat_sentences(self):
        cfg = {"folders": {os.path.normpath("/pics"): {"sentences": ["a", "b"]}}}
        result = server._get_folder_sections(cfg, "/pics")
        assert len(result) == 1
        assert result[0]["name"] == ""
        assert result[0]["sentences"] == ["a", "b"]

    def test_default_sections_fallback(self):
        cfg = {
            "default_sections": [{"name": "## Test", "sentences": ["x"]}],
            "folders": {},
        }
        result = server._get_folder_sections(cfg, "/new/folder")
        assert result[0]["name"] == "## Test"

    def test_default_sentences_migration(self):
        cfg = {"default_sentences": ["old1", "old2"], "folders": {}}
        result = server._get_folder_sections(cfg, "/new/folder")
        assert len(result) == 1
        assert result[0]["sentences"] == ["old1", "old2"]


class TestSetFolderSections:
    def test_set_new_folder(self):
        cfg = {"folders": {}}
        sections = [{"name": "## A", "sentences": ["s1"], "groups": []}]
        server._set_folder_sections(cfg, "/pics", sections)
        key = os.path.normpath("/pics")
        stored_sections = cfg["folders"][key]["sections"]
        assert len(stored_sections) == 1
        assert stored_sections[0]["name"] == "## A"
        assert stored_sections[0]["sentences"] == ["s1"]
        assert stored_sections[0]["groups"] == []
        assert stored_sections[0]["item_order"] == [{"type": "sentence", "sentence": "s1"}]

    def test_removes_legacy_sentences_key(self):
        key = os.path.normpath("/pics")
        cfg = {"folders": {key: {"sentences": ["old"]}}}
        server._set_folder_sections(cfg, "/pics", [{"name": "", "sentences": ["new"], "groups": []}])
        assert "sentences" not in cfg["folders"][key]
        assert "sections" in cfg["folders"][key]


class TestAllSentencesFromSections:
    def test_flatten(self):
        sections = [
            {"name": "", "sentences": ["a", "b"], "groups": []},
            {"name": "## X", "sentences": ["c"], "groups": [{"name": "Color", "sentences": ["red", "blue"]}]},
        ]
        assert server._all_sentences_from_sections(sections) == ["a", "b", "c", "red", "blue"]

    def test_empty(self):
        assert server._all_sentences_from_sections([]) == []


class TestRenameSentenceInSections:
    def test_rename_in_section_and_group(self):
        sections = [
            {"name": "", "sentences": ["old"], "groups": [{"name": "G", "sentences": ["x", "old-2"], "hidden_sentences": ["old-2"]}]},
        ]
        assert server._rename_sentence_in_sections(sections, "old", "new") is True
        assert sections[0]["sentences"] == ["new"]
        assert server._rename_sentence_in_sections(sections, "old-2", "new-2") is True
        assert sections[0]["groups"][0]["sentences"] == ["x", "new-2"]
        assert sections[0]["groups"][0]["hidden_sentences"] == ["new-2"]


class TestRenameSectionInSections:
    def test_rename_section_name(self):
        sections = [
            {"name": "Scene", "sentences": ["old"], "groups": []},
            {"name": "Other", "sentences": [], "groups": []},
        ]
        assert server._rename_section_in_sections(sections, "Scene", "Environment") is True
        assert sections[0]["name"] == "Environment"
        assert sections[1]["name"] == "Other"


class TestReadCaptionFile:
    def test_no_file(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        result = server._read_caption_file(img, ["a", "b"])
        assert result == {"enabled_sentences": [], "free_text": ""}

    def test_new_format_with_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt = tmp_path / "img.txt"
        txt.write_text("bright\n\n## Object\nred car\n\nFree text here\n", encoding="utf-8")

        result = server._read_caption_file(
            img, ["bright", "red car"], section_headers=["## Object"]
        )
        assert "bright" in result["enabled_sentences"]
        assert "red car" in result["enabled_sentences"]
        assert "Free text here" in result["free_text"]

    def test_old_format_backward_compat(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt = tmp_path / "img.txt"
        txt.write_text("bright\nred car\n\nFree text\n", encoding="utf-8")

        result = server._read_caption_file(img, ["bright", "red car"])
        assert "bright" in result["enabled_sentences"]
        assert "red car" in result["enabled_sentences"]
        assert "Free text" in result["free_text"]

    def test_only_free_text(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt = tmp_path / "img.txt"
        txt.write_text("This is just free text\nAnother line\n", encoding="utf-8")

        result = server._read_caption_file(img, [])
        assert result["enabled_sentences"] == []
        assert "This is just free text" in result["free_text"]

    def test_section_headers_consumed(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt = tmp_path / "img.txt"
        txt.write_text("**Shape**\nround\n\n## Lighting\nbright\n", encoding="utf-8")

        result = server._read_caption_file(
            img, ["round", "bright"], section_headers=["**Shape**", "## Lighting"]
        )
        assert result["enabled_sentences"] == ["round", "bright"]
        assert result["free_text"] == ""


class TestWriteCaptionFile:
    def test_simple_write_no_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)

        server._write_caption_file(img, ["a", "b"], "free stuff")
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "a\nb" in txt
        assert "free stuff" in txt

    def test_write_with_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "", "sentences": ["generic"], "groups": []},
            {"name": "## Lighting", "sentences": ["bright", "dark"], "groups": []},
            {"name": "**Shape**", "sentences": ["round"], "groups": []},
        ]
        server._write_caption_file(
            img, ["generic", "bright", "round"], "Free text here", sections
        )
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "generic" in txt
        assert "## Lighting\nbright" in txt
        assert "**Shape**\nround" in txt
        assert "Free text here" in txt
        # "dark" not enabled, should not appear
        assert "dark" not in txt

    def test_write_with_groups(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "## Vehicle", "sentences": [], "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}]},
        ]
        server._write_caption_file(img, ["Blue Car"], "", sections)
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "## Vehicle" in txt
        assert "\nCar\n" not in txt
        assert "Blue Car" in txt
        assert "Red Car" not in txt

    def test_write_skips_hidden_group_option(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "## Furniture", "sentences": [], "groups": [{"name": "Chair", "sentences": ["Chair visible", "Chair not in frame"], "hidden_sentences": ["Chair not in frame"]}]},
        ]
        server._write_caption_file(img, ["Chair not in frame"], "", sections)
        txt_path = tmp_path / "img.txt"
        assert not txt_path.exists() or txt_path.read_text(encoding="utf-8").strip() == ""

    def test_empty_clears_file(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt_path = tmp_path / "img.txt"
        txt_path.write_text("old content", encoding="utf-8")
        server._write_caption_file(img, [], "")
        assert txt_path.read_text(encoding="utf-8") == ""


class TestAutoCaptionMergeDuringRun:
    def test_stream_preserves_manual_caption_toggles(self, client, single_image, monkeypatch):
        folder = str(Path(single_image).parent)
        sections = [{"name": "", "sentences": ["cat", "dog"], "groups": []}]
        server._save_config({
            "last_folder": folder,
            "default_sentences": [],
            "crop_aspect_ratios": list(server.DEFAULT_CROP_ASPECT_RATIOS),
            "image_crops": {},
            "ollama_server": server.DEFAULT_OLLAMA_SERVER,
            "ollama_port": server.DEFAULT_OLLAMA_PORT,
            "ollama_timeout_seconds": 5,
            "ollama_host": f"http://{server.DEFAULT_OLLAMA_SERVER}:{server.DEFAULT_OLLAMA_PORT}",
            "ollama_model": "mock-model",
            "ollama_prompt_template": server.DEFAULT_OLLAMA_PROMPT_TEMPLATE,
            "ollama_group_prompt_template": server.DEFAULT_OLLAMA_GROUP_PROMPT_TEMPLATE,
            "ollama_enable_free_text": False,
            "ollama_free_text_prompt_template": server.DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE,
            "folders": {
                os.path.normpath(folder): {
                    "sections": sections,
                }
            },
        })

        call_count = {"value": 0}

        def fake_generate(host, payload, timeout):
            prompt = payload.get("prompt", "")
            call_count["value"] += 1
            if call_count["value"] == 2:
                server._write_caption_file(single_image, ["dog"], "", sections)
            if "Caption: cat" in prompt:
                return {"response": "YES"}
            if "Caption: dog" in prompt:
                return {"response": "YES"}
            raise AssertionError(f"Unexpected prompt: {prompt}")

        monkeypatch.setattr(server, "_encode_image_for_ollama", lambda image_path: "image-bytes")
        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        with client.stream(
            "POST",
            "/api/auto-caption/stream",
            json={
                "image_paths": [single_image],
                "model": "mock-model",
                "enable_free_text": False,
            },
        ) as response:
            assert response.status_code == 200
            events = [json.loads(line) for line in response.iter_lines() if line]

        completed = [event for event in events if event.get("type") == "image-complete"]
        assert len(completed) == 1
        assert completed[0]["enabled_sentences"] == ["dog"]

        final_caption = server._read_caption_file(single_image, ["cat", "dog"])
        assert final_caption["enabled_sentences"] == ["dog"]

    def test_skip_empty_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "## Empty", "sentences": ["not-enabled"], "groups": []},
            {"name": "## Has", "sentences": ["yes"], "groups": []},
        ]
        server._write_caption_file(img, ["yes"], "", sections)
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "## Empty" not in txt
        assert "## Has\nyes" in txt

    def test_roundtrip(self, tmp_path):
        """Write then read should produce the same enabled sentences."""
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "", "sentences": ["a"], "groups": []},
            {"name": "## Sec", "sentences": ["b", "c"], "groups": []},
        ]
        enabled = ["a", "b"]
        free = "Hello world"
        server._write_caption_file(img, enabled, free, sections)
        result = server._read_caption_file(
            img,
            server._all_sentences_from_sections(sections),
            section_headers=["## Sec"],
        )
        assert sorted(result["enabled_sentences"]) == sorted(enabled)
        assert result["free_text"].strip() == free


# ===== API ENDPOINT TESTS =====

class TestListImages:
    def test_list_images(self, client, img_dir):
        resp = client.get("/api/list-images", params={"folder": str(img_dir)})
        assert resp.status_code == 200
        data = resp.json()
        assert len(data["images"]) == 3
        names = {img["name"] for img in data["images"]}
        assert "photo1.jpg" in names
        assert "photo2.png" in names

    def test_invalid_folder(self, client):
        resp = client.get("/api/list-images", params={"folder": "/nonexistent/path"})
        assert resp.status_code == 400

    def test_has_caption_flag(self, client, img_dir):
        # Create a caption file for photo1
        (img_dir / "photo1.txt").write_text("caption", encoding="utf-8")
        resp = client.get("/api/list-images", params={"folder": str(img_dir)})
        images = resp.json()["images"]
        photo1 = next(i for i in images if i["name"] == "photo1.jpg")
        photo2 = next(i for i in images if i["name"] == "photo2.png")
        assert photo1["has_caption"] is True
        assert photo2["has_caption"] is False


class TestThumbnail:
    def test_get_thumbnail(self, client, single_image):
        resp = client.get("/api/thumbnail", params={"path": single_image, "size": 128})
        assert resp.status_code == 200
        assert resp.headers["content-type"] == "image/jpeg"
        assert len(resp.content) > 0

    def test_thumbnail_not_found(self, client):
        resp = client.get("/api/thumbnail", params={"path": "/no/such/file.jpg"})
        assert resp.status_code == 404


class TestPreview:
    def test_get_preview(self, client, single_image):
        resp = client.get("/api/preview", params={"path": single_image})
        assert resp.status_code == 200

    def test_preview_respects_exif_orientation(self, client, tmp_path):
        image_path = tmp_path / "rotated.jpg"
        img = Image.new("RGB", (30, 10), color="green")
        exif = piexif.dump({"0th": {piexif.ImageIFD.Orientation: 6}})
        img.save(str(image_path), exif=exif)

        resp = client.get("/api/preview", params={"path": str(image_path)})
        assert resp.status_code == 200
        preview = Image.open(BytesIO(resp.content))
        assert preview.width < preview.height

    def test_preview_not_found(self, client):
        resp = client.get("/api/preview", params={"path": "/no/such/file.jpg"})
        assert resp.status_code == 404


class TestCaptionAPI:
    def test_get_caption_no_file(self, client, single_image):
        resp = client.get("/api/caption", params={
            "path": single_image,
            "sentences": json.dumps(["a", "b"]),
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["enabled_sentences"] == []
        assert data["free_text"] == ""

    def test_save_and_get_caption(self, client, single_image):
        # Set up sections config so "bright" is a known sentence
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["bright"]}],
        })

        # Save
        resp = client.post("/api/caption/save", json={
            "image_path": single_image,
            "enabled_sentences": ["bright"],
            "free_text": "My notes",
        })
        assert resp.status_code == 200

        # Read back
        resp = client.get("/api/caption", params={
            "path": single_image,
            "sentences": json.dumps(["bright"]),
        })
        data = resp.json()
        assert "bright" in data["enabled_sentences"]
        assert "My notes" in data["free_text"]

    def test_save_caption_file_not_found(self, client):
        resp = client.post("/api/caption/save", json={
            "image_path": "/no/such/file.jpg",
            "enabled_sentences": [],
            "free_text": "",
        })
        assert resp.status_code == 404

    def test_rename_caption_preset_updates_config_and_files(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Old Caption"], "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}]}],
        })
        client.post("/api/caption/save", json={
            "image_path": image_path,
            "enabled_sentences": ["Old Caption", "Red Car"],
            "free_text": "notes",
        })

        resp = client.post("/api/caption/rename-preset", json={
            "folder": str(img_dir),
            "old_sentence": "Old Caption",
            "new_sentence": "New Caption",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["sections"][0]["sentences"] == ["New Caption"]

        caption = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "New Caption" in caption
        assert "Old Caption" not in caption
        assert "Red Car" in caption

    def test_rename_caption_preset_updates_free_text_occurrences(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Old Caption"]}],
        })
        client.post("/api/caption/save", json={
            "image_path": image_path,
            "enabled_sentences": ["Old Caption"],
            "free_text": "Old Caption appears again in notes.",
        })

        resp = client.post("/api/caption/rename-preset", json={
            "folder": str(img_dir),
            "old_sentence": "Old Caption",
            "new_sentence": "New Caption",
        })
        assert resp.status_code == 200

        caption = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "New Caption appears again in notes." in caption
        assert "Old Caption" not in caption

    def test_rename_section_updates_config_and_files(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "Scene", "sentences": ["Moon"]}],
        })
        client.post("/api/caption/save", json={
            "image_path": image_path,
            "enabled_sentences": ["Moon"],
            "free_text": "Scene appears in notes.",
        })

        resp = client.post("/api/section/rename", json={
            "folder": str(img_dir),
            "old_name": "Scene",
            "new_name": "Environment",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["sections"][0]["name"] == "Environment"

        caption = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "Environment" in caption
        assert "Scene" not in caption
        assert "Moon" in caption


class TestBatchToggle:
    def test_enable_sentence(self, client, img_dir):
        paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        # Set up sections config for this folder
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["tag1", "tag2"]}],
        })
        resp = client.post("/api/caption/batch-toggle", json={
            "image_paths": paths,
            "sentence": "tag1",
            "enabled": True,
        })
        assert resp.status_code == 200
        results = resp.json()["results"]
        assert all(r.get("ok") for r in results)

        # Verify written
        txt1 = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "tag1" in txt1

    def test_disable_sentence(self, client, img_dir):
        img_path = str(img_dir / "photo1.jpg")
        # Set up config and pre-enable
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["tag1"]}],
        })
        client.post("/api/caption/batch-toggle", json={
            "image_paths": [img_path],
            "sentence": "tag1",
            "enabled": True,
        })
        # Disable
        client.post("/api/caption/batch-toggle", json={
            "image_paths": [img_path],
            "sentence": "tag1",
            "enabled": False,
        })
        txt = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "tag1" not in txt

    def test_missing_file_in_batch(self, client, img_dir):
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["x"]}],
        })
        resp = client.post("/api/caption/batch-toggle", json={
            "image_paths": [str(img_dir / "photo1.jpg"), "/nonexistent.jpg"],
            "sentence": "x",
            "enabled": True,
        })
        results = resp.json()["results"]
        assert results[0].get("ok") is True
        assert "error" in results[1]

    def test_group_toggle_is_exclusive(self, client, img_dir):
        img_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": [], "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}]}],
        })
        client.post("/api/caption/batch-toggle", json={
            "image_paths": [img_path],
            "sentence": "Red Car",
            "enabled": True,
        })
        client.post("/api/caption/batch-toggle", json={
            "image_paths": [img_path],
            "sentence": "Blue Car",
            "enabled": True,
        })
        txt = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "Blue Car" in txt
        assert "Red Car" not in txt


class TestBulkCaptions:
    def test_bulk_get(self, client, img_dir):
        paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        resp = client.get("/api/captions/bulk", params={
            "paths": json.dumps(paths),
            "sentences": json.dumps(["a"]),
        })
        assert resp.status_code == 200
        data = resp.json()
        assert len(data) == 2

    def test_bulk_invalid_json(self, client):
        resp = client.get("/api/captions/bulk", params={
            "paths": "not json",
            "sentences": "[]",
        })
        assert resp.status_code == 400

    def test_bulk_post(self, client, img_dir):
        paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        resp = client.post("/api/captions/bulk", json={
            "paths": paths,
            "sentences": ["a"],
        })
        assert resp.status_code == 200
        data = resp.json()
        assert len(data) == 2


class TestSettingsAPI:
    def test_get_default_settings(self, client):
        resp = client.get("/api/settings")
        assert resp.status_code == 200
        data = resp.json()
        assert "last_folder" in data
        assert data["thumb_size"] == 160
        assert data["crop_aspect_ratios"] == ["4:3", "16:9", "3:4", "1:1", "9:16", "2:3", "3:2"]
        assert data["ollama_host"] == "http://127.0.0.1:11434"
        assert data["ollama_server"] == "127.0.0.1"
        assert data["ollama_port"] == 11434
        assert data["ollama_timeout_seconds"] == 20
        assert data["ollama_model"] == "llava"
        assert "{caption}" in data["ollama_prompt_template"]
        assert "{group_name}" in data["ollama_group_prompt_template"]
        assert data["ollama_enable_free_text"] is True
        assert "{caption_text}" in data["ollama_free_text_prompt_template"]

    def test_save_and_load_last_folder(self, client):
        client.post("/api/settings", json={"last_folder": "/my/folder"})
        resp = client.get("/api/settings")
        assert resp.json()["last_folder"] == "/my/folder"

    def test_save_and_load_thumbnail_size(self, client):
        client.post("/api/settings", json={"thumb_size": 224})
        resp = client.get("/api/settings")
        assert resp.json()["thumb_size"] == 224

    def test_save_and_load_sections(self, client, tmp_path):
        folder = str(tmp_path / "test_folder")
        os.makedirs(folder, exist_ok=True)
        sections = [
            {"name": "## Lighting", "sentences": ["bright"]},
            {"name": "", "sentences": ["generic"], "groups": [{"name": "Chair", "sentences": ["visible", "not visible"], "hidden_sentences": ["not visible"]}]},
        ]
        client.post("/api/settings", json={
            "folder": folder,
            "sections": sections,
        })
        resp = client.get("/api/settings", params={"folder": folder})
        data = resp.json()
        assert len(data["sections"]) == 2
        assert data["sections"][0]["name"] == "## Lighting"
        assert data["sections"][0]["sentences"] == ["bright"]
        assert data["sections"][1]["groups"][0]["hidden_sentences"] == ["not visible"]

    def test_settings_persistence(self, client, tmp_path):
        """Settings survive a config reload."""
        folder = str(tmp_path / "persist")
        os.makedirs(folder, exist_ok=True)
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "A", "sentences": ["s1"]}],
        })
        # Force reload
        cfg = server._load_config()
        sections = server._get_folder_sections(cfg, folder)
        assert sections[0]["sentences"] == ["s1"]


class TestConfigPersistence:
    def test_save_and_load(self, tmp_path, monkeypatch):
        config_path = str(tmp_path / "cfg.json")
        monkeypatch.setattr(server, "CONFIG_PATH", config_path)

        cfg = {"last_folder": "/test", "folders": {}, "default_sentences": []}
        server._save_config(cfg)
        loaded = server._load_config()
        assert loaded["last_folder"] == "/test"

    def test_load_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(server, "CONFIG_PATH", str(tmp_path / "missing.json"))
        cfg = server._load_config()
        assert cfg["last_folder"] == ""
        assert cfg["folders"] == {}
        assert cfg["crop_aspect_ratios"] == ["4:3", "16:9", "3:4", "1:1", "9:16", "2:3", "3:2"]
        assert cfg["ollama_host"] == "http://127.0.0.1:11434"
        assert cfg["ollama_server"] == "127.0.0.1"
        assert cfg["ollama_port"] == 11434
        assert cfg["ollama_timeout_seconds"] == 20
        assert cfg["ollama_model"] == "llava"
        assert "{caption}" in cfg["ollama_prompt_template"]
        assert "{group_name}" in cfg["ollama_group_prompt_template"]
        assert cfg["ollama_enable_free_text"] is True
        assert "{caption_text}" in cfg["ollama_free_text_prompt_template"]


class TestOllamaHelpers:
    def test_parse_yes_no(self):
        assert server._parse_ollama_yes_no("YES") is True
        assert server._parse_ollama_yes_no("No.") is False
        assert server._parse_ollama_yes_no("YES - the object is visible") is True
        assert server._parse_ollama_yes_no("uncertain") is False

    def test_prompt_template_substitution(self):
        prompt = server._ollama_prompt_for_sentence("Moon", "Caption check: {caption} / {sentence}")
        assert prompt == "Caption check: Moon / Moon"

    def test_free_text_prompt_template_substitution(self):
        prompt = server._ollama_prompt_for_free_text("- Moon", "Current: {caption_text} / {current_caption}")
        assert prompt == "Current: - Moon / - Moon"

    def test_group_prompt_template_substitution(self):
        prompt = server._ollama_prompt_for_group("Car", ["Red Car", "Blue Car"], "Group {group_name}\n{options}\n{count}")
        assert "Car" in prompt
        assert "1. Red Car" in prompt
        assert prompt.endswith("2")

    def test_parse_group_selection(self):
        assert server._parse_ollama_selection("Answer: 2", ["Red Car", "Blue Car"]) == 2
        assert server._parse_ollama_selection("Blue Car", ["Red Car", "Blue Car"]) == 2

    def test_merge_free_text_avoids_duplicates(self):
        merged, added = server._merge_free_text(
            "Already there\nNight sky",
            "Night sky\nMoon visible\n- Already there\nBright stars",
            ["Moon"],
        )
        assert added == ["Moon visible", "Bright stars"]
        assert merged == "Already there\nNight sky\nMoon visible\nBright stars"

    def test_auto_caption_sentences(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            assert timeout == 17
            prompt = payload["prompt"]
            if "Moon" in prompt:
                return {"response": "YES"}
            return {"response": "NO"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        enabled, results = server._auto_caption_sentences(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            ["Moon", "Night", "Car"],
            "Caption? {caption}",
            17,
        )
        assert enabled == ["Moon"]
        assert len(results) == 3
        assert results[0]["enabled"] is True
        assert results[1]["enabled"] is False

    def test_auto_caption_sections(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            prompt = payload["prompt"]
            if "Caption:" in prompt:
                return {"response": "YES" if "Moon" in prompt else "NO"}
            if "Group:" in prompt:
                return {"response": "2"}
            raise AssertionError("unexpected prompt")

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        sections = [{
            "name": "",
            "sentences": ["Moon"],
            "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}],
        }]
        enabled, results = server._auto_caption_sections(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            sections,
            "Caption: {caption}",
            "Group: {group_name}\n{options}",
            15,
        )
        assert enabled == ["Moon", "Blue Car"]
        assert results[0]["type"] == "sentence"
        assert results[1]["type"] == "group"
        assert results[1]["selected_sentence"] == "Blue Car"

    def test_auto_caption_sections_marks_hidden_group_selection(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            if "Group:" in payload["prompt"]:
                return {"response": "2"}
            return {"response": "NO"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        sections = [{
            "name": "",
            "sentences": [],
            "groups": [{"name": "Chair", "sentences": ["Chair visible", "Chair not in frame"], "hidden_sentences": ["Chair not in frame"]}],
        }]
        enabled, results = server._auto_caption_sections(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            sections,
            "Caption: {caption}",
            "Group: {group_name}\n{options}",
            15,
        )
        assert enabled == ["Chair not in frame"]
        assert results[0]["selected_sentence"] == "Chair not in frame"
        assert results[0]["selected_hidden"] is True

    def test_suggest_free_text(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            assert timeout == 9
            assert "Current caption text" in payload["prompt"]
            return {"response": "Moon visible\nBright stars"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        result = server._suggest_free_text(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            "- Night",
            timeout=9,
        )
        assert "Moon visible" in result


class TestCropAPI:
    def test_get_crop_default_none(self, client, single_image):
        resp = client.get("/api/crop", params={"path": single_image})
        assert resp.status_code == 200
        crop = resp.json()["crop"]
        assert crop["applied"] is False
        assert crop["current_width"] == 100
        assert crop["current_height"] == 100

    def test_save_and_get_crop(self, client, single_image):
        crop = {"x": 10, "y": 5, "w": 40, "h": 30, "ratio": "4:3"}
        resp = client.post("/api/crop", json={"image_path": single_image, "crop": crop})
        assert resp.status_code == 200
        saved = resp.json()["crop"]
        assert saved["applied"] is True
        assert saved["current_width"] == 40
        assert saved["current_height"] == 30

        with Image.open(single_image) as img:
            assert img.size == (40, 30)

        resp = client.get("/api/crop", params={"path": single_image})
        assert resp.json()["crop"]["original_width"] == 100
        assert os.path.isfile(server._get_crop_backup_path(single_image))

    def test_clear_crop(self, client, single_image):
        client.post("/api/crop", json={"image_path": single_image, "crop": {"x": 1, "y": 1, "w": 20, "h": 20}})
        resp = client.post("/api/crop", json={"image_path": single_image, "crop": None})
        assert resp.status_code == 200
        assert resp.json()["crop"]["applied"] is False

        with Image.open(single_image) as img:
            assert img.size == (100, 100)
        assert not os.path.exists(server._get_crop_backup_path(single_image))

    def test_thumbnail_respects_crop(self, client, single_image):
        with Image.open(single_image) as img:
            img = Image.new("RGB", (100, 100))
            for x in range(100):
                for y in range(100):
                    img.putpixel((x, y), (255, 0, 0) if x < 50 else (0, 0, 255))
            img.save(single_image)

        client.post("/api/crop", json={
            "image_path": single_image,
            "crop": {"x": 0, "y": 0, "w": 50, "h": 100, "ratio": "1:2"},
        })
        resp = client.get("/api/thumbnail", params={"path": single_image, "size": 64})
        assert resp.status_code == 200
        thumb = Image.open(BytesIO(resp.content))
        r, g, b = thumb.getpixel((thumb.width // 2, thumb.height // 2))
        assert r > b

    def test_rotate_image_right(self, client, tmp_path):
        image_path = str(tmp_path / "rotate.jpg")
        Image.new("RGB", (80, 50), color="red").save(image_path)

        resp = client.post("/api/rotate", json={"image_path": image_path, "direction": "right"})
        assert resp.status_code == 200
        crop = resp.json()["crop"]
        assert crop["current_width"] == 50
        assert crop["current_height"] == 80

        with Image.open(image_path) as img:
            assert img.size == (50, 80)

    def test_rotate_image_preserves_crop_undo(self, client, tmp_path):
        image_path = str(tmp_path / "rotate_crop.jpg")
        Image.new("RGB", (120, 80), color="blue").save(image_path)

        crop_resp = client.post("/api/crop", json={
            "image_path": image_path,
            "crop": {"x": 10, "y": 10, "w": 60, "h": 40, "ratio": "3:2"},
        })
        assert crop_resp.status_code == 200

        rotate_resp = client.post("/api/rotate", json={"image_path": image_path, "direction": "left"})
        assert rotate_resp.status_code == 200
        assert rotate_resp.json()["crop"]["applied"] is True

        clear_resp = client.post("/api/crop", json={"image_path": image_path, "crop": None})
        assert clear_resp.status_code == 200

        with Image.open(image_path) as img:
            assert img.size == (80, 120)


class TestAutoCaptionAPI:
    def test_auto_caption_single_image(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["Moon", "Night"], "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}]}],
            "ollama_server": "localhost",
            "ollama_port": 11435,
            "ollama_timeout_seconds": 12,
            "ollama_model": "llava",
            "ollama_prompt_template": "Caption: {caption}",
            "ollama_group_prompt_template": "Group: {group_name}\n{options}",
            "ollama_enable_free_text": True,
            "ollama_free_text_prompt_template": "Current caption text:\n{caption_text}",
        })

        def fake_auto_caption(host, model, image_path, sections, prompt_template, group_prompt_template, timeout):
            assert host == "http://localhost:11435"
            assert model == "llava"
            assert image_path == single_image
            assert server._all_sentences_from_sections(sections) == ["Moon", "Night", "Red Car", "Blue Car"]
            assert prompt_template == "Caption: {caption}"
            assert group_prompt_template == "Group: {group_name}\n{options}"
            assert timeout == 12
            return ["Moon", "Night", "Blue Car"], [
                {"sentence": "Moon", "enabled": True, "answer": "YES"},
                {"sentence": "Night", "enabled": True, "answer": "YES"},
                {"type": "group", "group_name": "Car", "selected_sentence": "Blue Car", "selected_hidden": False, "selection_index": 2, "answer": "2"},
            ]

        monkeypatch.setattr(server, "_auto_caption_sections", fake_auto_caption)
        monkeypatch.setattr(server, "_suggest_free_text", lambda *args: "Moon visible\nNight sky")

        resp = client.post("/api/auto-caption", json={
            "image_path": single_image,
            "model": "llava",
            "prompt_template": "Caption: {caption}",
            "group_prompt_template": "Group: {group_name}\n{options}",
            "enable_free_text": True,
            "free_text_prompt_template": "Current caption text:\n{caption_text}",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["enabled_sentences"] == ["Moon", "Night", "Blue Car"]
        assert data["results"][2]["selected_sentence"] == "Blue Car"
        assert data["host"] == "http://localhost:11435"
        assert data["timeout_seconds"] == 12
        assert data["prompt_template"] == "Caption: {caption}"
        assert data["group_prompt_template"] == "Group: {group_name}\n{options}"
        assert data["added_free_text_lines"] == ["Moon visible", "Night sky"]

        caption = (Path(single_image).with_suffix(".txt")).read_text(encoding="utf-8")
        assert "Moon" in caption
        assert "Night" in caption
        assert "Blue Car" in caption
        assert "Moon visible" in caption

    def test_auto_caption_requires_sentences(self, client, single_image):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": []}],
        })
        resp = client.post("/api/auto-caption", json={"image_path": single_image, "model": "llava"})
        assert resp.status_code == 400

    def test_auto_caption_ollama_error(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["Moon"]}],
        })

        def fake_auto_caption(host, model, image_path, sections, prompt_template, group_prompt_template, timeout):
            raise RuntimeError("connection refused")

        monkeypatch.setattr(server, "_auto_caption_sections", fake_auto_caption)
        resp = client.post("/api/auto-caption", json={"image_path": single_image, "model": "llava"})
        assert resp.status_code == 502

    def test_auto_caption_free_text_only(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["Moon"], "groups": []}],
            "ollama_enable_free_text": True,
        })
        client.post("/api/caption/save", json={
            "image_path": single_image,
            "enabled_sentences": ["Moon"],
            "free_text": "Night sky",
        })

        monkeypatch.setattr(server, "_suggest_free_text", lambda *args: "Bright stars")

        resp = client.post("/api/auto-caption", json={
            "image_path": single_image,
            "model": "llava",
            "free_text_only": True,
            "enable_free_text": True,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["free_text_only"] is True
        assert data["enabled_sentences"] == ["Moon"]
        assert data["added_free_text_lines"] == ["Bright stars"]

        caption = (Path(single_image).with_suffix(".txt")).read_text(encoding="utf-8")
        assert "Moon" in caption
        assert "Night sky" not in caption
        assert "Bright stars" in caption

    def test_auto_caption_target_group_preserves_other_captions(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{
                "name": "",
                "sentences": ["Moon", "Night"],
                "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}],
            }],
            "ollama_enable_free_text": True,
        })
        client.post("/api/caption/save", json={
            "image_path": single_image,
            "enabled_sentences": ["Moon", "Red Car"],
            "free_text": "Existing notes",
        })

        def fake_generate(host, payload, timeout=120):
            assert "Group:" in payload["prompt"] or "1." in payload["prompt"]
            return {"response": "2"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption", json={
            "image_path": single_image,
            "model": "llava",
            "target_section_index": 0,
            "target_group_index": 0,
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["enabled_sentences"] == ["Moon", "Blue Car"]
        assert data["results"][0]["selected_sentence"] == "Blue Car"
        assert data["free_text"].strip() == "Existing notes"

        caption = (Path(single_image).with_suffix(".txt")).read_text(encoding="utf-8")
        assert "Moon" in caption
        assert "Blue Car" in caption
        assert "Red Car" not in caption
        assert "Existing notes" in caption

    def test_auto_caption_target_sentence_preserves_other_captions(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{
                "name": "",
                "sentences": ["Moon", "Night"],
                "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}],
            }],
            "ollama_enable_free_text": True,
        })
        client.post("/api/caption/save", json={
            "image_path": single_image,
            "enabled_sentences": ["Night", "Red Car"],
            "free_text": "Keep this",
        })

        def fake_generate(host, payload, timeout=120):
            assert "Caption:" in payload["prompt"] or "Answer:" in payload["prompt"]
            return {"response": "YES"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption", json={
            "image_path": single_image,
            "model": "llava",
            "target_sentence": "Moon",
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["enabled_sentences"] == ["Moon", "Night", "Red Car"]
        assert data["results"][0]["type"] == "sentence"
        assert data["results"][0]["sentence"] == "Moon"
        assert data["free_text"].strip() == "Keep this"

        caption = (Path(single_image).with_suffix(".txt")).read_text(encoding="utf-8")
        assert "Moon" in caption
        assert "Night" in caption
        assert "Red Car" in caption
        assert "Keep this" in caption

    def test_auto_caption_hidden_group_option_not_written(self, client, single_image, monkeypatch):
        folder = str(os.path.dirname(single_image))
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{
                "name": "",
                "sentences": [],
                "groups": [{"name": "Chair", "sentences": ["Chair visible", "Chair not in frame"], "hidden_sentences": ["Chair not in frame"]}],
            }],
            "ollama_enable_free_text": False,
        })

        monkeypatch.setattr(server, "_ollama_generate", lambda *args, **kwargs: {"response": "2"})

        resp = client.post("/api/auto-caption", json={
            "image_path": single_image,
            "model": "llava",
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["results"][0]["selected_hidden"] is True

        caption_path = Path(single_image).with_suffix(".txt")
        assert not caption_path.exists() or "Chair not in frame" not in caption_path.read_text(encoding="utf-8")

    def test_auto_caption_stream_multiple_images(self, client, img_dir, monkeypatch):
        paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Moon", "Night"], "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}]}],
            "ollama_enable_free_text": True,
        })

        def fake_generate(host, payload, timeout=120):
            prompt = payload["prompt"]
            if "Current caption text" in prompt:
                return {"response": "Bright stars\nMoon visible"}
            if "Group:" in prompt or ("Car" in prompt and "1." in prompt):
                return {"response": "2"}
            if "Moon" in prompt:
                return {"response": "YES"}
            return {"response": "NO"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption/stream", json={
            "image_paths": paths,
            "model": "llava",
            "enable_free_text": True,
        })
        assert resp.status_code == 200
        events = [json.loads(line) for line in resp.text.splitlines() if line.strip()]
        assert events[0]["type"] == "start"
        assert any(e["type"] == "caption-check" and e["sentence"] == "Moon" for e in events)
        assert any(e["type"] == "group-selection" and e["selected_sentence"] == "Blue Car" for e in events)
        assert any(e["type"] == "free-text" and "Bright stars" in e["answer"] for e in events)
        assert events[-1]["type"] == "done"
        assert events[-1]["processed"] == 2

        caption1 = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        caption2 = (img_dir / "photo2.txt").read_text(encoding="utf-8")
        assert "Moon" in caption1
        assert "Blue Car" in caption1
        assert "Bright stars" in caption1
        assert "Moon" in caption2

    def test_auto_caption_stream_free_text_only(self, client, img_dir, monkeypatch):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Moon"], "groups": []}],
            "ollama_enable_free_text": True,
        })
        client.post("/api/caption/save", json={
            "image_path": image_path,
            "enabled_sentences": ["Moon"],
            "free_text": "Night sky",
        })

        def fake_generate(host, payload, timeout=120):
            if "Current caption text" in payload["prompt"]:
                return {"response": "Bright stars"}
            raise AssertionError("structured caption check should not run")

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption/stream", json={
            "image_paths": [image_path],
            "model": "llava",
            "free_text_only": True,
            "enable_free_text": True,
        })
        assert resp.status_code == 200
        events = [json.loads(line) for line in resp.text.splitlines() if line.strip()]
        assert events[0]["type"] == "start"
        assert events[0]["free_text_only"] is True
        assert any(e["type"] == "free-text" and e["free_text"] == "Bright stars" for e in events)
        assert not any(e["type"] == "caption-check" for e in events)

        caption = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "Moon" in caption
        assert "Night sky" not in caption
        assert "Bright stars" in caption

    def test_auto_caption_stream_target_group_preserves_other_captions(self, client, img_dir, monkeypatch):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{
                "name": "",
                "sentences": ["Moon", "Night"],
                "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}],
            }],
        })
        client.post("/api/caption/save", json={
            "image_path": image_path,
            "enabled_sentences": ["Night", "Red Car"],
            "free_text": "Keep this",
        })

        def fake_generate(host, payload, timeout=120):
            if "Group:" in payload["prompt"] or "1." in payload["prompt"]:
                return {"response": "2"}
            raise AssertionError("only target group should run")

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption/stream", json={
            "image_paths": [image_path],
            "model": "llava",
            "target_section_index": 0,
            "target_group_index": 0,
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        events = [json.loads(line) for line in resp.text.splitlines() if line.strip()]
        assert any(e["type"] == "group-selection" and e["selected_sentence"] == "Blue Car" for e in events)
        assert not any(e["type"] == "caption-check" for e in events)

        caption = (img_dir / "photo1.txt").read_text(encoding="utf-8")
        assert "Night" in caption
        assert "Blue Car" in caption
        assert "Red Car" not in caption
        assert "Keep this" in caption

    def test_auto_caption_stream_hidden_group_option_not_written(self, client, img_dir, monkeypatch):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{
                "name": "",
                "sentences": [],
                "groups": [{"name": "Chair", "sentences": ["Chair visible", "Chair not in frame"], "hidden_sentences": ["Chair not in frame"]}],
            }],
            "ollama_enable_free_text": False,
        })

        monkeypatch.setattr(server, "_ollama_generate", lambda *args, **kwargs: {"response": "2"})

        resp = client.post("/api/auto-caption/stream", json={
            "image_paths": [image_path],
            "model": "llava",
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        events = [json.loads(line) for line in resp.text.splitlines() if line.strip()]
        group_event = next(e for e in events if e["type"] == "group-selection")
        assert group_event["selected_hidden"] is True

        caption_path = img_dir / "photo1.txt"
        assert not caption_path.exists() or "Chair not in frame" not in caption_path.read_text(encoding="utf-8")
