"""Unit tests for the Image Captioning Tool backend."""

import json
import os
import shutil
import tempfile

import pytest
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

    def test_folder_with_sections(self):
        sections = [
            {"name": "## Lighting", "sentences": ["bright", "dark"]},
            {"name": "", "sentences": ["generic"]},
        ]
        cfg = {"folders": {os.path.normpath("/pics"): {"sections": sections}}}
        result = server._get_folder_sections(cfg, "/pics")
        assert len(result) == 2
        assert result[0]["name"] == "## Lighting"
        assert result[0]["sentences"] == ["bright", "dark"]

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
        sections = [{"name": "## A", "sentences": ["s1"]}]
        server._set_folder_sections(cfg, "/pics", sections)
        key = os.path.normpath("/pics")
        assert cfg["folders"][key]["sections"] == sections

    def test_removes_legacy_sentences_key(self):
        key = os.path.normpath("/pics")
        cfg = {"folders": {key: {"sentences": ["old"]}}}
        server._set_folder_sections(cfg, "/pics", [{"name": "", "sentences": ["new"]}])
        assert "sentences" not in cfg["folders"][key]
        assert "sections" in cfg["folders"][key]


class TestAllSentencesFromSections:
    def test_flatten(self):
        sections = [
            {"name": "", "sentences": ["a", "b"]},
            {"name": "## X", "sentences": ["c"]},
        ]
        assert server._all_sentences_from_sections(sections) == ["a", "b", "c"]

    def test_empty(self):
        assert server._all_sentences_from_sections([]) == []


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
        txt.write_text("- bright\n\n## Object\n- red car\n\nFree text here\n", encoding="utf-8")

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
        txt.write_text("**Shape**\n- round\n\n## Lighting\n- bright\n", encoding="utf-8")

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
        assert "- a\n- b" in txt
        assert "free stuff" in txt

    def test_write_with_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "", "sentences": ["generic"]},
            {"name": "## Lighting", "sentences": ["bright", "dark"]},
            {"name": "**Shape**", "sentences": ["round"]},
        ]
        server._write_caption_file(
            img, ["generic", "bright", "round"], "Free text here", sections
        )
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "- generic" in txt
        assert "## Lighting\n- bright" in txt
        assert "**Shape**\n- round" in txt
        assert "Free text here" in txt
        # "dark" not enabled, should not appear
        assert "dark" not in txt

    def test_empty_clears_file(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        txt_path = tmp_path / "img.txt"
        txt_path.write_text("old content", encoding="utf-8")
        server._write_caption_file(img, [], "")
        assert txt_path.read_text(encoding="utf-8") == ""

    def test_skip_empty_sections(self, tmp_path):
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "## Empty", "sentences": ["not-enabled"]},
            {"name": "## Has", "sentences": ["yes"]},
        ]
        server._write_caption_file(img, ["yes"], "", sections)
        txt = (tmp_path / "img.txt").read_text(encoding="utf-8")
        assert "## Empty" not in txt
        assert "## Has\n- yes" in txt

    def test_roundtrip(self, tmp_path):
        """Write then read should produce the same enabled sentences."""
        img = str(tmp_path / "img.jpg")
        Image.new("RGB", (10, 10)).save(img)
        sections = [
            {"name": "", "sentences": ["a"]},
            {"name": "## Sec", "sentences": ["b", "c"]},
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


class TestSettingsAPI:
    def test_get_default_settings(self, client):
        resp = client.get("/api/settings")
        assert resp.status_code == 200
        data = resp.json()
        assert "last_folder" in data

    def test_save_and_load_last_folder(self, client):
        client.post("/api/settings", json={"last_folder": "/my/folder"})
        resp = client.get("/api/settings")
        assert resp.json()["last_folder"] == "/my/folder"

    def test_save_and_load_sections(self, client, tmp_path):
        folder = str(tmp_path / "test_folder")
        os.makedirs(folder, exist_ok=True)
        sections = [
            {"name": "## Lighting", "sentences": ["bright"]},
            {"name": "", "sentences": ["generic"]},
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
