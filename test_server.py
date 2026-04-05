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
from tag2_images import _get_video_mask_path


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

def make_upload_file(name: str, size: tuple[int, int] = (24, 24), color: str = "blue") -> tuple[str, BytesIO, str]:
    buffer = BytesIO()
    extension = Path(name).suffix.lower()
    if extension in {".mp4", ".m4v", ".mov", ".webm", ".mkv", ".avi"}:
        buffer.write(b"fake-video-bytes")
        buffer.seek(0)
        return (name, buffer, "video/mp4")

    image = Image.new("RGB", size, color=color)
    format_name = {
        ".jpg": "JPEG",
        ".jpeg": "JPEG",
        ".png": "PNG",
        ".gif": "GIF",
        ".bmp": "BMP",
        ".webp": "WEBP",
        ".tif": "TIFF",
        ".tiff": "TIFF",
    }.get(extension, "PNG")
    image.save(buffer, format=format_name)
    buffer.seek(0)
    media_type = "image/jpeg" if format_name == "JPEG" else f"image/{format_name.lower()}"
    return (name, buffer, media_type)

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


class TestHttpsConfig:
    def test_get_https_uvicorn_kwargs_returns_empty_without_https_config(self):
        assert server._get_https_uvicorn_kwargs({}) == {}

    def test_get_https_uvicorn_kwargs_resolves_relative_paths(self, tmp_path, monkeypatch):
        config_path = tmp_path / "config.json"
        cert_path = tmp_path / "certs" / "localhost.pem"
        key_path = tmp_path / "certs" / "localhost-key.pem"
        cert_path.parent.mkdir(parents=True, exist_ok=True)
        cert_path.write_text("cert", encoding="utf-8")
        key_path.write_text("key", encoding="utf-8")
        monkeypatch.setattr(server, "CONFIG_PATH", str(config_path))

        result = server._get_https_uvicorn_kwargs({
            "https_certfile": os.path.join("certs", "localhost.pem"),
            "https_keyfile": os.path.join("certs", "localhost-key.pem"),
        })

        assert result == {
            "ssl_certfile": str(cert_path),
            "ssl_keyfile": str(key_path),
        }

    def test_get_https_uvicorn_kwargs_requires_both_files(self):
        with pytest.raises(RuntimeError, match="requires both https_certfile and https_keyfile"):
            server._get_https_uvicorn_kwargs({"https_certfile": "cert.pem"})

    def test_get_https_uvicorn_kwargs_rejects_missing_cert_file(self, tmp_path):
        key_path = tmp_path / "localhost-key.pem"
        key_path.write_text("key", encoding="utf-8")

        with pytest.raises(RuntimeError, match="certificate file not found"):
            server._get_https_uvicorn_kwargs({
                "https_certfile": str(tmp_path / "missing-cert.pem"),
                "https_keyfile": str(key_path),
            })


class TestAllSentencesFromSections:
    def test_flatten(self):
        sections = [
            {"name": "", "sentences": ["a", "b"], "groups": []},
            {"name": "## X", "sentences": ["c"], "groups": [{"name": "Color", "sentences": ["red", "blue"]}]},
        ]
        assert server._all_captions_from_sections(sections) == ["a", "b", "c", "red", "blue"]

    def test_empty(self):
        assert server._all_captions_from_sections([]) == []


class TestRenameSentenceInSections:
    def test_rename_in_section_and_group(self):
        sections = [
            {"name": "", "sentences": ["old"], "groups": [{"name": "G", "sentences": ["x", "old-2"], "hidden_sentences": ["old-2"]}]},
        ]
        assert server._rename_caption_in_sections(sections, "old", "new") is True
        assert sections[0]["sentences"] == ["new"]
        assert server._rename_caption_in_sections(sections, "old-2", "new-2") is True
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

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda image_path, **kwargs: ["image-bytes"])
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
            server._all_captions_from_sections(sections),
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

    def test_list_images_includes_videos(self, client, img_dir):
        (img_dir / "clip.mp4").write_bytes(b"fake-video-bytes")
        resp = client.get("/api/list-images", params={"folder": str(img_dir)})
        assert resp.status_code == 200
        images = resp.json()["images"]
        clip = next(item for item in images if item["name"] == "clip.mp4")
        assert clip["media_type"] == "video"

    def test_list_images_reports_video_mask_count(self, client, img_dir):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        Image.new("L", (16, 16), color=96).save(_get_video_mask_path(str(video_path), 0))
        Image.new("L", (16, 16), color=192).save(_get_video_mask_path(str(video_path), 24))

        resp = client.get("/api/list-images", params={"folder": str(img_dir)})
        assert resp.status_code == 200
        images = resp.json()["images"]
        clip = next(item for item in images if item["name"] == "clip.mp4")
        assert clip["has_mask"] is True
        assert clip["mask_count"] == 2

    def test_list_images_hides_mask_sidecars_and_reports_mask_flag(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        server._write_default_image_mask(image_path)

        resp = client.get("/api/list-images", params={"folder": str(img_dir)})
        assert resp.status_code == 200
        images = resp.json()["images"]
        names = {item["name"] for item in images}
        assert "photo1.jpg.mask.png" not in names
        photo1 = next(item for item in images if item["name"] == "photo1.jpg")
        assert photo1["has_mask"] is True


class TestFolderSuggestions:
    def test_folder_suggestions_match_partial_name(self, client, tmp_path):
        base = tmp_path / "folders"
        base.mkdir()
        (base / "alpha").mkdir()
        (base / "alpine").mkdir()
        (base / "beta").mkdir()

        resp = client.get("/api/folders/suggest", params={"query": str(base / "al")})

        assert resp.status_code == 200
        suggestions = resp.json()["suggestions"]
        paths = {item["path"] for item in suggestions}
        assert str(base / "alpha") in paths
        assert str(base / "alpine") in paths
        assert str(base / "beta") not in paths

    def test_folder_suggestions_list_children_for_trailing_separator(self, client, tmp_path):
        base = tmp_path / "folders"
        base.mkdir()
        (base / "alpha").mkdir()
        (base / "beta").mkdir()

        resp = client.get("/api/folders/suggest", params={"query": f"{base}{os.sep}"})

        assert resp.status_code == 200
        suggestions = resp.json()["suggestions"]
        names = {item["name"] for item in suggestions}
        assert names == {"alpha", "beta"}

    def test_folder_suggestions_return_empty_for_missing_parent(self, client, tmp_path):
        resp = client.get("/api/folders/suggest", params={"query": str(tmp_path / "missing" / "al")})

        assert resp.status_code == 200
        assert resp.json()["suggestions"] == []


class TestUploadImages:
    def test_upload_images_copies_files_into_loaded_folder(self, client, img_dir):
        files = [
            ("files", make_upload_file("fresh-a.jpg", color="green")),
            ("files", make_upload_file("fresh-b.png", color="yellow")),
        ]

        resp = client.post("/api/images/upload", data={"folder": str(img_dir)}, files=files)
        assert resp.status_code == 200

        data = resp.json()
        assert data["ok"] is True
        assert data["uploaded_count"] == 2
        assert data["skipped_count"] == 0
        uploaded_names = {item["name"] for item in data["uploaded"]}
        assert uploaded_names == {"fresh-a.jpg", "fresh-b.png"}
        assert (img_dir / "fresh-a.jpg").exists()
        assert (img_dir / "fresh-b.png").exists()

    def test_upload_images_auto_renames_conflicts(self, client, img_dir):
        existing_path = img_dir / "photo1.jpg"
        before_bytes = existing_path.read_bytes()

        resp = client.post(
            "/api/images/upload",
            data={"folder": str(img_dir)},
            files=[("files", make_upload_file("photo1.jpg", color="purple"))],
        )
        assert resp.status_code == 200

        data = resp.json()
        assert data["uploaded_count"] == 1
        assert data["renamed_count"] == 1
        uploaded = data["uploaded"][0]
        assert uploaded["source_name"] == "photo1.jpg"
        assert uploaded["name"] == "photo1 (1).jpg"
        assert uploaded["renamed"] is True
        assert (img_dir / "photo1 (1).jpg").exists()
        assert existing_path.read_bytes() == before_bytes

    def test_upload_images_accepts_video_files(self, client, img_dir):
        resp = client.post(
            "/api/images/upload",
            data={"folder": str(img_dir)},
            files=[("files", make_upload_file("clip.mp4"))],
        )
        assert resp.status_code == 200

        data = resp.json()
        assert data["uploaded_count"] == 1
        uploaded = data["uploaded"][0]
        assert uploaded["name"] == "clip.mp4"
        assert uploaded["media_type"] == "video"
        assert (img_dir / "clip.mp4").exists()

    def test_upload_images_skips_unsupported_files(self, client, img_dir):
        resp = client.post(
            "/api/images/upload",
            data={"folder": str(img_dir)},
            files=[
                ("files", make_upload_file("valid.webp", color="orange")),
                ("files", ("notes.txt", BytesIO(b"not-an-image"), "text/plain")),
            ],
        )
        assert resp.status_code == 200

        data = resp.json()
        assert data["uploaded_count"] == 1
        assert data["skipped_count"] == 1
        assert data["skipped"][0]["name"] == "notes.txt"
        assert data["skipped"][0]["reason"] == "Unsupported media file"
        assert (img_dir / "valid.webp").exists()
        assert not (img_dir / "notes.txt").exists()

    def test_upload_images_skips_mask_sidecars(self, client, img_dir):
        resp = client.post(
            "/api/images/upload",
            data={"folder": str(img_dir)},
            files=[("files", make_upload_file("photo1.jpg.mask.png", color="gray"))],
        )
        assert resp.status_code == 200
        data = resp.json()
        assert data["uploaded_count"] == 0
        assert data["skipped_count"] == 1
        assert data["skipped"][0]["name"] == "photo1.jpg.mask.png"
        assert not (img_dir / "photo1.jpg.mask.png").exists()

    def test_upload_images_rejects_invalid_folder(self, client):
        resp = client.post(
            "/api/images/upload",
            data={"folder": "/no/such/folder"},
            files=[("files", make_upload_file("fresh.jpg"))],
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Not a valid directory"


class TestVideoAPI:
    def test_get_video_meta(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {"width": 1920, "height": 1080, "duration": 12.5, "fps": 24.0})

        resp = client.get("/api/video/meta", params={"path": str(video_path)})
        assert resp.status_code == 200
        data = resp.json()
        assert data["width"] == 1920
        assert data["height"] == 1080
        assert data["duration"] == 12.5
        assert data["fps"] == 24.0
        assert data["mask_keyframes"] == []

    def test_get_video_meta_includes_mask_keyframes(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        Image.new("L", (16, 16), color=96).save(_get_video_mask_path(str(video_path), 12))
        Image.new("L", (16, 16), color=160).save(_get_video_mask_path(str(video_path), 48))
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {"width": 1920, "height": 1080, "duration": 12.5, "fps": 24.0})

        resp = client.get("/api/video/meta", params={"path": str(video_path)})

        assert resp.status_code == 200
        data = resp.json()
        assert data["mask_keyframes"] == [12, 48]

    def test_get_video_frame(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_resolve_ffmpeg_binaries", lambda cfg: ("ffmpeg", "ffprobe"))
        monkeypatch.setattr(server, "_extract_video_frame", lambda *args, **kwargs: b"jpeg-bytes")

        resp = client.get("/api/video/frame", params={"path": str(video_path), "time_seconds": 1.25})
        assert resp.status_code == 200
        assert resp.content == b"jpeg-bytes"
        assert resp.headers["content-type"].startswith("image/jpeg")

    def test_enqueue_video_crop_job(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_enqueue_video_job", lambda job: {"id": "job-1", **job, "status": "queued"})

        resp = client.post("/api/video/jobs/crop", json={
            "video_path": str(video_path),
            "crop": {"x": 10, "y": 12, "w": 300, "h": 200, "ratio": "3:2"},
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["job"]["type"] == "crop"
        assert data["job"]["video_path"] == str(video_path)
        assert Path(data["job"]["output_path"]).name.startswith("clip__crop_")

    def test_enqueue_video_clip_job(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_enqueue_video_job", lambda job: {"id": "job-2", **job, "status": "queued"})

        resp = client.post("/api/video/jobs/clip", json={
            "video_path": str(video_path),
            "start_seconds": 1.5,
            "end_seconds": 4.0,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["job"]["type"] == "clip"
        assert "clip" in data["job"]["id"] or data["job"]["id"] == "job-2"
        assert Path(data["job"]["output_path"]).name.startswith("clip__")

    def test_enqueue_video_clip_job_accepts_crop(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_enqueue_video_job", lambda job: {"id": "job-3", **job, "status": "queued"})

        resp = client.post("/api/video/jobs/clip", json={
            "video_path": str(video_path),
            "start_seconds": 1.5,
            "end_seconds": 4.0,
            "crop": {"x": 10, "y": 12, "w": 300, "h": 200, "ratio": "3:2"},
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["job"]["type"] == "clip"
        assert data["job"]["crop"]["w"] == 300
        assert "__crop_" in Path(data["job"]["output_path"]).name

    def test_enqueue_video_clip_job_rejects_invalid_range(self, client, img_dir):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")

        resp = client.post("/api/video/jobs/clip", json={
            "video_path": str(video_path),
            "start_seconds": 5.0,
            "end_seconds": 5.0,
        })
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Clip end time must be greater than start time"

    def test_enqueue_gif_convert_job(self, client, img_dir, monkeypatch):
        gif_path = img_dir / "loop.gif"
        Image.new("RGB", (24, 24), color="purple").save(gif_path, format="GIF")
        monkeypatch.setattr(server, "_enqueue_video_job", lambda job: {"id": "job-gif", **job, "status": "queued"})

        resp = client.post("/api/media/jobs/convert-gif", json={"media_path": str(gif_path)})

        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["job"]["type"] == "gif_to_mp4"
        assert data["job"]["video_path"] == str(gif_path)
        assert Path(data["job"]["output_path"]).name == "loop.mp4"


class TestDeleteImages:
    def test_delete_image_removes_sidecars_and_crop_state(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        caption_path = img_dir / "photo1.txt"
        caption_path.write_text("caption", encoding="utf-8")
        mask_path = server._get_image_mask_path(image_path)
        server._write_default_image_mask(image_path)
        backup_path = Path(server._get_crop_backup_path(image_path))
        backup_path.parent.mkdir(parents=True, exist_ok=True)
        backup_path.write_bytes(b"backup")

        cfg = server._load_config()
        cfg["image_crops"] = {server._normalize_image_key(image_path): {"x": 1, "y": 2, "w": 3, "h": 4}}
        server._save_config(cfg)

        resp = client.post("/api/images/delete", json={"image_paths": [image_path]})
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["deleted_count"] == 1
        assert data["errors"] == []
        assert not Path(image_path).exists()
        assert not caption_path.exists()
        assert not mask_path.exists()
        assert not backup_path.exists()
        assert server._normalize_image_key(image_path) not in server._load_config()["image_crops"]

    def test_delete_video_removes_video_mask_sidecars(self, client, img_dir):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        caption_path = img_dir / "clip.txt"
        caption_path.write_text("caption", encoding="utf-8")
        mask_a = _get_video_mask_path(str(video_path), 0)
        mask_b = _get_video_mask_path(str(video_path), 12)
        Image.new("L", (16, 16), color=96).save(mask_a)
        Image.new("L", (16, 16), color=160).save(mask_b)

        resp = client.post("/api/images/delete", json={"image_paths": [str(video_path)]})

        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["deleted_count"] == 1
        assert not video_path.exists()
        assert not caption_path.exists()
        assert not mask_a.exists()
        assert not mask_b.exists()

    def test_delete_images_requires_paths(self, client):
        resp = client.post("/api/images/delete", json={"image_paths": []})
        assert resp.status_code == 400
        assert resp.json()["detail"] == "No image paths provided"

    def test_delete_image_during_auto_caption_does_not_recreate_caption(self, client, img_dir, monkeypatch):
        image_path = str(img_dir / "photo1.jpg")
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Moon"], "groups": []}],
            "ollama_enable_free_text": False,
        })

        call_count = 0

        def fake_generate(host, payload, timeout=120):
            nonlocal call_count
            call_count += 1
            if call_count == 1:
                Path(image_path).unlink()
            return {"response": "YES"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption/stream", json={
            "image_paths": [image_path],
            "model": "llava",
            "enable_free_text": False,
        })

        assert resp.status_code == 200
        events = [json.loads(line) for line in resp.text.splitlines() if line.strip()]
        assert any(
            event["type"] == "error" and event["message"] == "Media was deleted during auto caption"
            for event in events
        )
        assert events[-1]["type"] == "done"
        assert events[-1]["processed"] == 0
        assert events[-1]["errors"] == 1
        assert not (img_dir / "photo1.txt").exists()


class TestOpenInExplorer:
    def test_linux_prefers_selection_aware_command(self, client, single_image, monkeypatch):
        calls = []

        monkeypatch.setattr(server.platform, "system", lambda: "Linux")
        monkeypatch.setattr(server.shutil, "which", lambda name: "cmd" if name == "dbus-send" else None)

        def fake_popen(command):
            calls.append(command)
            return object()

        monkeypatch.setattr(server.subprocess, "Popen", fake_popen)

        resp = client.get("/api/open-in-explorer", params={"path": single_image})
        assert resp.status_code == 200
        assert calls
        assert calls[0][0] == "dbus-send"
        assert "org.freedesktop.FileManager1.ShowItems" in calls[0]

    def test_open_file_uses_default_handler(self, client, single_image, monkeypatch):
        calls = []

        monkeypatch.setattr(server.platform, "system", lambda: "Linux")

        def fake_popen(command):
            calls.append(command)
            return object()

        monkeypatch.setattr(server.subprocess, "Popen", fake_popen)

        resp = client.get("/api/open-file", params={"path": single_image})
        assert resp.status_code == 200
        assert calls == [["xdg-open", os.path.normpath(single_image)]]


class TestCloneFolder:
    def test_clone_whole_folder_stream_copies_files_and_config(self, client, img_dir):
        source_folder = str(img_dir)
        (img_dir / "photo1.txt").write_text("caption", encoding="utf-8")
        server._write_default_image_mask(str(img_dir / "photo1.jpg"))
        (img_dir / "notes.md").write_text("hello", encoding="utf-8")
        client.post("/api/settings", json={
            "folder": source_folder,
            "sections": [{"name": "", "sentences": ["bright"], "groups": []}],
        })

        with client.stream(
            "POST",
            "/api/folder/clone/stream",
            json={
                "source_folder": source_folder,
                "new_folder_name": "images-copy",
                "image_paths": [],
            },
        ) as response:
            assert response.status_code == 200
            events = [json.loads(line) for line in response.iter_lines() if line]

        done = next(event for event in events if event.get("type") == "done")
        target_folder = Path(done["target_folder"])
        assert done["mode"] == "folder"
        assert target_folder.exists()
        assert (target_folder / "photo1.jpg").exists()
        assert (target_folder / "photo1.txt").exists()
        assert (target_folder / "photo1.jpg.mask.png").exists()
        assert (target_folder / "notes.md").exists()

        cloned_settings = client.get("/api/settings", params={"folder": str(target_folder)}).json()
        assert cloned_settings["sections"][0]["sentences"] == ["bright"]

    def test_clone_selected_images_stream_copies_subset_and_config(self, client, img_dir):
        source_folder = str(img_dir)
        selected_a = str(img_dir / "photo1.jpg")
        selected_b = str(img_dir / "photo2.png")
        (img_dir / "photo1.txt").write_text("caption one", encoding="utf-8")
        server._write_default_image_mask(selected_a)
        (img_dir / "photo3.txt").write_text("caption three", encoding="utf-8")
        client.post("/api/settings", json={
            "folder": source_folder,
            "sections": [{"name": "Scene", "sentences": ["indoor"], "groups": []}],
        })

        with client.stream(
            "POST",
            "/api/folder/clone/stream",
            json={
                "source_folder": source_folder,
                "new_folder_name": "images-subset",
                "image_paths": [selected_a, selected_b],
            },
        ) as response:
            assert response.status_code == 200
            events = [json.loads(line) for line in response.iter_lines() if line]

        done = next(event for event in events if event.get("type") == "done")
        target_folder = Path(done["target_folder"])
        assert done["mode"] == "selected"
        assert (target_folder / "photo1.jpg").exists()
        assert (target_folder / "photo2.png").exists()
        assert (target_folder / "photo1.txt").exists()
        assert (target_folder / "photo1.jpg.mask.png").exists()
        assert not (target_folder / "photo3.jpg").exists()
        assert not (target_folder / "photo3.txt").exists()

        cloned_settings = client.get("/api/settings", params={"folder": str(target_folder)}).json()
        assert cloned_settings["sections"][0]["name"] == "Scene"

    def test_clone_selected_media_copies_video_mask_sidecars(self, client, img_dir):
        source_folder = str(img_dir)
        video_path = img_dir / "clip.mp4"
        extra_video_path = img_dir / "clip-b.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        extra_video_path.write_bytes(b"fake-video-bytes")
        mask_a = _get_video_mask_path(str(video_path), 0)
        mask_b = _get_video_mask_path(str(video_path), 18)
        Image.new("L", (16, 16), color=96).save(mask_a)
        Image.new("L", (16, 16), color=160).save(mask_b)

        with client.stream(
            "POST",
            "/api/folder/clone/stream",
            json={
                "source_folder": source_folder,
                "new_folder_name": "video-subset",
                "image_paths": [str(video_path), str(extra_video_path)],
            },
        ) as response:
            assert response.status_code == 200
            events = [json.loads(line) for line in response.iter_lines() if line]

        done = next(event for event in events if event.get("type") == "done")
        target_folder = Path(done["target_folder"])
        assert (target_folder / "clip.mp4").exists()
        assert (target_folder / "clip-b.mp4").exists()
        assert (target_folder / mask_a.name).exists()
        assert (target_folder / mask_b.name).exists()


class TestDuplicateImage:
    def test_duplicate_image_copies_caption_and_mask(self, client, img_dir):
        image_path = str(img_dir / "photo1.jpg")
        (img_dir / "photo1.txt").write_text("caption one", encoding="utf-8")
        server._write_default_image_mask(image_path)

        resp = client.post("/api/image/duplicate", json={
            "image_path": image_path,
            "new_name": "photo1-copy",
        })

        assert resp.status_code == 200
        data = resp.json()
        duplicated_path = Path(data["image_path"])
        assert duplicated_path.name == "photo1-copy.jpg"
        assert duplicated_path.exists()
        assert (img_dir / "photo1-copy.txt").exists()
        assert (img_dir / "photo1-copy.jpg.mask.png").exists()


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

    def test_delete_caption_preset_updates_all_caption_files(self, client, img_dir):
        image_paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{"name": "", "sentences": ["Old Caption", "Keep Me"], "groups": []}],
        })
        for image_path in image_paths:
            client.post("/api/caption/save", json={
                "image_path": image_path,
                "enabled_sentences": ["Old Caption", "Keep Me"],
                "free_text": "Old Caption",
            })

        resp = client.post("/api/caption/delete-preset", json={
            "folder": str(img_dir),
            "sentence": "Old Caption",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["removed_sentences"] == ["Old Caption"]
        assert data["sections"][0]["sentences"] == ["Keep Me"]

        for image_path in image_paths:
            caption = Path(image_path).with_suffix(".txt").read_text(encoding="utf-8")
            assert "Old Caption" not in caption
            assert "Keep Me" in caption

    def test_delete_group_updates_all_caption_files(self, client, img_dir):
        image_paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [{
                "name": "",
                "sentences": ["Keep Me"],
                "groups": [{"name": "Car", "sentences": ["Red Car", "Blue Car"]}],
            }],
        })
        for image_path in image_paths:
            client.post("/api/caption/save", json={
                "image_path": image_path,
                "enabled_sentences": ["Keep Me", "Red Car"],
                "free_text": "Red Car",
            })

        resp = client.post("/api/group/delete", json={
            "folder": str(img_dir),
            "section_index": 0,
            "group_index": 0,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["removed_sentences"] == ["Red Car", "Blue Car"]
        assert data["sections"][0]["groups"] == []

        for image_path in image_paths:
            caption = Path(image_path).with_suffix(".txt").read_text(encoding="utf-8")
            assert "Red Car" not in caption
            assert "Blue Car" not in caption
            assert "Keep Me" in caption

    def test_delete_section_updates_all_caption_files(self, client, img_dir):
        image_paths = [str(img_dir / "photo1.jpg"), str(img_dir / "photo2.png")]
        client.post("/api/settings", json={
            "folder": str(img_dir),
            "sections": [
                {"name": "Scene", "sentences": ["Moon"], "groups": []},
                {"name": "", "sentences": ["Keep Me"], "groups": []},
            ],
        })
        for image_path in image_paths:
            client.post("/api/caption/save", json={
                "image_path": image_path,
                "enabled_sentences": ["Moon", "Keep Me"],
                "free_text": "Moon",
            })

        resp = client.post("/api/section/delete", json={
            "folder": str(img_dir),
            "section_index": 0,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["removed_sentences"] == ["Moon"]
        assert [section["name"] for section in data["sections"]] == [""]

        for image_path in image_paths:
            caption = Path(image_path).with_suffix(".txt").read_text(encoding="utf-8")
            assert "Moon" not in caption
            assert "Keep Me" in caption


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
        assert isinstance(data["video_training_presets"], list)
        assert len(data["video_training_presets"]) >= 1
        assert data["https_certfile"] == ""
        assert data["https_keyfile"] == ""
        assert data["https_port"] == 8900
        assert data["remote_http_mode"] == "redirect-to-https"
        assert data["ffmpeg_path"] == ""
        assert data["ffmpeg_threads"] >= 1
        assert data["ffmpeg_hwaccel"] == "auto"
        assert data["processing_reserved_cores"] >= 0
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

    def test_save_and_load_https_settings(self, client):
        client.post("/api/settings", json={
            "https_certfile": "certs/localhost.pem",
            "https_keyfile": "certs/localhost-key.pem",
            "https_port": 9443,
            "remote_http_mode": "block",
        })
        resp = client.get("/api/settings")
        data = resp.json()
        assert data["https_certfile"] == "certs/localhost.pem"
        assert data["https_keyfile"] == "certs/localhost-key.pem"
        assert data["https_port"] == 9443
        assert data["remote_http_mode"] == "block"

    def test_save_and_load_ffmpeg_path(self, client):
        client.post("/api/settings", json={"ffmpeg_path": "/usr/bin/ffmpeg"})
        resp = client.get("/api/settings")
        assert resp.status_code == 200
        assert resp.json()["ffmpeg_path"] == "/usr/bin/ffmpeg"

    def test_save_and_load_ffmpeg_processing_settings(self, client):
        client.post("/api/settings", json={
            "ffmpeg_threads": 12,
            "ffmpeg_hwaccel": "off",
            "processing_reserved_cores": 3,
        })
        resp = client.get("/api/settings")
        assert resp.status_code == 200
        data = resp.json()
        assert data["ffmpeg_threads"] == 12
        assert data["ffmpeg_hwaccel"] == "off"
        assert data["processing_reserved_cores"] == 3

    def test_get_ollama_models(self, client, monkeypatch):
        monkeypatch.setattr(server, "_list_ollama_models", lambda host, timeout=20: ["llava:latest", "qwen2.5vl"])
        resp = client.get("/api/ollama/models", params={"server": "localhost", "port": 11435})
        assert resp.status_code == 200
        data = resp.json()
        assert data["host"] == "http://localhost:11435"
        assert data["models"] == ["llava:latest", "qwen2.5vl"]

    def test_get_ollama_models_error(self, client, monkeypatch):
        def fail(host, timeout=20):
            raise RuntimeError("connection refused")

        monkeypatch.setattr(server, "_list_ollama_models", fail)
        resp = client.get("/api/ollama/models")
        assert resp.status_code == 502
        assert "Ollama error" in resp.json()["detail"]

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

    def test_save_and_load_video_training_presets_and_folder_profile(self, client, tmp_path):
        folder = str(tmp_path / "video_folder")
        os.makedirs(folder, exist_ok=True)
        presets = [
            {
                "key": "wan-fast",
                "label": "Wan Fast",
                "target_family": "wan",
                "num_frames": 40,
                "fps": 16,
                "preferred_extensions": ["mp4", ".mov"],
            },
            {
                "label": "Hunyuan Base",
                "target_family": "hunyuan",
                "num_frames": 129,
                "fps": 24,
                "preferred_extensions": [".mp4"],
                "description": "Longer Hunyuan motion clips.",
            },
        ]

        resp = client.post("/api/settings", json={
            "video_training_presets": presets,
            "folder": folder,
            "video_training_profile_key": "hunyuan-base",
        })

        assert resp.status_code == 200

        data = client.get("/api/settings", params={"folder": folder}).json()
        assert [preset["key"] for preset in data["video_training_presets"]] == ["wan-fast", "hunyuan-base"]
        assert data["video_training_profile_key"] == "hunyuan-base"
        assert data["video_training_profile"]["target_family"] == "hunyuan"
        assert data["video_training_profile"]["ideal_clip_seconds"] == pytest.approx(5.375)
        assert data["video_training_profile"]["preferred_extensions"] == [".mp4"]

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

    def test_remote_http_redirects_to_https_for_other_computers(self, client, monkeypatch):
        monkeypatch.setattr(server, "_load_config", lambda: {
            "https_certfile": "certs/localhost.pem",
            "https_keyfile": "certs/localhost-key.pem",
            "https_port": 9443,
            "remote_http_mode": "redirect-to-https",
        })
        monkeypatch.setattr(server, "_is_local_client_host", lambda host: False)

        resp = client.get("/", headers={"host": "192.168.0.50:8899"}, follow_redirects=False)
        assert resp.status_code == 307
        assert resp.headers["location"] == "https://192.168.0.50:9443/"

    def test_remote_http_can_be_blocked_for_other_computers(self, client, monkeypatch):
        monkeypatch.setattr(server, "_load_config", lambda: {
            "https_certfile": "certs/localhost.pem",
            "https_keyfile": "certs/localhost-key.pem",
            "https_port": 9443,
            "remote_http_mode": "block",
        })
        monkeypatch.setattr(server, "_is_local_client_host", lambda host: False)

        resp = client.get("/", headers={"host": "192.168.0.50:8899"}, follow_redirects=False)
        assert resp.status_code == 403
        assert "Use HTTPS instead" in resp.text

    def test_local_http_remains_allowed_with_https_enabled(self, client, monkeypatch):
        monkeypatch.setattr(server, "_load_config", lambda: {
            "https_certfile": "certs/localhost.pem",
            "https_keyfile": "certs/localhost-key.pem",
            "https_port": 9443,
            "remote_http_mode": "redirect-to-https",
        })
        monkeypatch.setattr(server, "_is_local_client_host", lambda host: True)

        resp = client.get("/", follow_redirects=False)
        assert resp.status_code == 200


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
        prompt = server._ollama_prompt_for_caption("Moon", "Caption check: {caption} / {sentence}")
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
            assert payload["images"] == ["image-bytes"]
            prompt = payload["prompt"]
            if "Moon" in prompt:
                return {"response": "YES"}
            return {"response": "NO"}

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["image-bytes"])
        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        enabled, results = server._auto_caption_captions(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            ["Moon", "Night", "Car"],
            encode_image_func=server._encode_media_for_ollama,
            generate_func=server._ollama_generate,
            prompt_template="Caption? {caption}",
            timeout=17,
        )
        assert enabled == ["Moon"]
        assert len(results) == 3
        assert results[0]["enabled"] is True
        assert results[1]["enabled"] is False

    def test_auto_caption_sections(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            assert payload["images"] == ["image-bytes"]
            prompt = payload["prompt"]
            if "Caption:" in prompt:
                return {"response": "YES" if "Moon" in prompt else "NO"}
            if "Group:" in prompt:
                return {"response": "2"}
            raise AssertionError("unexpected prompt")

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["image-bytes"])
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
            encode_image_func=server._encode_media_for_ollama,
            generate_func=server._ollama_generate,
            prompt_template="Caption: {caption}",
            group_prompt_template="Group: {group_name}\n{options}",
            timeout=15,
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

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["image-bytes"])
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
            encode_image_func=server._encode_media_for_ollama,
            generate_func=server._ollama_generate,
            prompt_template="Caption: {caption}",
            group_prompt_template="Group: {group_name}\n{options}",
            timeout=15,
        )
        assert enabled == ["Chair not in frame"]
        assert results[0]["selected_sentence"] == "Chair not in frame"
        assert results[0]["selected_hidden"] is True

    def test_suggest_free_text(self, single_image, monkeypatch):
        def fake_generate(host, payload, timeout=120):
            assert timeout == 9
            assert "Current caption text" in payload["prompt"]
            assert payload["images"] == ["image-bytes"]
            return {"response": "Moon visible\nBright stars"}

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["image-bytes"])
        monkeypatch.setattr(server, "_ollama_generate", fake_generate)
        result = server._suggest_free_text(
            "http://127.0.0.1:11434",
            "llava",
            single_image,
            "- Night",
            encode_image_func=server._encode_media_for_ollama,
            generate_func=server._ollama_generate,
            prompt_template=server.DEFAULT_OLLAMA_FREE_TEXT_PROMPT_TEMPLATE,
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


class TestMaskAPI:
    def test_get_mask_metadata_without_creation(self, client, single_image):
        resp = client.get("/api/mask", params={"path": single_image})
        assert resp.status_code == 200
        data = resp.json()
        assert data["exists"] is False
        assert data["created"] is False
        assert data["image_width"] == 100
        assert data["image_height"] == 100

    def test_get_mask_with_ensure_creates_default_gray_mask(self, client, single_image):
        resp = client.get("/api/mask", params={"path": single_image, "ensure": True})
        assert resp.status_code == 200
        data = resp.json()
        assert data["exists"] is True
        assert data["created"] is True

        mask_path = Path(data["path"])
        assert mask_path.exists()
        with Image.open(mask_path) as mask_image:
            assert mask_image.mode == "L"
            assert mask_image.size == (100, 100)
            assert mask_image.getpixel((50, 50)) == 128

    def test_get_mask_image_serves_png(self, client, single_image):
        client.get("/api/mask", params={"path": single_image, "ensure": True})
        resp = client.get("/api/mask/image", params={"path": single_image})
        assert resp.status_code == 200
        assert resp.headers["content-type"] == "image/png"

    def test_save_mask_normalizes_uploaded_image(self, client, single_image):
        buffer = BytesIO()
        Image.new("RGB", (40, 20), color=(255, 255, 255)).save(buffer, format="PNG")
        buffer.seek(0)

        resp = client.post(
            "/api/mask",
            data={"image_path": single_image},
            files={"mask": ("mask.png", buffer, "image/png")},
        )
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True

        mask_path = Path(data["path"])
        with Image.open(mask_path) as mask_image:
            assert mask_image.mode == "L"
            assert mask_image.size == (100, 100)
            assert mask_image.getpixel((50, 50)) == 255

    def test_get_video_mask_with_ensure_creates_keyframe_mask(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_resolve_ffmpeg_binaries", lambda cfg: ("ffmpeg", "ffprobe"))
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {
            "width": 96,
            "height": 54,
            "duration": 5.0,
            "fps": 24.0,
        })

        resp = client.get("/api/mask", params={"path": str(video_path), "frame_index": 48, "ensure": True})

        assert resp.status_code == 200
        data = resp.json()
        assert data["media_type"] == "video"
        assert data["frame_index"] == 48
        assert data["requested_frame_index"] == 48
        assert data["created"] is True
        assert data["mask_count"] == 1
        mask_path = Path(data["path"])
        assert mask_path.exists()
        with Image.open(mask_path) as mask_image:
            assert mask_image.mode == "L"
            assert mask_image.size == (96, 54)
            assert mask_image.getpixel((10, 10)) == 128

    def test_get_video_mask_reuses_previous_keyframe_without_create_new(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        previous_mask_path = _get_video_mask_path(str(video_path), 24)
        Image.new("L", (96, 54), color=200).save(previous_mask_path)
        monkeypatch.setattr(server, "_resolve_ffmpeg_binaries", lambda cfg: ("ffmpeg", "ffprobe"))
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {
            "width": 96,
            "height": 54,
            "duration": 5.0,
            "fps": 24.0,
        })

        resp = client.get("/api/mask", params={"path": str(video_path), "frame_index": 48})

        assert resp.status_code == 200
        data = resp.json()
        assert data["frame_index"] == 24
        assert data["requested_frame_index"] == 48
        assert data["source_frame_index"] == 24
        assert data["path"] == str(previous_mask_path)
        assert data["exists"] is True
        assert data["mask_count"] == 1

    def test_get_video_mask_with_create_new_clones_previous_keyframe(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        previous_mask_path = _get_video_mask_path(str(video_path), 24)
        Image.new("L", (96, 54), color=200).save(previous_mask_path)
        monkeypatch.setattr(server, "_resolve_ffmpeg_binaries", lambda cfg: ("ffmpeg", "ffprobe"))
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {
            "width": 96,
            "height": 54,
            "duration": 5.0,
            "fps": 24.0,
        })

        resp = client.get("/api/mask", params={
            "path": str(video_path),
            "frame_index": 48,
            "ensure": True,
            "create_new": True,
        })

        assert resp.status_code == 200
        data = resp.json()
        assert data["frame_index"] == 48
        assert data["requested_frame_index"] == 48
        assert data["source_frame_index"] == 24
        assert data["created"] is True
        assert data["mask_count"] == 2
        new_mask_path = Path(data["path"])
        assert new_mask_path.exists()
        with Image.open(new_mask_path) as mask_image:
            assert mask_image.size == (96, 54)
            assert mask_image.getpixel((10, 10)) == 200

    def test_save_video_mask_normalizes_uploaded_image(self, client, img_dir, monkeypatch):
        video_path = img_dir / "clip.mp4"
        video_path.write_bytes(b"fake-video-bytes")
        monkeypatch.setattr(server, "_resolve_ffmpeg_binaries", lambda cfg: ("ffmpeg", "ffprobe"))
        monkeypatch.setattr(server, "_probe_video_info", lambda path, ffprobe_path="ffprobe": {
            "width": 96,
            "height": 54,
            "duration": 5.0,
            "fps": 24.0,
        })

        buffer = BytesIO()
        Image.new("RGB", (40, 20), color=(255, 255, 255)).save(buffer, format="PNG")
        buffer.seek(0)

        resp = client.post(
            "/api/mask",
            data={"media_path": str(video_path), "frame_index": 12},
            files={"mask": ("mask.png", buffer, "image/png")},
        )

        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["media_type"] == "video"
        assert data["frame_index"] == 12
        assert data["mask_count"] == 1
        mask_path = Path(data["path"])
        with Image.open(mask_path) as mask_image:
            assert mask_image.mode == "L"
            assert mask_image.size == (96, 54)
            assert mask_image.getpixel((10, 10)) == 255


class TestImageEditAPI:
    def test_save_image_edit_normalizes_uploaded_image(self, client, single_image):
        buffer = BytesIO()
        Image.new("RGB", (32, 24), color=(24, 200, 80)).save(buffer, format="PNG")
        buffer.seek(0)

        resp = client.post(
            "/api/image/edit",
            data={"image_path": single_image},
            files={"image": ("edited.png", buffer, "image/png")},
        )

        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["image_width"] == 100
        assert data["image_height"] == 100
        with Image.open(single_image) as edited_image:
            assert edited_image.size == (100, 100)
            pixel = edited_image.convert("RGB").getpixel((50, 50))
            assert pixel[1] > pixel[0]
            assert pixel[1] > pixel[2]


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

        def fake_auto_caption(host, model, image_path, sections, **kwargs):
            assert host == "http://localhost:11435"
            assert model == "llava"
            assert image_path == single_image
            assert server._all_captions_from_sections(sections) == ["Moon", "Night", "Red Car", "Blue Car"]
            assert kwargs["encode_image_func"].func is server._encode_media_for_ollama
            assert kwargs["generate_func"] is server._ollama_generate
            assert kwargs["prompt_template"] == "Caption: {caption}"
            assert kwargs["group_prompt_template"] == "Group: {group_name}\n{options}"
            assert kwargs["timeout"] == 12
            return ["Moon", "Night", "Blue Car"], [
                {"sentence": "Moon", "enabled": True, "answer": "YES"},
                {"sentence": "Night", "enabled": True, "answer": "YES"},
                {"type": "group", "group_name": "Car", "selected_sentence": "Blue Car", "selected_hidden": False, "selection_index": 2, "answer": "2"},
            ]

        monkeypatch.setattr(server, "_auto_caption_sections", fake_auto_caption)
        monkeypatch.setattr(server, "_suggest_free_text", lambda *args, **kwargs: "Moon visible\nNight sky")

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

    def test_auto_caption_video_uses_sampled_frames(self, client, img_dir, monkeypatch):
        video_path = str(img_dir / "clip.mp4")
        Path(video_path).write_bytes(b"fake-video-bytes")
        folder = str(img_dir)
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["Walking"], "groups": []}],
            "ollama_model": "llava",
            "ollama_enable_free_text": False,
        })

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["frame-a", "frame-b", "frame-c"])

        def fake_generate(host, payload, timeout=120):
            assert payload["images"] == ["frame-a", "frame-b", "frame-c"]
            return {"response": "YES"}

        monkeypatch.setattr(server, "_ollama_generate", fake_generate)

        resp = client.post("/api/auto-caption", json={
            "media_path": video_path,
            "model": "llava",
            "enable_free_text": False,
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["enabled_sentences"] == ["Walking"]
        assert Path(video_path).with_suffix(".txt").read_text(encoding="utf-8").strip() == "Walking"

    def test_auto_caption_stream_accepts_video_media_paths(self, client, img_dir, monkeypatch):
        video_path = str(img_dir / "clip.mp4")
        Path(video_path).write_bytes(b"fake-video-bytes")
        folder = str(img_dir)
        client.post("/api/settings", json={
            "folder": folder,
            "sections": [{"name": "", "sentences": ["Walking"], "groups": []}],
            "ollama_model": "llava",
            "ollama_enable_free_text": False,
        })

        monkeypatch.setattr(server, "_encode_media_for_ollama", lambda path, **kwargs: ["frame-a", "frame-b"])
        monkeypatch.setattr(server, "_ollama_generate", lambda host, payload, timeout=120: {"response": "YES"})

        with client.stream(
            "POST",
            "/api/auto-caption/stream",
            json={"media_paths": [video_path], "model": "llava", "enable_free_text": False},
        ) as response:
            assert response.status_code == 200
            events = [json.loads(line) for line in response.iter_lines() if line]

        assert events[0]["type"] == "start"
        completed = [event for event in events if event.get("type") == "image-complete"]
        assert len(completed) == 1
        assert completed[0]["enabled_sentences"] == ["Walking"]

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

        def fake_auto_caption(host, model, image_path, sections, **kwargs):
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

        monkeypatch.setattr(server, "_suggest_free_text", lambda *args, **kwargs: "Bright stars")

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
