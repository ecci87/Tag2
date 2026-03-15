# Tag2 – Media Captioning Tool

A lightweight, browser-based tool for browsing images and videos and managing text captions for AI datasets. Most useful for LoRA type captioning.
Supports Ollama auto captioning using a check per user caption approach, instead of letting AI generate unrestricted free text, leading to more precise LoRA-training-friendly results.
Built with **FastAPI** and a single-page HTML frontend.

## Features

- Fast non-interrupted browsing of images and videos from any server local folder
- Thumbnail grid with adjustable sizes and video badges
- Zoomable image preview and video preview with timeline tools
- Predefined caption tags (per-folder configurable)
- Free-text captions
- Batch toggle captions across multiple selected media files
- Auto captioning with Ollama by checking media against the predefined captions
- Video auto captioning by sampling representative frames from each clip and sending them to the vision model
- Aspect ratio checking in the thumbnail grid against the configured allowed ratios
- Real image cropping with aspect-ratio snapping and reversible undo while the server is running
- Queued video crop and clip jobs with ffmpeg
- Caption files saved as `.txt` alongside each media file
- Complex filtering logic for the dataset based on the captions and other attributes

## Ollama Notes

- Use a multimodal vision model such as `llava`.
- Video auto captioning does not send the raw video file to Ollama. Tag2 samples a small set of representative frames from each clip and sends those frames as the vision input.
- ffmpeg and ffprobe must be available for video previews, thumbnails, editing, and video auto captioning.

## Quick Start

### Windows

```
run.bat
```

### Linux / macOS

```bash
chmod +x run.sh
./run.sh
```

The scripts automatically create a Python virtual environment and install dependencies. Once running, open **http://localhost:8899** in your browser.

### Manual Setup

```bash
python -m venv venv
source venv/bin/activate   # or venv\Scripts\activate on Windows
pip install -r requirements.txt
python server.py
```

### Custom Port

```bash
python server.py 9000
```

### Local HTTPS

For local encrypted traffic, add certificate paths to `config.json`:

```json
{
	"https_certfile": "certs/localhost.pem",
	"https_keyfile": "certs/localhost-key.pem",
	"https_port": 8900,
	"remote_http_mode": "redirect-to-https"
}
```

Relative paths are resolved from the project folder. If both files exist, the server keeps normal HTTP on the current port and also starts HTTPS on `https_port`.

- Same computer: keep using normal HTTP on the current port, for example `http://localhost:8899`
- Different computer: use HTTPS on the configured HTTPS port, for example `https://your-pc:8900`
- Remote HTTP policy:
	`redirect-to-https` redirects other computers from HTTP to HTTPS
	`block` rejects HTTP from other computers
	`allow` leaves remote HTTP enabled

Changing these network settings requires a server restart.

If you use UFW on Linux, allow only the HTTPS port from your LAN, for example:

```bash
sudo ufw allow from 192.168.0.0/16 to any port 8900 proto tcp
```

For a trusted local certificate, `mkcert` is the easiest option:

```bash
mkcert -install
mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1
```

## Requirements

- Python 3.10+
- Dependencies: FastAPI, Uvicorn, Pillow

## Configuration

On first use, a `config.json` file is created automatically to store:

- your last opened folder
- per-folder sentence presets
- Ollama connection and prompt settings
- ffmpeg path and video processing settings
- configured aspect ratios

This file is gitignored.

## License

MIT
