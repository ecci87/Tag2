# Tag2 – Media Captioning Tool

A lightweight, browser-based tool for browsing images and videos and managing captions, metadata sidecars, masks, and lightweight editing tasks for AI datasets. Most useful for LoRA-style captioning and dataset cleanup before training.
Supports Ollama auto captioning using a check-per-caption approach instead of unrestricted free-text generation, leading to more precise training-friendly tags.
Built with **FastAPI** and a single-page HTML frontend.

## Features

- Fast non-interrupted browsing of images and videos from any server local folder
- Thumbnail grid with adjustable sizes plus `VID`, `TXT`, `M`, and `AR` badges
- Zoomable image preview and video preview with timeline tools
- Predefined caption tags (per-folder configurable)
- Free-text captions
- Batch toggle captions across multiple selected media files
- Metadata tab for editing `.meta.json` sidecars per file or as sparse batch updates across a multi-selection
- Auto captioning with Ollama by checking media against the predefined captions
- Video auto captioning by sampling representative frames from each clip and sending them to the vision model
- Aspect ratio checking in the thumbnail grid against the configured allowed ratios
- Real image cropping with aspect-ratio snapping and reversible undo while the server is running
- Direct image edit mode for image files, with undo/redo/reset and save-back-to-source
- Grayscale mask editing for image files and video key frames
- Latent mask preview with configurable base size, divider, noise timestep, signal stats, and a signal probe rectangle for checking retained detail
- Queued video crop and clip jobs with ffmpeg
- GIF-to-MP4 conversion jobs for selected GIF files
- Editable video training preset library with per-folder profile selection for model families such as WAN, LTX, Hunyuan, or custom setups
- In-preview video training guidance showing target clip length, recommended range, current selection length, preferred extensions, and profile notes
- Folder cloning for the whole dataset or just the current multi-selection, including captions, masks, and per-folder config
- Duplicate-image workflow that also carries caption and mask sidecars
- Drag-and-drop upload queue for adding images and videos into the loaded folder
- Caption files saved as `.txt` alongside each media file
- Metadata files saved as `.meta.json` alongside each media file
- Complex filtering logic for the dataset based on the captions and other attributes

## Editing And Workflow Notes

- Double-click a thumbnail to reveal that file in the OS file manager.
- Double-click the `TXT` badge to open the caption sidecar directly in the default OS application.
- Double-click the preview panel to reset zoom.
- Video masking is key-frame based. The `+` button in the video toolbar adds a new key-frame mask at the current frame.
- The metadata panel supports `seed`, `min_t`, `max_t`, `sampling_frequency`, `caption_dropout_enabled`, and `caption_dropout_caption` fields, which is useful for per-sample training schedules and sample-specific caption dropout stored in `.meta.json` sidecars.
- Video training presets are edited as JSON in Settings and then selected per folder. The active profile drives the clip hints shown beside each video preview.

## Ollama Notes

- Use a multimodal vision model such as `llava`.
- Video auto captioning does not send the raw video file to Ollama. Tag2 samples a small set of representative frames from each clip and sends those frames as the vision input.
- ffmpeg and ffprobe must be available for video previews, thumbnails, clipping, GIF conversion, video editing, and video auto captioning.

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
- `ffmpeg` and `ffprobe` for video thumbnails, preview extraction, clipping, crop jobs, GIF conversion, and video-related editing

## Configuration

On first use, a `config.json` file is created automatically to store:

- your last opened folder
- per-folder sentence presets
- editable video training preset library and the per-folder selected video profile
- Ollama connection and prompt settings
- ffmpeg path and video processing settings
- configured aspect ratios and mask latent preview presets

This file is gitignored.

## License

MIT
