# Tag2 – Image Captioning Tool

A lightweight, browser-based tool for browsing images and managing text captions for AI datasets. Most useful for LoRA type captioning.
Supports Ollama auto captioning using a check per user caption approach, instead of letting AI generate a free from text leading to more precise LoRA training friendly results.
Built with **FastAPI** and a single-page HTML frontend.

## Features

- Browse images from any local folder
- Thumbnail grid with adjustable sizes
- Zoomable image preview
- Predefined caption tags (per-folder configurable)
- Free-text captionsa
- Batch toggle captions across multiple selected images
- Auto captioning with Ollama by checking images against the predefined captions
- Aspect ratio checking in the thumbnail grid against the configured allowed ratios
- Real image cropping with aspect-ratio snapping and reversible undo while the server is running
- Caption files saved as `.txt` alongside each image

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

## Requirements

- Python 3.10+
- Dependencies: FastAPI, Uvicorn, Pillow

## Configuration

On first use, a `config.json` file is created automatically to store:

- your last opened folder
- per-folder sentence presets
- Ollama connection and prompt settings
- configured aspect ratios

This file is gitignored.

## License

MIT
