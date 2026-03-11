# Tag2 – Image Captioning Tool

A lightweight, browser-based tool for browsing images and managing text captions. Built with **FastAPI** and a single-page HTML frontend.

## Features

- Browse images from any local folder
- Thumbnail grid with adjustable sizes
- Zoomable image preview
- Predefined sentence tags (per-folder configurable)
- Free-text captions
- Batch toggle sentences across multiple selected images
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

On first use, a `config.json` file is created automatically to store your last opened folder and per-folder sentence presets. This file is gitignored.

## License

MIT
