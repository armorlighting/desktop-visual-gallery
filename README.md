# Desktop Visual Gallery

<img src="assets/desktop-visual-gallery-icon.png" alt="Desktop Visual Gallery icon" width="160">

A private, local-first visual file browser for Windows. It turns a folder tree into a thumbnail gallery so you can recognize a file before you remember its name.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4)
![License](https://img.shields.io/badge/License-MIT-green)

## Why

Traditional search assumes you remember a filename or keyword. People often remember files by recognition instead: a PDF cover, spreadsheet layout, slide title, image, or approximate date. Desktop Visual Gallery adds that visual retrieval layer without uploading files anywhere.

## Features

- Recursively indexes a chosen folder (Desktop by default)
- Generates local thumbnails for images, PDFs, spreadsheets, Word documents, and PowerPoint files
- Filters by file type and directory
- Sorts by modification time, name, or size
- Opens the original file in its default application
- Reveals and selects the original file in Windows Explorer
- Copies the full file path to the clipboard
- Incremental refresh: unchanged thumbnails are reused
- Listens only on `127.0.0.1`
- Rejects open/reveal/copy requests outside the indexed root
- Excludes shortcuts, `desktop.ini`, and Office lock files by default

## Quick start

```powershell
git clone https://github.com/armorlighting/desktop-visual-gallery.git
cd desktop-visual-gallery
py -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\python visual_gallery.py
```

The gallery opens at `http://127.0.0.1:8765/`.

## Choose another folder

```powershell
.\.venv\Scripts\python visual_gallery.py --root "D:\Work Files"
```

Useful options:

```text
--root PATH          Folder to index (default: current user's Desktop)
--data-dir PATH      Local generated data directory
--port NUMBER        Local port (default: 8765)
--no-browser         Do not automatically open the browser
--refresh-only       Build/update the index and exit
```

## Create a desktop shortcut

After installing dependencies, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_windows.ps1
```

The shortcut refreshes the index before opening the gallery and uses the included custom gallery icon.

## Privacy and security

- No cloud service, analytics, telemetry, or external API is used.
- File names, paths, thumbnails, and metadata remain in the local data directory.
- The HTTP server binds to loopback only (`127.0.0.1`).
- File actions are validated against the configured root directory.
- Generated data is ignored by Git and should never be committed.

The gallery displays file names and document previews. Treat the generated data directory as sensitive when indexing confidential material.

## Supported previews

| Type | Preview |
|---|---|
| Images | Scaled image |
| PDF | First page |
| XLSX/XLSM | First non-empty worksheet |
| DOCX | Title and text excerpt |
| PPTX | First-slide text excerpt |
| Other files | Type card |

Legacy `.xls`, `.doc`, video, archive, and CAD files currently use type cards.

## License

MIT
