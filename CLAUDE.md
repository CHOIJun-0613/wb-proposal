# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Python automation tool for processing PPTX files (우리은행 AX AI Agent 구축 제안서 - Part III). Two main operations:
1. **Page numbering** — sets `firstSlideNum` across files so page numbers are continuous
2. **Merging** — combines 34 PPTX files into one via ZIP-level manipulation

## Setup

```bat
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Running

```bat
cmd\pagenum.bat      # Insert page numbers (outputs to output/YYYYMMDD-HH24MISS/)
cmd\pptmerge.bat     # Merge all PPTX into one file
runvenv.bat          # Open venv-activated CMD for manual execution
```

Or directly with venv active:

```bash
python src/add_page_numbers.py
python src/merge_pptx.py
```

## Architecture

### `src/add_page_numbers.py`
Reads each `.pptx` from `PPTX_DIR`, sets `firstSlideNum` on the presentation element (via `prs._element.set()`), saves to a timestamped subfolder under `OUTPUT_DIR`. Also generates `목차페이지정보.xlsx` with file-name / slide-count / start-page / end-page per file.

### `src/merge_pptx.py`
Uses `Merger` class that operates at the ZIP level (no python-pptx for merge, uses `zipfile` + `lxml`):
- `load_base()` — loads first PPTX as the base, tracks counters for slides, media, charts, diagrams, embeddings
- `append()` — for each subsequent PPTX, copies slides one by one via `_copy_slide()`
- `_copy_slide()` — rewrites `.rels` to point to renamed copies of all resources (media, charts, diagrams, embeddings), then registers the new slide in `presentation.xml` and `[Content_Types].xml`
- `save()` — writes all in-memory file bytes back to a new ZIP

Resource counters (`next_slide`, `next_media`, `next_chart`, `next_diag`, `next_embed`) are incremented globally to avoid name collisions across files.

## Configuration (`.env`)

| Variable | Default | Description |
|----------|---------|-------------|
| `PPTX_DIR` | `./pptx` | Source PPTX folder |
| `OUTPUT_DIR` | `./output` | Output root |
| `MERGE_OUTPUT` | `III.기술부문.pptx` | Merged output filename |
| `PAGE_NUM_FONT_SIZE` | `8` | Font size (pt) — reserved for text box rendering |
| `PAGE_NUM_COLOR` | `606060` | Font color (RGB hex) — reserved for text box rendering |
| `PAGE_NUM_MARGIN_RIGHT` | `0.3` | Right margin (inch) — reserved for text box rendering |
| `PAGE_NUM_MARGIN_BOTTOM` | `0.2` | Bottom margin (inch) — reserved for text box rendering |
| `PAGE_NUM_WIDTH` | `1.0` | Text box width (inch) — reserved for text box rendering |

> Note: `PAGE_NUM_*` variables are defined in `.env` but not yet consumed by `add_page_numbers.py` in its current form. The script currently only manipulates `firstSlideNum` on the presentation XML element.

## Key Notes

- The two scripts are independent: `merge_pptx.py` reads from `PPTX_DIR` directly, not from the page-numbered output folder. Run page numbering first if you need `firstSlideNum` set before merging.
- Source files live in `pptx/` (34 files, ~1,061 slides total); originals are never overwritten
- `pptx_backup/` holds the original backup copies
- Output always goes to a new `output/YYYYMMDD-HH24MISS/` folder
- File processing order is alphabetical sort — critical for correct page numbering sequence
- The merge script does not use python-pptx for the merge itself; it manipulates the OOXML ZIP directly to preserve all embedded assets (charts, SmartArt, OLE embeddings)
- `merge_pptx.py` always calls `set_first_slide_num(1)` on the merged output, resetting the start page to 1
