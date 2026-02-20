# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A Python CLI tool that splits large Screaming Frog "All Inlinks" Excel exports (900 MB+, 13+ tabs) into manageable files by geographical region, URL pattern, or both. Uses a two-pass streaming architecture to avoid loading entire files into memory.

## Running the Tool

```bash
# Local
pip install -r requirements.txt
python3 splitter.py [OPTIONS]

# Docker
./run.sh [OPTIONS]
```

**CLI flags:**
- `--split {region,url,both}` — splitting strategy (default: region)
- `--url-depth N` — path segments for URL grouping (default: 2)
- `--url-pattern REGEX` — custom regex override for URL grouping
- `--no-summary` — skip Summary sheet
- `--no-filter` — keep all row types (don't filter Sitemap Hreflang / XML Sitemap)
- `--input DIR` / `--output DIR` — override default directories

**Input/Output:**
- Place Excel files in `input/` directory
- Output goes to `output/<date>/`
- Auto-detects Docker (`/app/input`, `/app/output`) vs local paths

## Architecture

### Two-Pass Streaming

**Pass 1 (`analyze_workbook`):** Streams all sheets with `read_only=True`. Builds `Counter` objects for destination frequencies and bucket membership. No row storage — only counts.

**Pass 2 (`split_workbook`):** Streams all sheets again with `read_only=True`. Routes each row to an `OutputManager` bucket which writes using `write_only=True` workbooks. Summary sheets are created at close time using Pass 1 data.

### Key Components

- **`ProgressTracker`** — prints `[Step X of Y] description` format
- **`extract_url_group(url, depth, pattern)`** — strips locale prefix, takes N path segments as grouping key
- **`OutputManager`** — lazy creation of `write_only=True` output workbooks per bucket
- **`_resolve_buckets()`** — determines bucket assignment based on split mode
- **`find_column_index()`** — case-insensitive column lookup
- **`get_matching_regions()`** — regex region detection on Source Segments column

### Split Modes

- **region**: Groups by Source Segments → APAC, EU, LAC, MEISA, USA, OTHER
- **url**: Groups by URL path pattern from Source column
- **both**: Cross-product of region × URL group

### Priority Logic

Based on destination frequency across all data:
- **HIGH**: >= 100 occurrences
- **MEDIUM**: >= 10 and < 100 occurrences
- **LOW**: < 10 occurrences

### Output File Structure

Each output file has two sheets:
1. **Summary** — destinations ranked by frequency with Priority and Impact columns
2. **Data** — all original columns plus appended Priority column

### Multi-Tab Handling

Iterates `wb.sheetnames` to process every sheet. Headers are taken from the first sheet's first row; subsequent sheets' row 0 is skipped.

## Configuration

Top of `splitter.py`:
- `IGNORED_TYPES`: Types to exclude (default: Sitemap Hreflang, XML Sitemap)
- `HIGH_THRESHOLD` / `MEDIUM_THRESHOLD`: Priority thresholds (default: 100 / 10)
- `REGIONS`: Region keywords (default: APAC, MEISA, EU, LAC, USA)

## Dependencies

- Python 3.x
- openpyxl (see `requirements.txt`)
