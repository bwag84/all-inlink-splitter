# All Inlinks Splitter

Split large Screaming Frog "All Inlinks" Excel exports (900 MB+) into manageable, region- or URL-based files — directly from the command line.

## Why?

Screaming Frog crawl exports for large sites can exceed 900 MB and span 13+ Excel tabs. Most laptops can't even open them. This tool:

- **Streams** through every tab without loading the full file into memory
- **Splits** by geographical region, URL pattern, or both
- **Adds a Priority column** based on how often each destination URL appears (HIGH / MEDIUM / LOW)
- **Creates a Summary sheet** in each output file ranking destinations by impact

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Place .xlsx files in the input/ folder
cp my-crawl-export.xlsx input/

# 3. Run
python3 splitter.py
```

Output appears in `output/YYYY-MM-DD/`.

## CLI Options

```
python3 splitter.py [OPTIONS]
```

| Flag | Description | Default |
|------|-------------|---------|
| `--split {region,url,both}` | Splitting strategy | `region` |
| `--url-depth N` | Path segments used for URL grouping | `2` |
| `--url-pattern REGEX` | Custom regex for URL grouping (overrides `--url-depth`) | — |
| `--no-summary` | Skip the Summary sheet in output files | off |
| `--no-filter` | Keep all row types (don't filter Sitemap Hreflang / XML Sitemap) | off |
| `--input DIR` | Input directory | `input/` |
| `--output DIR` | Output base directory | `output/` |

## Split Modes

### Region (default)

Groups rows by the **Source Segments** column. Produces one file per region:

```
output/2025-06-15/
  demo_APAC.xlsx
  demo_EU.xlsx
  demo_LAC.xlsx
  demo_MEISA.xlsx
  demo_USA.xlsx
  demo_Canada.xlsx
  demo_OTHER.xlsx
```

Regions detected: APAC, MEISA, EU, LAC, USA, Canada. `US`, `USA`, `United States`, and `United States of America` source segments are grouped into the USA file. Rows matching no region go to OTHER. A row matching multiple regions appears in each.

### URL

Groups rows by the **Source** URL path. The `--url-depth` flag controls how many path segments are used:

```bash
python3 splitter.py --split url --url-depth 2
```

Example: `https://www.fedex.com/en-us/shipping/returns.html` → group key `shipping_returns`

The locale-like prefix (`en-us`) is automatically stripped.

For custom grouping, pass a regex:

```bash
python3 splitter.py --split url --url-pattern '/(\w+-\w+)/'
```

### Both

Cross-product of region and URL:

```bash
python3 splitter.py --split both
```

Produces files like `demo_APAC_shipping_returns.xlsx`.

## Output File Structure

Each output file has two sheets:

1. **Summary** — Destinations ranked by frequency, with Priority and Impact columns
2. **Data** — All original columns plus an appended **Priority** column

### Priority Levels

| Priority | Threshold | Meaning |
|----------|-----------|---------|
| HIGH | ≥ 100 occurrences | Quick wins — fixing one URL removes many errors |
| MEDIUM | ≥ 10 and < 100 | Moderate impact |
| LOW | < 10 | Low frequency |

## Filtering

By default, rows with Type = "Sitemap Hreflang" or "XML Sitemap" are excluded. Use `--no-filter` to keep them.

## Multi-Tab Support

Screaming Frog splits large exports across multiple Excel tabs. This tool reads **all tabs** in each file. Headers are taken from the first tab; subsequent tabs' header rows are skipped automatically.

## Docker

Build and run with Docker:

```bash
# Using the convenience wrapper
./run.sh

# With CLI arguments
./run.sh --split url --no-filter

# Manual Docker commands
docker build -t all-inlink-splitter .
docker run --rm \
  -v "$(pwd)/input:/app/input" \
  -v "$(pwd)/output:/app/output" \
  all-inlink-splitter --split both
```

## Project Structure

```
all-inlink-splitter/
├── splitter.py          # Main script
├── requirements.txt     # Python dependencies
├── Dockerfile           # Docker image definition
├── run.sh               # Docker convenience wrapper
├── input/               # Place .xlsx files here
├── output/              # Dated output directories
├── CLAUDE.md            # AI assistant guidance
└── README.md            # This file
```
