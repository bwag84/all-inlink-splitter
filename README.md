# All Inlinks Splitter

Split large Screaming Frog "All Inlinks" Excel exports (900 MB+) into manageable, region- or URL-based files — directly from the command line.

## Why?

Screaming Frog crawl exports for large sites can exceed 900 MB and span 13+ Excel tabs. Most laptops can't even open them. This tool:

- **Streams** through every tab without loading the full file into memory
- **Splits** by geographical region, URL pattern, or both
- **Creates optional regional subsets** for configured Source URL paths
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
| `--config FILE` | Regional URL subset configuration | `config.json` |
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

Regions detected: APAC, MEISA, EU, LAC, USA, Canada. `US`, `U.S.`, `USA`, `U.S.A.`, `United States`, and `United States of America` source segments are grouped into the USA file. When Source Segments contains no recognized geography, exact markers in an absolute HTTP(S) Source URL (`en-us`, `es-us`, or a `us-united-states` path segment) provide a USA fallback; an explicit geographic segment always takes precedence. Rows matching no region go to OTHER. A row matching multiple explicit regions appears in each.

### Regional URL subsets

Edit `config.json` to create smaller files inside selected regions while retaining the complete regional files:

```json
{
  "From / source": {
    "EU": ["campaign", "shipping/surcharges"],
    "USA": ["campaign"]
  },
  "To / destination": {
    "EU": ["folder"]
  }
}
```

Then run the normal regional split:

```bash
python3 splitter.py
```

A Source URL such as `https://www.fedex.com/en-gb/campaign/summer.html` remains in `demo_EU.xlsx` and is also copied to `demo_EU_campaign.xlsx`. A configured multi-folder path such as `shipping/surcharges` produces `demo_EU_shipping_surcharges.xlsx`.

Use **From / source** for paths on the page containing the link and **To / destination** for paths the link points to. For example, a link from `/nl-nl/customer-support` to `/nl-nl/folder/tracking.html` matches destination `folder` and is copied to `demo_EU_to_folder.xlsx`. Source and destination rules match independently; they are not an AND filter. Regions still come from Source Segments, so add a destination path to every source region whose inlinks you want to collect.

Matching is case-insensitive and checks complete, contiguous URL path segments after an optional language or language-country prefix. Domains and query strings are ignored. A row matching multiple configured paths is copied to every matching subset. Empty region arrays—or a missing config file—preserve the standard output exactly. Regional subsets apply only to the default `region` split mode; the existing `url` and `both` modes are unchanged.

The section names serve as headers because JSON does not support comments. Either section and any unused regions may be omitted. Legacy flat region-to-path configs still work as source-only rules. Output filename collisions between source and destination rules are rejected.

Use a different configuration when needed:

```bash
python3 splitter.py --config configs/campaign-review.json
```

Configuration errors stop the run before the workbook is processed. Region names must match the keys in the template, and paths must contain path segments only—not domains, query strings, `.` or `..` segments.

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

Each output file normally has two sheets:

1. **Summary** — Destinations ranked by frequency, with Priority and Impact columns
2. **Data** — The canonical union of input columns plus an appended **Priority** column

When a bucket contains more than 1,048,575 data rows, the file automatically continues into `Data 2`, `Data 3`, and so on. Summary entries likewise continue into numbered Summary sheets when needed. Every continuation sheet repeats its header, so no rows need to be discarded or repaired by Excel.

### Priority Levels

| Priority | Threshold | Meaning |
|----------|-----------|---------|
| HIGH | ≥ 100 occurrences | Quick wins — fixing one URL removes many errors |
| MEDIUM | ≥ 10 and < 100 | Moderate impact |
| LOW | < 10 | Low frequency |

## Filtering

By default, rows with Type = "Sitemap Hreflang" or "XML Sitemap" are excluded. Use `--no-filter` to keep them.

## Multi-Tab Support

Screaming Frog splits large exports across multiple Excel tabs. This tool reads **all tabs** in each file, resolves columns from each tab's own header row, and aligns reordered columns into one canonical output layout. Missing required headers, duplicate named headers, or data under unnamed headers reject that input with a precise error instead of silently misrouting or dropping rows; the CLI exits non-zero after processing any remaining inputs.

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

The Docker image includes the repository's `config.json`; `./run.sh` rebuilds the image so edits to that file are picked up on the next run.

## Project Structure

```
all-inlink-splitter/
├── splitter.py          # Main script
├── config.json          # Optional regional URL subset paths
├── requirements.txt     # Python dependencies
├── Dockerfile           # Docker image definition
├── run.sh               # Docker convenience wrapper
├── input/               # Place .xlsx files here
├── output/              # Dated output directories
├── CLAUDE.md            # AI assistant guidance
└── README.md            # This file
```
