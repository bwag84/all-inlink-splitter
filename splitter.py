#!/usr/bin/env python3
"""
All Inlinks Splitter — splits large Screaming Frog Excel crawl exports
into manageable files by geographical region, URL pattern, or both.

Two-pass streaming architecture for 900+ MB files:
  Pass 1: read_only — count destination frequencies (no row storage)
  Pass 2: read_only → write_only — route rows to output workbooks
"""

import re
import sys
import argparse
import json
from datetime import date
from pathlib import Path
from collections import Counter, defaultdict
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.xml.constants import MAX_COLUMN as EXCEL_MAX_COLUMNS
from openpyxl.xml.constants import MAX_ROW as EXCEL_MAX_ROWS

# ============================================================
# CONFIGURATION
# ============================================================

IGNORED_TYPES = ['Sitemap Hreflang', 'XML Sitemap']

HIGH_THRESHOLD = 100
MEDIUM_THRESHOLD = 10

REGION_ALIASES = {
    'APAC': ['APAC'],
    'MEISA': ['MEISA'],
    'EU': ['EU'],
    'LAC': ['LAC'],
    'USA': ['US', 'U.S.', 'USA', 'U.S.A.', 'United States', 'United States of America'],
    'Canada': ['Canada'],
}
SEGMENT_TOKEN_RE = re.compile(r'[A-Za-z0-9]+')
US_LOCALE_PREFIXES = {'en-us', 'es-us'}
US_PATH_SEGMENT = 'us-united-states'

# Locale-like path prefixes to skip when extracting URL groups
LOCALE_RE = re.compile(r'^[a-z]{2}(?:-[a-z]{2})?$', re.IGNORECASE)

# ============================================================
# PATH DETECTION
# ============================================================

SCRIPT_DIR = Path(__file__).parent.resolve()


def load_regional_subset_config(config_path):
    """Load configured URL path subsets keyed by canonical region name."""
    config_path = Path(config_path)
    if not config_path.exists():
        return {}

    with config_path.open(encoding='utf-8') as config_file:
        raw_config = json.load(config_file)
    if not isinstance(raw_config, dict):
        raise ValueError('configuration must be a JSON object keyed by region')

    sections = {'From / source', 'To / destination'}
    if sections.intersection(raw_config):
        unknown = raw_config.keys() - sections
        if unknown:
            raise ValueError(f"unknown configuration section(s): {sorted(unknown)}")
        result = {
            key: _validate_regional_paths(raw_config.get(key, {}), config_path)
            for key in ('From / source', 'To / destination')
        }
        for region, paths in result['To / destination'].items():
            source_names = {
                _subset_bucket_path(path).casefold()
                for path in result['From / source'].get(region, ())
            }
            if any('to_' + _subset_bucket_path(path).casefold() in source_names
                   for path in paths):
                raise ValueError(f"source and destination paths for {region!r} produce the same output name")
        return result
    return _validate_regional_paths(raw_config, config_path)


def _validate_regional_paths(raw_config, config_path):
    """Validate one region-to-path mapping, shared by both directions."""
    if not isinstance(raw_config, dict):
        raise ValueError('configuration section must be a JSON object keyed by region')
    subsets = {}
    for region, paths in raw_config.items():
        if region not in {*REGION_ALIASES, 'OTHER'}:
            raise ValueError(f"unknown region {region!r} in {config_path}")
        if not isinstance(paths, list):
            raise ValueError(f"region {region!r} must contain a JSON array of paths")
        normalized = []
        seen_paths = set()
        seen_output_names = set()
        for path in paths:
            if not isinstance(path, str):
                raise ValueError(f"region {region!r} paths must all be strings")
            clean_path = path.strip().strip('/')
            if not clean_path:
                continue
            segments = clean_path.split('/')
            if any(
                not segment
                or segment in ('.', '..')
                or not re.fullmatch(r'[A-Za-z0-9._~-]+', segment)
                for segment in segments
            ):
                raise ValueError(
                    f"unsafe path {path!r} for region {region!r}; "
                    "use URL path segments only"
                )
            path_key = clean_path.casefold()
            if path_key in seen_paths:
                raise ValueError(
                    f"duplicate path {path!r} for region {region!r}"
                )
            output_name = _subset_bucket_path(clean_path).casefold()
            if output_name in seen_output_names:
                raise ValueError(
                    f"paths for region {region!r} produce the same output name: "
                    f"{path!r}"
                )
            seen_paths.add(path_key)
            seen_output_names.add(output_name)
            normalized.append(clean_path)
        if normalized:
            subsets[region] = tuple(normalized)
    return subsets


def _subset_directions(config):
    """Accept both legacy source-only mappings and labeled configurations."""
    config = config or {}
    if 'From / source' in config or 'To / destination' in config:
        return config.get('From / source', {}), config.get('To / destination', {})
    return config, {}


def _subset_bucket_path(path):
    """Convert a validated configured path to its output bucket suffix."""
    return path.replace('/', '_')

def _detect_dirs(cli_input=None, cli_output=None):
    if cli_input:
        input_dir = Path(cli_input)
    elif Path('/app/input').exists():
        input_dir = Path('/app/input')
    else:
        input_dir = SCRIPT_DIR / 'input'

    if cli_output:
        output_base = Path(cli_output)
    elif Path('/app/output').exists():
        output_base = Path('/app/output')
    else:
        output_base = SCRIPT_DIR / 'output'

    return input_dir, output_base

# ============================================================
# PROGRESS TRACKER
# ============================================================

class ProgressTracker:
    def __init__(self, total_steps):
        self.total = total_steps
        self.current = 0

    def step(self, description):
        self.current += 1
        print(f"\n[Step {self.current} of {self.total}] {description}")
        print("-" * 50)

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def find_column_index(headers, column_name):
    """Find a column by its normalized, case-insensitive header name."""
    target = _header_key(column_name)
    for idx, header in enumerate(headers):
        if _header_key(header) == target:
            return idx
    return None


def _header_key(header):
    """Normalize harmless whitespace/case differences in workbook headers."""
    if header is None:
        return ''
    return re.sub(r'\s+', ' ', str(header).strip()).casefold()


def is_ignored_type(type_value):
    """Check if a row's Type should be filtered out."""
    if type_value is None:
        return False
    type_str = str(type_value).lower()
    return any(ignored.lower() in type_str for ignored in IGNORED_TYPES)


def get_priority(frequency):
    """Assign priority label based on destination frequency."""
    if frequency >= HIGH_THRESHOLD:
        return 'HIGH'
    elif frequency >= MEDIUM_THRESHOLD:
        return 'MEDIUM'
    return 'LOW'


def get_matching_regions(value, source_url=None):
    """Return regions from Source Segments, then a narrow Source URL fallback."""
    segment_tokens = [token.upper() for token in SEGMENT_TOKEN_RE.findall(str(value))]
    matches = [
        region
        for region, aliases in REGION_ALIASES.items()
        if any(_matches_segment_alias(segment_tokens, alias) for alias in aliases)
    ]
    if matches:
        return matches
    inferred_region = _infer_region_from_source_url(source_url)
    return [inferred_region] if inferred_region else ['OTHER']


def _matches_segment_alias(segment_tokens, alias):
    """Return True if alias appears as complete token(s) in Source Segments."""
    alias_tokens = [token.upper() for token in SEGMENT_TOKEN_RE.findall(alias)]
    if not alias_tokens:
        return False
    if len(alias_tokens) == 1:
        return alias_tokens[0] in segment_tokens
    end = len(segment_tokens) - len(alias_tokens) + 1
    return any(segment_tokens[i:i + len(alias_tokens)] == alias_tokens for i in range(end))


def _infer_region_from_source_url(source_url):
    """Infer USA only from exact US locale/content path segments observed in exports."""
    try:
        parsed = urlparse(str(source_url))
        if parsed.scheme.casefold() not in ('http', 'https') or not parsed.hostname:
            return None
        path_segments = [
            segment.casefold()
            for segment in parsed.path.split('/')
            if segment
        ]
    except (TypeError, ValueError):
        return None
    if path_segments and path_segments[0] in US_LOCALE_PREFIXES:
        return 'USA'
    if US_PATH_SEGMENT in path_segments:
        return 'USA'
    return None


def extract_url_group(url, depth=2, pattern=None):
    """
    Derive a grouping key from a URL.

    If *pattern* is given, use it as a regex and return the first match group
    (or the full match if there are no groups).

    Otherwise strip the locale-like first segment (e.g. "en-us") and join the
    next *depth* path segments with "_".
    """
    if pattern:
        m = re.search(pattern, str(url))
        if m:
            return m.group(1) if m.lastindex else m.group(0)
        return 'other'

    try:
        path = urlparse(str(url)).path.strip('/')
    except Exception:
        return 'other'

    if not path:
        return 'root'

    # Strip file extension from last segment (e.g. .html, .php)
    path = re.sub(r'\.[a-zA-Z]{2,5}$', '', path)

    segments = path.split('/')

    # Skip locale-like prefix
    if segments and LOCALE_RE.match(segments[0]):
        segments = segments[1:]

    if not segments:
        return 'root'

    key = '_'.join(segments[:depth])
    # Sanitise for use in filenames
    key = re.sub(r'[^\w\-]', '_', key)
    return key if key else 'other'


def _safe_cell(row, idx):
    """Safely get a cell value from a row tuple."""
    if idx is not None and idx < len(row):
        return row[idx]
    return None

# ============================================================
# PASS 1 — ANALYSIS (read_only, no row storage)
# ============================================================

def analyze_workbook(xlsx_path, args):
    """
    Stream through every sheet with read_only=True.
    Returns (headers, dest_counter, bucket_counter, total_rows, kept_rows).

    headers: canonical union of headers from every sheet
    dest_counter: Counter of destination URLs
    bucket_counter: Counter of bucket keys (for progress reporting)
    """
    wb = load_workbook(xlsx_path, read_only=True)
    headers = []
    dest_counter = Counter()
    bucket_counter = Counter()
    total_rows = 0
    kept_rows = 0

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            header_rows = tuple(
                ws.iter_rows(min_row=1, max_row=1, values_only=True)
            )
            if not header_rows:
                continue
            sheet_headers = header_rows[0]

            _merge_headers(headers, sheet_headers)
            type_col, dest_col, seg_col, source_col = _sheet_column_indices(
                sheet_headers, sheet_name, args
            )
            unnamed_columns = _unnamed_column_indices(sheet_headers)

            rows = ws.iter_rows(min_row=2, values_only=True)
            try:
                for row in rows:
                    # Skip empty rows
                    if all(c is None for c in row):
                        continue
                    _validate_unnamed_cells(row, unnamed_columns, sheet_name)

                    total_rows += 1

                    # Optional type filtering
                    if not args.no_filter and type_col is not None:
                        if is_ignored_type(_safe_cell(row, type_col)):
                            continue

                    kept_rows += 1

                    # Destination counting
                    dest_val = _safe_cell(row, dest_col)
                    if dest_val:
                        dest_counter[dest_val] += 1

                    # Bucket counting
                    buckets = _resolve_buckets(
                        row,
                        seg_col,
                        source_col,
                        args.split,
                        args.url_depth,
                        args.url_pattern,
                        getattr(args, 'regional_subsets', None),
                        dest_col=dest_col,
                    )
                    for b in buckets:
                        bucket_counter[b] += 1
            finally:
                rows.close()
    finally:
        wb.close()

    canonical_headers = tuple(headers) if headers else None
    return canonical_headers, dest_counter, bucket_counter, total_rows, kept_rows


def _merge_headers(canonical_headers, sheet_headers):
    """Extend canonical headers with named columns not seen on earlier sheets."""
    known = {_header_key(header) for header in canonical_headers}
    for header in sheet_headers:
        key = _header_key(header)
        if key and key not in known:
            canonical_headers.append(header)
            known.add(key)


def _column_remap(source_headers, target_headers):
    """Map each canonical output column to its index in one source sheet."""
    source_indices = {
        _header_key(header): idx
        for idx, header in enumerate(source_headers)
        if _header_key(header)
    }
    return tuple(source_indices.get(_header_key(header)) for header in target_headers)


def _align_row(row, column_remap):
    """Reorder one source row into the workbook's canonical column order."""
    return tuple(_safe_cell(row, idx) for idx in column_remap)


def _sheet_column_indices(headers, sheet_name, args):
    """Resolve and validate the columns needed to process one source sheet."""
    seen_headers = set()
    duplicate_headers = []
    for header in headers:
        key = _header_key(header)
        if not key:
            continue
        if key in seen_headers and str(header) not in duplicate_headers:
            duplicate_headers.append(str(header))
        seen_headers.add(key)
    if duplicate_headers:
        duplicates = ', '.join(duplicate_headers)
        raise ValueError(
            f"Sheet {sheet_name!r} has duplicate column header(s): {duplicates}"
        )

    indices = {
        'Type': find_column_index(headers, 'Type'),
        'Destination': find_column_index(headers, 'Destination'),
        'Source Segments': find_column_index(headers, 'Source Segments'),
        'Source': find_column_index(headers, 'Source'),
    }
    required = ['Destination']
    if not args.no_filter:
        required.append('Type')
    if args.split in ('region', 'both'):
        required.append('Source Segments')
    if args.split in ('url', 'both') or (
        args.split == 'region' and _subset_directions(
            getattr(args, 'regional_subsets', None)
        )[0]
    ):
        required.append('Source')

    missing = [name for name in required if indices[name] is None]
    if missing:
        missing_list = ', '.join(missing)
        raise ValueError(f"Sheet {sheet_name!r} is missing required column(s): {missing_list}")

    return (
        indices['Type'],
        indices['Destination'],
        indices['Source Segments'],
        indices['Source'],
    )


def _unnamed_column_indices(headers):
    """Return zero-based positions of blank header cells."""
    return tuple(idx for idx, header in enumerate(headers) if not _header_key(header))


def _validate_unnamed_cells(row, unnamed_columns, sheet_name):
    """Reject data that cannot be represented in the canonical named schema."""
    populated = [
        idx + 1
        for idx in unnamed_columns
        if _safe_cell(row, idx) not in (None, '')
    ]
    if populated:
        columns = ', '.join(str(idx) for idx in populated)
        raise ValueError(
            f"Sheet {sheet_name!r} has data in unnamed column(s): {columns}"
        )


def _resolve_buckets(
    row,
    seg_col,
    source_col,
    split_mode,
    url_depth,
    url_pattern,
    regional_subsets=None,
    dest_col=None,
):
    """Determine which output bucket(s) a row belongs to."""
    source_url = _safe_cell(row, source_col)
    if split_mode == 'region':
        regions = get_matching_regions(_safe_cell(row, seg_col), source_url)
        buckets = list(regions)
        source_subsets, destination_subsets = _subset_directions(regional_subsets)
        for subsets, url, prefix in (
            (source_subsets, source_url, ''),
            (destination_subsets, _safe_cell(row, dest_col), 'to_'),
        ):
            path_segments = _subset_path_segments(url)
            for region in regions:
                for path in subsets.get(region, ()):
                    configured_segments = tuple(
                        segment.casefold() for segment in path.split('/')
                    )
                    if _contains_path_segments(path_segments, configured_segments):
                        bucket_path = _subset_bucket_path(path)
                        buckets.append(f"{region}_{prefix}{bucket_path}")
        return buckets

    if split_mode == 'url':
        return [extract_url_group(source_url, url_depth, url_pattern)]

    # both
    regions = get_matching_regions(_safe_cell(row, seg_col), source_url)
    url_group = extract_url_group(source_url, url_depth, url_pattern)
    return [f"{r}_{url_group}" for r in regions]


def _subset_path_segments(source_url):
    """Return case-folded URL path segments after an optional locale prefix."""
    try:
        parsed = urlparse(str(source_url))
        if parsed.scheme.casefold() not in ('http', 'https') or not parsed.hostname:
            return ()
        segments = tuple(
            segment.casefold()
            for segment in parsed.path.split('/')
            if segment
        )
    except (TypeError, ValueError):
        return ()
    if segments and LOCALE_RE.fullmatch(segments[0]):
        return segments[1:]
    return segments


def _contains_path_segments(path_segments, configured_segments):
    """Return True when configured segments occur contiguously in a URL path."""
    width = len(configured_segments)
    if not width or width > len(path_segments):
        return False
    return any(
        path_segments[index:index + width] == configured_segments
        for index in range(len(path_segments) - width + 1)
    )

# ============================================================
# OUTPUT MANAGER — lazy write_only workbooks
# ============================================================

class OutputManager:
    """
    Manages one write_only workbook per bucket.  Workbooks are created lazily
    on first row append so we never create empty files.
    """

    def __init__(self, output_dir, base_name, headers, dest_counter, no_summary):
        if len(headers) + 1 > EXCEL_MAX_COLUMNS:
            raise ValueError(
                f"{len(headers)} input columns plus Priority exceed Excel's "
                f"{EXCEL_MAX_COLUMNS}-column limit"
            )
        self.output_dir = output_dir
        self.base_name = base_name
        self.headers = list(headers) + ['Priority']
        self.dest_counter = dest_counter
        self.no_summary = no_summary
        self._workbooks = {}    # bucket → Workbook
        self._data_sheets = {}  # bucket → worksheet
        self._data_sheet_parts = Counter()
        self._data_rows_in_sheet = Counter()
        self._dest_sets = defaultdict(set)  # bucket → set of destinations
        self._row_counts = Counter()

    def _init_bucket(self, bucket):
        wb = Workbook(write_only=True)
        self._workbooks[bucket] = wb
        self._start_data_sheet(bucket)

    def _start_data_sheet(self, bucket):
        """Create the next bounded Data sheet for a bucket."""
        self._data_sheet_parts[bucket] += 1
        part = self._data_sheet_parts[bucket]
        title = 'Data' if part == 1 else f'Data {part}'
        ws_data = self._workbooks[bucket].create_sheet(title)
        ws_data.append(self.headers)
        self._data_sheets[bucket] = ws_data
        self._data_rows_in_sheet[bucket] = 0

    def append(self, bucket, row, dest_value, priority):
        if bucket not in self._workbooks:
            self._init_bucket(bucket)
        if self._data_rows_in_sheet[bucket] >= EXCEL_MAX_ROWS - 1:
            self._start_data_sheet(bucket)
        self._data_sheets[bucket].append(list(row) + [priority])
        self._data_rows_in_sheet[bucket] += 1
        self._row_counts[bucket] += 1
        if dest_value:
            self._dest_sets[bucket].add(dest_value)

    def close_all(self):
        """Write summary sheets (unless --no-summary) and save all workbooks."""
        files_created = 0
        for bucket, wb in self._workbooks.items():
            if not self.no_summary:
                self._write_summary(wb, bucket)
            filename = f"{self.base_name}_{bucket}.xlsx"
            wb.save(self.output_dir / filename)
            count = self._row_counts[bucket]
            dests = len(self._dest_sets[bucket])
            print(f"    {bucket}: {count:,} rows, {dests} unique destinations")
            files_created += 1
        return files_created

    def abort_all(self):
        """Close streaming writers and remove their temporary XML files."""
        for wb in self._workbooks.values():
            for ws in wb.worksheets:
                rows = getattr(ws, '_rows', None)
                writer = getattr(ws, '_writer', None)
                if rows is not None:
                    try:
                        rows.close()
                    except Exception:
                        pass
                    ws._rows = None
                if writer is not None:
                    try:
                        writer.close()
                    except Exception:
                        pass
                    try:
                        writer.cleanup()
                    except Exception:
                        pass
                    writer.xf = None
                    ws._writer = None
            try:
                wb.close()
            except Exception:
                pass

        self._workbooks.clear()
        self._data_sheets.clear()

    def _write_summary(self, wb, bucket):
        """Create bounded Summary sheet(s) with destinations ranked by frequency."""
        part = 1
        ws = self._start_summary_sheet(wb, part)
        rows_in_sheet = 0

        dests_in_bucket = self._dest_sets[bucket]
        ranked = sorted(
            ((d, self.dest_counter[d]) for d in dests_in_bucket if d in self.dest_counter),
            key=lambda x: x[1],
            reverse=True,
        )
        for dest, freq in ranked:
            if rows_in_sheet >= EXCEL_MAX_ROWS - 1:
                part += 1
                ws = self._start_summary_sheet(wb, part)
                rows_in_sheet = 0
            priority = get_priority(freq)
            impact = f"Fix 1 URL → removes {freq} error{'s' if freq > 1 else ''}"
            ws.append([priority, dest, freq, impact])
            rows_in_sheet += 1

    @staticmethod
    def _start_summary_sheet(wb, part):
        """Insert the next Summary sheet before all Data sheets."""
        title = 'Summary' if part == 1 else f'Summary {part}'
        ws = wb.create_sheet(title, part - 1)
        ws.append(['Priority', 'Destination', 'Occurrences', 'Impact'])
        return ws

# ============================================================
# PASS 2 — SPLITTING (read_only → write_only)
# ============================================================

def split_workbook(xlsx_path, args, headers, dest_counter, output_dir):
    """
    Second streaming pass.  Routes each row to the correct OutputManager bucket,
    appending priority on the fly.
    """
    base_name = xlsx_path.stem
    om = OutputManager(output_dir, base_name, headers, dest_counter, args.no_summary)

    try:
        wb = load_workbook(xlsx_path, read_only=True)
        try:
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                header_rows = tuple(
                    ws.iter_rows(min_row=1, max_row=1, values_only=True)
                )
                if not header_rows:
                    continue
                sheet_headers = header_rows[0]

                type_col, dest_col, seg_col, source_col = _sheet_column_indices(
                    sheet_headers, sheet_name, args
                )
                column_remap = _column_remap(sheet_headers, headers)
                unnamed_columns = _unnamed_column_indices(sheet_headers)

                rows = ws.iter_rows(min_row=2, values_only=True)
                try:
                    for row in rows:
                        if all(c is None for c in row):
                            continue
                        _validate_unnamed_cells(row, unnamed_columns, sheet_name)

                        # Optional type filtering
                        if not args.no_filter and type_col is not None:
                            if is_ignored_type(_safe_cell(row, type_col)):
                                continue

                        dest_val = _safe_cell(row, dest_col)
                        freq = dest_counter.get(dest_val, 0) if dest_val else 0
                        priority = get_priority(freq)

                        buckets = _resolve_buckets(
                            row,
                            seg_col,
                            source_col,
                            args.split,
                            args.url_depth,
                            args.url_pattern,
                            getattr(args, 'regional_subsets', None),
                            dest_col=dest_col,
                        )
                        aligned_row = _align_row(row, column_remap)
                        for bucket in buckets:
                            om.append(bucket, aligned_row, dest_val, priority)
                finally:
                    rows.close()
        finally:
            wb.close()

        print(f"\n  Output files for {base_name}:")
        return om.close_all()
    except BaseException:
        om.abort_all()
        raise

# ============================================================
# CLI
# ============================================================

def build_parser():
    parser = argparse.ArgumentParser(
        description='Split large Screaming Frog All Inlinks exports into manageable files.',
    )
    parser.add_argument(
        '--split',
        choices=['region', 'url', 'both'],
        default='region',
        help='Splitting strategy (default: region)',
    )
    parser.add_argument(
        '--url-depth',
        type=int,
        default=2,
        metavar='N',
        help='Number of URL path segments for grouping in url/both mode (default: 2)',
    )
    parser.add_argument(
        '--url-pattern',
        type=str,
        default=None,
        metavar='REGEX',
        help='Custom regex for URL grouping (overrides --url-depth)',
    )
    parser.add_argument(
        '--config',
        type=str,
        default=str(SCRIPT_DIR / 'config.json'),
        metavar='FILE',
        help='Regional URL subset configuration (default: config.json)',
    )
    parser.add_argument(
        '--no-summary',
        action='store_true',
        help='Skip creating the Summary sheet in output files',
    )
    parser.add_argument(
        '--no-filter',
        action='store_true',
        help='Do not filter out ignored types (Sitemap Hreflang, XML Sitemap)',
    )
    parser.add_argument(
        '--input',
        type=str,
        default=None,
        metavar='DIR',
        help='Input directory (default: input/ or /app/input in Docker)',
    )
    parser.add_argument(
        '--output',
        type=str,
        default=None,
        metavar='DIR',
        help='Output base directory (default: output/ or /app/output in Docker)',
    )
    return parser

# ============================================================
# MAIN
# ============================================================

def main():
    parser = build_parser()
    args = parser.parse_args()

    config_path = Path(args.config)
    try:
        args.regional_subsets = (
            load_regional_subset_config(config_path)
            if config_path.exists()
            else {}
        )
    except (OSError, TypeError, ValueError) as error:
        parser.error(f"invalid regional subset config {config_path}: {error}")

    input_dir, output_base = _detect_dirs(args.input, args.output)
    input_dir.mkdir(parents=True, exist_ok=True)

    today = date.today().isoformat()
    output_dir = output_base / today
    output_dir.mkdir(parents=True, exist_ok=True)

    xlsx_files = sorted(input_dir.glob('*.xlsx'))
    if not xlsx_files:
        print(f"No .xlsx files found in {input_dir}")
        print("Place your Excel files in the input/ directory and run again.")
        sys.exit(1)

    # Each file needs 2 passes + final save → 3 steps per file
    total_steps = len(xlsx_files) * 3
    tracker = ProgressTracker(total_steps)

    print("=" * 60)
    print("All Inlinks Splitter")
    print("=" * 60)
    print(f"  Split mode : {args.split}")
    if args.regional_subsets:
        subset_count = sum(
            len(paths)
            for direction in _subset_directions(args.regional_subsets)
            for paths in direction.values()
        )
        print(f"  Subsets    : {subset_count} from {config_path}")
    if args.split in ('url', 'both'):
        if args.url_pattern:
            print(f"  URL pattern: {args.url_pattern}")
        else:
            print(f"  URL depth  : {args.url_depth}")
    print(f"  Filtering  : {'off' if args.no_filter else 'on (excluding ' + ', '.join(IGNORED_TYPES) + ')'}")
    print(f"  Files      : {len(xlsx_files)}")
    print(f"  Output     : output/{today}/")

    grand_total_files = 0
    failed_files = 0

    for xlsx_path in xlsx_files:
        # --- Pass 1: Analyse ---
        tracker.step(f"Analysing {xlsx_path.name}")
        try:
            headers, dest_counter, bucket_counter, total_rows, kept_rows = analyze_workbook(
                xlsx_path, args
            )
        except Exception as e:
            print(f"  ERROR: Could not read {xlsx_path.name} — {e}", file=sys.stderr)
            failed_files += 1
            # Skip the remaining 2 steps for this file
            tracker.current += 2
            continue

        if headers is None:
            print(f"  WARNING: No data found in {xlsx_path.name}, skipping.")
            tracker.current += 2
            continue

        print(f"  Total rows : {total_rows:,}")
        print(f"  After filter: {kept_rows:,}")
        print(f"  Unique dests: {len(dest_counter):,}")
        print(f"  Buckets     : {len(bucket_counter)}")

        # --- Priority summary ---
        tracker.step(f"Computing priorities for {xlsx_path.name}")
        high = sum(1 for f in dest_counter.values() if f >= HIGH_THRESHOLD)
        med = sum(1 for f in dest_counter.values() if MEDIUM_THRESHOLD <= f < HIGH_THRESHOLD)
        low = sum(1 for f in dest_counter.values() if f < MEDIUM_THRESHOLD)
        print(f"  HIGH  : {high:,} destinations")
        print(f"  MEDIUM: {med:,} destinations")
        print(f"  LOW   : {low:,} destinations")

        # --- Pass 2: Split & write ---
        tracker.step(f"Splitting {xlsx_path.name}")
        try:
            files_created = split_workbook(
                xlsx_path, args, headers, dest_counter, output_dir
            )
        except Exception as e:
            print(f"  ERROR: Could not write {xlsx_path.name} — {e}", file=sys.stderr)
            failed_files += 1
            continue
        grand_total_files += files_created

    if failed_files:
        print("\n" + "=" * 60, file=sys.stderr)
        print(
            f"Completed with {failed_files} error(s). Created {grand_total_files} "
            f"output file(s) in output/{today}/",
            file=sys.stderr,
        )
        print("=" * 60, file=sys.stderr)
        sys.exit(1)

    print("\n" + "=" * 60)
    print(f"Done. Created {grand_total_files} output file(s) in output/{today}/")
    print("=" * 60)


if __name__ == '__main__':
    main()
