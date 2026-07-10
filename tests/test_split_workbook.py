import unittest
from argparse import Namespace
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet._writer import ALL_TEMP_FILES

from splitter import analyze_workbook, split_workbook


class SplitWorkbookRegionOutputTests(unittest.TestCase):
    def test_usa_source_segment_writes_usa_output_file(self):
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            xlsx_path = tmp_path / 'crawl.xlsx'
            output_dir = tmp_path / 'output'
            output_dir.mkdir()

            wb = Workbook()
            ws = wb.active
            ws.title = 'All Inlinks'
            ws.append(['Type', 'Source', 'Destination', 'Source Segments'])
            ws.append(['Hyperlink', 'https://example.com/en-us/page', 'https://example.com/broken', 'USA'])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )
            headers, dest_counter, bucket_counter, total_rows, kept_rows = analyze_workbook(xlsx_path, args)

            self.assertEqual(bucket_counter['USA'], 1)
            self.assertEqual(total_rows, 1)
            self.assertEqual(kept_rows, 1)

            files_created = split_workbook(xlsx_path, args, headers, dest_counter, output_dir)

            self.assertEqual(files_created, 1)
            self.assertTrue((output_dir / 'crawl_USA.xlsx').exists())
            self.assertFalse((output_dir / 'crawl_US.xlsx').exists())

            result = load_workbook(output_dir / 'crawl_USA.xlsx', read_only=True)
            try:
                self.assertEqual(result.sheetnames, ['Summary', 'Data'])
                data_rows = list(result['Data'].iter_rows(values_only=True))
                self.assertEqual(data_rows[1][-1], 'LOW')
            finally:
                result.close()

    def test_analysis_uses_each_sheet_header_order(self):
        with TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / 'multi-tab.xlsx'
            wb = Workbook()
            eu = wb.active
            eu.title = 'First'
            eu.append(['Type', 'Source', 'Destination', 'Source Segments', 'Anchor'])
            eu.append([
                'Hyperlink',
                'https://example.com/europe',
                'https://example.com/missing-eu',
                'EU',
                'Europe link',
            ])
            usa = wb.create_sheet('Reordered')
            usa.append(['Destination', 'Source Segments', 'Type', 'Source', 'Link Position'])
            usa.append([
                'https://example.com/missing-page',
                'USA',
                'Hyperlink',
                'https://example.com/home',
                'Content',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )
            headers, dest_counter, bucket_counter, total_rows, kept_rows = analyze_workbook(
                xlsx_path, args
            )

            self.assertEqual(
                headers,
                (
                    'Type',
                    'Source',
                    'Destination',
                    'Source Segments',
                    'Anchor',
                    'Link Position',
                ),
            )
            self.assertEqual(dict(bucket_counter), {'EU': 1, 'USA': 1})
            self.assertEqual(
                dict(dest_counter),
                {
                    'https://example.com/missing-eu': 1,
                    'https://example.com/missing-page': 1,
                },
            )
            self.assertEqual(total_rows, 2)
            self.assertEqual(kept_rows, 2)

    def test_split_aligns_reordered_rows_to_canonical_headers(self):
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            xlsx_path = tmp_path / 'multi-tab.xlsx'
            output_dir = tmp_path / 'output'
            output_dir.mkdir()
            wb = Workbook()
            eu = wb.active
            eu.title = 'First'
            eu.append(['Type', 'Source', 'Destination', 'Source Segments', 'Anchor'])
            eu.append([
                'Hyperlink',
                'https://example.com/europe',
                'https://example.com/missing-eu',
                'EU',
                'Europe link',
            ])
            usa = wb.create_sheet('Reordered')
            usa.append(['Destination', 'Source Segments', 'Type', 'Source', 'Link Position'])
            usa.append([
                'https://example.com/missing-page',
                'USA',
                'Hyperlink',
                'https://example.com/home',
                'Content',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )
            headers, dest_counter, _, _, _ = analyze_workbook(xlsx_path, args)
            split_workbook(xlsx_path, args, headers, dest_counter, output_dir)

            self.assertFalse((output_dir / 'multi-tab_OTHER.xlsx').exists())
            usa_path = output_dir / 'multi-tab_USA.xlsx'
            self.assertTrue(usa_path.exists())
            result = load_workbook(usa_path, read_only=True)
            try:
                rows = list(result['Data'].iter_rows(values_only=True))
                self.assertEqual(
                    rows[0],
                    (
                        'Type',
                        'Source',
                        'Destination',
                        'Source Segments',
                        'Anchor',
                        'Link Position',
                        'Priority',
                    ),
                )
                self.assertEqual(
                    rows[1],
                    (
                        'Hyperlink',
                        'https://example.com/home',
                        'https://example.com/missing-page',
                        'USA',
                        None,
                        'Content',
                        'LOW',
                    ),
                )
            finally:
                result.close()

    def test_missing_region_column_is_rejected_with_sheet_name(self):
        with TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / 'bad-header.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.title = 'Broken tab'
            ws.append(['Type', 'Source', 'Destination', 'Source Segment'])
            ws.append([
                'Hyperlink',
                'https://example.com/home',
                'https://example.com/missing',
                'USA',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )

            with self.assertRaisesRegex(
                ValueError,
                r"Broken tab.*Source Segments",
            ):
                analyze_workbook(xlsx_path, args)

    def test_duplicate_named_columns_are_rejected(self):
        with TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / 'duplicate-header.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.title = 'Duplicate tab'
            ws.append([
                'Type',
                'Source',
                'Destination',
                'Destination',
                'Source Segments',
            ])
            ws.append([
                'Hyperlink',
                'https://example.com/home',
                'https://example.com/first',
                'https://example.com/second',
                'USA',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )

            with self.assertRaisesRegex(
                ValueError,
                r"Duplicate tab.*duplicate.*Destination",
            ):
                analyze_workbook(xlsx_path, args)

    def test_data_in_unnamed_column_is_rejected(self):
        with TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / 'unnamed-header.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.title = 'Unnamed tab'
            ws.append(['Type', 'Source', 'Destination', None, 'Source Segments'])
            ws.append([
                'Hyperlink',
                'https://example.com/home',
                'https://example.com/missing',
                'would otherwise be dropped',
                'USA',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )

            with self.assertRaisesRegex(
                ValueError,
                r"Unnamed tab.*unnamed column.*4",
            ):
                analyze_workbook(xlsx_path, args)

    def test_split_failure_cleans_write_only_temporary_files(self):
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            xlsx_path = tmp_path / 'changes-between-passes.xlsx'
            output_dir = tmp_path / 'output'
            output_dir.mkdir()
            wb = Workbook()
            first = wb.active
            first.title = 'First'
            first.append(['Type', 'Source', 'Destination', 'Source Segments'])
            first.append([
                'Hyperlink',
                'https://example.com/home',
                'https://example.com/missing',
                'USA',
            ])
            second = wb.create_sheet('Changed')
            second.append(['Type', 'Source', 'Destination', 'Source Segments'])
            second.append([
                'Hyperlink',
                'https://example.com/europe',
                'https://example.com/missing-eu',
                'EU',
            ])
            wb.save(xlsx_path)

            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )
            headers, dest_counter, _, _, _ = analyze_workbook(xlsx_path, args)

            changed = load_workbook(xlsx_path)
            changed['Changed'].cell(row=1, column=4, value='Wrong Header')
            changed.save(xlsx_path)
            changed.close()

            temp_files_before = set(ALL_TEMP_FILES)
            with self.assertRaisesRegex(ValueError, r"Changed.*Source Segments"):
                split_workbook(xlsx_path, args, headers, dest_counter, output_dir)

            self.assertEqual(set(ALL_TEMP_FILES), temp_files_before)
            self.assertEqual(list(output_dir.glob('*.xlsx')), [])


if __name__ == '__main__':
    unittest.main()
