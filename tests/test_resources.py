import gc
import os
import unittest
import warnings
from argparse import Namespace
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from splitter import analyze_workbook


class InputResourceCleanupTests(unittest.TestCase):
    def test_early_schema_error_closes_input_workbook_handle(self):
        with TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / 'bad.xlsx'
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append([
                'Type',
                'Source',
                'Destination',
                None,
                'Source Segments',
            ])
            worksheet.append([
                'Hyperlink',
                'https://example.com',
                'https://example.com/bad',
                'data without a header',
                'USA',
            ])
            workbook.save(xlsx_path)
            workbook.close()
            args = Namespace(
                split='region',
                url_depth=2,
                url_pattern=None,
                no_filter=False,
                no_summary=False,
            )

            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter('always', ResourceWarning)
                try:
                    analyze_workbook(xlsx_path, args)
                except ValueError:
                    pass
                else:
                    self.fail('analyze_workbook did not reject the invalid schema')

                leaked_descriptors = []
                descriptor_dir = Path('/proc/self/fd')
                if descriptor_dir.exists():
                    for descriptor in descriptor_dir.iterdir():
                        try:
                            target = Path(os.readlink(descriptor)).resolve()
                        except (FileNotFoundError, OSError):
                            continue
                        if target == xlsx_path.resolve():
                            leaked_descriptors.append(descriptor.name)
                gc.collect()

            resource_warnings = [
                warning
                for warning in caught
                if issubclass(warning.category, ResourceWarning)
            ]
            self.assertEqual(resource_warnings, [])
            self.assertEqual(leaked_descriptors, [])


if __name__ == '__main__':
    unittest.main()
