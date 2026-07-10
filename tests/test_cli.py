import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook


class CliFailureTests(unittest.TestCase):
    def test_invalid_workbook_returns_nonzero_exit(self):
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_dir = tmp_path / 'input'
            output_dir = tmp_path / 'output'
            input_dir.mkdir()
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Broken tab'
            worksheet.append(['Type', 'Source', 'Destination'])
            worksheet.append([
                'Hyperlink',
                'https://example.com/home',
                'https://example.com/missing',
            ])
            workbook.save(input_dir / 'bad.xlsx')

            result = subprocess.run(
                [
                    sys.executable,
                    'splitter.py',
                    '--input',
                    str(input_dir),
                    '--output',
                    str(output_dir),
                ],
                cwd=Path(__file__).resolve().parents[1],
                capture_output=True,
                text=True,
                check=False,
            )

            self.assertNotEqual(result.returncode, 0)
            self.assertIn('bad.xlsx', result.stderr)
            self.assertIn('Source Segments', result.stderr)
            self.assertIn('Completed with 1 error', result.stderr)
            self.assertNotIn('Done.', result.stdout)


if __name__ == '__main__':
    unittest.main()
