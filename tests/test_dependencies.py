import unittest
from pathlib import Path

from openpyxl.cell._writer import write_cell
from openpyxl.xml import LXML


class DependencyTests(unittest.TestCase):
    def test_requirements_enable_fast_openpyxl_xml_writer(self):
        requirements = Path('requirements.txt').read_text(encoding='utf-8').splitlines()
        packages = {
            line.split('==', 1)[0].strip().lower()
            for line in requirements
            if line.strip() and not line.lstrip().startswith('#')
        }

        self.assertIn('lxml', packages)
        self.assertTrue(LXML)
        self.assertEqual(write_cell.__name__, 'lxml_write_cell')


if __name__ == '__main__':
    unittest.main()
