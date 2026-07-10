import unittest

from splitter import find_column_index


class HeaderLookupTests(unittest.TestCase):
    def test_exact_header_wins_over_similarly_named_segment_header(self):
        headers = ('Source Segments', 'Source', 'Destination Segments', 'Destination')

        self.assertEqual(find_column_index(headers, 'Source'), 1)
        self.assertEqual(find_column_index(headers, 'Destination'), 3)


if __name__ == '__main__':
    unittest.main()
