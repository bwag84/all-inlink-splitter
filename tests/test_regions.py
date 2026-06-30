import unittest

from splitter import get_matching_regions


class RegionMatchingTests(unittest.TestCase):
    def test_matches_country_segment_names(self):
        cases = {
            'US': ['US'],
            'United States': ['US'],
            'United States of America': ['US'],
            'Canada': ['Canada'],
            'United States, Canada': ['US', 'Canada'],
        }

        for source_segments, expected in cases.items():
            with self.subTest(source_segments=source_segments):
                self.assertEqual(get_matching_regions(source_segments), expected)

    def test_us_does_not_match_inside_other_words(self):
        self.assertEqual(get_matching_regions('Australia'), ['OTHER'])

    def test_existing_region_segments_still_match(self):
        for region in ['APAC', 'MEISA', 'EU', 'LAC']:
            with self.subTest(region=region):
                self.assertEqual(get_matching_regions(region), [region])


if __name__ == '__main__':
    unittest.main()
