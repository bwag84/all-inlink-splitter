import unittest

from splitter import get_matching_regions


class RegionMatchingTests(unittest.TestCase):
    def test_matches_country_segment_names(self):
        cases = {
            'US': ['USA'],
            'U.S.': ['USA'],
            'United States': ['USA'],
            'U.S.A.': ['USA'],
            'United States of America': ['USA'],
            'Canada': ['Canada'],
            'United States, Canada': ['USA', 'Canada'],
        }

        for source_segments, expected in cases.items():
            with self.subTest(source_segments=source_segments):
                self.assertEqual(get_matching_regions(source_segments), expected)

    def test_us_does_not_match_inside_other_words(self):
        self.assertEqual(get_matching_regions('Australia'), ['OTHER'])

    def test_url_fallback_rejects_non_web_or_unobserved_locale_markers(self):
        source_urls = [
            'mailto:en-us',
            'en-us/page',
            '/en-us/page',
            'ftp://www.fedex.com/en-us/page',
            'https://www.fedex.com/xx-us/page',
            'https://www.fedex.com/us-us/page',
            'https://en-us.example.com/page',
            'https://www.fedex.com/page?locale=en-us',
            'https://:443/en-us/page',
            'https://user@/en-us/page',
            'https://@/en-us/page',
        ]

        for source_url in source_urls:
            with self.subTest(source_url=source_url):
                self.assertEqual(
                    get_matching_regions(None, source_url),
                    ['OTHER'],
                )

    def test_existing_region_segments_still_match(self):
        for region in ['APAC', 'MEISA', 'EU', 'LAC']:
            with self.subTest(region=region):
                self.assertEqual(get_matching_regions(region), [region])


if __name__ == '__main__':
    unittest.main()
