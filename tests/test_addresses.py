import unittest
from desktop_draft_automation import address_case, address_lines, address_value, postal_conflict


class AddressTests(unittest.TestCase):
    def test_address_words_are_capitalized_without_changing_short_acronyms(self):
        self.assertEqual(address_case('flat 12, lakeview residency, HDFC EGL office'), 'Flat 12, Lakeview Residency, HDFC EGL Office')
        self.assertEqual(address_case('Flat 12, Lakeview Residency'), 'Flat 12, Lakeview Residency')

    def test_existing_details_preserved_and_only_missing_added(self):
        c = dict(address1='Flat 12, Baner, PUNE, Maharashtra 411045', address2='', city='Pune', state='Maharashtra', country='India', postal='411045')
        self.assertEqual(address_lines(c), [c['address1'], 'India'])
        self.assertEqual(address_value(c), 'Flat 12, Baner, PUNE, Maharashtra 411045, India')

    def test_separate_fields_are_deduplicated(self):
        c = dict(address1='112 Maple Residency', address2='', city='Singapore', state='Singapore', country='Singapore', postal='048624')
        self.assertEqual(address_lines(c), ['112 Maple Residency', 'Singapore, 048624'])

    def test_spaces_and_punctuation_in_postal(self):
        c = dict(address1='House 4', address2='Pune / Maharashtra / India / 411-045', city='pune', state='Maharashtra', country='India', postal='411045')
        self.assertEqual(address_lines(c), [c['address1'], c['address2']])

    def test_substrings_do_not_hide_city(self):
        c = dict(address1='Yorkshire House', address2='', city='York', state='', country='', postal='')
        self.assertEqual(address_lines(c), ['Yorkshire House', 'York'])

    def test_labelled_conflict(self):
        c = dict(address1='Flat 12, PIN: 411046', address2='', country='India', postal='411045')
        self.assertTrue(postal_conflict(c))
        c['address1'] = 'Flat 12, PIN: 411-045'
        self.assertFalse(postal_conflict(c))
        c['address1'] = 'Building 411046'
        self.assertFalse(postal_conflict(c))
