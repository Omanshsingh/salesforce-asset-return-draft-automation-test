import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

import openpyxl

from desktop_draft_automation import read_cases
from draft_service import run_cases


class DuplicateDraftingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def read_rows(self, rows):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append([
            'Case ID', "User's Name", "User's Email", 'Line Address 1',
            'City', 'Country', 'Asset Serial Number', 'Duplicate',
        ])
        for row in rows:
            sheet.append(row)
        path = self.folder / 'cases.xlsx'
        workbook.save(path)
        workbook.close()
        cases, skipped = read_cases(path)
        self.assertEqual(skipped, [])
        return cases

    def row(self, case_id, flag, asset='Laptop SN-001', email='user@example.test'):
        return [case_id, 'Test User', email, '1 Test Street', 'Mumbai', 'India', asset, flag]

    def draft(self, cases):
        writer = Mock()
        writer.save.return_value = 'saved-draft-id'
        report = self.folder / 'Results.txt'
        count = run_cases(cases, writer, None, report, lambda _: None, lambda: False, delay=0)
        return count, writer, report.read_text(encoding='utf-8')

    def test_flagged_and_normal_cases_all_draft_without_review_errors(self):
        flags = ['Yes', 'y', 'true', '1', 'Duplicate', True, 1, 'No', None]
        cases = self.read_rows([self.row(f'CASE-{i}', flag) for i, flag in enumerate(flags)])
        for case in cases:
            self.assertEqual(case['status'], 'Ready for Draft')
            self.assertEqual(case['missing_fields'], [])
        count, writer, report = self.draft(cases)
        self.assertEqual(count, len(flags))
        self.assertEqual(writer.save.call_count, len(flags))
        self.assertEqual(report.count(': Draft created'), len(flags))
        self.assertNotIn('Needs review', report)
        self.assertNotIn('Duplicate', report)

    def test_flagged_row_in_group_drafts_all_distinct_assets_together(self):
        cases = self.read_rows([
            self.row('CASE-1', 'No'),
            self.row('CASE-1', 'Yes', asset='Monitor SN-002'),
            self.row('CASE-1', 'Duplicate'),
        ])
        self.assertEqual(len(cases), 1)
        self.assertEqual(cases[0]['source_row_count'], 3)
        count, writer, report = self.draft(cases)
        self.assertEqual(count, 1)
        body = writer.save.call_args.args[1]
        self.assertEqual(body.count('Laptop SN-001'), 1)
        self.assertEqual(body.count('Monitor SN-002'), 1)
        self.assertIn('CASE-1: Draft created', report)

    def test_flagged_case_with_invalid_email_still_uses_normal_validation(self):
        cases = self.read_rows([self.row('CASE-1', 'Yes', email='invalid-email')])
        self.assertEqual(cases[0]['missing_fields'], ["User's Email"])
        count, writer, report = self.draft(cases)
        self.assertEqual(count, 0)
        writer.save.assert_not_called()
        self.assertIn("Needs review: User's Email", report)
        self.assertNotIn('Duplicate', report)


if __name__ == '__main__':
    unittest.main()
