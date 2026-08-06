"""Offline regression checks for source-only Excel and Power Query assets.

The repository intentionally stores Excel logic as text.  The CI environment
does not provide an Excel or Power Query host, so these checks pin the fixtures
and source-level safety guards that prevent the audited regressions from being
silently reintroduced.  Run the fixture scenarios in Excel/Power Query before
releasing changes to the modules.
"""

import csv
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class TrialBalanceFixtureTests(unittest.TestCase):
    def test_code_less_parenthetical_account_is_present_in_both_layouts(self):
        with (ROOT / "samples" / "sample-xero-trial-balance.csv").open(newline="", encoding="utf-8") as f:
            combined = list(csv.reader(f))
        with (ROOT / "samples" / "sample-xero-trial-balance-columns.csv").open(
            newline="", encoding="utf-8"
        ) as f:
            separate = list(csv.reader(f))

        self.assertIn(["Rent (Sydney)", "Expense", "", "", "", ""], combined)
        self.assertIn(["", "Rent (Sydney)", "Expense", "", "", "", ""], separate)

    def test_combined_parser_guards_the_dash_fallback(self):
        source = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r'dashCandidate\s*=\s*if Text\.Contains\(trimmed, " - "\)\s*'
                r'then Text\.BeforeDelimiter\(trimmed, " - "\)',
                re.MULTILINE,
            ),
        )


class ReconResultSafetyTests(unittest.TestCase):
    def test_vba_requires_the_generated_sheet_marker_before_deleting(self):
        module = ROOT / "vba" / "modReconCompare.bas"
        source = module.read_text(encoding="utf-8")
        self.assertIn('Private Const RESULT_TAG_NAME As String = "__ReconCompareResultSheet"', source)
        self.assertIn("If Not IsGeneratedResultSheet(wb, stale) Then", source)
        self.assertIn("If Not marker Is Nothing Then", source)
        self.assertIn("previousAlerts = Application.DisplayAlerts", source)
        self.assertIn("Application.DisplayAlerts = previousAlerts", source)
        self.assertLess(source.index("If Not IsGeneratedResultSheet(wb, stale) Then"), source.index("stale.Delete"))
        self.assertNotIn('wb.Sheets("Recon Result").Delete', source)
        self.assertNotIn(b"\n", module.read_bytes().replace(b"\r\n", b""))


if __name__ == "__main__":
    unittest.main()
