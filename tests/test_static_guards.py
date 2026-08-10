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

    def test_abn_validator_rejects_unexpected_formatting_and_conversion_errors(self):
        source = (ROOT / "powerquery" / "Fx.ABNIsValid.pq").read_text(encoding="utf-8")
        self.assertIn('Text.From(abn ?? "", "en-AU")', source)
        self.assertIn('raw = if textAttempt[HasError] then "" else textAttempt[Value]', source)
        self.assertIn('allowedFormatting = Text.Select(raw, {"0".."9", " "}) = raw', source)
        self.assertIn('digitsOnly = Text.Remove(raw, {" "})', source)
        self.assertNotIn("Text.Trim(", source)

    def test_abn_format_guards_are_consulted_by_the_result_expression(self):
        """Binding a guard is not using it.

        The previous version of this test compared a Python set against a
        Python set and would have passed with every guard removed from the M.
        Pin each clause into the `result` conditional instead, so deleting one
        fails here.
        """
        source = (ROOT / "powerquery" / "Fx.ABNIsValid.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r"result\s*=\s*"
                r"if not allowedFormatting\s*"
                r"or Text\.Length\(digitsOnly\) <> 11\s*"
                r'or Text\.Select\(digitsOnly, \{"0"\.\."9"\}\) <> digitsOnly\s*'
                r'or Text\.StartsWith\(digitsOnly, "0"\) then\s*'
                r"false",
                re.MULTILINE,
            ),
        )

    def test_abn_reference_contract_matches_the_power_query_constants(self):
        """A Python port of the ATO checksum, pinned to the M's own constants.

        This does not execute the M - there is no Power Query host in CI - so
        it is a specification test: the vectors document the contract, and the
        constant assertions fail if the M drifts away from it.
        """
        source = (ROOT / "powerquery" / "Fx.ABNIsValid.pq").read_text(encoding="utf-8")
        weights = (10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19)
        self.assertIn(
            "weights = {" + ", ".join(str(w) for w in weights) + "}",
            source,
        )
        self.assertIn("Number.Mod(total, 89) = 0", source)

        def abn_is_valid(value: str) -> bool:
            if set(value) - set("0123456789 "):
                return False
            digits = value.replace(" ", "")
            if len(digits) != 11 or not digits.isdigit() or digits.startswith("0"):
                return False
            numbers = [int(d) for d in digits]
            numbers[0] -= 1
            return sum(n * w for n, w in zip(numbers, weights)) % 89 == 0

        for value in ("51824753556", "51 824 753 556", " 51 824 753 556 "):
            with self.subTest(accepted=value):
                self.assertTrue(abn_is_valid(value))
        for value in (
            "\t51 824 753 556",
            "51 824 753 556\n",
            "51\r824 753 556",
            "51\u00a0824 753 556",
            "51\u2007824 753 556",
            "51\u202f824 753 556",
            "51-824-753-556",
            "5182475355A",
            "5182475355",
            "51824753557",
            "00000090000",
        ):
            with self.subTest(rejected=repr(value)):
                self.assertFalse(abn_is_valid(value))

    def test_abn_checksum_is_implemented_once(self):
        """Match the implementation, not the topic.

        Forbidding the substring "ABN" everywhere else broke the moment any
        other query called or even named this helper, which is exactly the
        composition the README advertises.
        """
        implementations = sorted(
            path.name
            for path in (ROOT / "powerquery").glob("*.pq")
            if "Number.Mod(total, 89)" in path.read_text(encoding="utf-8")
        )
        self.assertEqual(implementations, ["Fx.ABNIsValid.pq"])

    def test_header_promoter_guards_run_before_any_row_is_examined(self):
        """M let-bindings are lazy.

        Bound only inside the Table.SelectRows predicate, neither guard fires
        for a table with no rows. Checked has to feed Table.AddIndexColumn.
        """
        source = (ROOT / "powerquery" / "Fx.PromoteHeaderAt.pq").read_text(encoding="utf-8")
        self.assertIn("if List.IsEmpty(ColumnNames) then", source)
        self.assertIn("Input table has no columns", source)
        self.assertIn("Table.AddIndexColumn(Checked,", source)
        # The blank check and the row matcher must use the same normalised
        # value, or a header pasted with a trailing space matches nothing.
        self.assertIn("Wanted = Text.Trim(FirstHeaderValue)", source)
        self.assertIn('else if Wanted = "" then', source)
        self.assertRegex(
            source,
            re.compile(
                r'each Text\.Trim\(Text\.From\(Record\.Field\(_, FirstColumn\) \?\? ""\)\) = Wanted'
            ),
        )
        self.assertNotIn("= FirstHeaderValue\n", source)


class ReconResultSafetyTests(unittest.TestCase):
    def test_reconciliation_refuses_missing_ranges_and_negative_tolerance(self):
        source = (ROOT / "vba" / "modReconCompare.bas").read_text(encoding="utf-8")
        self.assertIn("If rangeA Is Nothing Or rangeB Is Nothing Then", source)
        self.assertIn("If tolerance < 0 Then", source)

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

    def test_dead_marker_is_cleaned_up_instead_of_raised_at_the_user(self):
        """Deleting the generated sheet by hand used to brick the macro: the
        leftover marker is a HIDDEN defined name, invisible in Name Manager,
        and the error told the user to remove it.  The no-sheet branch must
        clean up a dead marker itself and only stop when the marker still
        resolves to a live (renamed, kept) sheet, with Excel-UI steps."""
        source = (ROOT / "vba" / "modReconCompare.bas").read_text(encoding="utf-8")
        branch = source[
            source.index("If stale Is Nothing Then") : source.index(
                "If Not IsGeneratedResultSheet(wb, stale) Then"
            )
        ]
        self.assertIn("Set keptSheet = MarkerTaggedSheet(wb, marker)", branch)
        self.assertIn("If keptSheet Is Nothing Then", branch)
        self.assertIn("marker.Delete", branch)
        self.assertLess(branch.index("If keptSheet Is Nothing Then"), branch.index("marker.Delete"))
        self.assertIn("Rename it back to", branch)
        self.assertNotIn("remove or rename the marker", source)

    def test_survived_delete_is_caught_before_the_marker_is_dropped(self):
        """Excel refuses to delete the last visible sheet, and with
        DisplayAlerts off it refuses silently.  The module must re-check the
        sheet after the delete and raise BEFORE removing the marker, so the
        surviving sheet stays recognisable as generated on the next run, and
        the raise must run with the DisplayAlerts-restoring handler armed."""
        source = (ROOT / "vba" / "modReconCompare.bas").read_text(encoding="utf-8")
        i_delete = source.index("stale.Delete")
        i_recheck = source.index("Set survivor = wb.Sheets(RESULT_SHEET_NAME)")
        i_raise = source.index("only visible sheet")
        i_marker_cleanup = source.index("If Not marker Is Nothing Then marker.Delete")
        self.assertLess(i_delete, i_recheck)
        self.assertLess(i_recheck, i_raise)
        self.assertLess(i_raise, i_marker_cleanup)
        self.assertIn("On Error GoTo DeleteFailed", source[i_recheck:i_raise])

    def test_marker_identity_survives_cell_edits_on_the_result_sheet(self):
        """The marker used to anchor ownership to the deletable cell $A$1:
        removing row 1 or column A broke it to #REF! and the module disowned
        its own sheet.  Identity must ride on the sheet's CodeName, stored as
        a text constant, with the legacy range anchor still recognised."""
        source = (ROOT / "vba" / "modReconCompare.bas").read_text(encoding="utf-8")
        self.assertIn('RefersTo:="=""" & ws.CodeName & """"', source)
        self.assertIn("Private Function StoredCodeName", source)
        self.assertIn("Private Function SheetWithCodeName", source)
        recognise = source[source.index("Private Function IsGeneratedResultSheet") :]
        recognise = recognise[: recognise.index("End Function")]
        self.assertIn("MarkerTaggedSheet(wb, marker)", recognise)
        self.assertNotIn("RefersToRange", recognise)
        legacy = source[source.index("Private Function MarkerTaggedSheet") :]
        legacy = legacy[: legacy.index("End Function")]
        self.assertIn("RefersToRange", legacy)


if __name__ == "__main__":
    unittest.main()
