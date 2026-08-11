"""Offline regression checks for source-only Excel and Power Query assets.

The repository intentionally stores Excel logic as text.  The CI environment
does not provide an Excel or Power Query host, so these checks pin the fixtures
and source-level safety guards that prevent the audited regressions from being
silently reintroduced.  Run the fixture scenarios in Excel/Power Query before
releasing changes to the modules.
"""

import csv
import re
import string
import unittest
from collections import namedtuple
from decimal import Decimal
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

COMBINED_FIXTURE = ROOT / "samples" / "sample-xero-trial-balance.csv"
SEPARATE_FIXTURE = ROOT / "samples" / "sample-xero-trial-balance-columns.csv"

# Text.Select(c, {"0".."9", "A".."Z", "a".."z"}) in the M: ASCII only, so
# str.isalnum() (which accepts every Unicode letter and digit) will not do.
ASCII_ALNUM = frozenset(string.ascii_letters + string.digits)
ASCII_DIGITS = frozenset(string.digits)

Account = namedtuple("Account", "code name type debit credit")


class TrialBalanceError(Exception):
    """What the ported parser raises where the M raises. A distinct type, so
    a test can assert the refusal without assertRaises(AssertionError)
    swallowing a genuine assertion failure from the same block."""


def _fixture_rows(path):
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.reader(f))


def _header_index(rows, name):
    """The header row is the one whose FIRST FIELD is exactly "Account" or
    "Account Code" - the same test Xero.TrialBalance.pq runs."""
    for index, row in enumerate(rows):
        if row and row[0].strip() in ("Account", "Account Code"):
            return index
    raise TrialBalanceError("no header row in %s" % name)


def _without_ytd_columns(rows, name):
    """The same rows as a pre-YTD Xero export: every YTD column dropped.

    Both committed fixtures carry both pairs, which is what makes the
    default and useYTD = true the same column choice on them. This is the
    export shape that tells those two branches apart.
    """
    header = [cell.strip() for cell in rows[_header_index(rows, name)]]
    keep = [index for index, cell in enumerate(header) if not cell.startswith("YTD ")]
    return [[row[index] for index in keep if index < len(row)] for row in rows]


def _is_code(candidate, whole_cell):
    """The isCode predicate from Xero.TrialBalance.pq, ported.

    Alphanumeric, at most 10 characters, at least one digit, and not the
    whole cell - so "Rent (Sydney)" keeps its parenthetical.
    """
    if candidate is None or candidate == "" or candidate == whole_cell:
        return False
    if len(candidate) > 10:
        return False
    if not any(character in ASCII_DIGITS for character in candidate):
        return False
    return all(character in ASCII_ALNUM for character in candidate)


def _split_combined_account(cell):
    """The combined-layout code/name split, ported: trailing " (code)" first,
    a "code - name" dash prefix as the fallback, otherwise no code."""
    trimmed = cell.strip()

    paren_candidate = None
    if trimmed.endswith(")"):
        paren_position = trimmed.rfind(" (")
        if paren_position >= 0:
            paren_candidate = trimmed[paren_position + 2 : len(trimmed) - 1]
    dash_candidate = trimmed.split(" - ")[0] if " - " in trimmed else None

    if _is_code(paren_candidate, trimmed):
        code = paren_candidate
    elif _is_code(dash_candidate, trimmed):
        code = dash_candidate
    else:
        return None, trimmed

    if trimmed.endswith(" (" + code + ")"):
        return code, trimmed[: len(trimmed) - len(code) - 3]
    return code, trimmed.split(" - ", 1)[1]


def _parse_amount(cell):
    """Blank is a genuine zero. Decimal, not float: the fixtures are money."""
    text = cell.strip()
    return Decimal("0") if text == "" else Decimal(text)


def _parse_trial_balance(path, use_ytd=None):
    return _parse_rows(_fixture_rows(path), path.name, use_ytd=use_ytd)


def _parse_rows(rows, name, use_ytd=None):
    """Xero.TrialBalance.pq's documented contract, ported for the fixtures.

    CI has no Power Query host, so this executes the specification, not the
    M. It is what makes the fixture assertions below about behaviour; the
    source pins in the tests that follow are what catch the M drifting away
    from it.
    """
    start = _header_index(rows, name)
    header = [cell.strip() for cell in rows[start]]

    split_layout = header[0] == "Account Code"
    has_period = "Debit" in header and "Credit" in header
    has_ytd = "YTD Debit" in header and "YTD Credit" in header

    # Debit/Credit are the current-period MOVEMENT, YTD Debit/YTD Credit the
    # AS-AT balances. Null (the default) prefers as-at; explicit true/false
    # demands its pair and errors when the export lacks it.
    if use_ytd is True:
        if not has_ytd:
            raise TrialBalanceError("useYTD = true but %s has no YTD pair" % name)
        take_ytd = True
    elif use_ytd is False:
        if not has_period:
            raise TrialBalanceError("useYTD = false but %s has no period pair" % name)
        take_ytd = False
    else:
        take_ytd = has_ytd

    debit_column = header.index("YTD Debit" if take_ytd else "Debit")
    credit_column = header.index("YTD Credit" if take_ytd else "Credit")
    type_column = header.index("Account Type")

    accounts = []
    for row in rows[start + 1 :]:
        if not row:
            continue
        if split_layout:
            code_cell, name_cell = row[0].strip(), row[1].strip()
            if code_cell == "" and name_cell == "":
                continue
            if code_cell == "Total" or name_cell == "Total":
                continue
            code, name = (code_cell or None), name_cell
        else:
            cell = row[0].strip()
            if cell in ("", "Total"):
                continue
            code, name = _split_combined_account(cell)
        accounts.append(
            Account(
                code,
                name,
                row[type_column].strip(),
                _parse_amount(row[debit_column]),
                _parse_amount(row[credit_column]),
            )
        )
    return accounts


def _column_totals(path):
    """Sum every amount column of a fixture, and read the fixture's own
    trailing Total row - the one the parser drops - for comparison."""
    rows = _fixture_rows(path)
    start = _header_index(rows, path.name)
    header = [cell.strip() for cell in rows[start]]
    amount_columns = ["Debit", "Credit", "YTD Debit", "YTD Credit"]

    summed = {name: Decimal("0") for name in amount_columns}
    declared = {}
    for row in rows[start + 1 :]:
        if not row:
            continue
        is_total_row = any(cell.strip() == "Total" for cell in row[:2])
        for name in amount_columns:
            value = _parse_amount(row[header.index(name)])
            if is_total_row:
                declared[name] = value
            else:
                summed[name] += value
    return summed, declared


class TrialBalanceFixtureTests(unittest.TestCase):
    def test_the_readme_promises_these_tests_quote_are_still_in_the_readme(self):
        """The docstrings below justify themselves by quoting the README.  A
        quote beats a line number - a one-line insert moves every number and
        makes the comment silently false - but a reworded sentence would rot
        a quote just as quietly, so the sentences are pinned here.  This pins
        the side that changes: if a rewrite lands, this fails and names the
        docstrings that have to be rewritten with it."""
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        for quoted in (
            # test_both_fixtures_balance_and_match_their_own_total_row
            "After loading a TB, the first check is always",
            # test_both_layouts_carry_the_same_accounts_and_amounts
            "carrying the same accounts and amounts in both shapes",
            # test_code_less_parenthetical_account_parses_with_no_code
            "it must load with a null `AccountCode` and its full name intact",
            # test_default_takes_the_as_at_pair_and_useYTD_false_the_movement
            "default returns as-at, `useYTD = false` for movement",
        ):
            with self.subTest(quoted=quoted):
                self.assertIn(quoted, readme)

    def test_code_less_parenthetical_account_is_present_in_both_layouts(self):
        with (ROOT / "samples" / "sample-xero-trial-balance.csv").open(newline="", encoding="utf-8") as f:
            combined = list(csv.reader(f))
        with (ROOT / "samples" / "sample-xero-trial-balance-columns.csv").open(
            newline="", encoding="utf-8"
        ) as f:
            separate = list(csv.reader(f))

        self.assertIn(["Rent (Sydney)", "Expense", "", "", "", ""], combined)
        self.assertIn(["", "Rent (Sydney)", "Expense", "", "", "", ""], separate)

    def test_both_fixtures_balance_and_match_their_own_total_row(self):
        """The README: "After loading a TB, the first check is always" that
        debits equal credits within tolerance.  Nothing checked that of the
        fixtures themselves, so one mistyped amount would send everyone
        following those instructions off to debug a parser that is working
        correctly."""
        for path in (COMBINED_FIXTURE, SEPARATE_FIXTURE):
            with self.subTest(fixture=path.name):
                summed, declared = _column_totals(path)
                self.assertEqual(summed["Debit"], summed["Credit"])
                self.assertEqual(summed["YTD Debit"], summed["YTD Credit"])
                self.assertEqual(summed["Debit"], Decimal("6995.00"))
                self.assertEqual(summed["YTD Debit"], Decimal("129934.50"))
                # The fixture's own Total row is a second opinion on the same
                # numbers; the parser drops it, so it can drift unnoticed.
                self.assertEqual(summed, declared)

    def test_both_layouts_carry_the_same_accounts_and_amounts(self):
        """The README calls the two fixtures "fabricated, balanced TBs
        carrying the same accounts and amounts in both shapes".  Stripping
        the 090 code from the separate-column fixture, or editing one amount
        in either, left the whole suite green."""

        def projection(path):
            as_at = _parse_trial_balance(path)
            movement = _parse_trial_balance(path, use_ytd=False)
            self.assertEqual(len(as_at), len(movement))
            rows = [
                (
                    account.code,
                    account.name,
                    account.type,
                    str(account.debit),
                    str(account.credit),
                    str(period.debit),
                    str(period.credit),
                )
                for account, period in zip(as_at, movement)
            ]
            # A code-less account carries None, which will not sort against
            # a str, so order on a stringified copy and keep the None.
            return sorted(rows, key=lambda row: tuple(value or "" for value in row))

        combined = projection(COMBINED_FIXTURE)
        self.assertEqual(len(combined), 12)
        self.assertEqual(combined, projection(SEPARATE_FIXTURE))

    def test_default_takes_the_as_at_pair_and_useYTD_false_the_movement(self):
        """The contract the README states, in "picks the right Debit/Credit
        pair (plain pair = period movement, YTD pair = as-at balances;
        default returns as-at, `useYTD = false` for movement)".  The two
        answers both balance, so a swap is invisible to the balance check the
        README prescribes - only the totals tell them apart."""
        for path in (COMBINED_FIXTURE, SEPARATE_FIXTURE):
            with self.subTest(fixture=path.name):
                as_at = _parse_trial_balance(path)
                self.assertEqual(sum(a.debit for a in as_at), Decimal("129934.50"))
                self.assertEqual(sum(a.credit for a in as_at), Decimal("129934.50"))

                movement = _parse_trial_balance(path, use_ytd=False)
                self.assertEqual(sum(a.debit for a in movement), Decimal("6995.00"))
                self.assertEqual(sum(a.credit for a in movement), Decimal("6995.00"))
                self.assertNotEqual(movement, as_at)

                # Asserting useYTD = true equals the default on a fixture
                # that carries both pairs compares the port with itself:
                # both reach take_ytd = True down the same branch.  The
                # branch worth exercising is the export WITHOUT the YTD
                # pair, where the default falls back to the movement and
                # useYTD = true has to refuse rather than hand back the
                # movement labelled as an as-at balance.
                rows = _without_ytd_columns(_fixture_rows(path), path.name)
                self.assertEqual(_parse_rows(rows, path.name), movement)
                with self.assertRaises(TrialBalanceError) as refused:
                    _parse_rows(rows, path.name, use_ytd=True)
                self.assertIn("no YTD pair", str(refused.exception))

    def test_code_less_parenthetical_account_parses_with_no_code(self):
        """The other half of the same README sentence: the fixtures "include
        a code-less `Rent (Sydney)` account: it must load with a null
        `AccountCode` and its full name intact".  Split as code "Sydney" /
        name "Rent" it collides with the real Rent (469) account in any lead
        schedule keyed on AccountName."""
        for path in (COMBINED_FIXTURE, SEPARATE_FIXTURE):
            with self.subTest(fixture=path.name):
                parsed = {(a.code, a.name) for a in _parse_trial_balance(path)}
                self.assertIn((None, "Rent (Sydney)"), parsed)
                self.assertNotIn(("Sydney", "Rent"), parsed)
                self.assertIn(("469", "Rent"), parsed)
                # Leading zero survives because codes are text, not numbers.
                self.assertIn(("090", "Business Bank Account"), parsed)

    def test_combined_parser_pins_the_is_code_predicate(self):
        """Behaviour above is a specification; this is what fails when the M
        drifts.  Drop the at-least-one-digit rule, or widen the 10-character
        cap, and "Rent (Sydney)" splits into code "Sydney" / name "Rent"."""
        source = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r"isCode = \(c as nullable text\) as logical =>\s*"
                r"c <> null\s*"
                r'and c <> ""\s*'
                r"and c <> trimmed\s*"
                r"and Text\.Length\(c\) <= 10\s*"
                r'and Text\.Select\(c, \{"0"\.\."9"\}\) <> ""\s*'
                r'and Text\.Select\(c, \{"0"\.\."9", "A"\.\."Z", "a"\.\."z"\}\) = c,',
                re.MULTILINE,
            ),
        )
        # Binding the predicate is not using it.
        self.assertIn("if isCode(parenCandidate) then parenCandidate", source)
        self.assertIn("else if isCode(dashCandidate) then dashCandidate", source)

    def test_combined_parser_pins_the_code_taken_by_the_parenthetical_split(self):
        """The arm that actually extracts the code, pinned like the two it
        sits between.  isCode says what counts as a code and the AccountName
        arithmetic says what is left as the name, but neither looks at the
        offsets that lift "090" out of "Business Bank Account (090)", and
        those offsets have the same off-by-one shape - a " (" two characters
        wide and a " (" plus ")" three characters wide.

        Change the + 2 to a + 1 and every parenthetical candidate keeps a
        leading "(", which isCode rejects, so AccountCode goes null on every
        coded account and AccountName keeps the raw "Accounts Receivable
        (610)" cell - every lead schedule keyed on code or name breaks, and
        isCode, the dash fallback and the fixture assertions all stay green
        because the fixture assertions run the Python port, not the M.

        Occurrence.Last and the trailing-")" gate are pinned in the same
        regex because the three are one rule: on "Rent (Sydney) (469)" First
        takes "Sydney) (469" and the real 469 is lost, and the "- 3" length
        is only correct because the gate has already established that the
        last character is the ")" it subtracts."""
        source = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r"parenPos =\s*"
                r'if Text\.EndsWith\(trimmed, "\)"\)\s*'
                r'then Text\.PositionOf\(trimmed, " \(", Occurrence\.Last\)\s*'
                r"else -1,\s*"
                r"parenCandidate =\s*"
                r"if parenPos >= 0\s*"
                r"then Text\.Range\(trimmed, parenPos \+ 2, Text\.Length\(trimmed\) - parenPos - 3\)\s*"
                r"else null,",
                re.MULTILINE,
            ),
        )

    def test_combined_parser_pins_the_name_left_by_the_split(self):
        """The other arm of the same code/name split.  isCode decides which
        text is the code; this decides what is left as the name, and it is
        arithmetic on lengths - the " (" and the ")" it strips are three
        characters.  Change that 3 to a 2 and "Business Bank Account (090)"
        loads as "Business Bank Account " with a trailing space, which
        silently breaks every lead schedule and recon keyed on AccountName;
        the isCode pin above and every fixture assertion stay green, because
        the code itself is still "090"."""
        source = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r'WithName = Table\.AddColumn\(\s*WithCode,\s*"AccountName",\s*each\s*'
                r"let trimmed = Text\.Trim\(\[Account\]\)\s*in\s*"
                r"if \[AccountCode\] = null then trimmed\s*"
                r'else if Text\.EndsWith\(trimmed, " \(" & \[AccountCode\] & "\)"\)\s*'
                r"then Text\.Start\(\s*trimmed,\s*"
                r"Text\.Length\(trimmed\) - Text\.Length\(\[AccountCode\]\) - 3\s*\)\s*"
                r'else Text\.AfterDelimiter\(trimmed, " - "\),',
                re.MULTILINE,
            ),
        )
        # The derived name replaces the raw combined cell rather than
        # sitting beside it, so nothing downstream can read the unsplit one.
        self.assertIn('Table.RemoveColumns(WithName, {"Account"})', source)

    def test_combined_parser_pins_the_debit_credit_pair_selection(self):
        """Both pairs balance, so choosing the wrong one still passes the
        README's balance check and gives the accountant positive confirmation
        of a materially wrong trial balance.  Pin all three branches: the
        default prefers YTD, and each explicit choice errors rather than
        falling back to the pair the caller did not ask for - and the two
        predicates those branches read, because pinning a branch without its
        input leaves the default flippable from the other end.  Rename either
        YTD column header upstream and hasYTD goes false, so the default
        hands back the current-period MOVEMENT in columns still named Debit
        and Credit; that table balances too, so the README's prescribed check
        confirms it."""
        source = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertIn(
            'hasPeriod = List.Contains(cols, "Debit") and List.Contains(cols, "Credit"),',
            source,
        )
        self.assertIn(
            'hasYTD = List.Contains(cols, "YTD Debit") and List.Contains(cols, "YTD Credit"),',
            source,
        )
        self.assertRegex(
            source,
            re.compile(
                r"takeYTD =\s*if useYTD = true then\s*if hasYTD then true\s*else error",
                re.MULTILINE,
            ),
        )
        self.assertRegex(
            source,
            re.compile(
                r"else if useYTD = false then\s*if hasPeriod then false\s*else error",
                re.MULTILINE,
            ),
        )
        self.assertRegex(source, re.compile(r"else\s*hasYTD,\s*Selected =", re.MULTILINE))
        # ... and that takeYTD = true means the YTD pair, not the other one.
        self.assertRegex(
            source,
            re.compile(
                r"else if takeYTD then\s*Table\.RenameColumns\(\s*"
                r'Table\.RemoveColumns\(Promoted, \{"Debit", "Credit"\}, MissingField\.Ignore\),\s*'
                r'\{\{"YTD Debit", "Debit"\}, \{"YTD Credit", "Credit"\}\}',
                re.MULTILINE,
            ),
        )

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
        # The weights and the modulus were pinned; the subtract-one step was
        # not.  Turning it into "+ 1" keeps both of those assertions passing
        # and rejects the ATO's own published example, so every genuinely
        # valid ABN in a supplier list is flagged invalid.
        self.assertIn("adjusted = {digits{0} - 1} & List.Skip(digits, 1)", source)
        self.assertIn("adjusted{_} * weights{_}", source)

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

    def test_header_slice_takes_the_first_match_and_keeps_the_header_row(self):
        """Finding the header row is guarded above; consuming the index it
        found is not, and both halves of that are silent when wrong.

        List.Max instead of List.Min slices at the LAST row carrying the
        header value, so an export that repeats it - an account literally
        named "Account", a repeated section header in a bank export - loses
        every data row above the repeat with no error.  HeaderIdx + 1 in the
        RemoveFirstN drops the header row itself, so PromoteHeaders promotes
        the first DATA row as the column names: the first account disappears
        and every downstream column reference breaks.

        Xero.TrialBalance.pq runs the same two steps over its own match list
        and is pinned here with the promoter, so the pair cannot be corrected
        in one file and left wrong in the other."""
        promoter = (ROOT / "powerquery" / "Fx.PromoteHeaderAt.pq").read_text(encoding="utf-8")
        self.assertIn("List.Min(Matches),", promoter)
        self.assertIn("Sliced = Table.RemoveFirstN(Raw, HeaderIdx),", promoter)
        # The function's own header comment states this contract; pinning it
        # beside the code stops the two drifting apart.
        self.assertRegex(
            promoter,
            re.compile(
                r"Finds the first\s*//\s*row whose first cell equals FirstHeaderValue, "
                r"drops everything above it, and\s*//\s*promotes it to headers\.",
                re.MULTILINE,
            ),
        )

        trial_balance = (ROOT / "powerquery" / "Xero.TrialBalance.pq").read_text(encoding="utf-8")
        self.assertIn("List.Min(HeaderMatches),", trial_balance)
        self.assertIn("Sliced = Table.RemoveFirstN(Raw, HeaderIdx),", trial_balance)

    def test_financial_year_switches_a_datetimezone_to_australian_eastern(self):
        """Date.From on a datetimezone returns the date of the value's LOCAL
        equivalent, so the HOST's zone decided the answer: 9am +10:00 on
        1 July 2026 came out FY2027 on a Sydney desktop and FY2026 on a
        UTC-hosted scheduled refresh.  The offset has to be resolved before
        the date is taken, and CONVERTED rather than dropped: RemoveZone on
        its own is host-independent too, but it reads the UTC-stamped
        2026-06-30T14:30:00Z - the shape the Xero API returns for 12:30am on
        1 July in Sydney - as 30 June, which is the wrong FY on the AU
        desktop that the old code got right."""
        source = (ROOT / "powerquery" / "Fx.AUFinancialYear.pq").read_text(encoding="utf-8")
        self.assertRegex(
            source,
            re.compile(
                r"normalised =\s*"
                r"if Value\.Is\(d, type datetimezone\)\s*"
                r"then DateTimeZone\.RemoveZone\(DateTimeZone\.SwitchZone\(d, 10\)\)\s*"
                r"else d,",
                re.MULTILINE,
            ),
        )
        self.assertIn('asDate = Date.From(normalised, "en-AU"),', source)
        self.assertNotIn('Date.From(d, "en-AU")', source)
        self.assertLess(source.index("normalised ="), source.index("asDate ="))
        # The header comment has to say which clock the caller gets, and warn
        # the states that AEST is not their own local calendar.
        self.assertIn("CONVERTED to Australian Eastern Standard Time (+10:00)", source)
        self.assertIn("Residual:", source)
        self.assertIn("Adelaide/Darwin", source)


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
        # Both halves of CRLF: a CR-only module survives a bare-LF check
        # untouched and imports into the VBE as one line.
        stripped = module.read_bytes().replace(b"\r\n", b"")
        self.assertNotIn(b"\n", stripped)
        self.assertNotIn(b"\r", stripped)

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


class WorkpaperFormatSafetyTests(unittest.TestCase):
    """Pins modWorkpaperFormat's write-safety guards, which nothing pinned
    before this class existed.

    Not a claim about the sibling module. ReconResultSafetyTests covers only
    modReconCompare's Nothing and negative-tolerance checks and its
    result-sheet marker machinery; several of that module's other guards are
    still unpinned and could be deleted with the suite green.
    """

    def source(self):
        return (ROOT / "vba" / "modWorkpaperFormat.bas").read_text(encoding="utf-8")

    def sub_body(self, name):
        """One Public Sub's text, so a guard is pinned to the sub it belongs
        to instead of merely being counted somewhere in the module."""
        source = self.source()
        start = source.index("Public Sub " + name)
        return source[start : source.index("End Sub", start)]

    def test_header_cells_are_forced_to_text_before_the_write(self):
        """An entity name starting with "=" - "=Smith & Co Pty Ltd" is a real
        trading name shape once someone pastes from a formula cell - is stored
        as a live formula unless the cells are text-formatted FIRST, and the
        client name then renders as #NAME? on every page of the pack."""
        source = self.source()
        self.assertIn('.Range("A1:A4").NumberFormat = "@"', source)
        self.assertLess(
            source.index('.Range("A1:A4").NumberFormat = "@"'),
            source.index('.Range("A1").Value = entityName'),
        )

    def test_header_insert_clears_the_clipboard_and_matches_the_freeze_default(self):
        """A live cut/copy marquee turns Insert into a paste, so the clipboard
        block lands in rows 1:5 instead of blank rows.  The inserted block and
        FreezeBelowHeader's default have to stay the same five rows, or the
        freeze lands inside the header."""
        source = self.source()
        self.assertIn("Application.CutCopyMode = False", source)
        self.assertIn('ws.Rows("1:5").Insert Shift:=xlDown', source)
        self.assertLess(
            source.index("Application.CutCopyMode = False"),
            source.index('ws.Rows("1:5").Insert'),
        )
        self.assertIn("Optional ByVal headerRows As Long = 5", source)

    def test_each_worksheet_taking_sub_refuses_a_protected_sheet(self):
        """Every sub that takes a Worksheet refuses a protected one, checked
        inside each sub rather than by counting the guard module-wide.

        This deliberately does NOT claim every writing sub is guarded: the
        fourth public sub, FormatAsAccounting, takes a Range and carries no
        worksheet-level guard, so a name promising all four would be false.
        The module-wide count is a floor rather than an equality for the same
        reason - pinning it at exactly three would fail the very change that
        adds a guard to FormatAsAccounting.
        """
        guard = "If ws.ProtectContents Then Err.Raise 5"
        for name in ("ApplyWorkpaperHeader", "AddReviewerLine", "FreezeBelowHeader"):
            with self.subTest(sub=name):
                self.assertIn(guard, self.sub_body(name))
        self.assertGreaterEqual(self.source().count(guard), 3)
        self.assertIn("If headerRows < 1 Then Err.Raise 5", self.sub_body("FreezeBelowHeader"))
        self.assertIn(
            "If ws.Visible <> xlSheetVisible Then Err.Raise 5",
            self.sub_body("FreezeBelowHeader"),
        )

    def test_last_row_search_reads_formulas_and_pins_its_own_settings(self):
        """Find inherits the user's last Find-dialog settings when they are
        not passed.  With LookIn:=xlValues a formula-only footer row is
        invisible, and the reviewer sign-off is written over live data."""
        source = self.source()
        self.assertIn("LookIn:=xlFormulas", source)
        self.assertIn("LookAt:=xlPart", source)
        self.assertIn("SearchOrder:=xlByRows", source)
        self.assertIn("SearchDirection:=xlPrevious", source)

    def test_module_stays_crlf_for_the_vbe_import(self):
        """CRLF means both halves.  Stripping CRLF pairs and looking only for
        a bare LF passes a CR-only file untouched - nothing is stripped and
        there is no LF to find - and the VBE imports CR-only source as a
        single line, destroying the module."""
        module = ROOT / "vba" / "modWorkpaperFormat.bas"
        stripped = module.read_bytes().replace(b"\r\n", b"")
        self.assertNotIn(b"\n", stripped)
        self.assertNotIn(b"\r", stripped)


if __name__ == "__main__":
    unittest.main()
