# Contributing

Power Query functions and VBA modules for Australian month-end work. Everything lives as text so a reviewer can read your change in a diff. Ship the source as text, and use a workbook only to demonstrate it.

## Data boundary

Keep client data out. The `.gitignore` blocks `input/`, `output/`, `clients/`, `exports/` and the common export extensions, including `.aba`, `.myox`, `.ofx` and `.qif`. Fabricate any sample you need and keep it under `samples/`.

## Format rules

- Export VBA as text under `vba/`. The VBE expects CRLF on import, so `.gitattributes` marks `.bas`, `.cls` and `.frm` as `-text whitespace=cr-at-eol`, which keeps git from touching their line endings and stops a trailing CR being reported as a whitespace error. That means `git diff --check` will not flag a renormalised blob. `tools/check_vba_encoding.py` is the check that does, and CI runs it.
- Keep Power Query M under `powerquery/` as plain text.
- Do not commit `.xlsm` or `.xlsx` as the source of truth for a function or module.

## Traps already found here

- `Csv.Document` infers width from the first row, so a one-field title row collapses the parse. Pass `Columns`, and put a caveat in a full-width trailing row.
- Typed M parameters such as `(x as date)` reject the shapes Excel hands over. Widen the type and use `Date.From(d, "en-AU")` so financial-year results stop following the machine locale.
- CSV formula injection: guard `=` in every case, and `+`, `-`, `@` only when the remainder is not a plain code. Otherwise account codes like `-00123` stop joining back to payroll. An A1-style reference such as `+A1` starts a formula, so an "inert remainder" test has to be narrower than "anything alphanumeric".
- `IsNumeric(True)` returns `True` in VBA, so Booleans sum as -1. `Activate` on a hidden sheet activates the visible neighbour. `Scripting.Dictionary` runs on Windows Excel only.
- `Trim$` misses non-breaking spaces, tabs and CRLF in reconciliation keys.

## Local verification

Python 3.10 or newer drives the test suite.

```bash
python -B -m unittest discover -s tests -v
```

On Windows with desktop Excel installed, run the checked-in Power Query
function in Excel itself against both fabricated Xero fixtures:

```powershell
powershell -NoProfile -File tools/native_excel_acceptance.ps1
```

The workbook is saved only under the operating system's temporary directory so
Power Query has a stable host path, then it is deleted in the runner's `finally`
block. VBA is not imported by automation because doing so depends on Excel's
machine-wide **Trust access to the VBA project object model** policy. Test VBA
in a disposable workbook by importing both `.bas` files, running
`ApplyWorkpaperHeader` against an entity name beginning with `=`, and running
`CompareKeyedRanges` against fabricated two-column ranges with a leading-zero
key. Confirm the header is text rather than a formula, existing rows move down,
and `Recon Result` carries the expected difference. Do not weaken the Office
trust policy solely to run this check.

## Pull requests

Say which function or module you changed and show the fixture that exercises it. For an M change, state the locale and column shape you tested against.

For a potential security vulnerability, follow [SECURITY.md](SECURITY.md) rather than opening an issue.
