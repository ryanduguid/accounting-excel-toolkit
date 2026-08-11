# Contributing

Power Query functions and VBA modules for Australian month-end work. Everything here lives as text so it can be reviewed in a diff. Source buried inside a binary workbook is not a contribution.

## Data boundary

No client data. The `.gitignore` blocks `input/`, `output/`, `clients/`, `exports/` and every common export extension, including `.aba`, `.myox`, `.ofx` and `.qif`. Fabricate any sample you need and keep it under `samples/`.

## Format rules

- VBA is exported as text under `vba/`. The VBE expects CRLF on import, so `.bas`, `.cls` and `.frm` are pinned in `.gitattributes` and checked by `tools/check_vba_encoding.py`. A renormalised blob shows up in `git diff --check`.
- Power Query M lives under `powerquery/` as plain text.
- Do not commit `.xlsm` or `.xlsx` as the source of truth for any function or module.

## Traps this toolkit already hit

- `Csv.Document` infers width from the first row, so a one-field title row collapses the parse. Always pass `Columns`, and put a caveat in a full-width trailing row.
- Typed M parameters such as `(x as date)` reject the shapes Excel actually hands over. Widen the type and use `Date.From(d, "en-AU")` so financial-year results do not follow the machine locale.
- CSV formula injection: guard `=` always, and `+`, `-`, `@` only when the remainder is not a plain code. Otherwise account codes like `-00123` stop joining back to payroll. An A1-style reference such as `+A1` is a formula start, so an "inert remainder" test has to be narrower than "anything alphanumeric".
- `IsNumeric(True)` is `True` in VBA, so Booleans sum as -1. `Activate` on a hidden sheet activates the visible neighbour. `Scripting.Dictionary` is Windows-only.
- `Trim$` alone misses non-breaking spaces, tabs and CRLF in reconciliation keys.

## Local verification

Python 3.10 or newer drives the test suite.

```bash
python -B -m unittest discover -s tests -v
```

## Pull requests

Say which function or module changed and show the fixture that exercises it. For an M change, state the locale and column shape you tested against.

For a potential security vulnerability, follow [SECURITY.md](SECURITY.md) rather than opening an issue.
