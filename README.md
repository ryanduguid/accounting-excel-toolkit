# accounting-excel-toolkit

Power Query (M) functions and VBA modules for accountants working with Australian ledger exports: a Xero trial balance parser, an AU financial-year helper, ABN validation, workpaper formatting, keyed reconciliations.

Every function here solves a problem the Australian ledger-export formats create. Written independently, from scratch, in my own time and on my own equipment.

## Power Query functions

One `.pq` file per function. To use: Excel/Power BI → Get Data → Blank Query → Advanced Editor → paste the file's contents → name the query as per the file name (e.g. `Xero.TrialBalance`).

| Function | Category | What it does |
|---|---|---|
| [`Xero.TrialBalance`](powerquery/Xero.TrialBalance.pq) | Xero | Parse a Xero TB CSV: skips metadata rows, picks the right Debit/Credit pair (plain pair = period movement, YTD pair = as-at balances; default returns as-at, `useYTD = false` for movement), drops Total row, splits account code as text (alphanumeric up to 10 chars with at least one digit, leading zeros survive). Handles both export layouts: the combined account cell (`Business Bank Account (090)`) and the separate Account Code / Account Name columns of the current UI export |
| [`Fx.PromoteHeaderAt`](powerquery/Fx.PromoteHeaderAt.pq) | Generic | Find-and-promote the real header row in any ledger export that buries it below title rows; errors clearly when the format changed |
| [`Fx.AUFinancialYear`](powerquery/Fx.AUFinancialYear.pq) | AU helpers | FY label, start, end for any date (1 July – 30 June) |
| [`Fx.ABNIsValid`](powerquery/Fx.ABNIsValid.pq) | AU helpers | ABN checksum validation (ATO weighting algorithm); checksum ≠ registered, so check ABN Lookup for status |

Test against [`samples/sample-xero-trial-balance.csv`](samples/sample-xero-trial-balance.csv) (combined layout) and [`samples/sample-xero-trial-balance-columns.csv`](samples/sample-xero-trial-balance-columns.csv) (separate-column layout). Both are fabricated, balanced TBs carrying the same accounts and amounts in both shapes, plus both the period-movement and YTD (as-at) pairs so the pair selection gets exercised. They also include a code-less `Rent (Sydney)` account: it must load with a null `AccountCode` and its full name intact.

After loading a TB, the first check is always: `Number.Abs(List.Sum(result[Debit]) - List.Sum(result[Credit])) < 0.005`, a tolerance, because the sums are IEEE doubles and exact `=` can fail on a genuinely balanced TB.

## VBA modules

Importable `.bas` source in [`vba/`](vba/). See [`vba/README.md`](vba/README.md) for import steps.

| Module | What it does |
|---|---|
| [`modWorkpaperFormat`](vba/modWorkpaperFormat.bas) | Workpaper header block, reviewer sign-off line, accounting number format, freeze panes |
| [`modReconCompare`](vba/modReconCompare.bas) | Keyed two-way recon between (key, amount) ranges with tolerance; duplicate keys summed; subledger vs GL pattern |

## Principles

1. **Source as text.** M in `.pq` files, VBA in `.bas` files. Nothing lives only inside a binary workbook.
2. **Exports are hostile input.** Parsers find the header row instead of assuming row counts, force account codes to text, and fail loudly when the format changed.
3. **No client data, ever.** Fixtures are fabricated and follow the `samples/sample-*.csv` naming convention. The `.gitignore` allowlists only that pattern, so a real export dropped into `samples/` stays blocked. Real exports stay outside any repo.

## Roadmap

Aged receivables/payables parser, MYOB export parsers, Office Scripts ports of the VBA modules, auto-generated function catalog.

## Related

[`australian-accounting-skills`](https://github.com/ryanduguid/australian-accounting-skills): Claude Code skills for AU public-practice workflows; its `xero-exports` skill pairs with these parsers.

[`xero-trial-balance-export`](https://github.com/ryanduguid/xero-trial-balance-export): the API path for the same trial balance; use it when OAuth app access is on the table and skip the CSV cleanup entirely.

## Disclaimer

Utility code, MIT-licensed, no warranty. Nothing here is tax or financial advice; outputs need professional review like any workpaper input. Xero and MYOB are trademarks of their respective owners. This project is independent and not affiliated with or endorsed by either.

## Author

Ryan Duguid, accountant in Newcastle NSW, CA ANZ Provisional Member.
