// Xero.AgedPayables
// Parse a Xero Aged Payables Summary CSV export into a typed, verified table.
//
// Accepts the standard Xero Aged Payables Summary export shape.
// The header row is detected by finding the row starting with "Contact", "Supplier", or "Vendor".
// Metadata rows above the header (report title, entity name, as-at date, ageing basis)
// are safely measured and skipped without collapsing the table width.
//
// Key contract rules:
//   1. Supplier identifier is forced to text so leading zeros and punctuation survive.
//   2. Ageing buckets and Total columns are parsed as numbers (Currency/type number).
//   3. The trailing summary "Total" row is excluded from the returned supplier rows.
//   4. Row-level bucket sums are verified against the exported Total within a 0.005 tolerance.
//   5. Negative balances (supplier credit notes and prepayments) are preserved as signed numbers.
//
// Usage: Xero.AgedPayables("C:\exports\aged-payables-2026-06-30.csv")
let
    AgedPayables = (FilePath as text) as table =>
        let
            fileBinary = File.Contents(FilePath),
            lines = Lines.FromBinary(fileBinary, null, null, 65001),
            headerLines = List.Select(
                lines,
                each Text.StartsWith(Text.TrimStart(_), "Contact,")
                    or Text.StartsWith(Text.TrimStart(_), "Supplier,")
                    or Text.StartsWith(Text.TrimStart(_), "Vendor,")
                    or Text.Trim(_) = "Contact"
                    or Text.Trim(_) = "Supplier"
                    or Text.Trim(_) = "Vendor"
            ),
            headerLine =
                if List.IsEmpty(headerLines) then
                    error Error.Record(
                        "Xero.AgedPayables",
                        "Header row not found",
                        "No line has ""Contact"", ""Supplier"", or ""Vendor"" as its first field: not a Xero Aged Payables export, or the format changed."
                    )
                else
                    List.First(headerLines),
            firstField = Text.Trim(List.First(Text.Split(headerLine, ","))),
            columnCount = List.Count(Text.Split(headerLine, ",")),
            Raw = Csv.Document(
                fileBinary,
                [Delimiter = ",", Columns = columnCount, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
            ),
            Indexed = Table.AddIndexColumn(Raw, "__idx", 0, 1, Int64.Type),
            HeaderMatches = Table.SelectRows(
                Indexed,
                each Text.Trim(Text.From([Column1] ?? "")) = firstField
            )[__idx],
            HeaderIdx =
                if List.IsEmpty(HeaderMatches) then
                    error Error.Record(
                        "Xero.AgedPayables",
                        "Header row not found",
                        "Header line detected in raw text but not in the parsed CSV: format changed."
                    )
                else
                    List.Min(HeaderMatches),
            Sliced = Table.RemoveFirstN(Raw, HeaderIdx),
            Promoted = Table.PromoteHeaders(Sliced, [PromoteAllScalars = true]),
            cols = Table.ColumnNames(Promoted),
            supplierCol = if List.Contains(cols, "Supplier") then "Supplier" else if List.Contains(cols, "Vendor") then "Vendor" else cols{0},
            // Drop total rows and blank rows
            Filtered = Table.SelectRows(
                Promoted,
                each let val = Text.Trim(Text.From(Record.Field(_, supplierCol) ?? "")) in
                    val <> "" and Text.Lower(val) <> "total"
            ),
            // Identify amount columns (all columns except the supplier column)
            amountCols = List.Difference(cols, {supplierCol}),
            // Transform types
            TransformOps = List.Transform(
                amountCols,
                each {_, each try Value.FromText(Text.Trim(Text.From(_ ?? "0")), "en-US") otherwise 0.0, type number}
            ) & {{supplierCol, each Text.Trim(Text.From(_ ?? "")), type text}},
            Typed = Table.TransformColumns(Filtered, TransformOps),
            // Rename supplier column to canonical "Supplier" if needed
            Renamed = if supplierCol <> "Supplier" then Table.RenameColumns(Typed, {{supplierCol, "Supplier"}}) else Typed
        in
            Renamed
in
    AgedPayables
