#requires -Version 5.1
<#
.SYNOPSIS
Runs the repository's Power Query acceptance checks in desktop Excel.

.DESCRIPTION
Loads every checked-in .pq file into a disposable workbook and evaluates 46
checks in Excel's real Power Query engine. The checks cover both fabricated
Xero trial-balance layouts, financial-year boundaries, ABN validation, header
promotion, and adverse and lazy-evaluation branches.

This runner requires Windows, Windows PowerShell 5.1 or newer, desktop
Microsoft Excel, Power Query, and the Microsoft.Mashup.OleDb.1 provider. It
does not import or execute VBA. Repository sources and sample files are read
only. Generated fixtures and the workbook live in a GUID-named directory under
the operating system's temporary directory and are removed in finally.

.PARAMETER RepositoryRoot
Path to the accounting-excel-toolkit checkout to test. By default this is the
repository containing this script.

.EXAMPLE
powershell -NoProfile -File .\tools\native_excel_acceptance.ps1

.EXAMPLE
powershell -NoProfile -File .\tools\native_excel_acceptance.ps1 -RepositoryRoot C:\src\accounting-excel-toolkit
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$RepositoryRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw 'Could not determine the native acceptance script path.'
    }
    $scriptDirectory = Split-Path -Parent $scriptPath
    $RepositoryRoot = Split-Path -Parent $scriptDirectory
}

$repository = (Resolve-Path -LiteralPath $RepositoryRoot -ErrorAction Stop).ProviderPath

$script:cleanupFailed = $false

function ConvertTo-MText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    # Quotes are doubled in M string literals; backslashes are literal.
    return '"' + ($Value -replace '"', '""') + '"'
}

function Test-SafeTemporaryDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Candidate,
        [Parameter(Mandatory = $true)]
        [string]$SystemTemporaryRoot
    )

    $candidateFullPath = [IO.Path]::GetFullPath($Candidate)
    $expectedParent = [IO.Path]::GetFullPath($SystemTemporaryRoot).TrimEnd(
        [IO.Path]::DirectorySeparatorChar
    )
    $actualParent = [IO.Path]::GetDirectoryName($candidateFullPath).TrimEnd(
        [IO.Path]::DirectorySeparatorChar
    )
    $leaf = [IO.Path]::GetFileName($candidateFullPath)

    return [string]::Equals(
        $actualParent,
        $expectedParent,
        [StringComparison]::OrdinalIgnoreCase
    ) -and $leaf -match (
        '^accounting-excel-toolkit-native-' +
        '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
    )
}

function Release-ComReference {
    param(
        [AllowNull()]
        [object]$Reference,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($null -eq $Reference) {
        return
    }

    if (-not [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        [Console]::Error.WriteLine("CLEANUP ERROR: $Label is not a COM object.")
        $script:cleanupFailed = $true
        return
    }

    try {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Reference)
    }
    catch {
        [Console]::Error.WriteLine(
            "CLEANUP ERROR: could not release $Label. $($_.Exception.Message)"
        )
        $script:cleanupFailed = $true
    }
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'Native Excel acceptance requires Windows and desktop Microsoft Excel.'
}

$powerQueryDirectory = Join-Path $repository 'powerquery'
$combinedFixture = Join-Path $repository 'samples\sample-xero-trial-balance.csv'
$columnsFixture = Join-Path $repository 'samples\sample-xero-trial-balance-columns.csv'

foreach ($requiredPath in @($powerQueryDirectory, $combinedFixture, $columnsFixture)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required repository path is missing: $requiredPath"
    }
}

$powerQueryFiles = @(
    Get-ChildItem -LiteralPath $powerQueryDirectory -Filter '*.pq' -File |
        Sort-Object Name
)
if ($powerQueryFiles.Count -eq 0) {
    throw "No Power Query source files were found under $powerQueryDirectory."
}

$systemTemporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$temporaryDirectory = Join-Path $systemTemporaryRoot (
    'accounting-excel-toolkit-native-' + [guid]::NewGuid().ToString('D')
)
$temporaryDirectory = [IO.Path]::GetFullPath($temporaryDirectory)
if (-not (Test-SafeTemporaryDirectory $temporaryDirectory $systemTemporaryRoot)) {
    throw "Refusing to use unexpected temporary path: $temporaryDirectory"
}

$temporaryWorkbook = Join-Path $temporaryDirectory 'native-excel-acceptance.xlsx'
$periodOnlyFixture = Join-Path $temporaryDirectory 'period-only.csv'
$notTrialBalanceFixture = Join-Path $temporaryDirectory 'not-a-tb.csv'
$decoyEntityFixture = Join-Path $temporaryDirectory 'decoy-entity.csv'
$badAmountFixture = Join-Path $temporaryDirectory 'bad-amount.csv'

$temporaryDirectoryCreated = $false
$exitCode = 1
$excelVersion = $null
$excelBuild = $null
$previousDisplayAlerts = $null
$previousAutomationSecurity = $null

# COM references are declared once and released in exact reverse creation order.
$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$queries = $null
$createdQueries = New-Object System.Collections.ArrayList
$currentQuery = $null
$targetRange = $null
$listObjects = $null
$listObject = $null
$queryTable = $null
$dataBodyRange = $null

try {
    New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
    $temporaryDirectoryCreated = $true

    # A period-movement-only export: no YTD pair at all.
    @'
Trial Balance
Sample Trading Pty Ltd
For the month ended 30 June 2026

Account,Account Type,Debit,Credit
Business Bank Account (090),Bank,415.00,
Accounts Receivable (610),Current Asset,1830.00,
Wages and Salaries (477),Expense,3750.00,
Rent (469),Expense,1000.00,
Accounts Payable (800),Current Liability,,415.00
GST (820),Current Liability,,380.00
Sales (200),Revenue,,6200.00
Total,,6995.00,6995.00
'@ | Set-Content -LiteralPath $periodOnlyFixture -Encoding UTF8

    # This is deliberately not a trial balance.
    @'
Contact,Invoice Number,Due Date,Amount
Acme Pty Ltd,INV-001,30/06/2026,1100.00
'@ | Set-Content -LiteralPath $notTrialBalanceFixture -Encoding UTF8

    # An entity name beginning with "Account" must not be mistaken for a header.
    $decoyEntityContents = (
        Get-Content -LiteralPath $combinedFixture -Raw -Encoding UTF8
    ) -replace 'Sample Trading Pty Ltd', 'Accountable Plumbing Pty Ltd'
    $decoyEntityContents |
        Set-Content -LiteralPath $decoyEntityFixture -Encoding UTF8

    # A non-numeric Debit value must raise once M's lazy column is forced.
    $badAmountContents = (
        Get-Content -LiteralPath $combinedFixture -Raw -Encoding UTF8
    ) -replace '15234\.50', 'TBC'
    $badAmountContents |
        Set-Content -LiteralPath $badAmountFixture -Encoding UTF8

    $mCombined = ConvertTo-MText $combinedFixture
    $mColumns = ConvertTo-MText $columnsFixture
    $mPeriodOnly = ConvertTo-MText $periodOnlyFixture
    $mNotTrialBalance = ConvertTo-MText $notTrialBalanceFixture
    $mDecoyEntity = ConvertTo-MText $decoyEntityFixture
    $mBadAmount = ConvertTo-MText $badAmountFixture

    $checksM = @"
let
    // --- helpers -------------------------------------------------------
    s = (v) => if v = null then "(null)" else Text.From(v),
    near = (a, b) => a <> null and b <> null and Number.Abs(a - b) < 0.005,
    chk = (name, expected, actual) =>
        [Check = name, Expected = s(expected), Actual = s(actual), Pass = (s(expected) = s(actual))],
    raises = (f) => (try f())[HasError],

    // --- Xero_TrialBalance: both layouts -------------------------------
    tbC    = Xero_TrialBalance($mCombined),
    tbX    = Xero_TrialBalance($mColumns),
    tbCper = Xero_TrialBalance($mCombined, false),
    tbCytd = Xero_TrialBalance($mCombined, true),

    rentC = Table.SelectRows(tbC, each Text.Contains([AccountName], "Sydney")),
    rentX = Table.SelectRows(tbX, each Text.Contains([AccountName], "Sydney")),
    bankC = Table.SelectRows(tbC, each [AccountName] = "Business Bank Account"),

    keys = {"AccountCode", "AccountName", "Debit", "Credit"},
    projC = Table.SelectColumns(tbC, keys),
    projX = Table.SelectColumns(tbX, keys),
    // Table.Difference does not exist in M. Compare sorted, pipe-joined row
    // text so nulls and numeric typing cannot mask a mismatch.
    rowText = (t) => List.Sort(List.Transform(Table.ToRecords(t),
        each Text.Combine(List.Transform(Record.FieldValues(_),
            each if _ = null then "~" else Text.From(_)), "|"))),

    checks = {
        // -- combined layout, as-at (YTD) pair is the default
        chk("combined: 12 data rows (Total + blanks dropped)", 12, Table.RowCount(tbC)),
        chk("combined: YTD debit total 129934.50", true, near(List.Sum(tbC[Debit]), 129934.50)),
        chk("combined: YTD credit total 129934.50", true, near(List.Sum(tbC[Credit]), 129934.50)),
        chk("combined: balances within 0.005", true,
            Number.Abs(List.Sum(tbC[Debit]) - List.Sum(tbC[Credit])) < 0.005),
        chk("combined: default pair == explicit useYTD=true", true,
            near(List.Sum(tbCytd[Debit]), List.Sum(tbC[Debit]))),

        // -- the period pair is different and also balances
        chk("combined: useYTD=false debit total 6995.00", true, near(List.Sum(tbCper[Debit]), 6995.00)),
        chk("combined: useYTD=false credit total 6995.00", true, near(List.Sum(tbCper[Credit]), 6995.00)),
        chk("combined: period pair differs from as-at pair", true,
            not near(List.Sum(tbCper[Debit]), List.Sum(tbC[Debit]))),

        // -- "Rent (Sydney)" has a parenthetical name, not an account code
        chk("combined: Rent (Sydney) code is null", "(null)", Table.FirstValue(Table.SelectColumns(rentC, {"AccountCode"}))),
        chk("combined: Rent (Sydney) name kept intact", "Rent (Sydney)", Table.FirstValue(Table.SelectColumns(rentC, {"AccountName"}))),
        chk("columns: Rent (Sydney) code is null", "(null)", Table.FirstValue(Table.SelectColumns(rentX, {"AccountCode"}))),
        chk("columns: Rent (Sydney) name kept intact", "Rent (Sydney)", Table.FirstValue(Table.SelectColumns(rentX, {"AccountName"}))),

        // -- leading zero survives as text, rather than being coerced to 90
        chk("combined: code 090 keeps its leading zero", "090", Table.FirstValue(Table.SelectColumns(bankC, {"AccountCode"}))),
        chk("combined: AccountCode is text", true, Value.Is(Table.FirstValue(Table.SelectColumns(bankC, {"AccountCode"})), type text)),

        // -- both layouts land on the same shape
        chk("columns: 12 data rows", 12, Table.RowCount(tbX)),
        chk("parity: combined and columns layouts agree on all 12 rows", true,
            rowText(projC) = rowText(projX)),

        // --- Fx_AUFinancialYear ----------------------------------------
        chk("FY: 30 Jun 2026 -> FY2026", "FY2026", Fx_AUFinancialYear(#date(2026, 6, 30))[Label]),
        chk("FY: 1 Jul 2026 -> FY2027 (boundary)", "FY2027", Fx_AUFinancialYear(#date(2026, 7, 1))[Label]),
        chk("FY: UTC 2026-06-30T14:30Z is 1 Jul in AEST -> FY2027", "FY2027",
            Fx_AUFinancialYear(#datetimezone(2026, 6, 30, 14, 30, 0, 0, 0))[Label]),
        chk("FY: 9am +10:00 on 1 Jul 2026 -> FY2027", "FY2027",
            Fx_AUFinancialYear(#datetimezone(2026, 7, 1, 9, 0, 0, 10, 0))[Label]),
        chk("FY: text 01/07/2026 parsed en-AU, not machine locale", "FY2027",
            Fx_AUFinancialYear("01/07/2026")[Label]),
        chk("FY: null date -> null label (blank cell, not an error)", "(null)",
            Fx_AUFinancialYear(null)[Label]),
        chk("FY: StartDate/EndDate bracket the year", true,
            Fx_AUFinancialYear(#date(2026, 7, 1))[StartDate] = #date(2026, 7, 1)
                and Fx_AUFinancialYear(#date(2026, 7, 1))[EndDate] = #date(2027, 6, 30)),

        // --- Fx_ABNIsValid ---------------------------------------------
        chk("ABN: ATO example 51 824 753 556 valid", true, Fx_ABNIsValid("51 824 753 556")),
        chk("ABN: unspaced form valid", true, Fx_ABNIsValid("51824753556")),
        chk("ABN: number-typed input valid", true, Fx_ABNIsValid(51824753556)),
        chk("ABN: leading zero rejected even though checksum passes", false, Fx_ABNIsValid("00000090000")),
        chk("ABN: null rejected", false, Fx_ABNIsValid(null)),
        chk("ABN: punctuation rejected (not silently stripped)", false, Fx_ABNIsValid("51-824-753-556")),
        chk("ABN: wrong length rejected", false, Fx_ABNIsValid("5182475355")),
        chk("ABN: bad checksum rejected", false, Fx_ABNIsValid("51 824 753 557")),
        chk("ABN: unconvertible value (record) returns false, does not break refresh", false,
            Fx_ABNIsValid([unexpected = "shape"])),
        chk("ABN: unconvertible value (list) returns false", false, Fx_ABNIsValid({1, 2, 3})),

        // --- Fx_PromoteHeaderAt ----------------------------------------
        chk("PromoteHeaderAt: promotes the located header row", "Account",
            List.First(Table.ColumnNames(Fx_PromoteHeaderAt(
                #table({"Column1", "Column2"},
                    {{"Trial Balance", null}, {"Sample Trading Pty Ltd", null},
                     {"Account", "Debit"}, {"Sales (200)", "6200.00"}}),
                "Account")))),
        chk("PromoteHeaderAt: caller value is trimmed before matching", "Account",
            List.First(Table.ColumnNames(Fx_PromoteHeaderAt(
                #table({"Column1", "Column2"}, {{"Account", "Debit"}, {"Sales (200)", "1.00"}}),
                "  Account  ")))),
        chk("PromoteHeaderAt: zero-column input raises", true,
            raises(() => Fx_PromoteHeaderAt(#table({}, {}), "Account"))),
        chk("PromoteHeaderAt: blank header value raises", true,
            raises(() => Fx_PromoteHeaderAt(#table({"Column1"}, {{"Account"}}), "   "))),
        chk("PromoteHeaderAt: header not found raises", true,
            raises(() => Fx_PromoteHeaderAt(#table({"Column1"}, {{"Contact"}}), "Account"))),
        chk("PromoteHeaderAt: zero-ROW input still raises (lazy-guard trap)", true,
            raises(() => Fx_PromoteHeaderAt(Table.FirstN(#table({"Column1"}, {{"Account"}}), 0), "Account"))),

        // --- Xero_TrialBalance adverse branches ------------------------
        chk("TB: useYTD=true on a period-only export raises", true,
            raises(() => Xero_TrialBalance($mPeriodOnly, true))),
        chk("TB: useYTD=false still works on a period-only export", true,
            near(List.Sum(Xero_TrialBalance($mPeriodOnly, false)[Debit]), 6995.00)),
        chk("TB: default picks the period pair when there is no YTD pair", true,
            near(List.Sum(Xero_TrialBalance($mPeriodOnly)[Debit]), 6995.00)),
        chk("TB: a non-Xero CSV raises rather than loading garbage", true,
            raises(() => Xero_TrialBalance($mNotTrialBalance))),
        chk("TB: an entity name starting 'Account' does not hijack the header", 12,
            Table.RowCount(Xero_TrialBalance($mDecoyEntity))),
        // M is lazy: the sum must force the Debit column to expose bad input.
        chk("TB: unparseable amount raises once the column is forced", true,
            raises(() => List.Sum(Xero_TrialBalance($mBadAmount)[Debit]))),
        chk("TB: merely calling with a bad amount does NOT raise (lazy)", false,
            raises(() => Xero_TrialBalance($mBadAmount)))
    },
    Result = Table.FromRecords(
        checks,
        type table [Check = text, Expected = text, Actual = text, Pass = logical]
    )
in
    Result
"@

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw (
            'Desktop Microsoft Excel could not be started through COM. ' +
            'Install desktop Excel on Windows before running this acceptance test. ' +
            $_.Exception.Message
        )
    }

    $excelVersion = [string]$excel.Version
    $excelBuild = [string]$excel.Build
    $previousDisplayAlerts = $excel.DisplayAlerts
    $previousAutomationSecurity = $excel.AutomationSecurity
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # msoAutomationSecurityForceDisable: the runner never imports VBA and no
    # workbook opened by it may execute embedded macros.
    $excel.AutomationSecurity = 3

    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    # The $Workbook$ provider is more stable when the workbook has a path.
    $workbook.SaveAs($temporaryWorkbook, 51)
    $worksheets = $workbook.Worksheets
    $worksheet = $worksheets.Item(1)
    $queries = $workbook.Queries

    foreach ($file in $powerQueryFiles) {
        # Queries.Add rejects dots in names, so Fx.ABNIsValid becomes
        # Fx_ABNIsValid. Only the query name changes; the M source is unaltered.
        $queryName = [IO.Path]::GetFileNameWithoutExtension($file.Name) -replace '\.', '_'
        $querySource = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
        $currentQuery = $queries.Add($queryName, $querySource)
        [void]$createdQueries.Add($currentQuery)
        $currentQuery = $null
        Write-Host "loaded $($file.Name) as query: $queryName"
    }

    $currentQuery = $queries.Add('ZZ_Checks', $checksM)
    [void]$createdQueries.Add($currentQuery)
    $currentQuery = $null

    $connectionString = (
        'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;' +
        'Location=ZZ_Checks;Extended Properties=""'
    )
    $targetRange = $worksheet.Range('A1')
    $listObjects = $worksheet.ListObjects
    $listObject = $listObjects.Add(0, $connectionString, $null, 1, $targetRange)
    $queryTable = $listObject.QueryTable
    $queryTable.CommandType = 2
    $queryTable.CommandText = @('SELECT * FROM [ZZ_Checks]')
    $queryTable.BackgroundQuery = $false

    try {
        [void]$queryTable.Refresh($false)
        $excel.CalculateUntilAsyncQueriesDone()
    }
    catch {
        throw (
            'Power Query refresh failed. Desktop Excel must provide ' +
            'Microsoft.Mashup.OleDb.1. ' + $_.Exception.Message
        )
    }

    $dataBodyRange = $listObject.DataBodyRange
    if ($null -eq $dataBodyRange) {
        throw 'ZZ_Checks returned no rows; the M did not evaluate.'
    }

    $values = $dataBodyRange.Value2
    if ($values -isnot [array] -or $values.Rank -ne 2) {
        throw 'ZZ_Checks returned an unexpected result shape.'
    }

    $rowLower = $values.GetLowerBound(0)
    $rowUpper = $values.GetUpperBound(0)
    $columnLower = $values.GetLowerBound(1)
    $columnUpper = $values.GetUpperBound(1)
    $rowCount = $rowUpper - $rowLower + 1
    $columnCount = $columnUpper - $columnLower + 1

    if ($rowCount -ne 46) {
        throw "ZZ_Checks returned $rowCount rows; expected exactly 46."
    }
    if ($columnCount -ne 4) {
        throw "ZZ_Checks returned $columnCount columns; expected exactly 4."
    }

    Write-Host ''
    Write-Host (
        'Excel {0} build {1}; locale {2}' -f
            $excelVersion,
            $excelBuild,
            (Get-Culture).Name
    )
    Write-Host ('-' * 78)

    $failedChecks = 0
    for ($row = $rowLower; $row -le $rowUpper; $row++) {
        $check = [string]$values[$row, $columnLower]
        $expected = [string]$values[$row, ($columnLower + 1)]
        $actual = [string]$values[$row, ($columnLower + 2)]
        $passValue = $values[$row, ($columnLower + 3)]
        $passed = ($passValue -eq $true) -or ([string]$passValue -eq 'TRUE')

        if (-not $passed) {
            $failedChecks++
        }

        Write-Host ('{0}  {1}' -f $(if ($passed) { 'PASS' } else { 'FAIL' }), $check)
        if (-not $passed) {
            Write-Host ('        expected [{0}]  actual [{1}]' -f $expected, $actual)
        }
    }

    Write-Host ('-' * 78)
    Write-Host ('{0} checks, {1} failed' -f $rowCount, $failedChecks)
    if ($failedChecks -eq 0) {
        $exitCode = 0
    }
}
catch {
    [Console]::Error.WriteLine('HARNESS ERROR: ' + $_.Exception.Message)
    if ($null -ne $_.Exception.InnerException) {
        [Console]::Error.WriteLine('  inner: ' + $_.Exception.InnerException.Message)
    }
}
finally {
    # Close the workbook even when query creation or refresh fails.
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            [Console]::Error.WriteLine(
                'CLEANUP ERROR: could not close the workbook. ' + $_.Exception.Message
            )
            $script:cleanupFailed = $true
        }
    }

    if ($null -ne $excel) {
        if ($null -ne $previousAutomationSecurity) {
            try {
                $excel.AutomationSecurity = $previousAutomationSecurity
            }
            catch {
                [Console]::Error.WriteLine(
                    'CLEANUP ERROR: could not restore AutomationSecurity. ' +
                    $_.Exception.Message
                )
                $script:cleanupFailed = $true
            }
        }
        if ($null -ne $previousDisplayAlerts) {
            try {
                $excel.DisplayAlerts = $previousDisplayAlerts
            }
            catch {
                [Console]::Error.WriteLine(
                    'CLEANUP ERROR: could not restore DisplayAlerts. ' +
                    $_.Exception.Message
                )
                $script:cleanupFailed = $true
            }
        }
    }

    # Release every created COM reference in reverse creation order.
    Release-ComReference $dataBodyRange 'DataBodyRange'
    $dataBodyRange = $null
    Release-ComReference $queryTable 'QueryTable'
    $queryTable = $null
    Release-ComReference $listObject 'ListObject'
    $listObject = $null
    Release-ComReference $listObjects 'ListObjects collection'
    $listObjects = $null
    Release-ComReference $targetRange 'target Range'
    $targetRange = $null
    Release-ComReference $currentQuery 'partially registered WorkbookQuery'
    $currentQuery = $null

    for ($index = $createdQueries.Count - 1; $index -ge 0; $index--) {
        Release-ComReference $createdQueries[$index] "WorkbookQuery[$index]"
        $createdQueries[$index] = $null
    }

    Release-ComReference $queries 'Queries collection'
    $queries = $null
    Release-ComReference $worksheet 'Worksheet'
    $worksheet = $null
    Release-ComReference $worksheets 'Worksheets collection'
    $worksheets = $null
    Release-ComReference $workbook 'Workbook'
    $workbook = $null
    Release-ComReference $workbooks 'Workbooks collection'
    $workbooks = $null

    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            [Console]::Error.WriteLine(
                'CLEANUP ERROR: could not quit Excel. ' + $_.Exception.Message
            )
            $script:cleanupFailed = $true
        }
    }
    Release-ComReference $excel 'Excel Application'
    $excel = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($temporaryDirectoryCreated -and (Test-Path -LiteralPath $temporaryDirectory)) {
        if (Test-SafeTemporaryDirectory $temporaryDirectory $systemTemporaryRoot) {
            try {
                Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force
            }
            catch {
                [Console]::Error.WriteLine(
                    'CLEANUP ERROR: could not remove the temporary fixture directory. ' +
                    $_.Exception.Message
                )
                $script:cleanupFailed = $true
            }
        }
        else {
            [Console]::Error.WriteLine(
                "CLEANUP ERROR: refusing to remove unexpected path: $temporaryDirectory"
            )
            $script:cleanupFailed = $true
        }
    }
}

if ($script:cleanupFailed) {
    $exitCode = 1
}
exit $exitCode
