$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repository = Split-Path -Parent $PSScriptRoot
$excel = $null
$workbook = $null
$previousAutomationSecurity = $null
$previousDisplayAlerts = $null
$temporaryWorkbook = Join-Path (
    [IO.Path]::GetTempPath()
) ("accounting-excel-toolkit-native-" + [guid]::NewGuid().ToString("N") + ".xlsx")

function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) {
        throw "${Label}: expected '$Expected', got '$Actual'."
    }
}

function Add-PowerQueryTable {
    param(
        $Workbook,
        [string]$QueryName,
        [string]$CsvPath,
        [string]$FunctionSource
    )

    $escapedPath = $CsvPath.Replace('"', '""')
    $formula = "let Parser = ($FunctionSource), Result = Parser(`"$escapedPath`") in Result"
    [void]$Workbook.Queries.Add($QueryName, $formula)

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $QueryName
    $connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$QueryName;Extended Properties=`"`""
    $table = $sheet.ListObjects.Add(0, $connectionString, $null, 1, $sheet.Range("A1"))
    $table.QueryTable.CommandType = 2
    $table.QueryTable.CommandText = @("SELECT * FROM [$QueryName]")
    $table.QueryTable.BackgroundQuery = $false
    [void]$table.QueryTable.Refresh($false)
    $excel.CalculateUntilAsyncQueriesDone()
    return $table
}

function Assert-TrialBalanceTable {
    param($Table, [string]$Label)

    Assert-Equal $Table.ListRows.Count 12 "$Label row count"
    Assert-Equal $Table.HeaderRowRange.Cells.Item(1, 1).Value2 "AccountCode" "$Label first column"
    Assert-Equal $Table.HeaderRowRange.Cells.Item(1, 2).Value2 "AccountName" "$Label second column"

    $rentFound = $false
    $bankFound = $false
    for ($row = 1; $row -le $Table.ListRows.Count; $row++) {
        $code = $Table.DataBodyRange.Cells.Item($row, 1).Value2
        $name = $Table.DataBodyRange.Cells.Item($row, 2).Value2
        if ($name -eq "Rent (Sydney)") {
            $rentFound = $true
            Assert-Equal $code $null "$Label code-less account code"
        }
        if ($name -eq "Business Bank Account") {
            $bankFound = $true
            Assert-Equal ([string]$code) "090" "$Label leading-zero account code"
        }
    }
    Assert-Equal $rentFound $true "$Label code-less account"
    Assert-Equal $bankFound $true "$Label bank account"

    $debitColumn = $Table.ListColumns.Item("Debit").DataBodyRange
    $creditColumn = $Table.ListColumns.Item("Credit").DataBodyRange
    $debit = $excel.WorksheetFunction.Sum($debitColumn)
    $credit = $excel.WorksheetFunction.Sum($creditColumn)
    Assert-Equal $debit 129934.5 "$Label YTD debit total"
    Assert-Equal $credit 129934.5 "$Label YTD credit total"
}

function Run-PowerQueryAcceptance {
    $source = Get-Content -LiteralPath (Join-Path $repository "powerquery\Xero.TrialBalance.pq") -Raw
    $fixtures = @(
        @{ Name = "Combined"; Path = "samples\sample-xero-trial-balance.csv" },
        @{ Name = "Separate"; Path = "samples\sample-xero-trial-balance-columns.csv" }
    )
    foreach ($fixture in $fixtures) {
        $path = (Resolve-Path -LiteralPath (Join-Path $repository $fixture.Path)).Path
        $table = Add-PowerQueryTable $workbook $fixture.Name $path $source
        Assert-TrialBalanceTable $table $fixture.Name
    }
    Write-Output "PASS: native Excel Power Query accepted both fabricated Xero layouts."
}

try {
    $excel = New-Object -ComObject Excel.Application
    $previousDisplayAlerts = $excel.DisplayAlerts
    $previousAutomationSecurity = $excel.AutomationSecurity
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # msoAutomationSecurityForceDisable: no workbook opened by this Power Query
    # harness may execute embedded macros. This runner never imports VBA.
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Add()
    # Power Query's $Workbook$ provider is materially more stable when the
    # workbook has a path. The file contains fabricated data only, lives under
    # the OS temporary directory and is deleted in finally.
    $workbook.SaveAs($temporaryWorkbook, 51)

    Run-PowerQueryAcceptance
    Write-Output "Excel $($excel.Version) build $($excel.Build); native acceptance complete."
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch { }
    }
    if ($null -ne $excel) {
        if ($null -ne $previousAutomationSecurity) {
            try { $excel.AutomationSecurity = $previousAutomationSecurity } catch { }
        }
        if ($null -ne $previousDisplayAlerts) {
            try { $excel.DisplayAlerts = $previousDisplayAlerts } catch { }
        }
        try { $excel.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $temporaryWorkbook) {
        Remove-Item -LiteralPath $temporaryWorkbook -Force
    }
}
