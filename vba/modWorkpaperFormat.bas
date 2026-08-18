Attribute VB_Name = "modWorkpaperFormat"
Option Explicit

' Kept byte-identical to ACCOUNTING_NUMBER_FORMAT in modReconCompare.bas -
' each module imports stand-alone, so each carries its own copy of the
' accounting format; tests/test_static_guards.py pins the two together.
Private Const ACCOUNTING_NUMBER_FORMAT As String = "#,##0.00_);(#,##0.00);""-""??_)"

' modWorkpaperFormat
' Standard workpaper presentation: header block, prepared-by stamp, and
' number formatting that matches how a reviewer expects a workpaper to read.
' Import via VBE: File > Import File... (or drag into the Modules folder).

' Writes the standard four-line workpaper header at the top of a sheet:
'   entity name, workpaper title, period end, prepared-by / date stamp.
' Inserts rows so existing content is pushed down, not overwritten.
Public Sub ApplyWorkpaperHeader( _
    ByVal ws As Worksheet, _
    ByVal entityName As String, _
    ByVal wpTitle As String, _
    ByVal periodEnd As Date, _
    ByVal preparedBy As String)

    If ws.ProtectContents Then Err.Raise 5, , "Sheet '" & ws.Name & "' is protected - unprotect it first."

    ' A live cut/copy marquee turns Insert into a paste - the clipboard
    ' block would land in rows 1:5 instead of blank rows
    Application.CutCopyMode = False
    ws.Rows("1:5").Insert Shift:=xlDown

    With ws
        ' Text format BEFORE the write - an entity name or title starting
        ' with "=" would otherwise be stored as a live formula
        .Range("A1:A4").NumberFormat = "@"
        .Range("A1").Value = entityName
        .Range("A2").Value = wpTitle
        .Range("A3").Value = "For the period ended " & Format$(periodEnd, "d mmmm yyyy")
        .Range("A4").Value = "Prepared by: " & preparedBy & "    Date: " & Format$(Date, "d mmm yyyy")

        .Range("A1:A2").Font.Bold = True
        .Range("A1").Font.Size = 12
        .Range("A4").Font.Italic = True
        .Range("A5:H5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub

' Adds a reviewer sign-off line below the last used row.
Public Sub AddReviewerLine(ByVal ws As Worksheet)
    If ws.ProtectContents Then Err.Raise 5, , "Sheet '" & ws.Name & "' is protected - unprotect it first."

    ' Last used row across ALL columns - End(xlUp) on column A alone lands
    ' inside the data when the final rows only hold amounts in B onwards
    ' (totals blocks, formula-only footers). LookIn/LookAt are explicit
    ' because Find otherwise inherits the user's last Find-dialog settings.
    Dim c As Range
    Dim lastRow As Long
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then
        lastRow = 0
    Else
        lastRow = c.Row
    End If

    ws.Cells(lastRow + 2, 1).Value = "Reviewed by: ______________    Date: ____________"
    ws.Cells(lastRow + 2, 1).Font.Italic = True
End Sub

' Applies accounting number format (thousands separator, bracketed negatives,
' dash for zero) to a range - the format reviewers expect on workpapers.
Public Sub FormatAsAccounting(ByVal target As Range)
    target.NumberFormat = ACCOUNTING_NUMBER_FORMAT
End Sub

' Freezes panes below the header block (the 5 rows ApplyWorkpaperHeader
' writes). Pass headerRows = 6 if your data keeps its own column-header row
' directly under the block. Clears any existing freeze AND split first
' (both silently hijack the freeze position), and pins the scroll to the
' top so the freeze anchors where the selection says, not where the window
' happened to be scrolled.
Public Sub FreezeBelowHeader(ByVal ws As Worksheet, Optional ByVal headerRows As Long = 5)
    ' Validate before touching the window - a failure after the existing
    ' freeze is cleared would leave the sheet half-done, and headerRows = 0
    ' would not fail at all (FreezePanes with A1 active freezes at the
    ' centre of the visible window, an arbitrary split).
    If headerRows < 1 Then Err.Raise 5, , "headerRows must be at least 1"
    If ws.ProtectContents Then Err.Raise 5, , "Sheet '" & ws.Name & "' is protected - unprotect it first."
    ' Activate on a hidden sheet silently activates the nearest visible
    ' neighbour instead - the neighbour's panes would be wrecked and the
    ' anchor Select would fail
    If ws.Visible <> xlSheetVisible Then Err.Raise 5, , "Sheet is hidden - unhide it before freezing panes."

    ' Goto both activates the sheet and selects the anchor cell, so the
    ' window mutated below is guaranteed to be ws's own
    Application.Goto ws.Cells(headerRows + 1, 1)
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .FreezePanes = True
    End With
End Sub
