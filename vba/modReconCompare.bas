Attribute VB_Name = "modReconCompare"
Option Explicit

Private Const RESULT_SHEET_NAME As String = "Recon Result"
Private Const RESULT_TAG_NAME As String = "__ReconCompareResultSheet"

' modReconCompare
' Keyed two-way reconciliation between two ranges - the "why doesn't the
' subledger agree to the GL" workhorse. Late-bound Scripting.Dictionary,
' so no references need adding. Windows Excel only - Mac Excel has no
' Scripting.Dictionary and no reference can supply it.
' Import via VBE: File > Import File...

' Compares two two-column ranges (key, amount). Writes a "Recon Result"
' sheet listing: keys only in A or only in B whose summed amount exceeds
' tolerance, and keys in both where the amounts differ by more than
' tolerance. Duplicate keys within a side are summed before comparing
' (subledger detail vs GL balance pattern).
'
' Keys compare as trimmed text, case-insensitive. A key stored as TEXT
' "001234" on one side and as the NUMBER 1234 on the other normalises to
' "001234" vs "1234" - two different keys, reported as two one-sided
' exceptions. Format both key columns the same way before running.
'
' Invisible characters ride in with pasted data: non-breaking spaces, tabs
' and line breaks normalise to plain spaces before trimming, zero-width
' spaces drop out. Trim$ alone leaves them, and a trailing non-breaking
' space reports the same key unmatched on both sides with nothing visible
' to explain it.
'
' Example:
'   CompareKeyedRanges Sheet1.Range("A2:B500"), Sheet2.Range("A2:B300"), 0.01
Public Sub CompareKeyedRanges( _
    ByVal rangeA As Range, _
    ByVal rangeB As Range, _
    Optional ByVal tolerance As Double = 0.005)

#If Mac Then
    ' Scripting.Dictionary lives in the Windows-only scripting runtime - name
    ' the platform instead of failing with a bare 429 on the first CreateObject.
    Err.Raise 5, , "modReconCompare needs Scripting.Dictionary - Windows Excel only."
#End If

    Dim skippedRows As Long
    skippedRows = 0

    Dim dictA As Object, dictB As Object
    Set dictA = SumByKey(rangeA, skippedRows)
    Set dictB = SumByKey(rangeB, skippedRows)

    ' Refuse to run when a source range lives on the result sheet - the
    ' delete below would destroy caller data.
    If StrComp(rangeA.Worksheet.Name, RESULT_SHEET_NAME, vbTextCompare) = 0 _
        Or StrComp(rangeB.Worksheet.Name, RESULT_SHEET_NAME, vbTextCompare) = 0 Then
        Err.Raise 5, , "Source range is on the 'Recon Result' sheet - move the data or rename the sheet."
    End If

    Dim wb As Workbook
    Set wb = rangeA.Worksheet.Parent
    If wb.ProtectStructure Then
        Err.Raise 5, , "Workbook structure is protected - unprotect it before running the recon."
    End If

    ' Only a sheet marked by this module is replaceable. An unrelated sheet
    ' called "Recon Result" can contain real user work and must never be
    ' deleted just because it shares the result name.
    DeletePreviousGeneratedResult wb

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = RESULT_SHEET_NAME
    MarkGeneratedResultSheet wb, ws

    ws.Range("A1:D1").Value = Array("Key", "Side A", "Side B", "Difference")
    ws.Range("A1:D1").Font.Bold = True

    Dim outRow As Long
    outRow = 2

    Dim k As Variant, amtA As Double, amtB As Double, diff As Double
    ' Keys in A (matched and A-only)
    For Each k In dictA.Keys
        amtA = dictA(k)
        amtB = 0#
        If dictB.Exists(k) Then amtB = dictB(k)
        diff = amtA - amtB
        If Abs(diff) > tolerance Then
            ' Text format BEFORE the write - .Value into a General cell
            ' re-parses "001234" to 1234 and "3-10" to a date
            ws.Cells(outRow, 1).NumberFormat = "@"
            ws.Cells(outRow, 1).Value = k
            ws.Cells(outRow, 2).Value = amtA
            If dictB.Exists(k) Then ws.Cells(outRow, 3).Value = amtB
            ws.Cells(outRow, 4).Value = diff
            outRow = outRow + 1
        End If
    Next k
    ' Keys only in B
    For Each k In dictB.Keys
        If Not dictA.Exists(k) Then
            amtB = dictB(k)
            If Abs(amtB) > tolerance Then
                ws.Cells(outRow, 1).NumberFormat = "@"
                ws.Cells(outRow, 1).Value = k
                ws.Cells(outRow, 3).Value = amtB
                ws.Cells(outRow, 4).Value = -amtB
                outRow = outRow + 1
            End If
        End If
    Next k

    ' A clean recon leaves outRow = 2 - the reversed corner pair would then
    ' normalise to the B1:D2 bounding box and format the header row.
    If outRow > 2 Then
        ws.Range(ws.Cells(2, 2), ws.Cells(outRow - 1, 4)).NumberFormat = "#,##0.00_);(#,##0.00);""-""??_)"
    End If
    ws.Columns("A:D").AutoFit

    ws.Cells(outRow + 1, 1).Value = "Items: " & (outRow - 2) & _
        "   Skipped rows (errors/blanks): " & skippedRows & _
        "   Tolerance: " & tolerance & _
        "   Run: " & Format$(Now, "d mmm yyyy hh:mm")
End Sub

Private Sub DeletePreviousGeneratedResult(ByVal wb As Workbook)
    Dim stale As Object
    Dim marker As Name
    Dim previousAlerts As Boolean

    On Error Resume Next
    Set stale = wb.Sheets(RESULT_SHEET_NAME)
    On Error GoTo 0
    If stale Is Nothing Then
        ' A left-over marker would make the next sheet unmarkable. Stop
        ' before creating anything rather than leaving an unowned output
        ' sheet behind; a user-owned marker is also never removed silently.
        On Error Resume Next
        Set marker = wb.Names(RESULT_TAG_NAME)
        On Error GoTo 0
        If Not marker Is Nothing Then
            Err.Raise 5, , "The reserved recon result marker exists but there is no 'Recon Result' sheet. It was left untouched; remove or rename the marker before running the recon."
        End If
        Exit Sub
    End If

    If Not IsGeneratedResultSheet(wb, stale) Then
        Err.Raise 5, , "A sheet named 'Recon Result' already exists but was not generated by modReconCompare. It was left untouched; rename or remove it before running the recon."
    End If

    previousAlerts = Application.DisplayAlerts
    On Error GoTo DeleteFailed
    Application.DisplayAlerts = False
    stale.Delete
    Application.DisplayAlerts = previousAlerts

    ' Deleting a worksheet normally leaves the workbook-level marker as a
    ' #REF! name, but Excel may remove it itself in some workbook formats.
    ' Either outcome is safe; only a remaining marker needs explicit cleanup.
    On Error Resume Next
    Set marker = wb.Names(RESULT_TAG_NAME)
    On Error GoTo DeleteFailed
    If Not marker Is Nothing Then marker.Delete
    Exit Sub

DeleteFailed:
    Application.DisplayAlerts = previousAlerts
    Err.Raise Err.Number, , "Cannot replace the generated 'Recon Result' sheet: " & Err.Description
End Sub

Private Function IsGeneratedResultSheet(ByVal wb As Workbook, ByVal candidate As Object) As Boolean
    Dim marker As Name
    Dim taggedRange As Range

    If TypeName(candidate) <> "Worksheet" Then Exit Function

    On Error Resume Next
    Set marker = wb.Names(RESULT_TAG_NAME)
    If Not marker Is Nothing Then Set taggedRange = marker.RefersToRange
    On Error GoTo 0

    If taggedRange Is Nothing Then Exit Function
    IsGeneratedResultSheet = (taggedRange.Worksheet Is candidate)
End Function

Private Sub MarkGeneratedResultSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim existing As Name

    On Error Resume Next
    Set existing = wb.Names(RESULT_TAG_NAME)
    On Error GoTo 0
    If Not existing Is Nothing Then
        Err.Raise 5, , "Cannot mark the generated 'Recon Result' sheet because the reserved result marker already exists. Remove the stale marker before running the recon."
    End If

    wb.Names.Add Name:=RESULT_TAG_NAME, _
        RefersTo:="='" & Replace$(ws.Name, "'", "''") & "'!$A$1", _
        Visible:=False
End Sub

' Sums a two-column (key, amount) range into a dictionary, keyed on the
' trimmed text of column 1. Rows with error values (#N/A, #REF!...), blank
' keys, or blank/non-numeric amounts are skipped and counted, not crashed on.
Private Function SumByKey(ByVal source As Range, ByRef skippedRows As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' Shape checks - Cells(r, 2) on a one-column range would read the
    ' worksheet column beside it, and a Ctrl-selected union silently
    ' truncates to its first area.
    If source.Areas.Count > 1 Then
        Err.Raise 5, , "Pass a single contiguous range - a Ctrl-selected union would be truncated to its first area."
    End If
    If source.Columns.Count < 2 Then
        Err.Raise 5, , "Range needs at least two columns (key, amount)."
    End If

    ' Bound the loop to the used range - a whole-column selection (A:B) is
    ' 1,048,576 rows and two COM reads per row, which freezes Excel for
    ' minutes. Rows only; column geometry stays exactly as passed.
    Dim used As Range, lastR As Long
    Set used = Intersect(source, source.Worksheet.UsedRange)
    If used Is Nothing Then
        Set SumByKey = dict
        Exit Function
    End If
    lastR = used.Row + used.Rows.Count - source.Row

    Dim r As Long, k As String, keyVal As Variant, v As Variant
    For r = 1 To lastR
        keyVal = source.Cells(r, 1).Value
        v = source.Cells(r, 2).Value
        If IsError(keyVal) Or IsError(v) Then
            skippedRows = skippedRows + 1
        Else
            ' Trim$ only sees plain spaces - a non-breaking space, tab or
            ' line break pasted in with a key leaves it looking identical to
            ' a clean one and matching nothing.
            k = CStr(keyVal)
            k = Replace$(k, ChrW$(160), " ")
            k = Replace$(k, vbTab, " ")
            k = Replace$(k, vbCrLf, " ")
            k = Replace$(k, vbCr, " ")
            k = Replace$(k, vbLf, " ")
            k = Replace$(k, ChrW$(8203), "")
            k = Trim$(k)
            ' Not IsEmpty guards the VBA trap IsNumeric(Empty) = True - a
            ' blank amount must count as skipped, not sum as a silent zero.
            ' VarType guards the sibling trap IsNumeric(True) = True with
            ' CDbl(True) = -1 - a stray TRUE must skip, not sum as -1.00
            If Len(k) > 0 And Not IsEmpty(v) And VarType(v) <> vbBoolean And IsNumeric(v) Then
                If dict.Exists(k) Then
                    dict(k) = dict(k) + CDbl(v)
                Else
                    dict.Add k, CDbl(v)
                End If
            ElseIf Len(k) > 0 Or Not IsEmpty(v) Then
                ' counts blank/bad-amount rows AND blank-key rows with data;
                ' fully empty rows (oversized selections) stay uncounted
                skippedRows = skippedRows + 1
            End If
        End If
    Next r

    Set SumByKey = dict
End Function
