Option Explicit

Sub PrintColumnsWithFormulaAddresses(Optional ByVal ws As Worksheet, Optional ByVal maxRows As Long = 5, Optional ByVal reportSheetName As String = "Formula Report")
    ' ws       -> sheet to scan (defaults to ActiveSheet)
    ' maxRows  -> how many rows (from row 1 downward) to include in each column
    
    If ws Is Nothing Then Set ws = ActiveSheet
    If maxRows < 1 Then Exit Sub
    
    ' Create or clear report sheet
    Dim rpt As Worksheet
    On Error Resume Next
    Set rpt = ThisWorkbook.Worksheets(reportSheetName)
    On Error GoTo 0
    If rpt Is Nothing Then
        Set rpt = ThisWorkbook.Worksheets.Add
        rpt.Name = reportSheetName
    Else
        rpt.Cells.Clear
    End If
    
    ' Headers
    rpt.Range("A1:D1").Value = Array("Column", "Cell", "Formula", "Sheet")
    Dim outRow As Long: outRow = 2
    
    Dim lastCol As Long, colIdx As Long
    Dim colRange As Range, f As Range, c As Range
    Dim ColLetter As String
    
    ' Determine last column based on UsedRange
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    For colIdx = ws.UsedRange.Column To lastCol
        ' Restrict rows to maxRows or actual UsedRange rows, whichever is smaller
        Dim lastRowToCheck As Long
        lastRowToCheck = Application.Min(maxRows, ws.Rows.Count)
        Set colRange = ws.Range(ws.Cells(1, colIdx), ws.Cells(lastRowToCheck, colIdx))
        
        On Error Resume Next
        Set f = colRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number = 0 And Not f Is Nothing Then
            ColLetter = Split(ws.Cells(1, colIdx).Address(False, False), "1")(0)
            
            For Each c In f.Cells
                rpt.Cells(outRow, 1).Value = ColLetter & " (" & colIdx & ")"
                rpt.Cells(outRow, 2).Value = c.Address(False, False)
                rpt.Cells(outRow, 3).Value = Mid(c.Formula, 2) ' Remove leading '='
                rpt.Cells(outRow, 4).Value = ws.Name
                outRow = outRow + 1
            Next
        End If
        Err.Clear
        On Error GoTo 0
    Next colIdx
    
    rpt.Columns("A:D").AutoFit
End Sub

Sub PrintFormulaAddresses()
    Dim ws As Worksheet

    Set ws = ActiveSheet

    TurnOffCalcs

    PrintColumnsWithFormulaAddresses ws, 5, "Formula Report"

    TurnOnCalcs
End Sub
