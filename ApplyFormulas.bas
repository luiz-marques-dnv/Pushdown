Option Explicit

'== Entry point with prompts ==
Sub BuildFormulaBlocks_Prompt()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim startColLetter As String
    Dim startCol As Long
    Dim numBlocks As Long
    Dim numRows As Long
    Dim firstAnchorRow As Long
    
    startColLetter = InputBox("Start column letter for the first block (e.g., A):", "Start Column", "A")
    If Len(startColLetter) = 0 Then Exit Sub
    startCol = LetterToCol(startColLetter)
    If startCol < 1 Then
        MsgBox "Invalid start column.", vbExclamation
        Exit Sub
    End If
    
    numBlocks = Val(InputBox("Number of blocks (A–E is 1 block, F–J is 2nd, etc.):", "Blocks", "2"))
    If numBlocks < 1 Then Exit Sub
    
    numRows = Val(InputBox("How many data rows to fill (starting at row 2):", "Rows", "5000"))
    If numRows < 1 Then Exit Sub
    
    firstAnchorRow = Val(InputBox("Script_Usfos anchor starts at which row? (A2 in your sample => 2)", "Anchor Row", "2"))
    If firstAnchorRow < 1 Then Exit Sub
    
    BuildFormulaBlocks ws, startCol, numBlocks, numRows, firstAnchorRow
End Sub

'== Main worker: builds formulas + formatting ==
Sub BuildFormulaBlocks_PFP_Scenarios(ByVal ws As Worksheet, _
                       ByVal startCol As Long, _
                       ByVal numBlocks As Long, _
                       ByVal numRows As Long, _
                       Optional ByVal firstAnchorRow As Long = 2)
    If ws Is Nothing Then Set ws = ActiveSheet
    If startCol < 1 Or numBlocks < 1 Or numRows < 1 Then Exit Sub
    
    Dim scrUp As Boolean: scrUp = Application.ScreenUpdating
    Dim calcState As XlCalculation: calcState = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanFail
    
    Dim b As Long, blockStart As Long
    Dim colAnchor As Long, colData As Long, colMapped As Long, colValue As Long
    Dim LtrAnchor As String, LtrData As String, LtrMapped As String, LtrValue As String
    Dim lastRow As Long: lastRow = WorksheetFunction.Min(1 + numRows, ws.Rows.Count) ' rows 2..lastRow are body
    
    ' --- Lay down formulas block-by-block ---
    For b = 0 To numBlocks - 1
        blockStart = startCol + b * 5
        colAnchor = blockStart
        colData = blockStart + 2
        colMapped = blockStart + 3
        colValue = blockStart + 4
        
        LtrAnchor = ColLetter(colAnchor)
        LtrData = ColLetter(colData)
        LtrMapped = ColLetter(colMapped)
        LtrValue = ColLetter(colValue)
        
        ' Row 1 (headers)
        ws.Cells(1, colAnchor).Formula = "=Script_Usfos!A" & (firstAnchorRow + b)
        ws.Cells(1, colData).Formula = "=SUMPRODUCT((" & LtrData & "2:" & LtrData & lastRow & "<>"""")/" & _
                                       "COUNTIF(" & LtrData & "2:" & LtrData & lastRow & "," & LtrData & "2:" & LtrData & lastRow & "&""""))+0.001"
        ws.Cells(1, colMapped).Formula = "=XLOOKUP(" & LtrAnchor & "1,Script_Usfos!$A:$A,Script_Usfos!$H:$H,"""")"
        ws.Cells(1, colValue).Formula = "=SUM(" & LtrValue & "2:" & LtrValue & lastRow & ")"
        
        ' Body formulas (rows 2..lastRow)
        If lastRow >= 2 Then
            ws.Cells(2, colMapped).Formula2 = "=LET(valor, XLOOKUP(" & LtrData & "2, FEMElements!$B:$B, FEMElements!$A:$A, """", 0, 2), IF(valor=0, """", valor))"
            ws.Range(ws.Cells(2, colMapped), ws.Cells(lastRow, colMapped)).FillDown
            
            ws.Cells(2, colValue).Formula = "=IFNA(IF(" & LtrMapped & "2=" & LtrMapped & "1,"""", VLOOKUP(" & LtrMapped & "2, BeamCoordinate!$A:$F, 6, FALSE)),"""")"
            ws.Range(ws.Cells(2, colValue), ws.Cells(lastRow, colValue)).FillDown
        End If
        
        ' ----- Formatting for this block -----
        ' 1) Header row bold (only the 5 columns of the block)
        With ws.Range(ws.Cells(1, blockStart), ws.Cells(1, colValue)).Font
            .Bold = True
        End With

        ' 2) Value column (3rd and 5th col of block) show 0 and two decimals respectively  (include header total)
        ws.Range(ws.Cells(1, colData), ws.Cells(lastRow, colData)).NumberFormat = "0"
        ws.Range(ws.Cells(1, colValue), ws.Cells(lastRow, colValue)).NumberFormat = "0.00"
        
        ' 3) Block vertical borders on body (rows 2..lastRow) so row 1 keeps "only below"
        If lastRow >= 1 Then
            Dim bodyRange As Range
            Set bodyRange = ws.Range(ws.Cells(1, blockStart), ws.Cells(lastRow, colValue))
            
            With bodyRange.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With bodyRange.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
    Next b
    
    ' 4) Header row: only a bottom line across all blocks (no top/left/right on header cells)
    Dim lastCol As Long
    lastCol = startCol + (numBlocks - 1) * 5 + 4
    With ws.Range(ws.Cells(1, startCol), ws.Cells(1, lastCol))
        ' Clear other edges so it's "only below"
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
    
CleanExit:
    Application.Calculation = calcState
    Application.ScreenUpdating = scrUp
    Exit Sub
    
CleanFail:
    Application.Calculation = calcState
    Application.ScreenUpdating = scrUp
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

'== Helpers ==
Private Function ColLetter(ByVal colIdx As Long) As String
    ColLetter = Split(Cells(1, colIdx).Address(False, False), "1")(0)
End Function

Private Function LetterToCol(ByVal s As String) As Long
    Dim i As Long, ch As Long
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then Exit For
        LetterToCol = LetterToCol * 26 + (ch - 64)
    Next i
End Function

Option Explicit

' Reads counts from Script_Usfos and FEMElements, then builds the blocks on the ActiveSheet.
Sub BuildFormulaBlocks_AllScenarios()
    Dim ws_PFP_Scenarios As Worksheet, ws_PFP_Scenarios_3_Sided As Worksheet

    Dim wsScen As Worksheet, wsElem As Worksheet
    On Error Resume Next
    Set wsScen = ThisWorkbook.Worksheets("Script_Usfos")
    Set wsElem = ThisWorkbook.Worksheets("FEMElements")
    On Error GoTo 0
    
    If wsScen Is Nothing Or wsElem Is Nothing Then
        MsgBox "Missing required sheets: 'Script_Usfos' and/or 'FEMElements'.", vbExclamation
        Exit Sub
    End If
    
    Dim scenarioCount As Long, elementCount As Long
    scenarioCount = CountRows(wsScen, 1, 2) ' Column A, start at row 2
    elementCount = CountRows(wsElem, 1, 2)  ' Column A, start at row 2

    If scenarioCount < 1 Then
        MsgBox "No scenarios found in Script_Usfos!A (from A2 down).", vbExclamation
        Exit Sub
    End If
    If elementCount < 1 Then
        MsgBox "No elements found in FEMElements!A (from A2 down).", vbExclamation
        Exit Sub
    End If
    
    TurnOffCalcs

    Set ws_PFP_Scenarios = ThisWorkbook.Worksheets("PFP Scenarios")
    BuildFormulaBlocks_PFP_Scenarios ws_PFP_Scenarios, 1, scenarioCount, elementCount, 2

    Set ws_PFP_Scenarios_3_Sided = ThisWorkbook.Worksheets("PFP Scenario 3-Sided")
    BuildFormulaBlocks_PFP_Scenarios ws_PFP_Scenarios_3_Sided, 1, scenarioCount, elementCount, 2
    TurnOnCalcs
End Sub

' Counts non-blank cells in a column from startRow to the last used row in that column.
Private Function CountRows(ByVal ws As Worksheet, ByVal colIdx As Long, ByVal startRow As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).row
    If lastRow < startRow Then
        CountRows = 0
    Else
        CountRows = lastRow - startRow + 1
    End If
End Function