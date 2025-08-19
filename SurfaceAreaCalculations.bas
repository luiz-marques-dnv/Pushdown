Sub PFP_Rate()

    PFP_Rate_Lot.Range("A2", "E50").Clear 'new
    
    Count_all = 2
    Count_H60 = 2
    Count_H30 = 2
    Count_H15 = 2
    Count_J15 = 2
    Count_J30 = 2
    Count_J60 = 2
    
    ultlinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row
    CURRENT_FOLDER = ActiveWorkbook.Path
    
    For i = 1 To (ultlinha - 1) * 5 - 4
        For j = 2 To shtPFP.Cells(Rows.Count, i + 2).End(xlUp).row
            If shtPFP.Cells(1, i + 3) = "H60" Then
                PFP_Rate_Lot.Cells(Count_H60, 1) = shtPFP.Cells(j, i + 2)
                Count_H60 = Count_H60 + 1
            End If
            If shtPFP.Cells(1, i + 3) = "J60" Then
                PFP_Rate_Lot.Cells(Count_J60, 2) = shtPFP.Cells(j, i + 2)
                Count_J60 = Count_J60 + 1
            End If
            If shtPFP.Cells(1, i + 3) = "H30" Then
                PFP_Rate_Lot.Cells(Count_H30, 3) = shtPFP.Cells(j, i + 2)
                Count_H30 = Count_H30 + 1
            End If
            If shtPFP.Cells(1, i + 3) = "J30" Then
                PFP_Rate_Lot.Cells(Count_J30, 4) = shtPFP.Cells(j, i + 2)
                Count_J30 = Count_J30 + 1
            End If
            If shtPFP.Cells(1, i + 3) = "H15" Then
                PFP_Rate_Lot.Cells(Count_H15, 5) = shtPFP.Cells(j, i + 2)
                Count_H15 = Count_H15 + 1
            End If
            If shtPFP.Cells(1, i + 3) = "J15" Then
                PFP_Rate_Lot.Cells(Count_J15, 6) = shtPFP.Cells(j, i + 2)
                Count_J15 = Count_J15 + 1
            End If

 
            PFP_Rate_Lot.Cells(Count_all, 7) = shtPFP.Cells(j, i + 2)
            Count_all = Count_all + 1
        Next
        i = i + 4
    Next
    Call Filter_to_Calculate_Surface_Area_Total
End Sub

Sub Filter_to_Calculate_Surface_Area_Total()
'
' Filter_to_Calculate_Surface_Area Macro
'

'

    ShtPFP_Rate.Activate ' new
  
    ShtPFP_Rate.Columns("F:I").AutoFilter
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Clear
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("I1:I5000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ShtPFP_Rate.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ShtPFP_Rate.Columns("F:I").AutoFilter
    ShtPFP_Rate.Columns("P:S").AutoFilter
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Clear
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("S1:S5000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ShtPFP_Rate.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ShtPFP_Rate.Columns("P:S").AutoFilter
    ShtPFP_Rate.Columns("V:Y").AutoFilter
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Clear
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("Y1:Y5000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ShtPFP_Rate.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ShtPFP_Rate.Columns("V:Y").AutoFilter
    ShtPFP_Rate.Columns("AB:AE").AutoFilter
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Clear
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AE1:AE5000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ShtPFP_Rate.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ShtPFP_Rate.Columns("AB:AE").AutoFilter
    ShtPFP_Rate.Columns("AH:AK").AutoFilter
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Clear
    ShtPFP_Rate.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AK1:AK5000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ShtPFP_Rate.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ShtPFP_Rate.Columns("AH:AK").AutoFilter
    shtScriptUsfos.Activate

End Sub


Sub Filter_to_Calculate_Surface_Area()
'
' Filter_to_Calculate_Surface_Area Macro
'

'
num_scenarios = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row - 1
shtPFP.Activate
scenario = 1
For Column = 1 To num_scenarios * 5 - 4
    
    Columns(Column).Resize(, 4).EntireColumn.Select
    Selection.AutoFilter
    shtPFP.AutoFilter.Sort.SortFields.Add2 Key _
        :=Columns(Column + 3), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With shtPFP.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter

Column = Column + 4
scenario = scenario + 1
Next

shtScriptUsfos.Activate
End Sub

