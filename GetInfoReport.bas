Sub BrowseReportEXCELXML()
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    setCurrentFolder

    
    With fd
        .Title = "Selecione uma Pasta"
        .AllowMultiSelect = False
        .InitialFileName = CURRENT_FOLDER & "\"
        .Filters.Clear
        .Filters.Add "XML file", "*.xml"
        
        ' put content in link textbox
        If .Show = -1 Then
           
            shtReportInfo.Range("E9") = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
End Sub

Function xmlToDictionary(fileName)
    Dim xmlDoc As Object
    Dim wsName As String
    Dim i As Integer
    Dim rowData() As Object
    
    ' Initialize the XML Document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    
    With xmlDoc
        .async = False
        .validateOnParse = True
        .SetProperty "SelectionNamespaces", "xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'" ' Handle namespaces
        .Load fileName
    End With
    
    If xmlDoc.parseError.ErrorCode <> 0 Then
        
        errorMsg = "Parsing error in file " & fileName & Chr$(10) & Chr$(10) & _
                   "Error Description: " & xmlDoc.parseError.reason & Chr$(10) & _
                   "Error Code: " & xmlDoc.parseError.ErrorCode & Chr$(10) & Chr$(10) & _
                   "Line, Position: " & xmlDoc.parseError.line & ", " & xmlDoc.parseError.linepos & Chr$(10) & _
                   "File Position: " & xmlDoc.parseError.filepos & Chr$(10) & _
                   "Source Text: " & xmlDoc.parseError.srcText & Chr$(10) & Chr$(10) & _
                   "Please check the XML structure and syntax."
                   
        MsgBox errorMsg, vbCritical + vbOKOnly, "XML Validation Error"
    End If
    
    ' Create a dictionary to hold all worksheets data
    Set sheetData = CreateObject("Scripting.Dictionary")
    
    ' Loop through each Worksheet
    Set xmlNodeList = xmlDoc.SelectNodes("//ss:Worksheet")
    
    For Each xmlNode In xmlNodeList
        wsName = xmlNode.Attributes.getNamedItem("ss:Name").text
        Set xmlRowList = xmlNode.SelectNodes("ss:Table/ss:Row")
        headerRow = True
        
        ' Initialize dictionary for current worksheet
        Set rowDict = CreateObject("Scripting.Dictionary")
    
        For R = 0 To xmlRowList.Length - 1
            Set xmlCellList = xmlRowList.Item(R).SelectNodes("ss:Cell/ss:Data")
            
            If headerRow Then
                ReDim headerArray(xmlCellList.Length - 1)
                For c = 0 To xmlCellList.Length - 1
                    cellData = xmlCellList.Item(c).text
                    headerArray(c) = cellData ' Store header names in an array
                    ' Initialize an array for each header to store corresponding column data
                    rowDict.Add cellData, New Collection
                Next c
                headerRow = False
                columnCount = xmlCellList.Length
            Else
                ' Collect data for each column under its corresponding header
                For c = 0 To xmlCellList.Length - 1
                    If c < columnCount Then ' Ensure matching header exists
                        cellData = xmlCellList.Item(c).text
                        ' Add cell data to the collection under the matching header
                        rowDict(headerArray(c)).Add cellData
                    End If
                Next c
            End If
        Next R
        
        ' Ensure there is data beyond headers
        If Not headerRow Then
            ' Prepare final structure: header as key, array of cell data as value
            Dim finalDict As Object
            Set finalDict = CreateObject("Scripting.Dictionary")
            For Each Key In rowDict.Keys
                Dim dataArray() As String
                ReDim dataArray(rowDict(Key).Count - 1)
                For i = 1 To rowDict(Key).Count
                    dataArray(i - 1) = rowDict(Key).Item(i)
                Next i
                finalDict.Add Key, dataArray
            Next Key
    
            ' Add worksheet data to main dictionary
            sheetData.Add wsName, finalDict
        End If
    
    '        ' Add worksheet data to main dictionary
    '        sheetData.Add wsName, rowDict
    Next xmlNode
    
    Set xmlToDictionary = sheetData

End Function
Sub InfoFromReport()
    Start = Timer
    
    TurnOffCalcs
    fileName = shtReportInfo.Range("E9")
    
    Set Info = xmlToDictionary(fileName)
    
    WsNames = Info.Keys()
    
    ' Get the set names from worksheet
    Set CriticalSet = shtReportInfo.Range("C3")
    Set ColumnsSet = CriticalSet.Offset(1, 0)
    Set ThreesidedSet = ColumnsSet.Offset(1, 0)
    
    ' Matching the names from set and the names for Fem Elements in report
    FemCritical = FindBestMatchingName(WsNames, "FEMElements" & CriticalSet)
    FemColumns = FindBestMatchingName(WsNames, "FEMElements" & ColumnsSet)
    Fem3Sided = FindBestMatchingName(WsNames, "FEMElements" & ThreesidedSet)
    
    'Genie Database
    Criticals = Info(FemCritical)("FEM Elem")
    Column = Info(FemColumns)("FEM Elem")
    I_profile_3_sided = Info(Fem3Sided)("FEM Elem")
    'Box_profile_3_sided = info(Fem3Sided)("FEM Elem")
    
    
    Set Range1 = shtGenie.Range("A2").Resize(1 + UBound(Criticals), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Criticals)

    Set Range1 = shtGenie.Range("B2").Resize(1 + UBound(Column), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Column)
    
    Set Range1 = shtGenie.Range("C2").Resize(1 + UBound(I_profile_3_sided), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(I_profile_3_sided)
    
    'BeamCoordinate
    BeamCoordinate = "BeamCoordinate"
    BeamProperty = "BeamProperty"
    
    Name = Info(BeamProperty)("Name")
    Description = Info(BeamProperty)("Description")
    Section = Info(BeamProperty)("Section")
    Length = Info(BeamCoordinate)("Length [m]")
    
    
    Set Range1 = Sheet6.Range("A2").Resize(1 + UBound(Name), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Name)

    Set Range1 = Sheet6.Range("B2").Resize(1 + UBound(Description), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Description)
    
    Set Range1 = Sheet6.Range("C2").Resize(1 + UBound(Section), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Section)
    
    Set Range1 = Sheet6.Range("D2").Resize(1 + UBound(Length), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Length)
    
    ' FEM Elements
    FEMelements = "FEMElements"
    
    Name = Info(FEMelements)("Name")
    FEMElement = Info(FEMelements)("FEM Elem")
    TypeFEM = Info(FEMelements)("Type")
    
    Set Range1 = Sheet5.Range("A2").Resize(1 + UBound(Name), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(Name)

    Set Range1 = Sheet5.Range("B2").Resize(1 + UBound(FEMElement), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(FEMElement)
    
    Set Range1 = Sheet5.Range("C2").Resize(1 + UBound(TypeFEM), 1)
    Range1.Value = Application.WorksheetFunction.Transpose(TypeFEM)
    
    ' SectionProperties
    SectionProperties = "SectionProperties"
    lColumn = Sheet7.Cells(1, Columns.Count).End(xlToLeft).Column
    
    
    For col = 1 To lColumn
        Header = Sheet7.Cells(1, col)
        If Info(SectionProperties).Exists(Header) Then
            HeaderData = Info(SectionProperties)(Header)
            Set Range1 = Sheet7.Cells(2, col).Resize(1 + UBound(HeaderData), 1)
            Range1.Value = Application.WorksheetFunction.Transpose(HeaderData)
        End If

    Next col
    
    TurnOnCalcs
    
    Debug.Print Timer - Start
End Sub

Sub CopyCodeToClipboard()
    ' Define the code to be pasted on Genie to generate a report
    Dim code As String
    Dim Sets As String
    Dim currentRow As Integer
    Dim setName As Variant
    Dim setType As Range
    Dim FirstSet As Range

    
    ' Get the set names from worksheet
    Set CriticalSet = shtReportInfo.Range("C3")
    Set ColumnsSet = CriticalSet.Offset(1, 0)
    Set ThreesidedSet = ColumnsSet.Offset(1, 0)
    
    ' Format sets string for code command
    Sets = CriticalSet & "," & ColumnsSet & "," & ThreesidedSet


    ' Define the code string
    code = "XMLReport = Report("""");" & vbCrLf & _
           "XMLReport.add(ChapterStructure());" & vbCrLf & _
           "XMLReport.element(1).setLoopSets(true);" & vbCrLf & _
           "XMLReport.element(1).setSets(Array(" & Sets & "));" & vbCrLf & _
           "XMLReport.element(1).add(TableFEMElements());" & vbCrLf & _
           "XMLReport.add(ChapterStructure());" & vbCrLf & _
           "XMLReport.element(2).add(TableBeamCoordinate());" & vbCrLf & _
           "XMLReport.element(2).add(TableBeamProperty());" & vbCrLf & _
           "XMLReport.element(2).add(TableFEMElements());" & vbCrLf & _
           "XMLReport.add(ChapterProperties());" & vbCrLf & _
           "XMLReport.element(3).add(TablePropertySection());" & vbCrLf & _
           "XMLReport.saveAs(""Report.xml"", mrExcelXML);"
           
    ' Function to copy to clipboard
    Clipboard code
    
    MsgBox "Code copied to clipboard!", vbInformation
End Sub


Function Clipboard(Optional text As String) As String
    Dim variant_text: variant_text = text  'Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
    With .parentWindow.clipboardData
        Select Case True
            Case Len(text): .setData "text", variant_text
            Case Else:   Clipboard = .GetData("text")
        End Select
    End With
    End With
End Function

Function FindBestMatchingName(names As Variant, substring As String) As String
    Dim i As Long, k As Long
    Dim maxMatches As Long
    Dim currentMatches As Long
    Dim bestMatchName As String
    Dim currentChar As String
    
    ' Initialize variables
    maxMatches = 0
    bestMatchName = ""
    
    ' Iterate through each name
    For i = LBound(names) To UBound(names)
        currentMatches = 0
            
        ' Iterate through each character in the current substring
        For k = 1 To Len(substring)
            currentChar = Mid(substring, k, 1)
            
            ' Check if the character is found in the current name
            If InStr(1, UCase(names(i)), UCase(currentChar), vbTextCompare) > 0 Then
                currentMatches = currentMatches + 1
            End If
        Next k
    
        
        ' Check if the current name has more matches than the previous best
        If currentMatches > maxMatches Then
            maxMatches = currentMatches
            bestMatchName = names(i)
        End If
    Next i
    
    ' Return the name with the most matches
    FindBestMatchingName = bestMatchName
End Function

Sub TestFindBestMatchingName()
    Dim names As Variant
    Dim substring As String
    Dim bestName As String
    
    fileName = shtReportInfo.Range("E9")
    
    Set Info = xmlToDictionary(fileName)
    
    XLMnames = Info.Keys()
    
    Set SetNames = GetSetNamesCollection(shtReportInfo.Range("B3"))
    
    substring = SetNames(1)
    
    bestName = FindBestMatchingName(XLMnames, substring)
    
    MsgBox "The best matching name is: " & bestName
End Sub





