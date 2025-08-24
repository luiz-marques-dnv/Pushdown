Option Private Module       'Indicates that module is private and not seeing in macros option
Public CURRENT_FOLDER As String

Enum Program
    FAHTS
    Usfos
    Strain
    StrainNeg
End Enum
Function TurnOffCalcs()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Function

Function TurnOnCalcs()
   Application.ScreenUpdating = True
   Application.Calculation = xlCalculationAutomatic
   Application.EnableEvents = True
   Application.DisplayAlerts = True
End Function
Sub setCurrentFolder()
    CURRENT_FOLDER = Application.ActiveWorkbook.Path
End Sub

Function IsCompleted(row As Long) As Boolean
    IsCompleted = (LCase(shtScriptUsfos.Cells(row, "S")) = "completed")
End Function

Sub CheckPrograms()
    Dim prog As Program

    For prog = Program.FAHTS To Program.StrainNeg

        CheckAndAlertProgram prog
    Next prog
End Sub

Private Function GetProgramPath(ByVal prog As Program) As String
    Dim usfosFolder As String
    
    usfosFolder = shtScriptUsfos.Range("W1")
    fahtsExe = shtScriptUsfos.Range("W2")
    
    Select Case prog
        Case FAHTS
            GetProgramPath = fahtsExe
        Case Usfos
            GetProgramPath = usfosFolder & "\Usfos.exe"
        Case Strain
            GetProgramPath = usfosFolder & "\strain.exe"
        Case StrainNeg
            GetProgramPath = usfosFolder & "\strain_neg.exe"
        Case Else
            GetProgramPath = ""
    End Select
End Function

Sub CheckAndAlertProgram(ByVal prog As String)
    Dim FileSystem As Object
    Dim programPath As String
    
    programPath = GetProgramPath(prog)

    ' Create a FileSystemObject
    Set FileSystem = CreateObject("Scripting.FileSystemObject")

    ' Check if the path does not exists
    If Not FileSystem.FileExists(programPath) Then
        ' Path does not exist, prompt the user with options
        userResponse = MsgBox("It is not possible to find " & programPath & vbNewLine & "Would you like to continue anyway?", vbCritical + vbYesNo)

        If userResponse = vbNo Then
            ' User chose to exit
            End
        ElseIf userResponse = vbYes Then
            ' User chose to continue, you can add code here to handle this situation
            Exit Sub
        End If
    End If

End Sub



Function FileExist(Path As String) As Boolean
    If Dir(Path) <> vbNullString Then FileExist = True
End Function

Sub ValidatePath(ByVal Path As String)
    If Not FileExist(Path) Then
        userResponse = MsgBox("It is not possible to find " & programPath & vbNewLine & "Would you like to continue anyway?", vbCritical + vbYesNo)
    End If

End Sub

Sub CheckFilePath(ByVal Path As String)
    Dim invalidChars As String
    Dim i As Integer
    Dim newFilePath As String

    ' List of invalid characters
    invalidChars = "/ç% "

    ' Check each character in filePath
    For i = 1 To Len(Path)
        If InStr(invalidChars, Mid(Path, i, 1)) > 0 Then
            MsgBox "Directory path name has invalid characters such as '/' 'ç' '%' or space. Please move the files to a valid Path.", vbExclamation
            
            ' Exit all Subs
            End
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------
' Macros for Browse buttons
Sub BrowseUsfosFolderBtn()
    Dim fd As FileDialog
    
    ' Show a message box with a hint for the user
    MsgBox "Please select the Bin folder for Usfos. Normally in C:\Program Files\DNVGL\Usfos Vxx\Bin"
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    
    With fd
        .Title = "Selecione uma Pasta"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Program Files\"
        ' put content in link textbox
        If .Show = -1 Then
           
            Range("W1") = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
    
End Sub
Sub BrowseFahtsFileBtn()
    Dim fd As FileDialog
    
    ' Show a message box with a hint for the user
    MsgBox "Please select the exe for Fahts"
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    setCurrentFolder

    
    With fd
        .Title = "Selecione uma Pasta"
        .AllowMultiSelect = False
        .InitialFileName = CURRENT_FOLDER & "\"
        .Filters.Add "Exe Files", "*.exe"
        
        ' put content in link textbox
        If .Show = -1 Then
           
            Range("W2") = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------
' Macros to clear worksheets TODO: Place in a better module
Sub Clear_PFP()

TurnOffCalcs

Dim ultlinha As Integer
ultlinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row

For i = 2 To ultlinha
    shtPFP.Range(shtPFP.Cells(2, (i - 1) * 5 - 4), shtPFP.Cells(Rows.Count, (i - 1) * 5 - 2)).ClearContents
    shtPFP3sided.Range(shtPFP3sided.Cells(2, (i - 1) * 5 - 4), shtPFP3sided.Cells(Rows.Count, (i - 1) * 5 - 2)).ClearContents
    shtScriptUsfos.Range("J" & i, "N" & i).ClearContents
Next

TurnOnCalcs

End Sub

Sub Clear_All()
' TODO: Place in a better module
TurnOffCalcs

Dim ultlinha As Integer
ultlinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row

shtGenie.Range(shtGenie.Cells(2, "A"), shtGenie.Cells(Rows.Count, "D")).ClearContents
Sheet6.Range(Sheet6.Cells(2, "A"), Sheet6.Cells(Rows.Count, "D")).ClearContents
Sheet5.Range(Sheet5.Cells(2, "A"), Sheet5.Cells(Rows.Count, "C")).ClearContents
Sheet7.Range(Sheet7.Cells(2, "A"), Sheet7.Cells(Rows.Count, "K")).ClearContents

Clear_PFP

TurnOnCalcs

End Sub

Sub PFP_From_Fahts()
'TODO: Place in a better module

Dim textline As String
Dim Count_PFP As Integer
Dim Count_Case As Integer
Dim i As Integer
Dim UL As Integer
Dim ultlinha As Integer
    
    TurnOffCalcs
    
    Call Clear_PFP
    
    Set s = Sheets("Script_Faths")
    Set u = Sheets("Script_Usfos")

    ultlinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row
    CURRENT_FOLDER = ActiveWorkbook.Path
    
    For i = 2 To ultlinha
        PFP = ""
        Count_Case = i - 1
        Count_PFP = 1
        Caso = shtScriptUsfos.Cells(i, "B")
        FilenameBeltemp = "head_" & Caso & ".fem"
        Pasta = CURRENT_FOLDER & "\" & shtScriptUsfos.Cells(i, "A") & "_" & Caso
        myFileBeltemp = Pasta & "\" & FilenameBeltemp
        Open myFileBeltemp For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            If Left(textline, 8) = "ELMINSUL" Then
                PFP = PFP & textline & Chr(13) & Chr(10)
                shtPFP.Cells(Count_PFP + 1, Count_Case * 5 - 4) = textline
                shtPFP.Cells(Count_PFP + 1, Count_Case * 5 - 4).TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, Space:=True
                Count_PFP = Count_PFP + 1
            End If
            If Left(textline, 8) = "GroupDef" Then
                Sheet1.Cells(1, 1).Value = textline
                shtScriptUsfos.Cells(i, "M").Value = textline
                Sheet1.Cells(1, 1).TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, Space:=True
                Sheet1.Cells(1, 1).EntireRow.Copy
                Sheet1.Cells(2, 1).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
            End If
        Loop
        Close #1
        UL = Sheet1.Cells(Rows.Count, "A").End(xlUp).row
        Sheet1.Activate
        Sheet1.Range("A5", "A" & UL).Copy
        shtPFP3sided.Cells(2, Count_Case * 5 - 2).PasteSpecial
        Sheet1.Cells.ClearContents
        shtScriptUsfos.Activate
        shtScriptUsfos.Cells(i, "L").Value = PFP
    Next
    
    TurnOnCalcs
    
End Sub

Sub PFP_From_Fahts_new()
'TODO: Place in a better module
    Dim ultimaLinha As Long
    Dim row As Long
    Dim thisScenario As ScenarioType
    Dim HeadFAHTSFile As String
    Dim allElements As Collection
    Dim IthreeSidedElements As Collection
    Dim boxThreeSidedElements As Collection
    
    response = MsgBox("This will clear all PFPs in worksheet and get new ones from the FAHTS Head for each scenario" & _
                        vbNewLine & vbNewLine & _
                        "Would you like to continue?", vbYesNo)
        
    If response = vbNo Then
        End
    End If
    
    
    TurnOffCalcs
    
    Call Clear_PFP

    ultimaLinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row
    
    
    For row = 2 To ultimaLinha
        ' Get scenario attributes
        thisScenario = InitScenariosAttributes(row)

        HeadFAHTSFile = thisScenario.FolderPath & "\" & "head_" & thisScenario.Name & ".fem"
        
        ' Get elements from head file
        Set allElements = GetElementsfromHeadFile(HeadFAHTSFile, "All Elements")

        ' Get I 3-sided elements from head file
        Set IthreeSidedElements = GetElementsfromHeadFile(HeadFAHTSFile, "I 3-Sided")

        ' Get Box 3-sided elements from head file
        Set boxThreeSidedElements = GetElementsfromHeadFile(HeadFAHTSFile, "Box 3-Sided")

        ' Write the PFPs to the worksheet file
        WritePFPsToWorksheet allElements, IthreeSidedElements, boxThreeSidedElements, row

    Next row


    TurnOnCalcs

End Sub

Function GetElementsfromHeadFile(ByVal headFile As String, command As String) As Collection
    ' Get all 3-sided elements from the head file
    Dim HeadFileHandle As Long
    Dim currentLine As String
    Dim elementNumber As Variant
    Dim splitLine As Variant
    Dim i As Long
    Dim elements As Collection
    Dim ElementsinLine As Collection
    Dim entriesDiv As String
    Dim FAHTSCommand As String
    Dim fileContent As String
    
    Set elements = New Collection

    If command = "All Elements" Then
        FAHTSCommand = "ELMINSUL 1"
        
    ElseIf command = "I 3-Sided" Then
        FAHTSCommand = "GroupDef 1000 Elem"

    ElseIf command = "Box 3-Sided" Then
        FAHTSCommand = "GroupDef 1001 Elem"
    Else
        MsgBox "Invalid elementType when running GetElementsfromHeadFile"
        Exit Function
    End If
    
    ' Read the file content
    fileContent = ReadFile(headFile)
       
    Set elements = getElementsFromString(fileContent, FAHTSCommand)
    
    Set GetElementsfromHeadFile = elements
End Function

Function getElementsFromString(ByVal fileContent As String, ByVal FAHTSCommand As String) As Collection
    ' This function search for elements called after every FAHTSCommands
    ' It founds all FAHTSCommands and get the elements until a non-numeric values appears
    Dim filePath As String
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedValue As String
    Dim i As Long
    Dim results As String
    Dim elements As Collection
    
    Set elements = New Collection

    ' Formating the content trimming blank spaces and replacing linebreaks
    fileContent = WorksheetFunction.Trim(fileContent)
    fileContent = Replace(fileContent, Chr(13), " ")
    fileContent = Replace(fileContent, Chr(10), " ")
    
    
    ' Initialize the starting position
    startPos = 1
    
    ' Loop through the file content and extract all FAHTSCommand occurrences
    Do
        ' Find the position of the FAHTSCommand
        startPos = InStr(startPos, fileContent, FAHTSCommand)
        
        If startPos > 0 Then
            ' The elements start after command
            startPos = startPos + Len(FAHTSCommand)
            
            ' Initialize endPos
            endPos = Len(fileContent)
            
            ' Find the position of the first non-numeric character after the FAHTSCommand
            ' In this way consider elements that were written just on lines after the command
            For i = startPos To Len(fileContent)
                If Not IsNumeric(Mid(fileContent, i, 1)) And Mid(fileContent, i, 1) <> " " Then
                    endPos = i - 1
                    Exit For
                
                End If
            Next i
            
            ' Extract the elements and format
            extractedValue = Mid(fileContent, startPos, endPos - startPos + 1)
            extractedValue = WorksheetFunction.Trim(extractedValue)
            
            ' Splitting to get easily each value
            splitValues = Split(extractedValue)
            For i = 0 To UBound(splitValues)
                elementNumber = CLng(splitValues(i))
                elements.Add elementNumber
            Next i
            
            startPos = endPos + 1
        End If
    Loop Until startPos = 0
    
    Set getElementsFromString = elements
End Function

Function ReadFile(filePath As String) As String
    Dim fileNumber As Integer
    Dim fileContent As String
    Dim line As String
    
    ' Get a free file number
    fileNumber = FreeFile
    
    ' Open the file
    Open filePath For Input As fileNumber
    ' Read the file line by line
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        fileContent = fileContent & line & vbCrLf
    Loop
    
    ' Close the file
    Close fileNumber
    
    ' Return the content
    ReadFile = fileContent
End Function

Sub WritePFPsToWorksheet(allElements As Collection, IthreeSidedElements As Collection, boxThreeSidedElements As Collection, row As Long)
    Dim element As Variant
    Dim elementNumber As Variant

    ' Write all elements to the worksheet
    For Each elementNumber In allElements
        writeAllElement row, elementNumber

    Next elementNumber

    ' Write I 3-sided elements to the worksheet
    For Each elementNumber In IthreeSidedElements
        writeI3SidedElement row, elementNumber
    Next elementNumber

    ' Write Box 3-sided elements to the worksheet
    For Each elementNumber In boxThreeSidedElements
        writeBox3SidedElement row, elementNumber
    Next elementNumber

End Sub


    



