Const STRASPAS As String = """"   'defining aspas to be inside strings
Const LOG_FILE_NAME As String = "PFP_Analysis_Log.txt"
Const HEAD_DEFAULT_FILE As String = "\usfos_head_default.fem"

Public checkStrainNegCriteria As Boolean
Public pushDownMethod As String

Type ScenarioType
    Name As String
    Number As Long
    FolderPath As String
    BeltempFilePath As String
    FileStrain As String
    FileStrainNeg As String
    strainCriteria As Double
    row As Long
    Iteration As Long
    FileUsfosOut As String
End Type

Sub Run_analysis()
    ' Initialization
    Dim ultimaLinha As Long
    Dim iter As Long
    Dim row As Long
    
    ' Sub to Disable calculations and screen updating
    TurnOnCalcs
    
    ' Check if necessary programs are on correct path
    AuxMacros.CheckPrograms
    
    'set current directory to public variable CURRENT_FOLDER
    ' TODO: Set default variables
    AuxMacros.setCurrentFolder
    
    ultimaLinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row

    checkStrainNegCriteria = shtScriptUsfos.Range("W32").Value

    ' TODO: CHange this to a config module
    pushDownMethod = Worksheets("Config").Range("A2").Value ' "AtTime" or "UpToTime"

    ' Main processing loop
    For row = 2 To ultimaLinha
        Do Until CriteriaMet(iter, row)
            ' Retrieve iteration number in the worksheet
            iter = shtScriptUsfos.Cells(row, "D")
            iter = iter + 1
        
            ProcessIteration iter, row
            
            ' Update iteration number in the worksheet
            shtScriptUsfos.Cells(row, "D") = iter
        Loop
    Next row

    ' Finalization
    MsgBox "Usfos/Faths Analysis Finished!"
    
    ' Sub to Enable calculations and screen updating
    TurnOnCalcs
End Sub

Private Function InitializeIteration(row) As Integer
    Set iterationCell = shtScriptUsfos.Cells(row, "D")
    If iterationCell = "" Then
        InitializeIteration = 0
        
    Else
        InitializeIteration = shtScriptUsfos.Cells(17, 23)
    End If
    
End Function

Private Sub ProcessIteration(iter As Long, row As Long)
    Dim loadCaseData As Collection
    Dim HighestStrainTime As Double
    Dim methodOutput() As Variant
    Dim PFPElement As Long
    Dim thisScenario As ScenarioType
   

    ' Initialize values from the sheet
    ' TODO: Change this to a class
    thisScenario = InitScenariosAttributes(row)

    ' Extract the load case data from the Beltemp file
    Set loadCaseData = ExtractLoadCaseData(thisScenario.BeltempFilePath)

    ' Get desired load case time based on desired method
    methodOutput = GetMethodOutputs(loadCaseData, thisScenario, pushDownMethod)

    ' Getting outputs depending on method
    HighestStrainTime = methodOutput(0)
    PFPElement = methodOutput(1)

    ' Apply PFP on eligible element
    If PFPElement <> 0 Then
        ' Write PFP element to the script file
        writePFPElement row, PFPElement

        ' Update Fahts for on going cases
        UpdateHeadFahts thisScenario

        ' Running Fahts
        ExecuteProgram Program.FAHTS, thisScenario.FolderPath, thisScenario.Name

        ' Copying Fahts files just to check    TODO: Delete after checking
        CopyFahtsOutFiles thisScenario.Name, thisScenario.FolderPath, thisScenario.Iteration

        ' Update Completion Status
        shtScriptUsfos.Cells(row, "S") = "on going"

    ' If PFP Element is 0, then the criteria was met and no other element needs pfp.
    ' Finally update to the case with highest strain for plotting
    ElseIf PFPElement = 0 Then
        ' Update Head USFOS file
        'TODO: Handle better after Paula response
        ' UpdateHeadUsfos thisScenario, HighestStrainTime, "AtTime"

        ' ' Running Usfos
        ' ExecuteProgram Program.Usfos, thisScenario.FolderPath, thisScenario.Name

        ' ' Running strain and Straing neg to get values outside criteria
        ' ExecuteProgram Program.Strain, thisScenario.FolderPath, thisScenario.Name
    
        ' If checkStrainNegCriteria Then
        '     ExecuteProgram Program.StrainNeg, thisScenario.FolderPath, thisScenario.Name
        ' End If

        shtScriptUsfos.Cells(row, "S") = "completed"
    End If
    

    ' Verify and update criteria
'    VerifyAndUpdateCriteria iter
End Sub

Function InitScenariosAttributes(ByVal row As Long) As ScenarioType
    ' Update the scenario attributes based on the row values TODO: Change this to a class
    Dim BeltempFilename As String
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim BeltempFilePath As String
    Dim scenarioAttributes As ScenarioType
    Dim thisScenario As ScenarioType
    
    ' Set the current folder
    AuxMacros.setCurrentFolder
    
    ' Initialize and fetch values from the worksheet
    thisScenario.row = row
    thisScenario.Number = shtScriptUsfos.Cells(row, "A")
    thisScenario.Name = shtScriptUsfos.Cells(row, "B")
    thisScenario.Iteration = shtScriptUsfos.Cells(row, "D")
    thisScenario.strainCriteria = shtScriptUsfos.Cells(row, "I")


    ' Paths
    thisScenario.FolderPath = CURRENT_FOLDER & "\" & thisScenario.Number & "_" & thisScenario.Name

    ' Beltemp file path
    BeltempFilename = thisScenario.Name & "_output_beltemp"
    thisScenario.BeltempFilePath = thisScenario.FolderPath & "\" & BeltempFilename & ".fem"

    ' Usfos Output paths
    thisScenario.FileStrain = thisScenario.FolderPath & "\" & thisScenario.Name & "_strain_output"
    thisScenario.FileStrainNeg = thisScenario.FolderPath & "\" & thisScenario.Name & "_strain_neg_output"
    thisScenario.FileUsfosOut = thisScenario.FolderPath & "\" & thisScenario.Name & "_usfos_output.out"

    InitScenariosAttributes = thisScenario
End Function

Private Function GetMethodOutputs(ByVal loadCaseData As Collection, ByRef thisScenario As ScenarioType, ByVal method As String) As Variant()
    ' Get the load case time based on the desired method
    
    Select Case method
        Case "AtTime"
            MethodsOutput = HighestStrainMethod(loadCaseData, thisScenario)
    End Select

    Select Case method
        Case "UpToTime"
            MethodsOutput = upToTimeMethod(loadCaseData, thisScenario)
    End Select
    
    GetMethodOutputs = MethodsOutput
End Function

Private Function HighestStrainMethod(ByVal loadCaseData As Collection, ByRef thisScenario As ScenarioType) As Variant()
    ' Get the time of the highest strain element above criteria
    Dim loadCaseTimeDetails As Variant
    Dim loadcase As Long
    Dim LoadCaseTime As Double
    Dim HighestStrainTime As Double
    Dim HighestPFPStrain As Double
    Dim HighestStrain As Double
    Dim HighestElementStrain As Long
    Dim HighestStrainCriteria As Double
    Dim ElementsAboveCriteria As Object
    Dim ElementsAbove0Criteria As Object
    Dim pfpEligibleElement As elementType
    Dim blankElement As elementType
    
    HighestStrain = 0
    HighestElementStrain = 0
    HighestStrainCriteria = 0

    For Each loadCaseTimeDetails In loadCaseData
        loadcase = loadCaseTimeDetails(0)
        LoadCaseTime = loadCaseTimeDetails(1)

        ' Update Head USFOS file
        UpdateHeadUsfos thisScenario, LoadCaseTime, pushDownMethod

        ' Running Usfos
        ExecuteProgram Program.Usfos, thisScenario.FolderPath, thisScenario.Name

        ' Running strain to get values outside criteria
        ExecuteProgram Program.Strain, thisScenario.FolderPath, thisScenario.Name

        ' Process positive strain criteria
        Set ElementsAboveCriteria = GetElementsAboveCriteriaStrain(thisScenario.FileStrain, thisScenario.strainCriteria)
        
        ' Find the highest strain element eligible for PFP (critical element)
        pfpEligibleElement = FindPFPEligibleElement(ElementsAboveCriteria, thisScenario)


        If checkStrainNegCriteria And pfpEligibleElement.Number = 0 Then
            ' Running straing neg to get values outside criteria
            ExecuteProgram Program.StrainNeg, thisScenario.FolderPath, thisScenario.Name

            ' Process negative strain criteria
            Set ElementsAboveCriteria = GetElementsAboveCriteriaStrain(thisScenario.FileStrainNeg, thisScenario.strainCriteria * -1)

            ' Find the highest strain element eligible for PFP (critical element)
            pfpEligibleElement = FindPFPEligibleElement(ElementsAboveCriteria, thisScenario)
        End If
        

        ' Process strain criteria with 0 to store the highest strain of all
        Set ElementsAbove0Criteria = GetElementsAboveCriteriaStrain(thisScenario.FileStrain, 0)
        
        ' Get the highest strain considering that ElementsAboveCriteria is sorted from highest to lowest strain
        If ElementsAbove0Criteria.Count > 0 Then
            HighestStrainCriteria = ElementsAbove0Criteria.Items(0)
            
        End If

        ' Read elements above buckling criteria
        'ProcessBucklingCriteria iter

        ' Update the highest eligeble strain and element if the current strain is higher
        If pfpEligibleElement.Strain > HighestPFPStrain Then
            HighestPFPStrain = pfpEligibleElement.Strain
            HighestElementStrain = pfpEligibleElement.Number
        End If
        

        ' Update the highest strain and time if the current strain is higher
        If HighestStrainCriteria >= HighestStrain Then
            HighestStrain = HighestStrainCriteria
            HighestStrainTime = LoadCaseTime
        End If
        
        'Copying starin files just to check    TODO: Delete after checking
        CopyStrainFiles thisScenario.Name, thisScenario.FolderPath, thisScenario.Iteration, loadcase

        ' Copying Status files just to check buckling    TODO: Delete after checking
        CopyStatusFiles thisScenario.Name, thisScenario.FolderPath, thisScenario.Iteration, loadcase
        
        ' Reset Eligible Element to default values
        pfpEligibleElement = blankElement
    Next loadCaseTimeDetails
    
    HighestStrainMethod = Array(HighestStrainTime, HighestElementStrain)
End Function

Private Function upToTimeMethod(ByVal loadCaseData As Collection, ByRef thisScenario As ScenarioType) As Variant()
    ' Get the time of the highest strain element above criteria
    Dim loadCaseTimeDetails As Variant
    Dim loadcase As Long
    Dim LoadCaseTime As Double
    Dim HighestStrainTime As Double
    Dim HighestPFPStrain As Double
    Dim HighestStrain As Double
    Dim HighestElementStrain As Long
    Dim HighestStrainCriteria As Double
    Dim ElementsAboveCriteria As Object
    Dim pfpEligibleElement As elementType
    Dim blankElement As elementType
    
    HighestStrain = 0
    HighestElementStrain = 0
    HighestStrainCriteria = 0

    ' Get last time and load case from the loadCaseData
    If loadCaseData.Count > 0 Then
        loadcase = loadCaseData(loadCaseData.Count)(0)
        LoadCaseTime = loadCaseData(loadCaseData.Count)(1)
    End If

        ' Update Head USFOS file
        UpdateHeadUsfos thisScenario, LoadCaseTime, pushDownMethod

        ' Running Usfos
        ExecuteProgram Program.Usfos, thisScenario.FolderPath, thisScenario.Name

        ' Running strain to get values outside criteria
        ExecuteProgram Program.Strain, thisScenario.FolderPath, thisScenario.Name

        ' Process positive strain criteria
        Set ElementsAboveCriteria = GetElementsAboveCriteriaStrain(thisScenario.FileStrain, thisScenario.strainCriteria)
        
        ' Find the highest strain element eligible for PFP (critical element)
        pfpEligibleElement = FindPFPEligibleElement(ElementsAboveCriteria, thisScenario)


        If checkStrainNegCriteria And pfpEligibleElement.Number = 0 Then
            ' Running straing neg to get values outside criteria
            ExecuteProgram Program.StrainNeg, thisScenario.FolderPath, thisScenario.Name

            ' Process negative strain criteria
            Set ElementsAboveCriteria = GetElementsAboveCriteriaStrain(thisScenario.FileStrainNeg, thisScenario.strainCriteria * -1)

            ' Find the highest strain element eligible for PFP (critical element)
            pfpEligibleElement = FindPFPEligibleElement(ElementsAboveCriteria, thisScenario)
        End If
    

        ' Read elements above buckling criteria
        'ProcessBucklingCriteria iter

        ' Update the highest eligeble strain and element if the current strain is higher
        If Abs(pfpEligibleElement.Strain) > Abs(HighestPFPStrain) Then
            HighestPFPStrain = pfpEligibleElement.Strain
            HighestElementStrain = pfpEligibleElement.Number
            HighestStrainTime = LoadCaseTime
        End If

        'Copying starin files just to check    TODO: Delete after checking
        CopyStrainFiles thisScenario.Name, thisScenario.FolderPath, thisScenario.Iteration, loadcase

        ' Copying Status files just to check buckling    TODO: Delete after checking
        CopyStatusFiles thisScenario.Name, thisScenario.FolderPath, thisScenario.Iteration, loadcase
        
        ' Reset Eligible Element to default values
        pfpEligibleElement = blankElement

        ' Check if any error happened in USFOS analysis
        errorMessage = checkErrorInFile(thisScenario.FileUsfosOut, thisScenario)

        If errorMessage <> "" Then
            HighestStrainTime = 0
            HighestElementStrain = 0
            handleError thisScenario.FileUsfosOut, thisScenario, errorMessage
        End If

    upToTimeMethod = Array(HighestStrainTime, HighestElementStrain)
End Function

Private Function GetElementsAboveCriteriaStrain(ByVal filePath As String, strainCriteria) As Object
    ' Process the strain file (either positive or negative) and returns a dinamic array of elements with number and strain that have a strain above the criteria
    ' It will overwrites the array if high loadcase is found
    
    Dim strainFile As Integer
    Dim lineFromFile As String
    Dim elementLine As String
    Dim elementLine_Array() As String
    Dim elementsStrain As Object

    Set elementsStrain = CreateObject("Scripting.Dictionary")

    strainFile = FreeFile
    
    Open filePath For Input As #strainFile
    Do Until EOF(strainFile)
        Line Input #strainFile, lineFromFile

        ' Finding beginning of Elements
        If InStr(lineFromFile, "Element no") Then
            Line Input #strainFile, lineFromFile
            elementsStrain.RemoveAll                        ' Way to reset a dictionary if elements are set
            
            ' loop to store elements above criteria
            While Len(Trim(lineFromFile)) <> 0
                elementLine = WorksheetFunction.Trim(lineFromFile)             ' removing extra spaces to avoid errors
                elementLine_Array = Split(elementLine, " ")                    ' the array should be: Element No, Strain, %
                elementNumber = elementLine_Array(0)
                ElementStrain = elementLine_Array(1)
                If Abs(ElementStrain) >= Abs(strainCriteria) Then
                    elementsStrain.Add elementNumber, ElementStrain
                End If
                Line Input #strainFile, lineFromFile
            Wend
        End If
    Loop

    Close #strainFile
    Set GetElementsAboveCriteriaStrain = elementsStrain

End Function

Private Sub UpdateCompletionStatus(ByVal row As Long, ByVal PFPElement As Long)
'TODO: Change this condition check
    If PFPElement <> Empty Then
        shtScriptUsfos.Cells(row, "S") = "on going"
    ElseIf PFPElement = Empty Then
        shtScriptUsfos.Cells(row, "S") = "completed"
    End If
End Sub

Private Function FindPFPEligibleElement(ByRef elementsStrain As Object, ByRef thisScenario As ScenarioType) As elementType
    'Verificar se elemento é elegível para receber PFP e retorna o numero do elemento elegível
    Dim test_strain As Long
    Dim element As elementType
    Dim blankElement As elementType
    Dim i As Long
    
    If elementsStrain.Count > 0 Then
        For i = 0 To elementsStrain.Count - 1
            element.Number = elementsStrain.Keys()(i)
            element.Strain = elementsStrain.Items()(i)
            On Error Resume Next
            test_strain = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(element.Number, shtGenie.Range("A2:A50000"), 1, False), "N")
            If Err = 0 Then
                ' Check if element is already protected and add if not

                If IsElementAlready3SidedProtected(thisScenario.row, element.Number) Then
                    ' Initialize scenario attributes

                    LogAlreadyProtectedWarning element.Number, thisScenario
                Else
                    FindPFPEligibleElement = element
                    Exit Function
                End If
            End If
        Next
    End If
    
    Debug.Print element.Number & "," & element.Strain
    element = blankElement
    Debug.Print element.Strain
    'FindPFPEligibleElement = element
End Function

Private Sub writePFPElement(ByVal row As Long, ByVal PFPElementNumber As Long)
    ' refactor all this code to separate the PFP 3-sided I and box profile finding and writing
    Dim ultlinhaPFP As Long
    Dim verify As Boolean
    Dim k As Long
    Dim test_3sided_I As String
    Dim test_3sided_Box As String
    Dim thisScenario As ScenarioType


    If PFPElementNumber <> 0 Then
        writeAllElement row, PFPElementNumber

        writeI3SidedElement row, PFPElementNumber

        writeBox3SidedElement row, PFPElementNumber

        thisScenario = InitScenariosAttributes(row)

        ' Log the element
        LogElement PFPElementNumber, thisScenario
    
    End If


End Sub

Sub writeAllElement(ByVal row As Long, ByVal PFPElementNumber As Long)
    Dim ultlinhaPFP As Long
    Dim verify As Boolean
    Dim k As Long
    

    shtScriptUsfos.Cells(row, "T") = "El. " & PFPElementNumber & " above criteria"
    ultlinhaPFP = shtPFP.Cells(Rows.Count, (row - 2) * 5 + 3).End(xlUp).row
    verify = True
    For k = 2 To ultlinhaPFP + 1
        If PFPElementNumber = shtPFP.Cells(k, (row - 2) * 5 + 3) Then
            verify = False
        End If
    Next
    If verify = True Then
        shtScriptUsfos.Cells(row, "J") = shtScriptUsfos.Cells(row, "J") & Chr(10) & "ELMINSUL 1   " & PFPElementNumber
        shtPFP.Cells(ultlinhaPFP + 1, (row - 2) * 5 + 1) = "ELMINSUL"
        shtPFP.Cells(ultlinhaPFP + 1, (row - 2) * 5 + 2) = "1"
        shtPFP.Cells(ultlinhaPFP + 1, (row - 2) * 5 + 3) = PFPElementNumber

    End If
End Sub

Sub writeI3SidedElement(ByVal row As Long, ByVal PFPElementNumber As Long)
    Dim ultlinhaPFP3Sided As Long
    Dim verify As Boolean
    Dim k As Long
    Dim test_3sided_I As String
    Dim thisScenario As ScenarioType



    'Verificar se elemento é elegível para receber PFP 3-sided I profile
    On Error Resume Next
    test_3sided_I = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(PFPElementNumber, shtGenie.Range("C2:C50000"), 1, False), "N")

    If Err = 0 Then
        ' Check if element is already protected and add if not
        If IsElementAlready3SidedProtected(row, PFPElementNumber) Then
            ' Initialize scenario attributes
            thisScenario = InitScenariosAttributes(row)
            
            LogAlreadyProtectedWarning PFPElementNumber, thisScenario
        Else
            scenarioColumnIndex = (row - 2) * 5 + 3
            ultlinhaPFP3Sided = shtPFP3sided.Cells(Rows.Count, scenarioColumnIndex).End(xlUp).row

            shtPFP3sided.Cells(ultlinhaPFP3Sided + 1, scenarioColumnIndex) = test_3sided_I
            shtScriptUsfos.Cells(row, "M") = shtScriptUsfos.Cells(row, "M") & "     " & PFPElementNumber
        End If

    End If
End Sub

Sub writeBox3SidedElement(ByVal row As Long, ByVal PFPElementNumber As Long)
    Dim ultlinhaPFP3Sided As Long
    Dim verify As Boolean
    Dim k As Long
    Dim test_3sided_Box As String
    Dim thisScenario As ScenarioType
    
    'Verificar se elemento é elegível para receber PFP 3-sided box profile
    On Error Resume Next
    test_3sided_Box = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(PFPElementNumber, shtGenie.Range("D2:D50000"), 1, False), "N")
    If Err = 0 Then

        ' Check if element is already protected and add if not
        If IsElementAlready3SidedProtected(row, PFPElementNumber) Then
            ' Initialize scenario attributes
            thisScenario = InitScenariosAttributes(row)

            LogAlreadyProtectedWarning PFPElementNumber, thisScenario
        Else
            scenarioColumnIndex = (row - 2) * 5 + 3
            ultlinhaPFP3Sided = shtPFP3sided.Cells(Rows.Count, scenarioColumnIndex).End(xlUp).row

            shtPFP3sided.Cells(ultlinhaPFP3Sided + 1, scenarioColumnIndex) = test_3sided_Box
            shtScriptUsfos.Cells(row, "N") = shtScriptUsfos.Cells(row, "N") & "     " & PFPElementNumber
        End If

    End If

End Sub

Private Function IsElementAlready3SidedProtected(ByVal row As Long, ByVal elementNumber As Long) As Boolean
    Dim lastRowWithData As Long
    Dim k As Long
    Dim columnIndex As Long
    
    columnIndex = (row - 2) * 5 + 3
    lastRowWithData = shtPFP3sided.Cells(Rows.Count, columnIndex).End(xlUp).row
    
    For k = 2 To lastRowWithData
        If elementNumber = shtPFP3sided.Cells(k, columnIndex) Then
            IsElementAlready3SidedProtected = True
            Exit Function
        End If
    Next k
    
    IsElementAlready3SidedProtected = False
End Function
Sub testIsElementAlready3SidedProtected()
    Dim row As Long
    Dim elementNumber As Long
    Dim isProtected As Boolean
    
    row = 2 ' Example row number
    elementNumber = 996 ' Example element number to check
    
    isProtected = IsElementAlready3SidedProtected(row, elementNumber)
    
    If isProtected Then
        MsgBox "Element " & elementNumber & " is already protected in row " & row
    Else
        MsgBox "Element " & elementNumber & " is not protected in row " & row
    End If
End Sub

Private Sub ProcessBucklingCriteria(ByVal iter As Integer)
    ' Logic to read elements above buckling criteria
End Sub


Private Sub UpdateHeadUsfos(ByRef scenario As ScenarioType, ByVal LoadCaseTime As String, ByVal pushDownMethod As String)
    ' Updates the Head USFOS file for each row in the specified range.
    ' TODO: CHange the fucntion to receive the files path
    Dim HeadFilename As String
    Dim HeadDefaultFile As String
    Dim HeadCaseFile As String

    Dim currentLoadCaseTime As Double
    Dim FSObject As Object

    
    ' Define the paths to Head file
    HeadDefaultFile = scenario.FolderPath & HEAD_DEFAULT_FILE

    HeadFilename = "usfos_head_" & scenario.Name & ".fem"
    HeadCaseFile = scenario.FolderPath & "\" & HeadFilename

    ' Copy and modify head file
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile HeadDefaultFile, HeadCaseFile, True
    
    ' Update the head file
    If pushDownMethod = "UpToTime" Then
        ModifyHeadFileUpToTime headFile:=HeadCaseFile, LoadCaseTime:=LoadCaseTime
    ElseIf pushDownMethod = "AtTime" Then
        ModifyHeadFileAtTime headFile:=HeadCaseFile, LoadCaseTime:=LoadCaseTime
    Else
        ' Handle other methods if needed
        ' TODO: Implement handling for other methods
        MsgBox "Unhandled method: " & pushDownMethod
    End If

End Sub

Private Sub CopyFirstRunFiles(ByVal ultimaLinha As Long)
    ' Copies files from the first USFOS run to a designated folder.
    Dim row As Long
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim firstRunFolderPath As String
    Dim sourcePath As String
    
    For row = 2 To ultimaLinha
        ' Get the case name and create the case folder path
        scenarioName = shtScriptUsfos.Cells(row, "B")
        caseFolderPath = CURRENT_FOLDER & "\" & shtScriptUsfos.Cells(row, "A") & "_" & scenarioName
        
        ' Create the folder path for the first run
        firstRunFolderPath = caseFolderPath & "\" & scenarioName & "_FirstRun"
        
        ' Create a directory for the first run if it doesn't exist
        If Dir(firstRunFolderPath, vbDirectory) = "" Then ' Check if the directory already exists
            On Error Resume Next
            MkDir firstRunFolderPath
            On Error GoTo 0
        End If

        ' Copy files from the case folder to the first run folder
        ' TODO: CHange the way this is called so it does not have to put scenarioName
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_usfos_output.out"
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_usfos_output.raf"
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_usfos.trm"
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_usfos_output_status.text"
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_strain_output"
        CopyCaseFile caseFolderPath, firstRunFolderPath, scenarioName & "_output.raf"
    Next row
End Sub

Private Sub ExecuteShellCommand(ByVal command As String)
    'On Error GoTo ErrorHandler
    ' Todo: Check if error handler is necessayr
    
    Dim shellObject As Object
    Set shellObject = CreateObject("WScript.Shell")
    Dim ShellCommand As Variant
    
    ' Execute the command
    ShellCommand = shellObject.Run(command, 1, True)

    'Exit Sub

'ErrorHandler:
    'MsgBox "An error occurred: " & Err.Description
    'End
End Sub

Private Sub ExecuteProgramBackup(ByVal programName As Program, ByVal workingDirectory As String, ByVal scenarioName As String)
    Dim executable As String
    Dim outputFileName As String
    Dim inputFileName As String
    Dim usfosBin As String
    Dim fullCommand As String
    Dim startTime As Double
    Dim endTime As Double
    Dim executionTime As Double
    Dim minimumExpectedTime As Double
        
    minimumExpectedTime = 3 ' Set this to the minimum expected execution time in seconds
    
    usfosBin = shtScriptUsfos.Range("W1")
    fahtsExe = shtScriptUsfos.Range("W2")

    Select Case programName
        Case Program.FAHTS
            executable = fahtsExe
            inputFileName = "input.inp"

        Case Program.Usfos
            executable = usfosBin & "\Usfos.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_usfos.trm"
            inputFileName = "usfos_input.inp"

        Case Program.Strain
            executable = usfosBin & "\strain.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain.trm"
            inputFileName = "strain_input.inp"

        Case Program.StrainNeg
            executable = usfosBin & "\strain_neg.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain_neg.trm"
            inputFileName = "strain_neg_input.inp"

        Case Else
            MsgBox "Invalid Program Name when calling ExecuteProgram sub"
            Exit Sub
    End Select

    
    fullCommand = """" & executable & """" & executionCommand & outputFileName & " < " & inputFileName
    
    ' Get the start time
    startTime = Timer
    
    ExecuteShellCommandHidden fullCommand, workingDirectory
    
    ' Get the end time
    endTime = Timer

    ' Calculate the execution time
    executionTime = endTime - startTime

    ' Check if the execution time is below the minimum expected time
    If executionTime < minimumExpectedTime Then
        MsgBox "Warning: The executable '" & executable & "' closed too quickly. This might indicate a problem." & vbNewLine & _
               vbNewLine & "Details:" & vbNewLine & _
               "Working Directory: " & workingDirectory & vbNewLine & _
               "Scenario Name: " & scenarioName & vbNewLine & _
               vbNewLine & "Tip: Check for program license or path issues."
    End If

    
End Sub
Sub ExecuteProgram(ByVal prog As Program, ByVal workingDirectory As String, ByVal scenarioName As String)
    Dim executable As String
    Dim outputFileName As String
    Dim inputFileName As String
    Dim usfosBin As String
    Dim fullCommand As String
    Dim PreviousTimestamp As Date
    Dim trmFileName As String
    Dim OutputCheckPath As String
    Dim HasFileChanged As Boolean
        

    usfosBin = shtScriptUsfos.Range("W1")
    fahtsExe = shtScriptUsfos.Range("W2")

    Select Case prog
        Case Program.FAHTS
            executable = fahtsExe
            inputFileName = "input.inp"
            outputFileName = scenarioName & "_output_beltemp.fem"
            programName = "FAHTS"

        Case Program.Usfos
            executable = usfosBin & "\Usfos.exe"
            executionCommand = " 400 > "
            trmFileName = scenarioName & "_usfos.trm"
            inputFileName = "usfos_input.inp"
            outputFileName = scenarioName & "_usfos_output.out"
            programName = "USFOS"

        Case Program.Strain
            executable = usfosBin & "\strain.exe"
            executionCommand = " 400 > "
            trmFileName = scenarioName & "_strain.trm"
            inputFileName = "strain_input.inp"
            outputFileName = scenarioName & "_strain_output"
            programName = "STRAIN"

        Case Program.StrainNeg
            executable = usfosBin & "\strain_neg.exe"
            executionCommand = " 400 > "
            trmFileName = scenarioName & "_strain_neg.trm"
            inputFileName = "strain_neg_input.inp"
            outputFileName = scenarioName & "_strain_neg_output"
            programName = "STRAIN NEG"

        Case Else
            MsgBox "Invalid Program Name when calling ExecuteProgram sub"
            Exit Sub
    End Select

    
    fullCommand = """" & executable & """" & executionCommand & trmFileName & " < " & inputFileName
    
    ' File path used for checking if the program run as expected
    OutputCheckPath = workingDirectory & "\" & outputFileName
    
    ' Get the timestamp before running the program
    PreviousTimestamp = ScriptsHandling.GetFileLastModified(OutputCheckPath)
    
    ExecuteShellCommandHidden fullCommand, workingDirectory
    
    HasFileChanged = CheckFileChange(OutputCheckPath, PreviousTimestamp)

    ' Check if the file has changed. If it has, the program ran as expected
    If Not HasFileChanged Then
        response = MsgBox("Warning: The " & OutputCheckPath & " file was not updated." & _
                            vbNewLine & vbNewLine & _
                            "This means that " & programName & " did not run as expected" & _
                            vbNewLine & vbNewLine & _
                            "Tip: Check for program license or path issues." & _
                            vbNewLine & vbNewLine & _
                            "Would you like to continue?", vbYesNo)
        If response = vbNo Then
            End
        End If
    End If

    
End Sub
Private Sub ExecuteShellCommandHidden(ByVal command As String, workingDirectory As String)
    ' Calls Shell to execute outside programs in hidden mode (intWindowStyle=7)

    Dim shellObject As Object
    Dim ExecuteShellCommandHidden As Variant
    Set shellObject = CreateObject("WScript.Shell")

    ' Construct the command with proper quoting
    Dim fullCommand As String
    fullCommand = "cmd.exe /c " & command

    ' Change the working directory. Same as cd command
    shellObject.CurrentDirectory = workingDirectory

    ' Execute the command
    ExecuteShellCommandHidden = shellObject.Run(fullCommand, 7, True)

    Exit Sub

'ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ExecuteShellCommandHidden = Nothing
End Sub


Private Function ExtractLoadCaseData(ByVal BeltempFilePath As String) As Collection
    ' Read and process the Beltemp file to get all load cases steps and times.
    Dim currentLine As String
    Dim loadCaseLine As String
    Dim splitLoadCaseDetails() As String
    Dim loadCaseData As Collection
    Dim loadCaseTimeDetails(1) As String
    
    Set loadCaseData = New Collection
    
    BeltempFileHandle = FreeFile
    Open BeltempFilePath For Input As #BeltempFileHandle
    Do Until EOF(BeltempFileHandle)
        Line Input #BeltempFileHandle, currentLine
        If InStr(currentLine, "LCASETIM") Then
            loadCaseLine = WorksheetFunction.Trim(currentLine)          ' remove extra spaces
            splitLoadCaseDetails = Split(loadCaseLine, " ")             ' split into array
            
            ' Creating an array of size 2
            loadCaseTimeDetails(0) = splitLoadCaseDetails(1)            ' get the load case info
            loadCaseTimeDetails(1) = splitLoadCaseDetails(2)            ' get the time info

            ' Adding the pair to the collection
            loadCaseData.Add loadCaseTimeDetails
        End If
    Loop
    Close #belTempFilePathHandle
    
    Set ExtractLoadCaseData = loadCaseData
End Function

Private Function GetNextLoadCaseTime(ByVal loadCaseData As Collection, ByVal lastLoadCaseTime As Double) As Double
    ' Get the next load case time based on the last load case time
    Dim loadCaseTimeDetails As Variant
    
    loadCaseTimeFound = False
    
    For Each loadCaseTimeDetails In loadCaseData
        If CDbl(loadCaseTimeDetails(1)) > CDbl(lastLoadCaseTime) Then
            GetNextLoadCaseTime = loadCaseTimeDetails(1)
            Exit Function
        End If
    Next

End Function

Sub ModifyHeadFileAtTime(ByVal headFile As String, ByVal LoadCaseTime As Double)
    ' Modify the head file based on the case information
    Dim HeadFileHandle As Long
    
    HeadFileHandle = FreeFile
    Open headFile For Input As HeadFileHandle
    HeadFileContent = Input(LOF(HeadFileHandle), HeadFileHandle)
    Close #HeadFileHandle
    
    ' Replace the specific string
    HeadFileContent = Replace(HeadFileContent, "'###PUSHDOWN COMMAND###", "Pushdown AtTime " & LoadCaseTime)

    ' Open the file for writing
    HeadFileHandle = FreeFile
    Open headFile For Output As HeadFileHandle
    Print #HeadFileHandle, HeadFileContent
    Close #HeadFileHandle
End Sub

Sub ModifyHeadFileUpToTime(ByVal headFile As String, ByVal LoadCaseTime As Double)
    ' Modify the head file based on the case information
    Dim HeadFileHandle As Long
    
    HeadFileHandle = FreeFile
    Open headFile For Input As HeadFileHandle
    HeadFileContent = Input(LOF(HeadFileHandle), HeadFileHandle)
    Close #HeadFileHandle
    
    ' Replace the specific string
    HeadFileContent = Replace(HeadFileContent, "'###PUSHDOWN COMMAND###", "Pushdown UpToTime " & LoadCaseTime)

    ' Open the file for writing
    HeadFileHandle = FreeFile
    Open headFile For Output As HeadFileHandle
    Print #HeadFileHandle, HeadFileContent
    Close #HeadFileHandle
End Sub


Sub UpdateHeadFahts(ByRef scenario As ScenarioType)
    Dim HeadFahtsPath As String

    
    
    HeadFahtsPath = scenario.FolderPath & "\head_" & scenario.Name & ".fem"

    shtScriptUsfos.Cells(scenario.row, "L") = shtScriptUsfos.Cells(scenario.row, "J") & vbCrLf & shtScriptUsfos.Cells(scenario.row, "K")

    HeadFahtsHandle = FreeFile
    Open HeadFahtsPath For Output As #HeadFahtsHandle
    WriteHeadFileContent HeadFahtsHandle, scenario.row
    Close #HeadFahtsHandle

End Sub

Private Sub WriteHeadFileContent(ByVal fileNum As Integer, ByVal row As Long)
    ' Write the content to the file
    Print #fileNum, "HEAD             " & shtScriptFaths.Cells(row, "L")
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " TOLMERGE 0.0001"
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, "'            nlengt   nCIRC    Iel1  iel2 ... "
'                    Print #fileNum, " MESHPIPE       4      16                   "
    Print #fileNum, " MESHPIPE       2      8                   "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, "'       nlength n_side  n_top   n_bott  iel1  iel2  iel3  ...."
'                    Print #fileNum, " MESHIPRO  4      2       4        4"
    Print #fileNum, " MESHIPRO  2      2       4        4"
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, "         n_12  n_14  iel1  iel2  iel3  ...."
    Print #fileNum, "  MESHQUAD      2       2"
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " MOVFIRE  " & shtScriptFaths.Cells(row, "G") & "   " & shtScriptFaths.Cells(row, "H") & "   " & shtScriptFaths.Cells(row, "I")
    Print #fileNum, " KAMELEON  "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " Lin_Rule " & shtScriptFaths.Cells(row, "J")
    Print #fileNum, " ' "
    Print #fileNum, " ' TimeUnit   Minute     ! Minutes is the normal time unit in Safety"
    Print #fileNum, " ' "
    Print #fileNum, " '         end-time(s)   nstep   resinc(s) "
    Print #fileNum, " TEMPSIM      " & Round(shtScriptFaths.Cells(row, "C"), 1) & "   " & shtScriptFaths.Cells(row, "E") & "   " & shtScriptFaths.Cells(row, "F")
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " '           id    rho(kg/m3)  c (J/kgK)  k (W/mK)   emiss   convection "
    Print #fileNum, "  THERMPAR    1     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    2     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    3     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    4     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    5     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    6     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    7     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    8     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR    9     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR   10     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR   11     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR   12     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR   13     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, "  THERMPAR   14     7850.0      460.0      54.0      0.8       0.0 "
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " PUSHDOWN "
    Print #fileNum, ""
    Print #fileNum, "INSIDPAR    1   1.273   0.7"
    Print #fileNum, " ' "
    If shtScriptUsfos.Cells(row, "H") = "H60" Then
        Print #fileNum, "INSULPAR 1 1 1 0.2 15"  ' Modificado para os cenários mais longos de poça
    Else
        'Print #fileNum, "INSULPAR 1 1 5 0.8 15"   ' default modificado para todos os cenários
        Print #fileNum, "INSULPAR 1 1 1 0.2 15"   ' default modificado para todos os cenários
    End If
    Print #fileNum, " ' "
    Print #fileNum, " ' "
    Print #fileNum, " ' "

    Print #fileNum, shtScriptUsfos.Cells(row, "L") & ""
    If shtScriptUsfos.Cells(row, "M") <> "" Then
        Print #fileNum, "'                 * Exposure  *      * Passive Fire Protection *"
        Print #fileNum, "     '    GeoTyp   top  web  bott        top  web  bott         ID_gruop"
        Print #fileNum, " EXP_GRUP  IProf     3    3    3           2    3    3            1000"
        Print #fileNum, " ' "
        Print #fileNum, "'elementos com 3 lados (perfis I)"
        Print #fileNum, "GroupDef    1000    Elem     " & shtScriptUsfos.Cells(row, "M") & ""
    End If
    If shtScriptUsfos.Cells(row, "N") <> "" Then
        Print #fileNum, "'                top  Lside Rside bott   top  Lside Rside bott  "
        Print #fileNum, "' EXP_GRUP  Box    1    1      1    1      0     1     1     1     1001     'perfil box"
        Print #fileNum, " ' "
        Print #fileNum, "'elementos com 3 lados (perfis box)"
        Print #fileNum, "'GroupDef   1001    Elem    " & shtScriptUsfos.Cells(row, "N") & ""
    End If
    Close #fileNum

End Sub

Private Sub VerifyAndUpdateCriteria(ByVal iter As Integer)
    ' Logic to verify and update criteria
End Sub


Private Function CriteriaMet(ByRef iter As Long, row As Long) As Boolean
    ' Determine if the criteria are met
    ' Return True if criteria are met, False otherwise
    
    screening = "Screening PFP Phase" = shtScriptUsfos.Range("W29")
    
    If iter > 0 And screening Then CriteriaMet = True
    
    If iter = 500 Then CriteriaMet = True

    If shtScriptUsfos.Cells(row, "S") = "completed" Then CriteriaMet = True
        
End Function

Sub CopyStrainFiles(ByVal scenarioName, ByVal caseFolderPath As String, ByVal iter As Long, loadcase As Long)
        ' Create a directory for strain files if it doesn't exist
    Dim StrainDirectory As String
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim FSObject As Object
    Dim StrainFiles As String
    
    StrainFoldePath = caseFolderPath & "\StrainFiles"
    If Dir(StrainFoldePath, vbDirectory) = "" Then  ' Check if the directory already exists
        On Error Resume Next
        MkDir StrainFoldePath
        On Error GoTo 0
    End If

    ' Copy the strain files to the directory
    SourceFilePath = caseFolderPath & "\" & scenarioName & "_strain_output"
    DestinationFilePath = StrainFoldePath & "\" & scenarioName & "_strain_output_" & "iter_" & iter & "_LC_" & loadcase

    ' Copy strain files to the directory
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile SourceFilePath, DestinationFilePath, True
End Sub

Sub CopyStatusFiles(ByVal scenarioName, ByVal caseFolderPath As String, ByVal iter As Long, loadcase As Long)
    ' Create a directory for statusfiles if it doesn't exist
    Dim StrainDirectory As String
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim FSObject As Object
    Dim StrainFiles As String
    
    StatusFoldePath = caseFolderPath & "\StatusFiles"
    If Dir(StatusFoldePath, vbDirectory) = "" Then  ' Check if the directory already exists
        On Error Resume Next
        MkDir StatusFoldePath
        On Error GoTo 0
    End If

    ' Copy the status files to the directory

    SourceFilePath = caseFolderPath & "\" & scenarioName & "_usfos_output_status.text"
    DestinationFilePath = StatusFoldePath & "\" & scenarioName & "_usfos_output_status_" & "iter_" & iter & _
                        "_LC_" & loadcase & ".text"

    ' Copy strain files to the directory
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile SourceFilePath, DestinationFilePath, True
End Sub

Sub CopyFahtsOutFiles(ByVal scenarioName, ByVal caseFolderPath As String, ByVal iter As Long)
        ' Create a directory for strain files if it doesn't exist
    Dim StrainDirectory As String
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim FSObject As Object
    Dim StrainFiles As String
    
    FahtsFoldePath = caseFolderPath & "\FahtsOutFiles"
    If Dir(FahtsFoldePath, vbDirectory) = "" Then  ' Check if the directory already exists
        On Error Resume Next
        MkDir FahtsFoldePath
        On Error GoTo 0
    End If

    ' Copy the strain files to the directory
    SourceFilePath = caseFolderPath & "\" & scenarioName & "_output.out"
    DestinationFilePath = FahtsFoldePath & "\" & scenarioName & "_output_" & "iter_" & iter & ".out"

    ' Copy strain files to the directory
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile SourceFilePath, DestinationFilePath, True
End Sub

Sub Run_fahts_only()
    ' Updates the head file and executes the FAHTS program for each scenario.
    '

    Dim row As Long
    Dim ultimaLinha As Long
    Dim thisScenario As ScenarioType
    
    TurnOffCalcs
    
    ultimaLinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row
    
    AuxMacros.setCurrentFolder
    
    AuxMacros.CheckAndAlertProgram Program.FAHTS
    
    For row = 2 To ultimaLinha
        If LCase(shtScriptUsfos.Cells(row, "S")) <> "completed" Then
            thisScenario = InitScenariosAttributes(row)

            ' Update Fahts for on going cases
            UpdateHeadFahts thisScenario
                    
            ' Running Fahts
            ExecuteProgram Program.FAHTS, thisScenario.FolderPath, thisScenario.Name

        End If
    Next
    
    ' Show message that the process is done
    MsgBox "FAHTS process completed for all scenarios.", vbInformation, "Process Completed"

    TurnOnCalcs
End Sub

Sub WriteToLogFile(message As String)
    Dim logFilePath As String
    Dim logFile As Integer
    
    logFilePath = CURRENT_FOLDER & "\" & LOG_FILE_NAME
    logFile = FreeFile
    
    Open logFilePath For Append As #logFile
    Print #logFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & message
    Close #logFile
End Sub

Private Sub LogAlreadyProtectedWarning(ByVal elementNumber As Long, ByRef scenario As ScenarioType)
    WriteToLogFile "WARN: Element " & elementNumber & " (3-sided I) already protected but still above criteria - " & _
                   "Scenario " & scenario.Number & ": " & scenario.Name & " at Iteration: " & scenario.Iteration
End Sub

Private Sub LogError(ByVal errorMessage As String, ByRef scenario As ScenarioType)
    WriteToLogFile "ERROR: " & errorMessage & " - Scenario " & scenario.Number & ": " & scenario.Name & _
                   " at Iteration: " & scenario.Iteration
End Sub
  

Private Function checkErrorInFile(ByVal filePath As String, ByRef scenario As ScenarioType)
    Dim fileContent As String
    Dim line As String
    Dim lineNumber As Long
    Dim errorMessage As String
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber

    lineNumber = 0
    errorMessage = ""

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        lineNumber = lineNumber + 1

        ' Check for specific error patterns in the line
        ' Check for "ERROR" keyword
        If InStr(UCase(line), "ERROR") > 0 Then
            errorMessage = "Error found"
            Exit Do
        ElseIf InStr(UCase(line), "F R A C T U R E                 F A I L U R E") > 0 Then
            errorMessage = "Fracture Failure"
            Exit Do
        End If
    Loop

    Close #fileNumber

    checkErrorInFile = errorMessage
End Function

Sub handleError(filePath As String, ByRef scenario As ScenarioType, Optional ByVal errorMessage As String)
    ' Handle error scenario
    LogError errorMessage, scenario
    
    ' Write error in column U
    shtScriptUsfos.Cells(scenario.row, "U").Value = "ERROR"
End Sub

Sub LogElement(ByVal elementNumber As Long, ByRef scenario As ScenarioType)
    WriteToLogFile "INFO: Element " & elementNumber & " protected - Scenario " & scenario.Number & ": " & scenario.Name & _
                   " at Iteration: " & scenario.Iteration
End Sub

Sub testCheckErrorInFile()
    Dim scenario As ScenarioType
    Dim filePath As String
    
    ' Initialize a test scenario
    scenario.Number = 21
    scenario.Name = "TestScenario"
    scenario.Iteration = 1
    scenario.FolderPath = "D:\LUIMAR\SBM\TS063\Teste_AtTime_UptoTime\Automatizado_Planilha_UpToTime\21_L06_seg_TS063_L_Pool_P1_tests"
    
    ' Set the file path to check
    filePath = scenario.FolderPath & "\L06_seg_TS063_L_Pool_P1_usfos_output.out"
    
    CURRENT_FOLDER = "D:\LUIMAR\SBM\TS063\Teste_AtTime_UptoTime"
    
    ' Call the function to check for errors in the file
    checkErrorInFile filePath, scenario
End Sub

Sub run2()
Run_fahts_only
Run_analysis
End Sub





