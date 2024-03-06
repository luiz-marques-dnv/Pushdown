Const HEAD_DEFAULT_FILE As String = "\usfos_head_default.fem"
Const STRASPAS As String = """"   'defining aspas to be inside strings

Public Enum Program
    FAHTS
    USFOS
    Strain
    SRAINNeg
End Enum

Public Type ElementType
    Number As Long
    Strain As Double

End Type
Sub Run_analysis()
    ' Initialization
    Dim ultimaLinha As Long
    Dim iter As Long
    
    ' Sub to Disable calculations and screen updating
    TurnOnCalcs
    
    ' Check if necessary programs are on correct path
    AuxMacros.CheckPrograms
    
    'set current directory to public variable CURRENT_FOLDER
    ' TODO: Set default variables
    AuxMacros.setCurrentFolder
    
    ultimaLinha = shtScriptUsfos.Cells(Rows.Count, "A").End(xlUp).row
    iter = InitializeIteration()
    
    ' Main processing loop
    Do Until CriteriaMet(iter)
        iter = iter + 1
        ' Update iteration number in the worksheet
        shtScriptUsfos.Cells(17, 23) = iter
        
        
        
        ProcessIteration iter, ultimaLinha
    Loop

    ' Finalization
    MsgBox "Usfos/Faths Analysis Finished!"
    
    ' Sub to Enable calculations and screen updating
    TurnOnCalcs
End Sub

Private Function InitializeIteration() As Integer
    If shtScriptUsfos.Cells(17, 23) = "" Then
        InitializeIteration = 0
        
    Else
        InitializeIteration = shtScriptUsfos.Cells(17, 23)
    End If
    
End Function

Private Sub ProcessIteration(iter As Long, ultimaLinha As Long)
    Dim row As Long
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim BeltempFilePath As String
    Dim loadCaseData As Collection
    Dim loadCaseTimeDetails As Variant
    Dim loadcase As Long
    Dim LoadCaseTime As Double
    Dim pfpEligibleElement As ElementType
    Dim HighestStrain As Double
    Dim HighestElementStrain As Long
    Dim shellCommandUsfos As String
    Dim shellCommandStrain As String
    Dim shellCommandStrainNeg As String
    Dim UsfosShell As Variant
    Dim StrainShell As Variant
    Dim StrainNegShell As Variant
    Dim HighestPFPStrain As Double
    Dim HighestElementStrain As Long
    Dim ElementsPositiveCriteria As Scripting.Dictionary
    
    If iter = 1 Then
        'Defining commands paths to be run
        shellCommandUsfos = """" & CURRENT_FOLDER & "\" & "script_rodar_USFOS.bat"
        shellCommandStrain = """" & CURRENT_FOLDER & "\" & "script_rodar_STRAIN.bat" & """"
        shellCommandStrainNeg = """" & CURRENT_FOLDER & "\" & "script_rodar_STRAIN_neg.bat" & """"
        
        ' Running Usfos
        UsfosShell = ExecuteShellCommand(shellCommandUsfos)

        ' Running strain and Straing neg to get values outside criteria
        StrainShell = ExecuteShellCommand(shellCommandStrain)
        StrainNegShell = ExecuteShellCommand(shellCommandStrainNeg)
        
        For row = 2 To ultimaLinha
             ' Process strain criteria
            pfpEligibleElement = ProcessStrainCriteria(row)
    
            ' TODO: hange thi condition handling
            If pfpEligibleElement.Number <> 0 Then
                ' Update Completion Status
                shtScriptUsfos.Cells(row, "S") = "on going"
            
    
            ElseIf pfpEligibleElement.Number = 0 Then
                shtScriptUsfos.Cells(row, "S") = "completed"
            End If
        Next row

        ' Create a folder and Copy first Run files
        CopyFirstRunFiles ultimaLinha
    
        
    ElseIf iter > 1 Then
        For row = 2 To ultimaLinha
            If Not AuxMacros.IsCompleted(row) Then
                ' Initialize values from the sheet
                'TODO: Change this to a class
                InitializeCaseValues row, scenarioName, caseFolderPath, BeltempFilePath
                
                ' Reset values for new scenario
                HighestStrain = 0
                HighestElementStrain = 0
    
    
                ' Extract the load case data from the Beltemp file
                Set loadCaseData = ExtractLoadCaseData(BeltempFilePath)
                
                startTime = Timer
                For Each loadCaseTimeDetails In loadCaseData
                    loadcase = loadCaseTimeDetails(0)
                    LoadCaseTime = loadCaseTimeDetails(1)

                    
                    ' Update Head USFOS file
                    UpdateHeadUsfos row, LoadCaseTime
    
                    ' Running Usfos
                    ExecuteProgram Program.USFOS, caseFolderPath, scenarioName
    
                    ' Running strain and Straing neg to get values outside criteria
                    ExecuteProgram Program.Strain, caseFolderPath, scenarioName
                    'ExecuteProgram Program.SRAINNeg, caseFolderPath, scenarioName
                    
                    ' Process strain criteria
                    Set ElementsAboveCriteria = ProcessStrainCriteria(scenarioName, caseFolderPath, strainCriteria)

                    ' Get the highest strain considering that ElementsAboveCriteria is sorted from highest to lowest strain
                    HighestStrainCriteria = ElementsAboveCriteria.Items(0)

                    ' Find the highest strain element eligible for PFP (critical element)
                    pfpEligibleElement = FindPFPEligibleElement(ElementsPositiveCriteria)
                    
                    ' Read elements above buckling criteria
                    ProcessBucklingCriteria iter
    
                    ' Update the highest eligeble strain and element if the current strain is higher
                    If pfpEligibleElement.Strain > HighestPFPStrain Then
                        HighestPFPStrain = pfpEligibleElement.Strain
                        HighestElementStrain = pfpEligibleElement.Number
                    End If

                    ' Update the highest strain and time if the current strain is higher
                    if HighestStrainCriteria > HighestStrain Then
                        Highestrain = HighestStrainCriteria
                        HighestStrainTime = LoadCaseTime
                    End If
                    
                    'Copying starin files just to check    TODO: Delete after checking
                    CopyStrainFiles scenarioName, caseFolderPath, iter, loadcase
NextIteration:
                Next loadCaseTimeDetails
                

                If HighestStrain <> 0 Then
                    ' Write PFP element to the script file
                    writePFPElement row, HighestElementStrain
    
                    ' Update Fahts for on going cases
                    UpdateHeadFahts row
                    
                    ' Running Fahts
                    ExecuteProgram Program.FAHTS, caseFolderPath, scenarioName
                    
                    CopyFahtsOutFiles scenarioName, caseFolderPath, iter
    
                    ' Update Completion Status
                    shtScriptUsfos.Cells(row, "S") = "on going"

                ' If the highest strain is 0, then the criteria was met. Finally update to the case with highest strain
                ElseIf HighestElementStrain = 0 Then
                    ' Update Head USFOS file
                    UpdateHeadUsfos HighestStrainTime, scenarioName, caseFolderPath

                    ' Running Usfos
                    ExecuteProgram Program.USFOS, caseFolderPath, scenarioName
                    
                End If
                
            End If
        Next row
    End If

    ' Verify and update criteria
    VerifyAndUpdateCriteria iter

    ' Run Faths if criteria not met
    If Not CriteriaMet(iter) Then
        'UpdateAndRunFahts iter
    End If
End Sub

Private Function ProcessStrainCriteria(ByVal row As Long) As ElementType
    Dim FileStrain As String
    Dim FileStrainNeg As String
    Dim strainCriteria As Double
    Dim ElementsPositiveCriteria As Scripting.Dictionary
    Dim ElementsNegativeCriteria As Scripting.Dictionary
    Dim pfpEligibleElement As ElementType
    'Dim ElementsNegativeCriteria As Collection
    
    strainCriteria = shtScriptUsfos.Cells(row, "I")

    ' Common setup for strain and strain_neg
    StrainsFilePaths row, FileStrain, FileStrainNeg

    ' Process positive strain criteria
    Set ElementsPositiveCriteria = GetElementsAboveCriteriaStrain(FileStrain, strainCriteria)
    
    ' Process negative strain criteria                                                              TODO: make negative criteria an option
    'set ElementsNegativeCriteria = GetElementsAboveCriteriaStrain FileStrainNeg, StrainCriteria*-1

    
    'pfpEligibleElement ElementsPositiveCriteria, ElementsNegativeCriteria                          TODO: make negative criteria an option
    pfpEligibleElement = FindPFPEligibleElement(ElementsPositiveCriteria)
    
    ' Upadate Complete or on going status
    'UpdateCompletionStatus row, PFPElement

    ProcessStrainCriteria = pfpEligibleElement
End Function

Private Sub StrainsFilePaths(ByVal row As Long, ByRef FileStrain As String, ByRef FileStrainNeg As String)
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim FilenameStrain As String
    Dim FilenameStrain_neg As String

    scenarioName = shtScriptUsfos.Cells(row, "B")
    FilenameStrain = scenarioName & "_strain_output"
    FilenameStrain_neg = scenarioName & "_strain_neg_output"
    caseFolderPath = CURRENT_FOLDER & "\" & shtScriptUsfos.Cells(row, "A") & "_" & scenarioName
    FileStrain = caseFolderPath & "\" & FilenameStrain
    FileStrainNeg = caseFolderPath & "\" & FilenameStrain_neg
End Sub
Private Function GetElementsAboveCriteriaStrain(ByVal filePath As String, strainCriteria) As Scripting.Dictionary
    ' Process the strain file (either positive or negative) and returns a dinamic array of elements with number and strain that have a strain above the criteria
    ' It will overwrites the array if high loadcase is found
    
    Dim strainFile As Integer
    Dim lineFromFile As String
    Dim elementLine As String
    Dim elementLine_Array() As String
    Dim elementsStrain As Scripting.Dictionary

    Set elementsStrain = New Scripting.Dictionary

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
                ElementNumber = elementLine_Array(0)
                ElementStrain = elementLine_Array(1)
                If Abs(ElementStrain) > strainCriteria Then
                    elementsStrain.Add ElementNumber, ElementStrain
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

Private Function FindPFPEligibleElement(ByRef elementsStrain As Scripting.Dictionary) As ElementType
    'Verificar se elemento é elegível para receber PFP e retorna o numero do elemento elegível
    Dim test_strain As Long
    Dim Element As ElementType
    Dim i As Long
    
    If elementsStrain.Count > 0 Then
        For i = 0 To elementsStrain.Count - 1
            Element.Number = elementsStrain.Keys(i)
            Element.Strain = elementsStrain.Items(i)
            On Error Resume Next
            test_strain = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(Element.Number, shtGenie.Range("A2:A50000"), 1, False), "N")
            If Err = 0 Then
                  FindPFPEligibleElement = Element
                  Exit Function
            End If

        Next
    End If
End Function

Private Sub writePFPElement(ByVal row As Long, ByVal PFPElementNumber As Long)
    ' refactor all this code to separate the PFP 3-sided I and box profile finding and writing
    Dim ultlinhaPFP As Long
    Dim verify As Boolean
    Dim k As Long
    Dim test_3sided_I As String
    Dim test_3sided_box As String

   If PFPElementNumber <> 0 Then
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

        'Verificar se elemento é elegível para receber PFP 3-sided I profile
        On Error Resume Next
        test_3sided_I = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(ShtStrain.Cells(PFPElementNumber, "B"), shtGenie.Range("C2:C50000"), 1, False), "N")

        If Err = 0 Then
            shtScriptUsfos.Cells(row, "M") = shtScriptUsfos.Cells(row, "M") & "     " & test_3sided_I
            ultlinhaPFP = shtPFP3sided.Cells(Rows.Count, (row - 2) * 5 + 1).End(xlUp).row

            If verify = True Then
                shtPFP3sided.Cells(ultlinhaPFP + 1, (row - 2) * 5 + 3) = test_3sided_I
            End If
        End If
        
        
        'Verificar se elemento é elegível para receber PFP 3-sided box profile
        On Error Resume Next
        test_3sided_box = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(ShtStrain.Cells(PFPElementNumber, "B"), shtGenie.Range("D2:D50000"), 1, False), "N")

        If Err = 0 Then
            shtScriptUsfos.Cells(row, "N") = shtScriptUsfos.Cells(row, "N") & "     " & test_3sided_box
            ultlinhaPFP = shtPFP3sided.Cells(Rows.Count, (row - 2) * 5 + 1).End(xlUp).row

            If verify = True Then
                shtPFP3sided.Cells(ultlinhaPFP + 1, (row - 2) * 5 + 3) = test_3sided_box
            End If
        End If
    End If
End Sub

Private Sub ProcessBucklingCriteria(ByVal iter As Integer)
    ' Logic to read elements above buckling criteria
End Sub


Private Sub UpdateHeadUsfos(ByVal row As Long, ByVal LoadCaseTime As String)
    ' Updates the Head USFOS file for each row in the specified range.
    ' TODO: CHange the fucntion to receive the files path and not the row
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim HeadFilename As String
    Dim HeadDefaultFile As String
    Dim HeadCaseFile As String

    Dim currentLoadCaseTime As Double
    Dim FSObject As Object

    ' Initialize values from the sheet
    scenarioName = shtScriptUsfos.Cells(row, "B")
    caseFolderPath = CURRENT_FOLDER & "\" & shtScriptUsfos.Cells(row, "A") & "_" & scenarioName

    ' Define the paths to Head and Beltemp files
    HeadFilename = "usfos_head_" & scenarioName
    HeadDefaultFile = CURRENT_FOLDER & HEAD_DEFAULT_FILE
    HeadCaseFile = caseFolderPath & "\" & HeadFilename & ".fem"
    
    ' Copy and modify head file
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile HeadDefaultFile, HeadCaseFile, True
    
    ' Update the head file
    ModifyHeadFile HeadFile:=HeadCaseFile, LoadCaseTime:=LoadCaseTime
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

Private Sub CopyCaseFile(ByVal sourcePath As String, ByVal destinationPath As String, ByVal fileName As String)
    ' Copy a file from the source path to the destination path and overwites if it already exists.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim SourceFile As String
    Dim DestinationFile As String
    
    SourceFile = sourcePath & "\" & scenarioName & fileName
    DestinationFile = destinationPath & "\" & scenarioName & fileName
    
    On Error Resume Next
    fso.CopyFile SourceFile, DestinationFile  ' Copy source to target and overwrite if the file exists.
    If Err.Number <> 0 Then
        Debug.Print "Source:" & SourceFile
        'MsgBox "An error occurred while copying the file " & fileName & ". Error: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Sub


Private Function ExecuteShellCommand(ByVal command As String) As Variant
    On Error GoTo ErrorHandler

    Dim shellObject As Object
    Set shellObject = CreateObject("WScript.Shell")
    ExecuteShellCommand = shellObject.Run(command, 1, True)

     Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ExecuteShellCommand = Nothing
End Function

Private Sub ExecuteProgram(ByVal programName As Program, ByVal workingDirectory As String, ByVal scenarioName As String)
    Dim executable As String
    Dim outputFileName As String
    Dim inputFileName As String
    Dim usfosBin As String
    

    usfosBin = shtScriptUsfos.Range("W1")

    Select Case programName
        Case Program.FAHTS
            executable = "..\Fahts\fahts_v670.exe"
            inputFileName = "input.inp"

        Case Program.USFOS
            executable = usfosBin & "\Usfos.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_usfos.trm"
            inputFileName = "usfos_input.inp"

        Case Program.Strain
            executable = usfosBin & "\strain.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain.trm"
            inputFileName = "strain_input.inp"

        Case Program.SRAINNeg
            executable = usfosBin & "\strain_neg.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain_neg.trm"
            inputFileName = "strain_neg_input.inp"

        Case Else
            MsgBox "Invalid Program Name when calling ExecuteProgram sub"
            Exit Sub
    End Select

    Dim fullCommand As String
    fullCommand = """" & executable & """" & executionCommand & outputFileName & " < " & inputFileName
    Debug.Print (fullCommand)
    ExecuteShellCommandHidden fullCommand, workingDirectory
End Sub


Private Function ExecuteShellCommandHidden(ByVal command As String, workingDirectory As String) As Variant
    On Error GoTo ErrorHandler

    Dim shellObject As Object
    Set shellObject = CreateObject("WScript.Shell")

    ' Construct the command with proper quoting
    Dim fullCommand As String
    fullCommand = "cmd.exe /c " & command

    ' Change the working directory. Same as cd command
    shellObject.CurrentDirectory = workingDirectory

    ' Execute the command
    ExecuteShellCommandHidden = shellObject.Run(fullCommand, 7, True)

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ExecuteShellCommandHidden = Nothing
End Function


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

Private Sub ModifyHeadFile(ByVal HeadFile As String, ByVal LoadCaseTime As Double)
    ' Modify the head file based on the case information
    Dim HeadFileHandle As Long
    
    HeadFileHandle = FreeFile
    Open HeadFile For Input As HeadFileHandle
    HeadFileContent = Input(LOF(HeadFileHandle), HeadFileHandle)
    Close #HeadFileHandle
    
    ' Replace the specific string
    HeadFileContent = Replace(HeadFileContent, "'###PUSHDOWN COMMAND###", "Pushdown AtTime " & LoadCaseTime)

    ' Open the file for writing
    HeadFileHandle = FreeFile
    Open HeadFile For Output As HeadFileHandle
    Print #HeadFileHandle, HeadFileContent
    Close #HeadFileHandle
End Sub

Sub UpdateHeadFahts(ByVal row As Long)
    Dim scenarioName As String
    Dim caseFolderPath As String
    Dim HeadFahtsPath As String

    scenarioName = shtScriptUsfos.Cells(row, "B")
    caseFolderPath = CURRENT_FOLDER & "\" & shtScriptUsfos.Cells(row, "A") & "_" & scenarioName
    
    HeadFahtsPath = caseFolderPath & "\head_" & scenarioName & ".fem"

    shtScriptUsfos.Cells(row, "L") = shtScriptUsfos.Cells(row, "J") & vbCrLf & shtScriptUsfos.Cells(row, "K")

    HeadFahtsHandle = FreeFile
    Open HeadFahtsPath For Output As #HeadFahtsHandle
    WriteHeadFileContent HeadFahtsHandle, row
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


Private Function CriteriaMet(iter) As Boolean
    ' Determine if the criteria are met
    ' Return True if criteria are met, False otherwise
    
    screening = "Screening PFP Phase" = shtScriptUsfos.Range("W29")
    
    If iter > 0 And screening Then CriteriaMet = True
    
    If iter = 50 Then CriteriaMet = True
    
    
End Function

Private Sub InitializeCaseValues(ByVal row As Long, ByRef scenarioName, ByRef caseFolderPath As String, ByRef BeltempFilePath As String)
    Dim BeltempFilename As String

    ' Initialize and fetch values from the worksheet
    scenarioName = shtScriptUsfos.Cells(row, "B")
    
    caseFolderPath = CURRENT_FOLDER & "\" & shtScriptFaths.Cells(row, "A") & "_" & scenarioName
    BeltempFilename = scenarioName & "_output_beltemp"
    BeltempFilePath = caseFolderPath & "\" & BeltempFilename & ".fem"
End Sub

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




