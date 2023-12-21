Const HEAD_DEFAULT_FILE As String = "\usfos_head_default.fem"
Const STRASPAS As String = """"   'defining aspas to be inside strings

Public Enum Program
    FAHTS
    USFOS
    Strain
    SRAINNeg
End Enum

Private Type ElementType
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
    Dim FahtsOutPath As String
    Dim FahtsData As SimulationData
    Dim BeltempData As Scripting.Dictionary
    Dim BeltempTimes As Collection
    Dim BeltempTemperatures As Collection
    Dim HighestElementStrain As Long
    Dim strainCriteria As Double
    Dim pfpEligibleElement As ElementType
    Dim filteredloadCaseData As Collection
    Dim remainingloadCaseData As Collection

    For row = 2 To ultimaLinha
        If Not AuxMacros.IsCompleted(row) Then
            ' Initialize values from the sheet
            'TODO: Change this to a class
            InitializeCaseValues row, scenarioName, caseFolderPath, FahtsOutPath
            
            ' Get strain criteria from sheet
            strainCriteria = shtScriptUsfos.Cells(row, "I")
    
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
                
                
                ' Process strain criteria
                pfpEligibleElement = ProcessStrainCriteria(scenarioName, caseFolderPath, strainCriteria)
        
                ' TODO: hange thi condition handling
                If pfpEligibleElement.Number <> 0 Then
                    ' Update Completion Status
                    shtScriptUsfos.Cells(row, "S") = "on going"
                
        
                ElseIf pfpEligibleElement.Number = 0 Then
                    shtScriptUsfos.Cells(row, "S") = "completed"
                End If
                

                ' Create a folder and Copy first Run files
                CopyFirstRunFiles ultimaLinha
            
                
            ElseIf iter > 1 Then
                ' Retrievieng some information from fahts output
                FahtsData = ExtractDataFromFahtsOut(FahtsOutPath)
                
                ' Storing Beltemp Times and Temperatures
                Set BeltempData = GetBeltempData(FahtsData)
                Set BeltempTimes = BeltempData("Times")
                Set BeltempTemperatures = BeltempData("Temperatures")

                ' Get a subset from loadCaseData so less load cases are processed
                Set FilteredTimes = GetFilteredTimes(BeltempTimes, BeltempTemperatures)


                HighestElementStrain = ExecuteCriteriasCheck(FilteredTimes, scenarioName, caseFolderPath, iter, strainCriteria)

                ' Check if the filtered load case data was enough to find a PFP element
                If HighestElementStrain = 0 Then

                    ' Get the remaining load case data
                    Set RemainingFilteredTime = GetRemainingFilteredTime(BeltempTimes, FilteredTimes)

                    HighestElementStrain = ExecuteCriteriasCheckAll(RemainingFilteredTime, scenarioName, caseFolderPath, iter, strainCriteria)
                End If

                If HighestElementStrain <> 0 Then
                    ' Write PFP element to the script file
                    writePFPElement row, HighestElementStrain
    
                    ' Update Fahts for on going cases
                    UpdateHeadFahts row
                    
                    ' Running Fahts
                    ExecuteProgram Program.FAHTS, caseFolderPath, scenarioName
                    
                    CopyFahtsOutFiles scenarioName, caseFolderPath, iter
    
                    ' Update Completion Status
                    shtScriptUsfos.Cells(row, "S") = "on going"
    
                ElseIf HighestStrain = 0 Then
                    shtScriptUsfos.Cells(row, "S") = "completed"
                End If

            End If
        End If
    Next row
    ' Verify and update criteria
    VerifyAndUpdateCriteria iter

    ' Run Faths if criteria not met
    If Not CriteriaMet(iter) Then
        'UpdateAndRunFahts iter
    End If
End Sub

Private Function ExecuteCriteriasCheck(BeltempTimes, scenarioName, caseFolderPath, iter, strainCriteria) As Long
    Dim time As Variant
    Dim loadcase As Long
    Dim pfpEligibleElement As ElementType
    Dim HighestStrain As Double
    Dim HighestElementStrain As Long
    
    For Each time In BeltempTimes
        ' Update Head USFOS file
        UpdateHeadUsfos time, scenarioName, caseFolderPath

        ' Running Usfos
        ExecuteProgram Program.USFOS, caseFolderPath, scenarioName

        ' Running strain and Straing neg to get values outside criteria
        ExecuteProgram Program.Strain, caseFolderPath, scenarioName
        'ExecuteProgram Program.SRAINNeg, caseFolderPath, scenarioName
        
        ' Process strain criteria
        pfpEligibleElement = ProcessStrainCriteria(scenarioName, caseFolderPath, strainCriteria)
        
        ' Read elements above buckling criteria
        ProcessBucklingCriteria iter

        If pfpEligibleElement.Strain <> 0 Then
            HighestStrain = pfpEligibleElement.Strain
            HighestElementStrain = pfpEligibleElement.Number

            'Copying starin files just to check    TODO: Delete after checking
            CopyStrainFiles scenarioName, caseFolderPath, iter, loadcase

            ExecuteCriteriasCheck = pfpEligibleElement.Number

            Exit Function
        End If
    Next time

    ExecuteCriteriasCheck = pfpEligibleElement.Number

End Function

Private Function ExecuteCriteriasCheckAll(BeltempTimes, scenarioName, caseFolderPath, iter, strainCriteria) As Long
    For Each time In BeltempTimes
        ' Update Head USFOS file
        UpdateHeadUsfos time, scenarioName, caseFolderPath

        ' Running Usfos
        ExecuteProgram Program.USFOS, caseFolderPath, scenarioName

        ' Running strain and Straing neg to get values outside criteria
        ExecuteProgram Program.Strain, caseFolderPath, scenarioName
        'ExecuteProgram Program.SRAINNeg, caseFolderPath, scenarioName
        
        ' Process strain criteria
        pfpEligibleElement = ProcessStrainCriteria(scenarioName, caseFolderPath, strainCriteria)
        
        ' Read elements above buckling criteria
        ProcessBucklingCriteria iter

        If pfpEligibleElement.Strain > HighestStrain Then
            HighestStrain = pfpEligibleElement.Strain
            HighestElementStrain = pfpEligibleElement.Number
        End If
        
        'Copying starin files just to check    TODO: Delete after checking
        CopyStrainFiles scenarioName, caseFolderPath, iter, loadcase

    Next time

    ExecuteCriteriasCheckAll = HighestElementStrain

End Function

Private Function ProcessStrainCriteria(ByVal scenarioName As String, ByVal caseFolderPath As String, ByVal strainCriteria As Double) As ElementType
    Dim FileStrain As String
    Dim FileStrainNeg As String
    Dim ElementsPositiveCriteria As Scripting.Dictionary
    Dim ElementsNegativeCriteria As Scripting.Dictionary
    Dim pfpEligibleElement As ElementType
    'Dim ElementsNegativeCriteria As Collection

    ' Common setup for strain and strain_neg
    StrainsFilePaths scenarioName, caseFolderPath, FileStrain, FileStrainNeg

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

Private Sub StrainsFilePaths(ByVal scenarioName As String, ByVal caseFolderPath As String, ByRef FileStrain As String, ByRef FileStrainNeg As String)
    Dim FilenameStrain As String
    Dim FilenameStrain_neg As String

    FilenameStrain = scenarioName & "_strain_output"
    FilenameStrain_neg = scenarioName & "_strain_neg_output"
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
            shtScriptUsfos.Cells(i, "M") = shtScriptUsfos.Cells(i, "M") & "     " & test_3sided_I
            ultlinhaPFP = shtPFP3sided.Cells(Rows.Count, (i - 2) * 5 + 1).End(xlUp).row

            If verify = True Then
                shtPFP3sided.Cells(ultlinhaPFP + 1, (i - 2) * 5 + 3) = test_3sided_I
            End If
        End If
        
        
        'Verificar se elemento é elegível para receber PFP 3-sided box profile
        On Error Resume Next
        test_3sided_box = Application.WorksheetFunction.IfNa(WorksheetFunction.VLookup(ShtStrain.Cells(PFPElementNumber, "B"), shtGenie.Range("D2:D50000"), 1, False), "N")

        If Err = 0 Then
            shtScriptUsfos.Cells(i, "N") = shtScriptUsfos.Cells(i, "N") & "     " & test_3sided_box
            ultlinhaPFP = shtPFP3sided.Cells(Rows.Count, (i - 2) * 5 + 1).End(xlUp).row

            If verify = True Then
                shtPFP3sided.Cells(ultlinhaPFP + 1, (i - 2) * 5 + 3) = test_3sided_box
            End If
        End If
    End If
End Sub

Private Sub ProcessBucklingCriteria(ByVal iter As Integer)
    ' Logic to read elements above buckling criteria
End Sub


Private Sub UpdateHeadUsfos(ByVal LoadCaseTime As Long, ByVal scenarioName As String, ByVal caseFolderPath As String)
    ' Updates the Head USFOS file for each row in the specified range.
    ' TODO: CHange the fucntion to receive the files path and not the row

    Dim HeadFilename As String
    Dim HeadDefaultFile As String
    Dim HeadCaseFile As String

    Dim currentLoadCaseTime As Double
    Dim FSObject As Object


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
    Dim executable, outputFileName, inputFileName, usfosBinFolder, fullCommand As String
    Dim ShellResult As Variant
    Dim InitShellTimer, elapsedShellTime As Double
    Dim userResponse As Variant
    
    ' Get the path to the usfos bin folder
    usfosBinFolder = shtScriptUsfos.Range("W1")

    Select Case programName
        Case Program.FAHTS
            executable = "..\Fahts\fahts_v670.exe"
            inputFileName = "input.inp"
        
        Case Program.USFOS
            executable = usfosBinFolder & "\Usfos.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_usfos.trm"
            inputFileName = "usfos_input.inp"

        Case Program.Strain
            executable = usfosBinFolder & "\strain.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain.trm"
            inputFileName = "strain_input.inp"

        Case Program.SRAINNeg
            executable = usfosBinFolder & "\strain_neg.exe"
            executionCommand = " 400 > "
            outputFileName = scenarioName & "_strain_neg.trm"
            inputFileName = "strain_neg_input.inp"

        Case Else
            MsgBox "Invalid Program Name when calling ExecuteProgram sub"
            Exit Sub
    End Select

    
    fullCommand = """" & executable & """" & executionCommand & outputFileName & " < " & inputFileName
    
    ' Initiate time to check if the program will execute normally.
    ' This is an alternative choice since getting the output from shell would need much more unnecessary effort
    InitShellTimer = Timer
    
    ' Use shell in hidden mode to execute a command
    ShellResult = ExecuteShellCommandHidden(fullCommand, workingDirectory)
    
    ' Retrieve elepased time
    elapsedShellTime = Timer - InitShellTimer
    
    ' Check if shell executed normally
    If elapsedShellTime < 3 Then
        ' The shell execution failured
        userResponse = MsgBox("It is not possible to execute " & executable & vbNewLine & "Check license or path. Would you like to continue anyway?", vbCritical + vbYesNo)

        If userResponse = vbNo Then
            ' User chose to exit
            End
        ElseIf userResponse = vbYes Then
            ' User chose to continue, you can add code here to handle this situation
            Exit Sub
        End If
    End If
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

Private Sub InitializeCaseValues(ByVal row As Long, ByRef scenarioName, ByRef caseFolderPath As String, ByRef FahtsOutPath As String)
    ' Initialize and fetch values from the worksheet
    scenarioName = shtScriptUsfos.Cells(row, "B")
    
    caseFolderPath = CURRENT_FOLDER & "\" & shtScriptFaths.Cells(row, "A") & "_" & scenarioName

    FahtsOutPath = caseFolderPath & "\" & scenarioName & "_output.out"
 
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

' Function to extract data from the file
Function ExtractDataFromFahtsOut(filePath As String) As SimulationData
    Dim fileNum As Integer
    Dim currentLine As String
    Dim data As SimulationData
    Dim record As StepRecord
    Dim TrimLine As String
    Dim parts() As String
    Dim i As Integer


    fileNum = FreeFile()
    Open filePath For Input As #fileNum

    Do Until EOF(fileNum)
        Line Input #fileNum, currentLine

        ' Retrive Number of time steps from the file
        If InStr(currentLine, "Number of time steps") > 0 Then
            ' remove extra spaces
            TrimLine = WorksheetFunction.Trim(currentLine)
            data.NumberOfTimeSteps = CDbl(Split(TrimLine, ":")(1))

        ' Retrive Beltemp time steps from the file
        ElseIf InStr(currentLine, "Time between *BELTEMP*") > 0 Then
            ' remove extra spaces
            TrimLine = WorksheetFunction.Trim(currentLine)
            data.BeltempTimeInterval = CDbl(Split(TrimLine, ":")(1))

        ' Retrive total simulation time from the file
        ElseIf InStr(currentLine, "Total simulation time") > 0 Then
            ' remove extra spaces
            TrimLine = WorksheetFunction.Trim(currentLine)
            data.TotalSimTime = CDbl(Split(TrimLine, ":")(1))

        ' Retrive the data for each time step
        ElseIf InStr(currentLine, "(s)      (J)          (W/m2)     (degr C)    (degr C)") > 0 Then
            Line Input #fileNum, currentLine
            Line Input #fileNum, currentLine
            Line Input #fileNum, currentLine
            ' remove extra spaces
            TrimLine = WorksheetFunction.Trim(currentLine)
            While Len(TrimLine) > 1
                ' Split the line into different parts
                parts = Split(TrimLine, " ")
                ' Get the step number
                record.StepNo = CInt(parts(0))
                ' Get the time in seconds
                record.TimeInSeconds = CDbl(parts(1))
                ' Get the accumulated energy
                record.AccumulatedEnergy = CDbl(parts(2))
                ' Get the max heat input
                record.MaxHeatInput = CDbl(parts(3))
                ' Get the max temperature
                record.MaxTemp = CDbl(parts(4))
                ' Get the min temperature
                record.MinTemp = CDbl(parts(5))

                ' Add the record to the data
                ReDim Preserve data.StepData(i)
                data.StepData(i) = record
                i = i + 1
                Line Input #fileNum, currentLine
                TrimLine = WorksheetFunction.Trim(currentLine)
            Wend
        End If
    Loop

    Close #fileNum

    ExtractDataFromFahtsOut = data
End Function

' Function to calculate beltemp step
Function CalculateBeltempStep(NumberOfTimeSteps As Double, TotalSimTime As Double, BeltempTimeInterval As Double) As Long
    CalculateBeltempStep = CInt(Application.WorksheetFunction.RoundUp(NumberOfTimeSteps * BeltempTimeInterval / TotalSimTime, 0))
End Function

' Function to get simulation info
Function GetSimulationInfo(data As SimulationData) As Variant
    Dim result(2) As Double
    result(0) = data.TotalSimTime
    result(1) = data.NumberOfTimeSteps
    result(2) = data.BeltempTimeInterval
    GetSimulationInfo = result
End Function

' Function to get step temperature
Function GetTemperatureFromTime(data As SimulationData, time As Double) As Double
    ' using linear search
    ' TODO: use binary search
    Dim timeStep As Integer
    For i = 0 To UBound(data.StepData)
        If data.StepData(i).TimeInSeconds = time Then
            GetTemperatureFromTime = data.StepData(i).MaxTemp
            Exit Function
        End If
End Function

' Beltemp steps are subset of Fahts steps
Function IsBeltempStep(FahtsStep As Long, BeltempStep As Long) As Boolean
    IsBeltempStep = (FahtsStep Mod BeltempStep = 0)
End Function

Function GetBeltempData(FahtsData As SimulationData) As Scripting.Dictionary
    ' Main function to get beltemp temperatures from Fahst output data
    ' Beltemp are presented in different steps than fahts so it is necessary to convert
    Dim BeltempStep As Long
    Dim lastFahtsStep As Long
    Dim i As Long
    Dim BeltempData As New Scripting.Dictionary
    Dim BeltempTemperatures As New Collection
    Dim BeltempTimes As New Collection

    ' Calculate the beltmp step to know ich values from fahts are considered
    BeltempStep = CalculateBeltempStep(FahtsData.NumberOfTimeSteps, FahtsData.TotalSimTime, FahtsData.BeltempTimeInterval)
    
    lastFahtsStep = UBound(FahtsData.StepData) + 1
 
    For i = 1 To lastFahtsStep
    ' Fahts Initializes in step one
        If IsBeltempStep(i, BeltempStep) Then
            ' Add data to collections
            BeltempTimes.Add Round((FahtsData.StepData(i - 1).TimeInSeconds), 1)
            BeltempTemperatures.Add FahtsData.StepData(i - 1).MaxTemp

        End If
    Next i

    If Not IsBeltempStep(lastFahtsStep, BeltempStep) Then
        ' Add data to collections
        BeltempTimes.Add Round((FahtsData.StepData(lastFahtsStep - 1).TimeInSeconds), 1)
        BeltempTemperatures.Add FahtsData.StepData(lastFahtsStep - 1).MaxTemp
       
    End If

    ' Add data to dictionary
    BeltempData.Add "Times", BeltempTimes
    BeltempData.Add "Temperatures", BeltempTemperatures

    Set GetBeltempData = BeltempData
End Function

Function FindPeaksIdx(values As Collection) As Collection
    Dim peaks As New Collection
    Dim i As Long

    ' Check for peaks and decrease shape after plateau in values excluding the boundaries
    For i = 2 To values.Count - 1
        If values(i) >= values(i - 1) And values(i) > values(i + 1) Then
            peaks.Add i
        End If
    Next i

    Set FindPeaksIdx = peaks
End Function

Function FindMaxTempIdx(Temperatures As Collection) As Long
    Dim maxTempIdx As Long
    Dim i As Long

    maxTempIdx = 1
    For i = 2 To Temperatures.Count
        If Temperatures(i) > Temperatures(maxTempIdx) Then
            maxTempIdx = i
        End If
    Next i

    FindMaxTempIdx = maxTempIdx
End Function


Private Function GetFilteredTimes(ByVal BeltempTimes As Collection, ByVal BeltempTemperatures As Collection) As Collection
    ' Get the next load case time based on the last load case time
    Dim loadCaseTimeDetails As Variant
    Dim FilteredTimes As New Collection
    Dim maxTempIdx As Long
    Dim PeaksIdx As Collection
    Dim idx As Variant

    maxTempIdx = FindMaxTempIdx(BeltempTemperatures)

    FilteredTimes.Add BeltempTimes(maxTempIdx)
    
    Set PeaksIdx = FindPeaksIdx(BeltempTemperatures)
    For Each idx In PeaksIdx
        If idx <> maxTempIdx Then
            FilteredTimes.Add BeltempTimes(idx)
        End If
    Next

    Set GetFilteredTimes = FilteredTimes
End Function

Private Function GetRemainingFilteredTime(ByVal BeltempTimes As Collection, ByVal FilteredTimes As Collection) As Collection
    ' Get the next load case time based on the last load case time
    Dim loadCaseTimeDetails As Variant
    Dim RemainingFilteredTimes As New Collection
    Dim idx As Variant

    For Each idx In BeltempTimes
        If Not IsInCollection(FilteredTimes, idx) Then
            RemainingFilteredTimes.Add idx
        End If
    Next

    Set GetRemainingFilteredTime = RemainingFilteredTimes
End Function
