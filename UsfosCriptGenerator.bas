Const SCRIPT_DIR As String = "\script_rodar_"
Const STRASPAS As String = """"   'defining aspas to be inside strings
Const HEAD_DEFAULT_FILE As String = "\usfos_head_default.fem"

Sub GerarHeadfiles_usfos()
    TurnOffCalcs
    
    Dim ultimaLinha As Long
    Dim row As Long
    Dim scriptUsfos As String
    Dim scriptStrain As String
    Dim scriptStrainNeg As String
    Dim response As VbMsgBoxResult
    
    ' Initialize Current Folder and check programs
    AuxMacros.setCurrentFolder
    AuxMacros.CheckPrograms

    ' Asks the user if they want to only update files or if they want to update the formulas of the worksheet
    ' Update formulas is only necessary when the number of scenarios or FEM elements has changed
    response = AskToUpdateFormulas()

    InitializeDefaults

   scriptUsfos = "script_rodar_USFOS"
   scriptStrain = "script_rodar_STRAIN"
   scriptStrainNeg = "script_rodar_STRAIN_neg"

   ' Creating Script files for Usfos, Strain and Strain neg as txt
   ScriptsHandling.CreateScriptFiles commandName:="USFOS", ScriptName:=scriptUsfos
   ScriptsHandling.CreateScriptFiles commandName:="STRAIN", ScriptName:=scriptStrain
   ScriptsHandling.CreateScriptFiles commandName:="STRAIN NEG", ScriptName:=scriptStrainNeg
  
    ultimaLinha = shtScriptFaths.Cells(Rows.Count, "A").End(xlUp).row
    For row = 2 To ultimaLinha
        If Not AuxMacros.IsCompleted(row) Then
            ProcessCase row
        
        End If
    Next row

   'Renaming scripts from txt to bat
   ScriptsHandling.FinalizeScripts ScriptName:=scriptUsfos
   ScriptsHandling.FinalizeScripts ScriptName:=scriptStrain
   ScriptsHandling.FinalizeScripts ScriptName:=scriptStrainNeg


    If response Then
    ApplyFormulas.BuildFormulaBlocks_AllScenarios
    End If

    MsgBox "Script and input files generated!"
    
    TurnOnCalcs
End Sub

Private Sub InitializeDefaults()
' Initialize strain criteria, Factor Thermal Loads  and Drive
    If IsEmpty(shtScriptUsfos.Cells(20, "W")) Then shtScriptUsfos.Cells(20, "W") = 2
    If IsEmpty(shtScriptUsfos.Cells(23, "W")) Then shtScriptUsfos.Cells(23, "W") = 1
    
End Sub

Private Sub ProcessCase(ByVal row As Long)
    Dim scenarioName As String
    Dim CaseFolder As String
    Dim ModelFileName As String
    Dim BeltempFilename As String
    Dim BeltempFile As String
    Dim fire As String
    Dim headFile As String
    Dim HeadFilename As String
    Dim InputFile As String


    ' Check for necessary conditions before proceeding
    'If Not IsCaseValid(row) Then Exit Sub  'TODO: Consider when cases are not valid to stop macro

    
    ' Write Data on shtScriptUsfos
    shtScriptUsfos.Cells(row, "A") = shtScriptFaths.Cells(row, "A")  'Case Number
    shtScriptUsfos.Cells(row, "B") = shtScriptFaths.Cells(row, "B")  'Case Name
    shtScriptUsfos.Cells(row, "C") = shtScriptFaths.Cells(row, "C")  'End Time
    
    If shtScriptUsfos.Cells(row, "I") = "" Then shtScriptUsfos.Cells(row, "I") = shtScriptUsfos.Cells(20, "W") ' Strain Criteria
    If shtScriptUsfos.Cells(row, "R") = "" Then shtScriptUsfos.Cells(row, "R") = shtScriptUsfos.Cells(23, "W") ' Factor Termal Loads
    
    ' Initialize values from the sheet
    ScriptsHandling.InitializeCaseValues row, scenarioName, CaseFolder, ModelFileName, BeltempFilename, fire, HeadFilename, InputFile 'TODO: Change this to a function
    
    
    ' Generate Strain and Strain Neg Input
    GerarStrainsInput scenarioName, CaseFolder
    
    ' Get info about last load case and last time
    BeltempFilePath = CaseFolder & "\" & BeltempFilename & ".fem"
    Set LastLoadCaseInfo = GetLastLoadCaseInfos(BeltempFilePath)
    LastLoadCase = LastLoadCaseInfo("load case")                    'last loadcase number
     
    ProcessFireInfo fire, row, scenarioName
    
    
    ' Copy and modify head file
    UpdateHeadUsfos row, LastLoadCase

    ' Copy and modify model file
    ModelFile = ModelFileName & ".fem"
    CopyCaseFile CURRENT_FOLDER, CaseFolder, ModelFile
    
    ' Creating Input file to be used for scripts files
    CreateInputFile InputFile, HeadFilename, ModelFileName, BeltempFilename, scenarioName

    ' Write to script files
    WriteToScriptFiles row, scenarioName, CaseFolder

End Sub

Private Function GetLastLoadCaseInfos(ByVal BeltempFile As String) As Object
    ' Read and process the Beltemp file.
    Dim textline As String
    Dim loadcase_line As String
    Dim loadcase_Array() As String
    Dim loadcase_infos As Object
    
    Set loadcase_infos = CreateObject("Scripting.Dictionary")
    
    BelTempFileNum = FreeFile
    Open BeltempFile For Input As #BelTempFileNum
    Do Until EOF(BelTempFileNum)
        Line Input #BelTempFileNum, textline
        If InStr(textline, "LCASETIM") Then
            loadcase_line = WorksheetFunction.Trim(textline)                    ' remove extra spaces
            loadcase_Array = Split(loadcase_line, " ")                          ' split into array
            loadcase_infos("load case") = loadcase_Array(1)                     ' get the load case info
            loadcase_infos("time") = loadcase_Array(2)                          ' get the time info
        End If
    Loop
    Close #BelTempFileNum
    Set GetLastLoadCaseInfos = loadcase_infos
End Function

Private Sub ProcessFireInfo(ByVal fire As String, ByVal row As Long, ByVal scenarioName As String)
    Dim fireDurationInfo As String

    ' Check for missing fire information
    If fire = "" Or fire = "?" Then
        MsgBox "Please include Pool or Jet information in Sheet Script_Faths Column X for Case " & scenarioName & "! Then, run the script again."
        Exit Sub
    End If
    
    exposureTime = shtScriptUsfos.Cells(row, "C").Value   'End time in sheet
    
    ' Determine fire duration based on fire type and exposure time
    fireDurationInfo = DetermineFireDuration(fire, exposureTime)

    ' Update the worksheet with the fire duration info
    shtScriptUsfos.Cells(row, "H").Value = fireDurationInfo
End Sub

Private Function DetermineFireDuration(ByVal fireType As String, ByVal exposureTime As Long) As String
    ' Example logic for determining fire duration info
    If exposureTime <= 15 * 60 Then        ' Less than or equal to 15 minutes
        If fireType = "Pool" Then
            DetermineFireDuration = "H15"
        ElseIf fireType = "Jet" Then
            DetermineFireDuration = "J15"
        End If
    ElseIf exposureTime <= 30 * 60 Then    ' Between 15 and 30 minutes
        If fireType = "Pool" Then
            DetermineFireDuration = "H30"
        ElseIf fireType = "Jet" Then
            DetermineFireDuration = "J30"
        End If
    ElseIf exposureTime > 30 * 60 Then     ' More than 30 minutes
        If fireType = "Pool" Then
            DetermineFireDuration = "H60"
        ElseIf fireType = "Jet" Then
            DetermineFireDuration = "J60"
        End If
    End If
End Function

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

    ' Define the paths to Head
    HeadFilename = "usfos_head_" & scenarioName
    HeadDefaultFile = CURRENT_FOLDER & HEAD_DEFAULT_FILE

    ' Define the paths to Head for specific Case
    HeadCaseDefaultFile = caseFolderPath & HEAD_DEFAULT_FILE
    HeadCaseFile = caseFolderPath & "\" & HeadFilename & ".fem"

    ' Copy Head File Default to Case (No modifications)
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    FSObject.CopyFile HeadDefaultFile, HeadCaseDefaultFile, True
    
    ' Copy and modify head file
    FSObject.CopyFile HeadDefaultFile, HeadCaseFile, True
    
    ' Update the head file
    ModifyHeadFile headFile:=HeadCaseFile, LoadCaseTime:=LoadCaseTime
End Sub


Private Sub ModifyHeadFile(ByVal headFile As String, ByVal LoadCaseTime As Double)
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
Private Sub CreateInputFile(ByVal InputFile As String, ByVal HeadFilename As String, ByVal ModelFileName As String, ByVal BeltempFilename As String, ByVal scenarioName As String)
Dim InputFileNum As Long

InputFileNum = FreeFile
Open InputFile For Output As #InputFileNum
        Print #InputFileNum, HeadFilename
        Print #InputFileNum, ModelFileName
        Print #InputFileNum, BeltempFilename
        Print #InputFileNum, scenarioName & "_usfos_output"
Close #InputFileNum
End Sub

Private Sub WriteToScriptFiles(ByVal row As Long, ByVal scenarioName As String, ByVal CaseFolder As String)
    Dim UsfosScriptFile As String, StrainScriptFile As String, StrainNegScriptFile As String
    Dim Usfos_directory, Strain_directory, Strain_neg_directory As String
    Dim fileNum As Integer
    Dim Drive As String
    
    
    ' Initialize Drive. Default is C
    If shtScriptUsfos.Range("W26") = "d:\" Then Drive = "/D " Else Drive = ""
    
    ' Define the paths to the script files
    UsfosScriptText = CURRENT_FOLDER & SCRIPT_DIR & "USFOS.txt"
    StrainScriptText = CURRENT_FOLDER & SCRIPT_DIR & "STRAIN.txt"
    StrainNegScriptText = CURRENT_FOLDER & SCRIPT_DIR & "STRAIN_neg.txt"

    ' Define directory paths for executables
    Usfos_directory = shtScriptUsfos.Range("W1") & "\Usfos.exe"
    Strain_directory = shtScriptUsfos.Range("W1") & "\strain.exe"
    Strain_neg_directory = shtScriptUsfos.Range("W1") & "\strain_neg.exe"
    fahtsExe = shtScriptUsfos.Range("W2")

    ' Write commands to USFOS script file #TODO: CHange this to be called by a single functionpassing just if it is usfos, strain or strain neg
    fileNum = FreeFile()
    Open UsfosScriptText For Append As #fileNum
    Print #fileNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #fileNum, STRASPAS & Usfos_directory & STRASPAS; " 400 > " & STRASPAS & scenarioName & "_usfos.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\usfos_input.inp" & STRASPAS
    Print #fileNum, "::"
    Close #fileNum

    ' Write commands to STRAIN script file #TODO: CHange this to be called by a single functionpassing just if it is usfos, strain or strain neg
    fileNum = FreeFile()
    Open StrainScriptText For Append As #fileNum
    Print #fileNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #fileNum, STRASPAS & Strain_directory & STRASPAS; " 400 > " & STRASPAS & scenarioName & "_strain.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\strain_input.inp" & STRASPAS
    Print #fileNum, "::"
    Close #fileNum

    ' Write commands to STRAIN Neg script file #TODO: CHange this to be called by a single functionpassing just if it is usfos, strain or strain neg
    fileNum = FreeFile()
    Open StrainNegScriptText For Append As #fileNum
    Print #fileNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #fileNum, STRASPAS & Strain_neg_directory & STRASPAS; " 400 > " & STRASPAS & scenarioName & "_strain_neg.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\strain_neg_input.inp" & STRASPAS
    Print #fileNum, "::"
    Close #fileNum
    
    ' Initialize File paths to text and bat fot Script case
    ScriptFileCaseText = CaseFolder & "\script_case.txt"
    ScriptFileCaseBat = CaseFolder & "\script_case.bat"
    
    ' Write commands to script_case #TODO: change this to be a inside a funciton
    ScriptFileCaseNum = FreeFile
    Open ScriptFileCaseText For Output As #ScriptFileCaseNum
    Print #ScriptFileCaseNum, ":::: Shell to start FAHTS run"
    Print #ScriptFileCaseNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #ScriptFileCaseNum, STRASPAS & fahtsExe & STRASPAS & " <input.inp"
    Print #ScriptFileCaseNum, "cd .."
    Print #ScriptFileCaseNum, ":::: Shell to start USFOS run"
    Print #ScriptFileCaseNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #ScriptFileCaseNum, STRASPAS & Usfos_directory & STRASPAS; " 400 > " & STRASPAS & scenarioName & "_usfos.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\usfos_input.inp" & STRASPAS
    Print #ScriptFileCaseNum, "::"
    Print #ScriptFileCaseNum, ":::: Shell to start STRAIN run"
    Print #ScriptFileCaseNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #ScriptFileCaseNum, STRASPAS & Strain_directory & STRASPAS; STRASPAS & scenarioName & "_strain.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\strain_input.inp" & STRASPAS

    Print #ScriptFileCaseNum, "::"
    Print #ScriptFileCaseNum, ":::: Shell to start STRAIN neg run"
    Print #ScriptFileCaseNum, "cd " & Drive & STRASPAS & CaseFolder & STRASPAS
    Print #ScriptFileCaseNum, STRASPAS & Strain_neg_directory & STRASPAS; " 400 > " & STRASPAS & scenarioName & "_strain_neg.trm" & STRASPAS & " < " & STRASPAS & CaseFolder & "\strain_neg_input.inp" & STRASPAS
    Print #ScriptFileCaseNum, "::"
    Close #ScriptFileCaseNum
    
    If ScriptsHandling.FileExists(ScriptFileCaseBat) Then
         Kill ScriptFileCaseBat
    End If
    ScriptsHandling.RenameFile ScriptFileCaseText, ScriptFileCaseBat
    
End Sub

Private Sub GerarStrainsInput(ByVal scenarioName As String, ByVal CaseFolder As String)
    ' Generate strain input file. Different from PUSHDOWN the loadcase will alway be one so it could be generated first
    InputFileStrainPath = CaseFolder & "\strain_input.inp"
    
    strainFile = FreeFile
    Open InputFileStrainPath For Output As strainFile
    Print #strainFile, scenarioName & "_usfos_output"
    Print #strainFile, scenarioName & "_strain_output"
    Print #strainFile, 1
    Print #strainFile, "500000"
    Print #strainFile, "0"
    Print #strainFile, "n"
    Close #strainFile
    
    StrainNegFile = FreeFile
    InputFileStrainNegPath = CaseFolder & "\strain_neg_input.inp"
    Open InputFileStrainNegPath For Output As StrainNegFile
    Print #StrainNegFile, scenarioName & "_usfos_output"
    Print #StrainNegFile, scenarioName & "_strain_neg_output"
    Print #StrainNegFile, 1
    Print #StrainNegFile, "500000"
    Print #StrainNegFile, "-1000"
    Print #StrainNegFile, "n"
    Close #StrainNegFile



End Sub

Public Function AskToUpdateFormulas(Optional autoAnswerYes As Variant) As Boolean
    If Not IsMissing(autoAnswerYes) Then
        AskToUpdateFormulas = CBool(autoAnswerYes)
        Exit Function
    End If
    
    Dim msg As String, resp As VbMsgBoxResult
    msg = "Do you want to update the worksheet formulas now?" & vbNewLine & vbNewLine & _
          "Updating is only necessary if:" & vbNewLine & _
          " - This is your first time running the tool," & vbNewLine & _
          " - Or the number of scenarios or FEM elements has changed." & vbNewLine & vbNewLine & _
          "(Choose Yes to refresh formulas, or No to keep them as they are.)"

    resp = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Update Formulas")
    AskToUpdateFormulas = (resp = vbYes)
End Function




