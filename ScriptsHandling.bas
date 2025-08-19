Option Private Module       'Indicates that module is private and not seeing in macros option
Const HEAD_DEFAULT_FILE As String = "\usfos_head_default.fem"

Sub CreateScriptFiles(ByVal commandName As String, ScriptName As String)
    ' This sub creates the bat first as txt to be written
    'TODO: Eliminate command Name
    Dim ScriptPath As String
    
    ' Define file paths
    ScriptPath = CURRENT_FOLDER & "\" & ScriptName & ".txt"
    
    ' Create Script File
    ScriptFile = FreeFile
    Open ScriptPath For Output As #ScriptFile
    Print #ScriptFile, ":::: Shell to start " & commandName & " run"
    Close #ScriptFile

End Sub

Sub InitializeCaseValues(ByVal row As Long, ByRef scenarioName As String, ByRef CaseFolder As String, ByRef ModelFileName As String, ByRef BeltempFilename As String, ByRef fire As String, ByRef HeadFilename As String, ByRef InputFile As String)
    ' Initialize and fetch values from the worksheet
    scenarioName = shtScriptUsfos.Cells(row, "B")
    ModelFileName = Replace(expression:=shtScriptFaths.Cells(row, "P"), Find:=".fem", Replace:="")
    fire = shtScriptFaths.Cells(row, "X")
    
    CaseFolder = CURRENT_FOLDER & "\" & shtScriptFaths.Cells(row, "A") & "_" & scenarioName
    BeltempFilename = scenarioName & "_output_beltemp"
    HeadFilename = "usfos_head_" & scenarioName
    InputFile = CaseFolder & "\usfos_input.inp"
End Sub

Sub FinalizeScripts(ScriptName As String)
    Dim UsfosScriptFile As String, StrainScriptFile As String, StrainNegScriptFile As String
    Dim UsfosScriptBat As String, StrainScriptBat As String, StrainNegScriptBat As String

    ' Define file paths based on ScriptName
    ScriptPathTxt = CURRENT_FOLDER & SCRIPT_DIR & "\" & ScriptName & ".txt"

    ' Define BAT file names
    ScriptPathBat = Replace(ScriptPathTxt, ".txt", ".bat")

    ' Finalize Script File
    If FileExists(ScriptPathBat) Then
        Kill ScriptPathBat
    End If
    RenameFile ScriptPathTxt, ScriptPathBat
    
End Sub

Sub RenameFile(ByVal oldName As String, ByVal newName As String)
    If FileExists(oldName) Then
        Name oldName As newName
    End If
End Sub

Function FileExists(ByVal filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Sub CopyCaseFile(ByVal sourcePath As String, ByVal destinationPath As String, ByVal fileName As String)
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

Function GetFileLastModified(filePath As String) As Date
    Dim FileSystem As Object
    Dim File As Object
    
    ' Create a FileSystemObject
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If FileSystem.FileExists(filePath) Then
        ' Get the file object
        Set File = FileSystem.GetFile(filePath)
        
        ' Return the last modified date and time
        GetFileLastModified = File.DateLastModified
    Else
        ' Return 0 date if file not found
        GetFileLastModified = 0
    End If
    
    ' Clean up
    Set File = Nothing
    Set FileSystem = Nothing
End Function

Function CheckFileChange(filePath As String, PreviousTimestamp As Date) As Boolean
    Dim CurrentTimestamp As Date

    CurrentTimestamp = GetFileLastModified(filePath)

    ' Return False if file not found (timestamp = 0) or if no change
    CheckFileChange = (CurrentTimestamp > 0) And (CurrentTimestamp > PreviousTimestamp)
End Function
