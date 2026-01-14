Option Explicit

'Change this if your worksheet tab name is different.
Private Const SHEET_NAME As String = "Script_Faths"

'========================
' 1) Browse Source Folder -> AJ2
'========================
Public Sub BrowseSourceFolder_ToAJ2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim p As String
    p = PickFolder("Select SOURCE folder")
    If Len(p) = 0 Then Exit Sub

    ws.Range("AJ2").Value = p
End Sub

'========================
' 2) Browse Destination Folder -> AJ3
'========================
Public Sub BrowseDestinationFolder_ToAJ3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim p As String
    p = PickFolder("Select DESTINATION folder")
    If Len(p) = 0 Then Exit Sub

    ws.Range("AJ3").Value = p
End Sub

'========================
' 3) Copy all files/folders from:
'    SourceBase\[A]_[B]   -> DestinationBase\[A]_[B]
'    for rows 2..last in columns A and B
'========================
Public Sub Copy_AllCaseFolders_Contents()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim sourceBase As String, destBase As String
    sourceBase = Trim$(ws.Range("AJ2").Value)
    destBase = Trim$(ws.Range("AJ3").Value)

    If Len(sourceBase) = 0 Or Len(destBase) = 0 Then
        MsgBox "Please set Source (AJ2) and Destination (AJ3) first.", vbExclamation
        Exit Sub
    End If

    sourceBase = RemoveTrailingSlash(sourceBase)
    destBase = RemoveTrailingSlash(destBase)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(sourceBase) Then
        MsgBox "Source base folder does not exist:" & vbCrLf & sourceBase, vbCritical
        Exit Sub
    End If
    If Not fso.FolderExists(destBase) Then
        MsgBox "Destination base folder does not exist:" & vbCrLf & destBase, vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No data found in column A starting at row 2.", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    Dim numVal As String, caseVal As String, folderName As String
    Dim srcPath As String, dstPath As String
    Dim copiedCount As Long, missingCount As Long, skippedCount As Long, errCount As Long

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        numVal = Trim$(CStr(ws.Cells(i, "A").Value))
        caseVal = Trim$(CStr(ws.Cells(i, "B").Value))

        If Len(numVal) = 0 And Len(caseVal) = 0 Then
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If

        folderName = SanitizeFolderName(numVal & "_" & caseVal)

        srcPath = sourceBase & "\" & folderName
        dstPath = destBase & "\" & folderName

        If Not fso.FolderExists(srcPath) Then
            missingCount = missingCount + 1
            GoTo NextRow
        End If

        If Not fso.FolderExists(dstPath) Then
            CreateFullPath fso, dstPath
        End If

        'Per-row error handling:
        On Error GoTo RowErr
        CopyFolderContents srcPath, dstPath
        On Error GoTo 0

        copiedCount = copiedCount + 1

NextRow:
        On Error GoTo 0
    Next i

    Application.ScreenUpdating = True

    MsgBox "Done!" & vbCrLf & _
           "Cases copied: " & copiedCount & vbCrLf & _
           "Missing source folders: " & missingCount & vbCrLf & _
           "Skipped blank rows: " & skippedCount & vbCrLf & _
           "Rows with copy errors: " & errCount, vbInformation
    Exit Sub

RowErr:
    errCount = errCount + 1
    Resume NextRow
End Sub


'========================
' Helpers
'========================
Private Function PickFolder(ByVal dialogTitle As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    Dim startPath As String
    startPath = ThisWorkbook.Path

    'If the workbook was never saved, Path is empty ? fall back
    If Len(startPath) = 0 Then startPath = Application.DefaultFilePath

    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False

        'This makes the dialog open in the workbook folder
        .InitialFileName = startPath & "\"

        If .Show <> -1 Then
            PickFolder = ""
        Else
            PickFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Function RemoveTrailingSlash(ByVal p As String) As String
    Do While Right$(p, 1) = "\" Or Right$(p, 1) = "/"
        p = Left$(p, Len(p) - 1)
    Loop
    RemoveTrailingSlash = p
End Function

Private Function SanitizeFolderName(ByVal s As String) As String
    'Windows invalid filename chars: \ / : * ? " < > |
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    s = Trim$(s)
    For Each ch In badChars
        s = Replace$(s, CStr(ch), "_")
    Next ch

    'Optional: collapse double spaces/underscores a bit
    Do While InStr(s, "__") > 0
        s = Replace$(s, "__", "_")
    Loop

    SanitizeFolderName = s
End Function

Private Sub CreateFullPath(ByVal fso As Object, ByVal fullPath As String)
    'Creates nested folders if needed
    Dim parts() As String, i As Long, build As String
    parts = Split(fullPath, "\")
    build = parts(0) 'e.g. "C:" or server share start

    For i = 1 To UBound(parts)
        build = build & "\" & parts(i)
        If Not fso.FolderExists(build) Then
            fso.CreateFolder build
        End If
    Next i
End Sub


'--- Copy files + subfolders from srcPath INTO dstPath
Private Sub CopyFolderContents(ByVal srcPath As String, ByVal dstPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Copy files (including extensionless)
    On Error Resume Next
    fso.CopyFile srcPath & "\*", dstPath & "\", True
    On Error GoTo 0

    ' Copy SubFolders
    On Error Resume Next
    fso.CopyFolder srcPath & "\*", dstPath & "\", True
    On Error GoTo 0
End Sub
