Attribute VB_Name = "AuxMacros"
Option Private Module       'Indicates that module is private and not seeing in macros option
Public CURRENT_FOLDER As String

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
    CURRENT_FOLDER = Application.ActiveWorkbook.path
End Sub

Function IsCompleted(row As Long) As Boolean
    IsCompleted = (shtScriptUsfos.Cells(row, "S") = "completed")
End Function

Sub CheckPrograms()
    Dim prog As Program

    For prog = FAHTS To SRAINNeg
        Dim programPath As String
        programPath = GetProgramPath(prog)
        CheckAndAlertProgramPath programPath
    Next prog
End Sub

Private Function GetProgramPath(ByVal prog As Program) As String
    Dim usfosFolder As String
    
    setCurrentFolder
    
    usfosFolder = shtScriptUsfos.Range("W1")
    
    Select Case prog
        Case FAHTS
            GetProgramPath = CURRENT_FOLDER & "\Fahts\fahts_v670.exe"
        Case USFOS
            GetProgramPath = usfosFolder & "\Usfos.exe"
        Case Strain
            GetProgramPath = usfosFolder & "\strain.exe"
        Case SRAINNeg
            GetProgramPath = usfosFolder & "\strain_neg.exe"
        Case Else
            GetProgramPath = ""
    End Select
End Function

Private Sub CheckAndAlertProgramPath(ByVal programPath As String)
    Dim FileSystem As Object
    Dim path As String

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


