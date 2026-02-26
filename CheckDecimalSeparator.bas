Option Explicit

'--- Windows API: read user locale decimal separator (Windows only)
#If VBA7 Then
    Private Declare PtrSafe Function GetLocaleInfoW Lib "kernel32" ( _
        ByVal Locale As Long, _
        ByVal lcType As Long, _
        ByVal lpLCData As LongPtr, _
        ByVal cchData As Long) As Long
#Else
    Private Declare Function GetLocaleInfoW Lib "kernel32" ( _
        ByVal Locale As Long, _
        ByVal lcType As Long, _
        ByVal lpLCData As Long, _
        ByVal cchData As Long) As Long
#End If

Private Const LOCALE_USER_DEFAULT As Long = &H400
Private Const LOCALE_SDECIMAL As Long = &HE

Private Function GetLocaleString(ByVal lcType As Long) As String
    Dim n As Long
    n = GetLocaleInfoW(LOCALE_USER_DEFAULT, lcType, 0, 0)
    If n <= 1 Then Exit Function

    Dim buff As String
    buff = String$(n, vbNullChar)
    GetLocaleInfoW LOCALE_USER_DEFAULT, lcType, StrPtr(buff), n

    GetLocaleString = Left$(buff, n - 1) 'remove trailing null
End Function

Public Function SystemDecimalSeparator() As String
    SystemDecimalSeparator = GetLocaleString(LOCALE_SDECIMAL)
End Function

Public Sub CheckDecimalSeparatorsMustBeDot()
    Dim sysDec As String, xlDec As String
    sysDec = SystemDecimalSeparator()
    xlDec = Application.International(xlDecimalSeparator)

    If sysDec <> "." Or xlDec <> "." Then
        Dim msg As String
        msg = "This workbook (and another program) requires '.' as the decimal separator." & vbCrLf & vbCrLf & _
              "Windows decimal: '" & sysDec & "'" & vbCrLf & _
              "Excel decimal:    '" & xlDec & "'" & vbCrLf & vbCrLf & _
              "Fix in Windows:" & vbCrLf & _
              "1) Press Win + R" & vbCrLf & _
              "2) Type: intl.cpl  and press Enter" & vbCrLf & _
              "3) Click 'Additional settings...'" & vbCrLf & _
              "4) Set 'Decimal symbol' to: ." & vbCrLf & _
              "5) Click OK > Apply > OK" & vbCrLf & _
              "6) Close and reopen Excel (and the other program)" & vbCrLf & vbCrLf & _
              "If Excel still shows comma after that:" & vbCrLf & _
              "Excel: File > Options > Advanced > check 'Use system separators'."

        MsgBox msg, vbCritical, "Decimal separator must be '.'"

        'Optional: open the Windows Region dialog automatically:
        'Shell "control.exe intl.cpl", vbNormalFocus

        'Stop your macro if needed:
        Err.Raise vbObjectError + 513, "CheckDecimalSeparatorsMustBeDot", "Wrong decimal separator"
    End If
End Sub

Private Sub Workbook_Open()
    CheckDecimalSeparatorsMustBeDot
End Sub


