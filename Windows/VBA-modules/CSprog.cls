VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SprogNr As Integer
Private SA(900) As String
'0=WordLanguage, 1=Dansk, 2=English, 3=Spanish

Private Sub Class_Initialize()
    RunFirst
    CheckSetting
    If Sprog.SprogNr = 1 Then
        AppNavn = "WordMat"
    Else
        AppNavn = "WordMath"
    End If
    LoadSprogArray
End Sub
Public Sub CheckSetting()
    If LanguageSetting = 0 Then
        SprogNr = GetOfficeLanguage
    Else
        SprogNr = LanguageSetting
    End If
End Sub
Function GetOfficeLanguage() As Integer
'Reads the UI-language for Word
Select Case Application.International(wdProductLanguageID)
Case 1030 ' Dansk
    GetOfficeLanguage = 1
Case 1034 ' Spanish
    GetOfficeLanguage = 3
Case Else ' 1033 is english
    GetOfficeLanguage = 2
End Select

End Function
Public Function A(i As Long) As String
    On Error Resume Next
    If SA(0) = vbNullString Then LoadSprogArray
    A = SA(i)
End Function
Public Sub LoadSprogArray()
' method from ReadTextFile
' also check ReadTextfileToString
    Dim FileNo As Integer
    Dim linje As String
    Dim Arr() As String
    Dim FilePath As String, slash As String
    
#If Mac Then
    slash = "/"
#Else
    slash = "\"
#End If
    FilePath = GetWordMatDir
    If SprogNr = 1 Then
        FilePath = FilePath & "languages" & slash & "da.csv"
    ElseIf SprogNr = 2 Then
        FilePath = FilePath & "languages" & slash & "en.csv"
    ElseIf SprogNr = 3 Then
        FilePath = FilePath & "languages" & slash & "sp.csv"
    End If
    FileNo = FreeFile

    Open FilePath For Input As FileNo
    Do While Not EOF(FileNo) ' Loop until end of file.
        Line Input #FileNo, linje
        Arr = Split(linje, "#£")
        If UBound(Arr) > 0 Then
            If IsNumeric(Arr(0)) Then
                SA(CInt(Arr(0))) = FormatText(Arr(1))
            End If
        End If
    Loop
    Close FileNo
End Sub

Function FormatText(ByVal t As String) As String
    t = Replace(t, "\r\n", VbCrLfMac)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\t", vbTab)
    t = Replace(t, "$Tab", vbTab)
    t = Replace(t, "$*", ChrW(183))
    t = Replace(t, "$C", ChrW(8450))
    t = Replace(t, "$R", ChrW(8477))
    t = Replace(t, "$X^2", ChrW(&H3C7) & ChrW(&HB2))
    t = Replace(t, "$Degree", ChrW(176))
    t = Replace(t, "$ArrowRight", ChrW(8594))
    
    FormatText = t
End Function

''''''''''''''Often used translation texts'''''''''
Public Property Get SyntaxError() As String
    SyntaxError = A(751)
End Property
Public Property Get Error() As String
    Error = A(822)
End Property
Public Property Get OK() As String
    OK = A(824)
End Property
Public Property Get Cancel() As String
    Cancel = A(825)
End Property
Public Property Get Reset() As String
    Reset = A(838)
End Property
Public Property Get ErrorGeneral() As String
    ErrorGeneral = A(844)
End Property

