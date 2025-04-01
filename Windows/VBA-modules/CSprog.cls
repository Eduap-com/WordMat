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
Private SA(999) As String
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
Case 1030 ' Danish
    GetOfficeLanguage = 1
Case 1033 ' English
    GetOfficeLanguage = 2
Case 1034 ' Spanish
    GetOfficeLanguage = 3
Case 3082 ' Spanish international
    GetOfficeLanguage = 3
Case 2058 ' Spanish mexican
    GetOfficeLanguage = 3
Case 11274 ' Spanish Argentina
    GetOfficeLanguage = 3
Case 9226 ' Spanish Columbia
    GetOfficeLanguage = 3
Case 14346 ' Spanish Uruguay
    GetOfficeLanguage = 3
Case 1031 ' German
    GetOfficeLanguage = 4
Case 2055 ' German Switzerland
    GetOfficeLanguage = 4
Case 3079 ' German Austria
    GetOfficeLanguage = 4
Case 1036 ' French
    GetOfficeLanguage = 5
Case 3084 ' French Canada
    GetOfficeLanguage = 5
Case 4108 ' French Switzerland
    GetOfficeLanguage = 5
Case 2060 ' French Belgium
    GetOfficeLanguage = 5
Case 5132 ' French Luxemburg
    GetOfficeLanguage = 5
Case 1040 ' Italian
    GetOfficeLanguage = 6
Case 2064 ' Italian Switzerland
    GetOfficeLanguage = 6
Case 1053 ' Swedish
    GetOfficeLanguage = 7
Case 2077 ' Swedish finnish
    GetOfficeLanguage = 7
Case 1044 ' Norwegian
    GetOfficeLanguage = 8
Case 2067 ' Norwegian Nynorsk
    GetOfficeLanguage = 8
Case 1035 ' Finnish
    GetOfficeLanguage = 9
Case 1043 ' Dutch
    GetOfficeLanguage = 10
Case 1045 ' Polish
    GetOfficeLanguage = 11
Case 1046 ' Portuguese Brazil
    GetOfficeLanguage = 12
Case 2070 ' Portuguese Portugal
    GetOfficeLanguage = 12
Case 1080 ' Faroese
    GetOfficeLanguage = 13
Case 1061 ' estonian
    GetOfficeLanguage = 14
Case 1062 ' latvian
    GetOfficeLanguage = 15
Case 1063 ' lithuanian
    GetOfficeLanguage = 16
Case 1029 ' Check
    GetOfficeLanguage = 17
Case 1032 ' Greek
    GetOfficeLanguage = 2 ' there are charset problems with greek letters
Case 1039 ' Icelandic - not supported in Excel
    GetOfficeLanguage = 1
Case Else
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
    ElseIf SprogNr = 4 Then
        FilePath = FilePath & "languages" & slash & "de.csv"
    ElseIf SprogNr = 5 Then
        FilePath = FilePath & "languages" & slash & "fr.csv"
    ElseIf SprogNr = 6 Then
        FilePath = FilePath & "languages" & slash & "it.csv"
    ElseIf SprogNr = 7 Then
        FilePath = FilePath & "languages" & slash & "sv.csv"
    ElseIf SprogNr = 8 Then
        FilePath = FilePath & "languages" & slash & "nb.csv"
    ElseIf SprogNr = 9 Then
        FilePath = FilePath & "languages" & slash & "fi.csv"
    ElseIf SprogNr = 10 Then
        FilePath = FilePath & "languages" & slash & "nl.csv"
    ElseIf SprogNr = 11 Then
        FilePath = FilePath & "languages" & slash & "pl.csv"
    ElseIf SprogNr = 12 Then
        FilePath = FilePath & "languages" & slash & "pt-pt.csv"
    ElseIf SprogNr = 13 Then
        FilePath = FilePath & "languages" & slash & "fo.csv"
    ElseIf SprogNr = 14 Then
        FilePath = FilePath & "languages" & slash & "et.csv"
    ElseIf SprogNr = 15 Then
        FilePath = FilePath & "languages" & slash & "lv.csv"
    ElseIf SprogNr = 16 Then
        FilePath = FilePath & "languages" & slash & "lt.csv"
    ElseIf SprogNr = 17 Then
        FilePath = FilePath & "languages" & slash & "cs.csv"
    ElseIf SprogNr = 18 Then
        FilePath = FilePath & "languages" & slash & "el.csv"
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

