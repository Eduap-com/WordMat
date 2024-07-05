Attribute VB_Name = "ModuleP"
Option Private Module
Option Explicit
' Her kaldes funktioner, der kræves WordMat partnerskab
' Application.run "Funktion"  kan error håndteres, overføre parametre og returnere værdier

Sub TestPP()

    MsgBox QActivePartnership
    MsgBox mSkoleNavn
End Sub

Public Function QActivePartnership(Optional Force As Boolean = False, Optional ShowForm As Boolean = False) As Boolean
' Returns true if the user has active partnership
    Err.Clear
    On Error Resume Next
    
    Application.Run macroname:="PQActivePartnership", varg1:=Force, varg2:=True, varg3:=ShowForm
    If Err.Number = 513 Then
        QActivePartnership = True
        mSkoleNavn = Err.Description
        mPartnerSkab = 1
    Else
        QActivePartnership = False
        mSkoleNavn = vbNullString
        mPartnerSkab = 2
    End If
    Err.Clear
    
    GoTo slut
fejl:
    QActivePartnership = False
    mSkoleNavn = vbNullString
slut:
End Function
Public Function QCheckPartnerShip() As Boolean
' Returns true if the user has active partnership, and ask the user to change dll setting if not
    Dim t As String
    QCheckPartnerShip = QActivePartnership
    If Not QCheckPartnerShip Then
        If Sprog.SprogNr = 1 Then
            t = "Forbindelse til Maxima med direkte dll eller WSH kræver WordMat+ (Partnerskab med Eduap)" & vbCrLf & vbCrLf & "Vil du skifte til standardindstillingen: registeret dll?"
        Else
            t = "Connecting to Maxima with direct dll or WSH requires Partnership with Eduap" & vbCrLf & vbCrLf & "Do you want to change the setting to use the default option: registered dll?"
        End If
        If MsgBox2(t, vbYesNo, "Error") = vbYes Then
            DllConnType = 0
        End If
    End If
End Function
Function SkoleNavn() As String
    On Error GoTo fejl
    If mSkoleNavn <> vbNullString Then
        SkoleNavn = mSkoleNavn
    ElseIf mPartnerSkab = 0 Then
        If QActivePartnership Then
            mPartnerSkab = 1
            SkoleNavn = mSkoleNavn
        Else
            mPartnerSkab = 2
            SkoleNavn = vbNullString
        End If
    End If
        
    GoTo slut
fejl:
    SkoleNavn = vbNullString
slut:

End Function


#If Mac Then
#Else
Public Function PGetMaxProc() As Object
    Dim DllDir As String, RK As String
    On Error GoTo slut
    RK = GetReg("InstallLocation")
    If RK = "All" Then
        On Error Resume Next
        DllDir = GetProgramFilesDir & "\WordMat\"
        On Error GoTo slut
        If Dir(DllDir & "MathMenu.dll") = vbNullString Then
            DllDir = Environ("AppData") & "\WordMat\"
        End If
    Else
        DllDir = Environ("AppData") & "\WordMat\"
        If Dir(DllDir & "MathMenu.dll") = vbNullString Then
            On Error Resume Next
            DllDir = GetProgramFilesDir & "\WordMat\"
            On Error GoTo 0
        End If
    End If
    If Dir(DllDir & "MathMenu.dll") <> vbNullString Then
        Set PGetMaxProc = GetObjectFromDll(DllDir, "MathMenu.dll", "MaximaProcessClass")
    End If
slut:
End Function
Public Function PGetWebView() As Object
    Dim DllDir As String, CLRdllDir As String, RK As String, DllName As String
    DllName = "WebViewWrap.dll"
    On Error GoTo slut
    
    RK = GetReg("InstallLocation")
    If RK = "All" Then
        On Error Resume Next
        CLRdllDir = "GetProgramFilesDir" & "\WordMat\"
        DllDir = CLRdllDir & "WebViewWrap\"
        On Error GoTo slut
        If Dir(DllDir & DllName) = vbNullString Then
            CLRdllDir = Environ("AppData") & "\WordMat\"
            DllDir = CLRdllDir & "WebViewWrap\"
        End If
    Else
        CLRdllDir = Environ("AppData") & "\WordMat\"
        DllDir = CLRdllDir & "WebViewWrap\"
        If Dir(DllDir & DllName) = vbNullString Then
            On Error Resume Next
            CLRdllDir = "GetProgramFilesDir" & "\WordMat\"
            DllDir = CLRdllDir & "WebViewWrap\"
            On Error GoTo slut
        End If
    End If
    Set PGetWebView = GetObjectFromDll(DllDir, DllName, "WebViewWrap.Browser", CLRdllDir)
slut:
End Function

#End If
