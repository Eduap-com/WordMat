Attribute VB_Name = "ModuleP"
Option Private Module
Option Explicit
' Application.run "Function" can handle error, parameters and return values

Sub TestPP()
    MsgBox QActivePartnership
    MsgBox mSkoleNavn
End Sub

Public Function QActivePartnership(Optional Force As Boolean = False, Optional ShowForm As Boolean = False) As Boolean
' Returns true if the user has active partnership
    Err.Clear
    On Error Resume Next
    
    QActivePartnership = Application.Run("PQActivePartnership", Force, False, ShowForm)
    
    If QActivePartnership Then
        If mSkoleNavn = vbNullString Then mSkoleNavn = GetRegSettingString("Skolenavn")
    End If
    Exit Function
    
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
Public Function QCheckForWordMatPlus()
    If QActivePartnership(True, False) Then
        MsgBox TT.A(120), vbOKOnly, "OK"
    Else
        MsgBox TT.A(121), vbOKOnly, "Sorry"
    End If
End Function
Function SkoleNavn() As String
    On Error GoTo fejl
    If mSkoleNavn <> vbNullString Then
        SkoleNavn = mSkoleNavn
    ElseIf mPartnerSkab = 0 Then
        If QActivePartnership Then
            mPartnerSkab = 1
            mSkoleNavn = GetRegSetting("Skolenavn")
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

