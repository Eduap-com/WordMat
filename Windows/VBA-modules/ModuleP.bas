Attribute VB_Name = "ModuleP"
Option Private Module
Option Explicit
' Functions that require WordMat partnership
' Application.run "Function" can handle error, parameters and return values

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
Fejl:
    QActivePartnership = False
    mSkoleNavn = vbNullString
slut:
End Function

Function SkoleNavn() As String
    On Error GoTo Fejl
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
Fejl:
    SkoleNavn = vbNullString
slut:

End Function

