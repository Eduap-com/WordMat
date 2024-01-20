Attribute VB_Name = "ModuleP"
Option Private Module
Option Explicit
' Her kaldes funktioner, der kræves WordMat partnerskab
' Application.run "Funktion"  kan error håndteres, overføre parametre og returnere værdier

Sub WMS(SubName As String)
    On Error Resume Next
    Application.Run (SubName)
End Sub

Function WMFs(FunctionName As String) As String
    Err.Clear
    On Error Resume Next
    WMFs = Application.Run(FunctionName)
    If Err.Number <> 0 Then
        WMFs = "Error"
    End If
End Function

Public Function QActivePartnership() As Boolean
' Returns true if the user has active partnership
    On Error GoTo Fejl
    QActivePartnership = Application.Run("PQActivePartnership")
    
    GoTo Slut
Fejl:
    QActivePartnership = False
Slut:
End Function
Sub Test67()
On Error Resume Next
    Application.Run "TestWordMatP2"
End Sub
