Attribute VB_Name = "ModuleP"
Option Explicit
' Her kaldes funktioner, der kr�ves WordMat partnerskab
' Application.run "Funktion"  kan error h�ndteres, overf�re parametre og returnere v�rdier

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
    Application.Run "TestWordMatP"

End Sub
