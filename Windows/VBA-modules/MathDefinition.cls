VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MathDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum MathDefinitionType
    Variable
    Func
End Enum

Public DefinitionType As MathDefinitionType
Public Name As String
Public Var As String
Public Vars As New Collection
Public Value As String
Public MathObject As OMath


Function GetDefinitionString() As String
    If DefinitionType = Variable Then
        GetDefinitionString = Name & ":" & Value
    Else
        Dim s As String, v As Variant
        s = Variable
        For Each v In Variables
            s = s & "," & v
        Next
        GetDefinitionString = Name & "(" & s & "):=" & Value
    End If
End Function

Sub LoadFromString(s As String)

    

End Sub
