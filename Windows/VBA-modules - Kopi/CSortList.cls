VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSortList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dname() As String
Private dval() As String
Public Length As Integer

Private Sub Class_Initialize()
    Length = 0
End Sub

Public Function GetName(i As Integer)
    GetName = dname(i)
End Function
Public Function GetVal(i As Integer)
    GetVal = dval(i)
End Function
Public Sub Add(Name As String, val As String, Optional i As Integer = -1)
    Length = Length + 1
    ReDim Preserve dname(Length)
    ReDim Preserve dval(Length)
    If i = -1 Then i = Length

    MoveUp (i)
    dname(i) = Name
    dval(i) = val

End Sub

Private Sub MoveUp(i As Integer)
For j = Length - 1 To i Step -1
    dname(j + 1) = dname(j)
    dval(j + 1) = dval(j)
Next
End Sub
