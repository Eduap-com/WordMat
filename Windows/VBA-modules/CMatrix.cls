VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private tabel() As String
Private Tabel As Variant

Private Sub Class_Initialize()
    'ReDim tabel(0, 0)
    Tabel = Array(0, 0)
End Sub

Public Property Get Element(r As Integer, c As Integer) As Variant
On Error GoTo Fejl
If c > Columns Or r > Rows Then GoTo Slut

Element = Tabel(r, c)

GoTo Slut
Fejl:
Slut:

End Property

Public Property Let Element(r As Integer, c As Integer, nval As Variant)
    If r > Rows Then
        Rows = r
    End If
    If c > Columns Then
        Columns = c
    End If
    Tabel(r, c) = nval
End Property

Public Property Get Rows() As Integer
    Rows = UBound(Tabel, 1)
End Property
Public Property Let Rows(nval As Integer)
    If nval > Rows Then
        Set Tabel = ExpandArray(Tabel, 1, nval - Rows, 0)
'        ReDim tabel(nval, UBound(tabel, 2))
    End If
End Property
Public Property Get Columns() As Integer
    Columns = UBound(Tabel, 2)
End Property
Public Property Let Columns(nval As Integer)
    If nval > Columns Then
        Set Tabel = ExpandArray(Tabel, 2, nval - Columns, 0)
'    ReDim Preserve tabel(UBound(tabel, 1), nval)
    End If
End Property

Public Sub SetSpaceVector()
    ReDim Tabel(3, 1)
End Sub


