VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGrupper 
   Caption         =   "Grupper i intervaller"
   ClientHeight    =   4440
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   7695
   OleObjectBlob   =   "UserFormGrupper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGrupper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub CommandButton_nulstil_Click()
    TextBox_data.Text = ""
    TextBox_intervaller.Text = ""
End Sub

Private Sub CommandButton_ok_Click()
Dim Data As String
Dim intervaller As String
Dim v As Double
Dim Min As Double
Dim Max As Double
Dim n As Integer
Dim Arr As Variant
Dim arrdata As Variant
Dim arrint As Variant
Dim i As Integer
Dim j As Integer
On Error GoTo Fejl
Me.Hide
'Application.ScreenUpdating = False
Data = TextBox_data.Text
intervaller = TextBox_intervaller.Text
arrdata = Split(Data, vbCrLf)
arrint = Split(intervaller, vbCrLf)
n = UBound(arrint)

If n > 0 Then
' indsæt data
If CheckBox_data.Value = True Then
    Dim t As String
    t = AppNavn & " grupperede talmængden: { " & Replace(Data, vbCrLf, " " & ListSeparator & " ") & " }" & vbCrLf
    t = t & "i intervallerne: " & Replace(intervaller, vbCrLf, " " & ListSeparator & " ") & vbCrLf
    Selection.TypeText t
    Selection.Collapse wdCollapseEnd
End If

' indsæt tabel
Dim Tabel As Table
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=UBound(arrint) + 2, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
Set Tabel = Selection.Tables(1)
    With Tabel
'        If .Style <> "Tabel - Gitter" Then
'            .Style = "Tabel - Gitter"
'        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With

Tabel.Cell(1, 1).Range.Text = "Fra"
Tabel.Cell(1, 2).Range.Text = "Til"
Tabel.Cell(1, 3).Range.Text = "Hyppighed"
Tabel.Columns.Width = 50

' fyld tabel med data
For i = 0 To n
    Arr = Split(arrint(i), "-")
    Min = Arr(0)
    If UBound(Arr) = 0 Then
        If Trim(Arr(0)) = "" Then
            
        Else
            GoTo Fejl
        End If
    End If
    Max = Arr(1)
    Tabel.Cell(i + 2, 1).Range.Text = Min
    Tabel.Cell(i + 2, 2).Range.Text = Max
    Tabel.Cell(i + 2, 3).Range.Text = 0
    For j = 0 To UBound(arrdata)
        On Error Resume Next
        If Trim(arrdata(j)) <> "" Then
            arrdata(j) = Replace(arrdata(j), ",", ".")
            v = CDbl(arrdata(j))
            v = val(arrdata(j))
            On Error GoTo Fejl
            If v > Min And v <= Max Then
                Tabel.Cell(i + 2, 3).Range.Text = val(Tabel.Cell(i + 2, 3).Range.Text) + 1
            End If
        End If
    Next
Next
End If

GoTo Slut
Fejl:
    MsgBox "Der er en fejl i dine intervaller. Hver linje skal indeholde interval f.eks. 5-10.", vbOKOnly, Sprog.Error
Slut:
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Unload Me
'  If CloseMode = 0 Then
'    Cancel = 1
'    Me.Hide
'  End If
End Sub
