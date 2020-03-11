Attribute VB_Name = "Grafik"
Option Explicit
Sub InsertBitmap()
On Error GoTo Fejl
#If Mac Then
    MsgBox "Drawing using Paint is not supported on Mac" & vbCrLf & "Use Insert | Autoshapes " & vbCrLf & vbCrLf & "There is a freehand drawing tool"
#Else
    Selection.InlineShapes.AddOLEObject ClassType:="Paint.Picture", FileName:="", LinkToFile:=False, DisplayAsIcon:=False
#End If
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub

