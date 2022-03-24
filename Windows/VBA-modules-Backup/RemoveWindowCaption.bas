Attribute VB_Name = "RemoveWindowCaption"
Option Explicit
#If Mac Then
    Sub RemoveCaption(objForm As Object)
    End Sub
#Else
Private Declare PtrSafe Function FindWindow Lib "user32" _
Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As LongPtr
 
Private Declare PtrSafe Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" ( _
ByVal hwnd As LongPtr, _
ByVal nIndex As LongPtr) As LongPtr
 
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" _
Alias "GetWindowLongA" ( _
ByVal hwnd As LongPtr, _
ByVal nIndex As LongPtr) As LongPtr
 
Private Declare PtrSafe Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As LongPtr, _
ByVal nIndex As LongPtr, _
ByVal dwNewLong As LongPtr) As LongPtr
 
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As LongPtr, _
ByVal nIndex As LongPtr, _
ByVal dwNewLong As LongPtr) As LongPtr
 
Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
ByVal hwnd As LongPtr) As LongPtr

Sub RemoveCaption(objForm As Object)
' skal køres i initialize af form
' Call RemoveCaption(Me)
Dim mhWndForm As Variant
Dim lStyle As Variant
On Error GoTo slut

'    Dim lStyle          As LongPtr
'    Dim hMenu           As LongPtr
'    Dim mhWndForm       As Long ' bedre som variant da den understøttes af alle word

    If val(Application.Version) < 14 Then
'        mhWndForm = FindWindow("ThunderXFrame", objForm.Caption) 'XL97
        mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
        lStyle = GetWindowLong(mhWndForm, -16)
        lStyle = lStyle And Not &HC00000
        SetWindowLong mhWndForm, -16, lStyle
        DrawMenuBar mhWndForm
    Else
'        RemoveCaption64bit (objForm)
    mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    lStyle = GetWindowLongPtr(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLongPtr mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
    End If
slut:
End Sub

Sub RemoveCaption64bit(objForm As Object)
' skal køres i initialize af form
' Call RemoveCaption(Me)
'Dim mhWndForm As Variant
Dim lStyle As Variant

'On Error GoTo slut
     
'    Dim lStyle          As LongPtr
'    Dim hMenu           As LongPtr
#If Win64 Then
    Dim mhWndForm       As LongPtr
#Else
    Dim mhWndForm       As Long
#End If

    mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    lStyle = GetWindowLongPtr(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLongPtr mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm

slut:
End Sub
#End If

