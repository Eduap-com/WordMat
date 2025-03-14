Attribute VB_Name = "VBAExtensions"
Option Explicit
Option Private Module
' Small general functions that extend the buildin VBA functions, or missing functions
' No functions specific to the project or dependent on other modules may appear in this module

Public Function GetCountOfChar(ByVal ar_sText As String, ByVal a_sChar As String) As Integer
    Dim l_iIndex As Integer
    Dim l_iMax As Integer
    Dim l_iLen As Integer

    GetCountOfChar = 0
    l_iMax = Len(ar_sText)
    l_iLen = Len(a_sChar)
    For l_iIndex = 1 To l_iMax
        If (Mid(ar_sText, l_iIndex, l_iLen) = a_sChar) Then 'found occurrence
            GetCountOfChar = GetCountOfChar + 1
            If (l_iLen > 1) Then l_iIndex = l_iIndex + (l_iLen - 1) 'if matching more than 1 char, need to move more than one char ahead to continue searching
        End If
    Next l_iIndex
End Function
Function TrimR(ByVal Text As String, c As String)
' fjerner c fra højre side af text
If Text = "" Then GoTo slut
Do While right(Text, 1) = c
    Text = Left(Text, Len(Text) - 1)
Loop
TrimR = Text
slut:
End Function
Function TrimL(ByVal Text As String, c As String)
' fjerner c fra venstre side af text
If Text = "" Then GoTo slut
Do While Left(Text, 1) = c
    Text = right(Text, Len(Text) - 1)
Loop
TrimL = Text
slut:
End Function

Function TrimB(ByVal Text As String, c As String)
' fjerner c fra Begge sider af text
TrimB = TrimL(Text, c)
TrimB = TrimR(TrimB, c)
slut:
End Function
Function TrimRenter(ByVal Text As String)
' removes crlf at right end
    TrimRenter = TrimR(TrimR(Text, vbLf), vbCr)
End Function

Sub Wait(pausetime As Variant)
'pausetime in senconds
Dim start
    start = Timer    ' Set start time.
    Do While Timer < start + pausetime
        DoEvents    ' Yield to other processes.
    Loop
End Sub

