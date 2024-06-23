Attribute VB_Name = "WolframAlpha"
Option Explicit
Sub OpenWolframAlpha(Optional pretext As String)
Dim text As String
On Error GoTo fejl
    PrepareMaxima
    omax.ReadSelection
    If omax.AntalKom = 1 Then
        text = omax.Kommando
    Else
        text = omax.KommandoerStreng
    End If
    text = omax.CodeForMaxima(text)
    text = ConvertToWolfram(text)
    text = pretext & " " & text
    
    OpenLink ("http://www.wolframalpha.com/input/?i=" & text)
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function ConvertToWolfram(text As String)
    
    text = Replace(text, "%", "") ' først fjernes evt. % fra maxima konstanter
    text = Replace(text, "+", "%2B") '
    text = Replace(text, "^", "%5E") ' ikke nødvendig?
    text = Replace(text, "(", "%28") '
    text = Replace(text, ")", "%29") '
'    text = Replace(text, "", "")

    

    ConvertToWolfram = text
End Function
