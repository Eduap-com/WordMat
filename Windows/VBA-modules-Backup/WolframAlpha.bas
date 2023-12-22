Attribute VB_Name = "WolframAlpha"
Option Explicit
Sub OpenWolframAlpha(Optional pretext As String)
Dim Text As String
On Error GoTo fejl
    PrepareMaxima
    omax.ReadSelection
    If omax.AntalKom = 1 Then
        Text = omax.Kommando
    Else
        Text = omax.KommandoerStreng
    End If
    Text = omax.CodeForMaxima(Text)
    Text = ConvertToWolfram(Text)
    Text = pretext & " " & Text
    
    OpenLink ("http://www.wolframalpha.com/input/?i=" & Text)
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function ConvertToWolfram(Text As String)
    
    Text = Replace(Text, "%", "") ' først fjernes evt. % fra maxima konstanter
    Text = Replace(Text, "+", "%2B") '
    Text = Replace(Text, "^", "%5E") ' ikke nødvendig?
    Text = Replace(Text, "(", "%28") '
    Text = Replace(Text, ")", "%29") '
'    text = Replace(text, "", "")

    

    ConvertToWolfram = Text
End Function
