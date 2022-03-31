Attribute VB_Name = "TestModul"
Option Explicit
    Private visok As Boolean
Sub TestTabel()
    UserFormTest.Show
End Sub
Sub TestSub()
    Application.ScreenUpdating = False
    Dim scrollpos As Double
    
    PrepareMaxima
    
    visok = True
    MaximaExact = 1

    TestSolve "x^2=9", "x", "x=-3    " & VBA.ChrW(8744) & "    x=3"
    TestSolve VBA.ChrW(8747) & "_0^a" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "x^2 dx=4""", "a", "a=9334/4077"
    TestBeregn VBA.ChrW(8747) & "_(-" & VBA.ChrW(8734) & ")^" & VBA.ChrW(8734) & "" & VBA.ChrW(9618) & "1/(" & VBA.ChrW(8730) & "2" & VBA.ChrW(960) & "á3)áe^(-1/2á((y-1)/3)^2 ) dy", "=1"
    
    ' slut
    Selection.TypeText (" - Test f*ae*rdig - ")
    Selection.TypeParagraph
    
End Sub
Sub TestBeregn(komm As String, resul As String)
    omax.Kommando = komm
    omax.beregn
    If Not omax.MaximaOutput = resul Then
        Selection.TypeText (" - Fejl - " & omax.Kommando & " - " & omax.MaximaOutput)
        Selection.TypeParagraph
    ElseIf visok Then
        Selection.TypeText (" - ok - " & omax.Kommando)
        Selection.TypeParagraph
    End If
    
End Sub

Sub TestSolve(komm As String, var As String, resul As String)
    omax.Kommando = komm
    omax.MaximaSolve (var)
    If Not omax.MaximaOutput = resul Then
        Selection.TypeText (" - Fejl - " & omax.Kommando & " - " & omax.MaximaOutput)
        Selection.TypeParagraph
    ElseIf visok Then
        Selection.TypeText (" - ok - " & omax.Kommando)
        Selection.TypeParagraph
    End If

End Sub

Sub testGetListItem()
Dim ea As New ExpressionAnalyser

ea.text = "dette er[ad;sdfs] en test ; hej(a;b{1;2}) ;{a;d} hallo"

MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()

End Sub

Sub ShowDef()
Dim s As String
Dim i As Integer

For i = 0 To omax.defindex - 1
    s = s & omax.DefName(i) & " - " & omax.DefValue(i) & vbCrLf
Next

MsgBox s

End Sub


Sub OpenJanis()
Dim ReturnValue
Dim sti As String
    sti = """" & GetProgramFilesDir & "\WordMat\janis.jnlp"""

ReturnValue = Shell(sti, vbNormalFocus)
End Sub
