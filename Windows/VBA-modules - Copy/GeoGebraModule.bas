Attribute VB_Name = "GeoGebraModule"
Option Explicit
' duer ikke med 64bit
'#If VBA7 Then
'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
'#Else
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#End If

Public GeoGebraDefs As String
Public GeoGebraAssumes As String

Sub GeoGebraWeb(Optional Gtype As String = "", Optional CASfunc As String = "")
    ' gtype="", "3d", "CAS"
    Dim Cmd As String, UrlLink As String
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim k As Integer, i As Integer, j As Integer, p As Integer
    Dim DefS As String
    Dim Arr As Variant
    Dim fktnavn As String, udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String

    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
    ea2.SetNormalBrackets

    On Error GoTo fejl

    PrepareMaxima
    omax.ConvertLnLog = False
    omax.ReadSelection

    
    ' s*ae*t definitioner i rigtig r*ae*kkef*oe*lge
    For i = 0 To omax.defindex - 1
        DefList = DefList & "," & omax.DefName(i)
        ea.text = omax.DefValue(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next
    
    ' definer variable der ikke er defineret
    omax.FindVariable
    ea.text = DefList
    For i = 0 To sl.Length - 1
        fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
        If sl.GetVal(i) <> ReplacedVar Then
            DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
            p = InStr(sl.GetName(i), "(")
            If p > 0 Then
                Cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
            Else
                Cmd = sl.GetName(i) & "=" & fktudtryk
            End If
        Else
            Cmd = sl.GetName(i) & "=" & fktudtryk
        End If
        Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
        UrlLink = UrlLink & Cmd
    Next
    
    j = 1
    ' inds*ae*t de markerede funktioner
    For i = 0 To omax.KommandoArrayLength
        udtryk = omax.KommandoArray(i)
        udtryk = Replace(udtryk, "definer:", "")
        udtryk = Replace(udtryk, "Definer:", "")
        udtryk = Replace(udtryk, "define:", "")
        udtryk = Replace(udtryk, "Define:", "")
        udtryk = Replace(udtryk, VBA.ChrW(8788), "=") ' :=
        udtryk = Replace(udtryk, VBA.ChrW(8797), "=") ' tripel =
        udtryk = Replace(udtryk, VBA.ChrW(8801), "=") ' def =
        udtryk = Trim(udtryk)
        If Gtype <> "CAS" Then
            If Len(udtryk) > 0 Then
                If InStr(udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                    If InStr(udtryk, "=") > 0 Then
                        Arr = Split(udtryk, "=")
                        lhs = Arr(0)
                        rhs = Arr(1)
                        ea.text = lhs
                        fktnavn = ea.GetNextVar(1)
                        varnavn = ea.GetNextBracketContent(1)
                    
                        If lhs = fktnavn & "(" & varnavn & ")" Then
                            ea.text = rhs
                            ea.pos = 1
                            ea.ReplaceVar varnavn, "x"
                            fktudtryk = ea.text
                            DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        
                            Cmd = fktnavn & "(x)=" & fktudtryk
                            Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
                            UrlLink = UrlLink & Cmd

                        Else
                            fktudtryk = ReplaceIndepvarX(rhs)
                            DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                            If fktnavn = "y" Then
                                Cmd = fktnavn & "=" & fktudtryk
                            Else
                                Cmd = fktnavn & "(x)=" & fktudtryk
                            End If
                            Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
                            UrlLink = UrlLink & Cmd
                            j = j + 1
                        End If
                    ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                        DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                        Cmd = "u" & j & "=" & udtryk
                        Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
                        UrlLink = UrlLink & Cmd
                        '                    geogebrafil.CreateFunction "u" & j, udtryk, True
                    Else
                        udtryk = ReplaceIndepvarX(udtryk)
                        DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                        Cmd = "f" & j & "=" & udtryk
                        Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
                        UrlLink = UrlLink & Cmd

                        '                    geogebrafil.CreateFunction "f" & j, udtryk, False
                        j = j + 1
                    End If
                End If
            End If
        Else 'CAS
            If CASfunc <> "" Then
                udtryk = CASfunc & "(" & udtryk & ")"
            End If
            Cmd = Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ";"
            UrlLink = UrlLink & Cmd
        End If
    Next

    'Punkter
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, setdata As String
        Cregr.GetTableData
        For j = 1 To UBound(Cregr.XValues)
            '            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
            '            cmd = "(" & Replace(Replace(geogebrafil.ConvertToGeoGebraSyntax(Cregr.XValues(j)), "+", "%2B"), ",", ".") & "," & Replace(Replace(geogebrafil.ConvertToGeoGebraSyntax(Cregr.YValues(j)), "+", "%2B"), ",", ".") & ")"
            Cmd = "(" & Replace(ConvertToGeogebraSyntax(Cregr.XValues(j)), "+", "%2B") & "," & Replace(ConvertToGeogebraSyntax(Cregr.YValues(j)), "+", "%2B") & ")"
            setdata = setdata & Cmd & ","
            UrlLink = UrlLink & Cmd & ";"
        Next
        setdata = Left(setdata, Len(setdata) - 1)
        setdata = "{" & setdata & "}"
        UrlLink = UrlLink & setdata & ";"
        '        geogebrafil.CreateList "punkter", setdata
    End If
            
    '    MsgBox UrlLink & cmd
    '    OpenLink UrlLink, True
    OpenGeoGebraWeb UrlLink, Gtype
fejl:

slut:
End Sub

Sub OpenGeoGebraWeb(ByVal Cmd As String, Gtype As String, Optional ConvertSyntax As Boolean = False, Optional UseDefs As Boolean = True)
    Dim UrlLink As String, Arr() As String, c As Variant, ArrDef() As String, ArrCas() As String, i As Integer, AssumeString As String
    Dim DefS As String
       
     If UseDefs Then
        FindGeoGebraDefsAndAssumes
        If GeoGebraDefs <> "" Then
            ArrDef = Split(GeoGebraDefs, ";")
            For i = 0 To UBound(ArrDef)
                DefS = DefS & ArrDef(i) & ";"
            Next
        End If
        If GeoGebraAssumes <> "" Then
            AssumeString = "Assume(" & GeoGebraAssumes
        End If
    End If
    
    If Gtype = "CAS" Then
        ArrCas = Split(Cmd, ";")
        Cmd = ""
        For i = 0 To UBound(ArrCas)
            ArrCas(i) = ConvertToGeogebraSyntax(ArrCas(i), ConvertSyntax)
            If AssumeString <> "" Then
                    Cmd = Cmd & AssumeString & " , " & ArrCas(i) & ");"
            Else
                    Cmd = Cmd & ArrCas(i) & ";"
            End If
        Next
    End If
    
    If Len(Cmd) > 0 Then If right(Cmd, 1) = ";" Then Cmd = Left(Cmd, Len(Cmd) - 1)
'    If ConvertSyntax Then Cmd = ConvertToGeogebraSyntax(Cmd, True)
    Cmd = DefS & Cmd
    Cmd = Replace(Cmd, "+", "%2B")
        
#If Mac Then
'    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html"
    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra" & Gtype & "Applet.html"
#Else
'    UrlLink = "https://geogebra.org/calculator"
'    UrlLink = "file:///C:/Program%20Files%20(x86)/WordMat/geogebra-math-apps/GeoGebraApplet.html"
    UrlLink = "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebra" & Gtype & "Applet.html"
#End If
    UrlLink = UrlLink & "?command=" & Cmd

    OpenLink UrlLink, True

End Sub

Sub FindGeoGebraDefsAndAssumes()
' s*ae*tter stregene GeoGebraDefs og GeoGebraAssumes ud fra omax
Dim Arr() As String, i As Integer
'    MsgBox omax.DefString
'    MsgBox omax.defstringtext ' alle definitioner og assumes, men i maxima syntaks
    
    GeoGebraAssumes = ""
    GeoGebraDefs = ""
    
    ' Tilf*oe*j definitioner og assumes der er skrevet i mat-felter
    Arr = Split(omax.DefString, "$")
    For i = 0 To UBound(Arr) - 1
        If Left(Arr(i), 7) = "assume(" Then
            GeoGebraAssumes = GeoGebraAssumes & Mid(Arr(i), 8, Len(Arr(i)) - 8) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(Arr(i), False) & ";"
        End If
    Next
    
    ' Tilf*oe*j midlertidige definitioner og assumes
    Arr = Split(omax.TempDefs, ";") ' alle midlertidige definitioner og assumes adskilt ad semikolon
    For i = 0 To UBound(Arr)
        If InStr(Arr(i), ">") > 0 Or InStr(Arr(i), "<") > 0 Then
            GeoGebraAssumes = GeoGebraAssumes & Arr(i) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(Arr(i)) & ";"
        End If
    Next
    If GeoGebraAssumes <> "" Then GeoGebraAssumes = Left(GeoGebraAssumes, Len(GeoGebraAssumes) - 1)
    If GeoGebraDefs <> "" Then GeoGebraDefs = Left(GeoGebraDefs, Len(GeoGebraDefs) - 1)

End Sub

Public Function GetDefsForGeoGebra(Optional ConvertHTML As Boolean = True) As String
' bruges ikke mere, men m*aa*ske er metoden bedre, da den g*oe*r noget med r*ae*kkef*oe*lgen
    Dim Cmd As String, i As Integer, k As Integer, j As Integer, fktudtryk As String, UrlLink As String, p As Integer
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim ea As New ExpressionAnalyser, Arr() As String
    Dim geogebrafil As New CGeoGebraFile
    
    ea.SetNormalBrackets

    ' s*ae*t definitioner i rigtig r*ae*kkef*oe*lge
    For i = 0 To omax.defindex - 1
        ea.text = omax.DefValue(i)
        DefList = DefList & "," & omax.DefName(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next
    
    
       
    ' definer variable der ikke er defineret
    ea.text = DefList
    For i = 0 To sl.Length - 1
'        If InStr(sl.GetVal(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
                p = InStr(sl.GetName(i), "(")
                If p > 0 Then
                    Cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
                Else
                    Cmd = sl.GetName(i) & "=" & fktudtryk
                End If
            Else
                Cmd = sl.GetName(i) & "=" & fktudtryk
            End If
            Cmd = ConvertToGeogebraSyntax(Cmd, False)
            If ConvertHTML Then Cmd = Replace(Cmd, "+", "%2B")
            Cmd = Cmd & ";"
            UrlLink = UrlLink & Cmd
'        End If
    Next
    
    If UrlLink <> "" Then
        If right(UrlLink, 1) = ";" Then UrlLink = Left(UrlLink, Len(UrlLink) - 1)
    End If
    GetDefsForGeoGebra = UrlLink
End Function
Function RunGeoGebraDirect(ByVal Cmd As String, Optional UseDefs As Boolean = True) As String
'        If Defliste <> "" Then Defliste = Defliste & ";"
        Dim Res As String, i As Integer, s As String
        Dim UfWait2 As UserFormWaitForMaxima
'        On Error Resume Next

        If UseDefs Then FindGeoGebraDefsAndAssumes
        
        If Cmd = "" Then Cmd = omax.Kommando
        Cmd = ConvertToGeogebraSyntax(Cmd)
        Res = ExecuteGeoGebraCasCommand(Cmd, UseDefs)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd)
        i = 0
        If Left(Res, 9) = "reloading" Then
            Set UfWait2 = New UserFormWaitForMaxima
            UfWait2.Label_tip.Font.Size = 10
            UfWait2.Label_tip.Font.Italic = False
            UfWait2.Show vbModeless
            UfWait2.Label_tip.Caption = Sprog.A(684)
            UfWait2.Label_progress.Caption = "*"
#If Mac Then
            Do
                Wait (0.2)
                UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
                If UfWait2.StopNow Then GoTo slut
                s = RunScript("IsGeoGebraAppReady", "")
                If Left(s, 3) = "yes" Then Exit Do
                i = i + 1
            Loop Until i > 50
            Wait (1)
#End If
            UfWait2.Label_tip.Caption = "Executing GeoGebra command"
            Res = ExecuteGeoGebraCasCommand(Cmd)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
        ElseIf Left(Res, 7) = "opening" Or Res = "null" Then
            Set UfWait2 = New UserFormWaitForMaxima
            UfWait2.Label_tip.Font.Size = 10
            UfWait2.Label_tip.Font.Italic = False
            UfWait2.Show vbModeless
            UfWait2.Label_tip.Caption = Sprog.A(684)
            UfWait2.Label_progress.Caption = "*"
#If Mac Then
            Do
                Wait (0.2)
                UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
                If UfWait2.StopNow Then GoTo slut
                s = RunScript("IsGeoGebraAppReady", "")
                If Left(s, 3) = "yes" Then Exit Do
                i = i + 1
            Loop Until i > 50
            Wait (1)
#Else
'            Wait (1)
#End If
            UfWait2.Label_tip.Caption = "Executing GeoGebra command"
            Res = ExecuteGeoGebraCasCommand(Cmd)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
        ElseIf Left(Res, 5) = "error" Then
'            Wait (1)
            GoTo slut
        End If
'        Res = Replace(Res, " ", "")
        Res = ConvertGeoGebraSyntaxToWord(Res)
        omax.MaximaOutput = Res
        RunGeoGebraDirect = Res
slut:
     If Not UfWait2 Is Nothing Then Unload UfWait2
End Function

Function ConvertToGeogebraSyntax(ByVal text As String, Optional ConvertMaxima As Boolean = True, Optional HtmlReady As Boolean = False) As String
'    Dim geogebrafil As New CGeoGebraFile
' definitioner vil allerede v*ae*re k*oe*rt igennem codeforMaxima, s*aa* der skal convertmaxima v*ae*re false

   Dim p As Integer, p2 As Integer, Arr() As String, p3 As Integer, sp As Integer, ep As Integer
   Dim ea As ExpressionAnalyser, s As String, t As String, gexpr As String, i As Integer, n As Integer
   Set ea = New ExpressionAnalyser
   ea.SetNormalBrackets
    
    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Trim(text)
    

 '  text = Replace(text, "log", "lg")
   If ConvertMaxima Then
      text = omax.CodeForMaxima(text, CASengine)
   End If

      'gr*ae*ske bogstaver
      text = Replace(text, "Delta", VBA.ChrW(916))
      text = Replace(text, "delta", VBA.ChrW(948))
      text = Replace(text, "alpha", VBA.ChrW(945))
      text = Replace(text, "beta", VBA.ChrW(946))
      text = Replace(text, "gamma_incomplete", "\Gamma_incomplete ")
      text = Replace(text, "gamma", VBA.ChrW(915))
      text = Replace(text, "gammaLB", VBA.ChrW(947))
      text = Replace(text, "theta", VBA.ChrW(952))
      text = Replace(text, "Theta", VBA.ChrW(920))
      text = Replace(text, "lambda", VBA.ChrW(955))
      text = Replace(text, "Lambda", VBA.ChrW(923))
      text = Replace(text, "rho", VBA.ChrW(961))
      text = Replace(text, "varphi", VBA.ChrW(966))
      text = Replace(text, "phi", VBA.ChrW(981))
      text = Replace(text, "Phi", VBA.ChrW(934))
      text = Replace(text, "varepsilon", VBA.ChrW(949))
      text = Replace(text, "epsilon", VBA.ChrW(1013))
      text = Replace(text, "psi", VBA.ChrW(968))
      text = Replace(text, "Psi", VBA.ChrW(936))
      text = Replace(text, "sigma", VBA.ChrW(963))
      text = Replace(text, "Sigma", VBA.ChrW(931))
      text = Replace(text, "mu", VBA.ChrW(956))
      text = Replace(text, "Ohm", VBA.ChrW(937))
      text = Replace(text, "Omega", VBA.ChrW(937))
      text = Replace(text, "omega", VBA.ChrW(969))
      text = Replace(text, "Xi", VBA.ChrW(926))
      text = Replace(text, "xi", VBA.ChrW(958))
      text = Replace(text, "Chi", VBA.ChrW(935))
      text = Replace(text, "chi", VBA.ChrW(967))
      text = Replace(text, "tau", VBA.ChrW(964))
      text = Replace(text, "Pi", VBA.ChrW(928))
      text = Replace(text, "greek-nu", VBA.ChrW(957))
      text = Replace(text, "kappa", VBA.ChrW(954))
      text = Replace(text, "zeta", VBA.ChrW(950))
      text = Replace(text, "eta", VBA.ChrW(951)) ' skal v*ae*re sidst da eta indg*aa*r i andre
      text = Replace(text, "increment", VBA.ChrW(8710))  ' specielt delta increment
      text = Replace(text, "Symhalf", VBA.ChrW(189)) ' _
      text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
      text = Replace(text, "<=", VBA.ChrW(8804))  '
      text = Replace(text, ">=", VBA.ChrW(8805)) '
      text = Replace(text, "CVinkelO", VBA.ChrW(8736)) '
      text = Replace(text, "CVinkel", VBA.ChrW(8736)) '
      text = Replace(text, "Symangle", VBA.ChrW(8736))  '
      text = Replace(text, "SymVecta", Sprog.A(683))
      
'      text = Replace(text, "Integrate", "Integral")  ' anden variabel end x godtages ikke i NIntegral, men ok i IntegralSymbolic
    
    
      text = Replace(text, "diff", "Derivative")  ' anden variabel end x godtages ikke i NIntegral, men ok i IntegralSymbolic
    
      'Else
      '    text = Replace(text, "log", "ln")
   
      ' integrate erstattes med Integral. I GeoGebra overs*ae*ttes Integral automatisk til NItegral eller IntegralSymbolic afh af omst*ae*ndigheder.
      ' Der ser dog ud til at v*ae*re et problem med Integral(f(x),x,0,1)  man kan ikke selv angive en variabel, selvom det st*aa*r at det er en mulighed
'   p = InStr(text, "Integrate")
'   Do While p > 0
'      If p > 0 Then
'        ea.text = text
'        s = ea.GetNextBracketContent(p + 9)
'        arr = Split(s, ",")
'        text = Left(text, p - 1) & "Integral(" & arr(0) & "," & arr(2) & "," & arr(3) & right(text, Len(text) - p - Len(s) - 9)
'      End If
'       p = InStr(text, "Integrate")
'   Loop
    
   
   ea.text = text
   ea.ReplaceVar "NIntegrate", "NIntegral"
   ea.ReplaceVar "Integrate", "Integral"
   ea.ReplaceVar "minf", "-infinity"
   ea.ReplaceVar "inf", "infinity"
   ea.ReplaceVar "log", "ln"
   ea.ReplaceVar "lg10", "log10"
   text = ea.text
   text = Replace(text, "%e", "exp(1)") 'VBA.ChrW(101)
   text = Replace(text, "%pi", VBA.ChrW(960)) '"pi"
   text = Replace(text, "%", "")
'   text = Replace(text, "[", "{")
'   text = Replace(text, "]", "}")
   text = Replace(text, "##", "*") 'prikprodukt
   text = Replace(text, "~", "*") ' vectorprodukt
'   text = Replace(text, "matrix", "")

'
      p = InStr(text, "logbase(")
      Do While p > 0
        If p > 0 Then
          ea.text = text
          s = ea.GetNextBracketContent(p + 7)
          Arr = Split(s, ",")
          If UBound(Arr) > 0 Then text = Left(text, p - 1) & "log(" & Arr(1) & "," & Arr(0) & right(text, Len(text) - p - Len(s) - 7)
        End If
        p = InStr(text, "logbase(")
      Loop


      p = InStr(text, "if")
      p2 = InStr(text, "then")
      If p > 0 And p2 > 0 Then
         sp = p
         ea.text = text
         ea.pos = p - 1
         s = ea.GetNextBracketContent()
         ep = p + Len(s) + 1
         p3 = 1
         Do
            p = InStr(p3, s, "if")
            If p <= 0 Then Exit Do
            p2 = InStr(p3, s, "then")
            p3 = InStr(p3, s, "else")
            If p3 <= 0 Then p3 = Len(s) + 1
            gexpr = gexpr & "If[" & Trim(Mid(s, p + 2, p2 - p - 2)) & "," & Trim(Mid(s, p2 + 4, p3 - p2 - 4)) & ","
            n = n + 1
            If p3 = Len(s) + 1 Then Exit Do
            p3 = p3 + 1
         Loop While p3 < Len(s)
         If right(gexpr, 1) = "," Then gexpr = Left(gexpr, Len(gexpr) - 1)
         For i = 1 To n
            gexpr = gexpr & "]"
         Next
         text = Left(text, sp - 1) & gexpr & right(text, Len(text) - ep + 1)
         If Left(text, 1) = "(" Then text = right(text, Len(text) - 1)
'         text = Replace(text, " and ", " ??¤ ") '&& der m*aa* v*ae*re sket noget fejlkonvertering
'         text = Replace(text, " or ", " ??¬ ") '||
         text = Replace(text, " and ", " && ") '&&
         text = Replace(text, " or ", " || ") '||
      End If

   ConvertToGeogebraSyntax = text
       
'    ConvertToGeogebraSyntax = geogebrafil.ConvertToGeogebraSyntax(s, True)
    If HtmlReady Then
        ConvertToGeogebraSyntax = Replace(ConvertToGeogebraSyntax, "+", "%2B")
    End If
'    Set geogebrafil = Nothing
End Function
Function ConvertGeoGebraSyntaxToWord(ByVal text As String) As String
    Dim p As Long, p2 As Long, ea As New ExpressionAnalyser, mtext As String, s As String, n As Integer
    ea.SetNormalBrackets
    
    text = TrimB(text, """")
    
    ' mellemrum mellem variable skal have *
    ea.text = text
    p = InStr(ea.text, " ")
    Do While p > 0
        If ea.IsLetterPos(p - 1) And ea.IsLetterPos(p + 1) Then
            ea.ReplaceStringAt "*", p
        End If
        p = InStr(p + 1, ea.text, " ")
    Loop
    
    'potenser kan v*ae*re skrevet som specialtegn for 0,1,2,3    p = InStr(ea.text, VBA.ChrW(176)) ' h*ae*vet 0
    ReplaceSuperScriptNo ea, 8314, "+"
    ReplaceSuperScriptNo ea, 8315, "-"
    ReplaceSuperScriptNo ea, 8317, "("
    ReplaceSuperScriptNo ea, 8318, ")"
    ReplaceSuperScriptNo ea, 8304, "0"
    ReplaceSuperScriptNo ea, 185, "1"
    ReplaceSuperScriptNo ea, 178, "2"
    ReplaceSuperScriptNo ea, 179, "3"
    For n = 4 To 9
        ReplaceSuperScriptNo ea, 8304 + n, CStr(n)
    Next
    
    
    text = ea.text
    text = Replace(text, "\u003C", "<") ' m*ae*rkeligt at lige pr*ae*cis denne st*aa*r s*aa*dan
    text = Replace(text, " = ", "=") ' mellemrum omkring =
    text = Replace(text, " < ", "<")
    text = Replace(text, " > ", ">")


    text = Replace(text, "sin^-1", "asin")
    text = Replace(text, "cos^-1", "acos")
    text = Replace(text, "tan^-1", "atan")

    text = omax.ConvertToWordSymbols(text)
    
    ea.text = text
    
    ReplaceTrigSuperscript ea, "sin"
    ReplaceTrigSuperscript ea, "cos"
    ReplaceTrigSuperscript ea, "tan"
        
    
    'matricer og vektorer
     ea.SetTuborgBrackets
    p = InStr(ea.text, "{{")
    Do While p > 0
        mtext = ea.GetNextBracketContent(p)
        mtext = Replace(mtext, " ", "") ' mellem efter ; giver problemer i n*ae*ste linje
        mtext = Replace(mtext, "}" & ListSeparator & "{", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = omax.matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & omax.matrixendbracket
        ea.text = Left(ea.text, p - 1) & mtext & right(ea.text, Len(ea.text) - ea.pos + 1)

        p = InStr(p + 1, ea.text, "{{")
    Loop
    
    p = InStr(ea.text, "_{")
    Do While p > 0
        s = ea.GetNextBracketContent(p)
        ea.text = Left(ea.text, p) & s & right(ea.text, Len(ea.text) - Len(s) - p - 2)
        p = InStr(p + 1, ea.text, "_{")
    Loop
    text = ea.text
    
    ConvertGeoGebraSyntaxToWord = text
End Function

Sub ReplaceSuperScriptNo(ByRef ea As ExpressionAnalyser, ChrNo As Integer, ChrS As String)
    Dim p As Integer
    p = InStr(ea.text, VBA.ChrW(ChrNo)) ' h*ae*vet -
    Do While p > 0
        If p > 0 Then
            ea.ReplaceStringAt ChrS, p
            If Not ea.IsNumberOrSign(p - 1) And Not ea.IsSuperScriptNoPos(p - 1) Then
                ea.InsertBeforePos "^", p
            End If
        End If
        p = InStr(ea.text, VBA.ChrW(ChrNo)) ' h*ae*vet -
    Loop
End Sub
Sub ReplaceTrigSuperscript(ByRef ea As ExpressionAnalyser, Trig As String)
' erstatter fx sin & chrw(8289) & "^2" med "sin" & "^2" & chrw(8289)
' alts*aa* retter op p*aa* output fra converttowordsymols
Dim p As Integer, p2 As Integer
    p = InStr(ea.text, Trig & VBA.ChrW(8289) & "^")
    Do While p > 0
        If p > 0 Then
            p2 = p
            p = p + Len(Trig) + 2
            Do While ea.IsNumberOrSign(p) And p <= ea.Length
                p = p + 1
            Loop
            ea.InsertTextAt VBA.ChrW(8289), p
            ea.RemoveChar (p2 + Len(Trig))
        End If
        p = InStr(ea.text, Trig & VBA.ChrW(8289) & "^")
    Loop
End Sub

Sub GeoGebra()
    On Error GoTo fejl
    Dim geogebrasti As String
    Dim geogebrafilersti As String
    Dim appnr As Integer
    Dim UFwait As New UserFormWaitForMaxima
    PrepareMaxima ' omax bliver brugt
    
    UFwait.Label_tip.Caption = Sprog.A(362)
    UFwait.Label_progress.Caption = "***"
    UFwait.CommandButton_stop.visible = False
    UFwait.Show vbModeless
    
    geogebrasti = GeoGebraPath()
    If geogebrasti = "" Then GoTo fejl ' hvis geogebra ikke installeret s*aa* fejl
'    geogebrafilersti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    geogebrafilersti = GetTempDir()
    
    CreateGeoGebraFil geogebrafilersti
    
    geogebrafilersti = geogebrafilersti & "geogebra.ggb"
    
    UFwait.Label_progress.Caption = "******"
    If FileExists(geogebrafilersti) Then ' check om geogebrafilen er lavet
        geogebrasti = geogebrasti & " """ & geogebrafilersti & """"
    Else
        MsgBox "The GeoGebra.ggb file cannot be located", vbOKOnly, Sprog.Error
        GoTo fejl
    End If
    
#If Mac Then
    RunScript "OpenGeoGebra", geogebrafilersti
#Else
    appnr = Shell(geogebrasti, vbNormalFocus)   'vbNormalFocus vbMinimizedFocus ' til offline installer
#End If
' til webstart:
'    appnr = Shell("javaws -system -open """ & geogebrafilersti & "geogebra.ggb""" & " http://www.geogebra.org/webstart/geogebra.jnlp", vbNormalFocus)    'vbNormalFocus vbMinimizedFocus
    
    UFwait.Label_progress.Caption = "*********"
    On Error Resume Next
    DoEvents
    Unload UFwait
    
    GoTo slut
fejl:
    UserFormGeoGebra.Show
slut:
End Sub
Function GeoGebraPath() As String
' path to the geogebra executable. Returns "" if not found
On Error GoTo fejl
#If Mac Then
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra 5.app"
    If FileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra.app"
    If FileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra 6.app"
    If FileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra Graphing Calculator.app"
    If FileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = ""
#Else
    
'    GeoGebraPath = GetProgramFilesDir & "\GeoGebra 4.2\GeoGebra.exe"
'    GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 5.*", vbDirectory)
    GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 5*", vbDirectory)
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 6*", vbDirectory)
    End If
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 4.*", vbDirectory)
    End If
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra Classic*", vbDirectory)
    End If
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra*", vbDirectory)
    End If
'    If GeoGebraPath = "" Then
'        GeoGebraPath = GetProgramFilesDir & "\WordMat\GeoGebra\GeoGebra.exe"
'    Else
    
    If Not GeoGebraPath = "" Then
        GeoGebraPath = GetProgramFilesDir & "\" & GeoGebraPath & "\GeoGebra.exe"
        GeoGebraPath = """" & GeoGebraPath & """"
    End If
#End If
    GoTo slut
fejl:
    GeoGebraPath = ""
slut:
End Function
Sub CreateGeoGebraFil(geogebrasti As String)
    Dim geogebrafil As New CGeoGebraFile
    Dim i As Integer, j As Integer
    Dim DefS As String
    Dim Arr As Variant
    Dim fktnavn As String, udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String
    Dim dd As New DocData
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    On Error GoTo fejl
    ea.SetNormalBrackets
    ea2.SetNormalBrackets
    geogebrafil.Show3D = False
    
    'geogebrasti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    omax.ConvertLnLog = False
#If Mac Then
    Wait 1 ' anden gang man *aa*bner geogebra p*aa* mac l*aa*ser word hvis denne ikke er der. Det har noget med finddefinitions at g*oe*re, men har ikke kunnet se hvad
#End If

    PrepareMaxima
'    omax.FindDefinitions ' bliver fundet i preparemaxima
    omax.ReadSelection

'    defs = ConvertToAscii(omax.defstring)
'    arr = Split(defs, ";")
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim k As Integer
    
    ' s*ae*t definitioner i rigtig r*ae*kkef*oe*lge
    For i = 0 To omax.defindex - 1
        ea.text = omax.DefValue(i)
        DefList = DefList & "," & omax.DefName(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next
    
    ' definer variable der ikke er defineret
    ea.text = DefList
    For i = 0 To sl.Length - 1
        If InStr(sl.GetVal(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), DefList, geogebrafil
            End If
            geogebrafil.CreateFunction sl.GetName(i), fktudtryk, False, False
        End If
    Next
    ' definer de definerede funktioner/variable
'    For i = 0 To sl.Length - 1
'    Next
    
    j = 1
    ' inds*ae*t de markerede funktioner
    For i = 0 To omax.KommandoArrayLength
        udtryk = omax.KommandoArray(i)
        udtryk = Replace(udtryk, "definer:", "")
        udtryk = Replace(udtryk, "Definer:", "")
        udtryk = Replace(udtryk, "define:", "")
        udtryk = Replace(udtryk, "Define:", "")
        udtryk = Replace(udtryk, VBA.ChrW(8788), "=") ' :=
        udtryk = Replace(udtryk, VBA.ChrW(8797), "=") ' tripel =
        udtryk = Replace(udtryk, VBA.ChrW(8801), "=") ' def =
        udtryk = Trim(udtryk)
        If Len(udtryk) > 0 Then
            If InStr(udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                If InStr(udtryk, "=") > 0 Then
                    Arr = Split(udtryk, "=")
                    lhs = Arr(0)
                    rhs = Arr(1)
                    ea.text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.text = rhs
                        ea.pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.text
                        DefinerKonstanter fktudtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction fktnavn, fktudtryk, False, True
                    Else
                        fktudtryk = ReplaceIndepvarX(rhs)
                        DefinerKonstanter udtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction "f" & j, udtryk, False
                        j = j + 1
                    End If
                ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                ' kan f*oe*rst bruges med GeoGebra 4.0
                    DefinerKonstanter udtryk, DefList, geogebrafil
                    geogebrafil.CreateFunction "u" & j, udtryk, True
                Else
                    udtryk = ReplaceIndepvarX(udtryk)
                    DefinerKonstanter udtryk, DefList, geogebrafil
                    geogebrafil.CreateFunction "f" & j, udtryk, False
                    j = j + 1
                End If

                
            End If
        End If
    Next
    
    'datapunkter
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, setdata As String
        Cregr.GetTableData
        For j = 1 To UBound(Cregr.XValues)
'            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
            setdata = setdata & "(" & Replace(Cregr.XValues(j), ",", ".") & "," & Replace(Cregr.YValues(j), ",", ".") & "),"
        Next
        setdata = Left(setdata, Len(setdata) - 1)
        setdata = "{" & setdata & "}"
        geogebrafil.CreateList "punkter", setdata
    End If
    
' Problemet med dette er at tabellen skal v*ae*re lodret
'    dd.ReadSelection
'    If Len(dd.GetSetForm) > 4 Then
'        geogebrafil.CreateList "punkter", dd.GetSetForm
'    End If

    If FileExists(geogebrasti & "geogebra.xml") Then
        Kill geogebrasti & "geogebra.xml"
    End If
    geogebrafil.Save geogebrasti & "geogebra.xml"
    
    DoEvents
    i = 0
    Do While i < 10 And Not (FileExists(geogebrasti & "geogebra.xml"))
        Wait 1
        i = i + 1
    Loop
'    CreateZipFile geogebrasti & "geogebra.zip", geogebrasti & "geogebra_thumbnail.png", geogebrasti & "geogebra.xml"
    If FileExists(geogebrasti & "geogebra.ggb") Then
        Kill geogebrasti & "geogebra.ggb"
    End If
#If Mac Then
    CreateZipFile geogebrasti & "geogebra.ggb", geogebrasti & "geogebra.xml"
#Else
    CreateZipFile geogebrasti & "geogebra.zip", geogebrasti & "geogebra.xml"
    Name geogebrasti & "geogebra.zip" As geogebrasti & "geogebra.ggb"
#End If
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
On Error Resume Next
    omax.ConvertLnLog = True
End Sub
Function DefinerKonstanter(Expr As String, ByRef DefList As String, Optional ByRef geogebrafil As CGeoGebraFile = Nothing, Optional ByRef UrlLink As String) As String
' definer variable der ikke er defineret i expr
' deflist er en liste af variable der er defineret
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim var As String, var2 As String, i As Integer, p As Integer
Dim varval As String
    ea.text = DefList
    ea2.text = Expr
    ea2.pos = 0
    Do
        var = ea2.GetNextVar
'        MsgBox AscW(var) & vbCrLf & VBA.ChrW(960)
        ea2.pos = ea2.pos + 1
'        If var = "z" Then geogebrafil.Show3D = True  ' ikke implementeret endnu
        If Not (ea2.ChrByIndex(ea2.pos) = "(") And Not (Left(var, 1) = "_") And Not (ea.IsFunction(var)) And Not (ea.ContainsVar(var)) And var <> "" And var <> "x" And var <> "y" And var <> "z" And var <> "e" And var <> "pi" And var <> "matrix" And var <> "if" And var <> "elseif" And var <> "then" And var <> "and" And var <> "or" And var <> "else" And var <> VBA.ChrW(960) Then  ' 960=pi
            varval = InputBox(Sprog.A(363) & " " & var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
            If Len(varval) > 0 Then
                If Not geogebrafil Is Nothing Then
                    geogebrafil.CreateFunction var, varval, False, False
                Else
                    UrlLink = UrlLink & var & "=" & varval & ";"
                End If
                DefList = DefList & "," & var
                i = i + 1
            End If
        End If
    Loop While var <> ""
    
    If i > 1 Then
        var = Split(DefList, ",")(0)
            varval = InputBox(Sprog.A(363) & " " & var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
            If Len(varval) > 0 Then
                If Not geogebrafil Is Nothing Then
                    geogebrafil.CreateFunction var, varval, False, False
                Else
                    UrlLink = UrlLink & var & "=" & varval & ";"
                End If
            End If
    End If
    
    DefinerKonstanter = var
End Function
Sub TestZipFile()

'    DateiZippen "C:\TEMP\Example.xml", "C:\TEMP\Example.zip"
    CreateZipFile "C:\WordMatGraf.zip", "C:\WordMatGraf.gif", "C:\WordMatGraf.jpg"

End Sub

 

Sub CreateZipFile(zipfilnavn As Variant, filnavn As Variant, Optional filnavn2 As Variant = "", Optional filnavn3 As Variant = "", Optional filnavn4 As Variant = "")
#If Mac Then
    RunScript "ZipFile", zipfilnavn & ";" & filnavn
#Else
'
' s. http://www.rondebruin.nl/windowsxpzip.htm

' Early binding, set reference to:
'   Microsoft Shell Controls and automation (C:\WINNT\systems32\SHELL32.dll)
        
' a)
    Dim oApp As Object 'Shell ' Early binding, late binding as in the example (Dim oApp As Object) didn't work. Fixed 16/4-2018. It works if the filename strings are defined as variants. zipfilnavn and filnavn
    Dim objFolder As Variant
'    Set oApp = New Shell
    Set oApp = CreateObject("Shell.Application")
    Dim i As Integer
' or b)
    
    NewZip (zipfilnavn)
    Set objFolder = oApp.Namespace(zipfilnavn)
    
    i = 1
    objFolder.CopyHere filnavn
    If filnavn2 <> "" Then
        objFolder.CopyHere filnavn2
        i = i + 1
    End If
    If filnavn3 <> "" Then
        objFolder.CopyHere filnavn3
        i = i + 1
    End If
    If filnavn4 <> "" Then
        objFolder.CopyHere filnavn4
        i = i + 1
    End If
    
    On Error Resume Next
    
    Do Until oApp.Namespace(zipfilnavn).Items.Count = i
       Wait 1
'        Sleep 1000 ' 64bit problemer
    Loop
    
'    On Error GoTo 0
#End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:

End Sub
 

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub



