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
    Dim cmd As String, UrlLink As String
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim k As Integer, i As Integer, j As Integer, p As Integer
    Dim DefS As String
    Dim arr As Variant, uvar As String
    Dim fktnavn As String, udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String
    Dim TempCas As Integer

    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
    ea2.SetNormalBrackets

    On Error GoTo fejl

    TempCas = CASengine
    CASengine = 1

    PrepareMaxima
    omax.ConvertLnLog = False
    omax.ReadSelection

    If Gtype <> "CAS" Then
    ' sæt definitioner i rigtig rækkefølge
    For i = 0 To omax.defindex - 1
        DefList = DefList & "," & omax.DefName(i)
        ea.Text = omax.DefValue(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.Text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next

'    ' definer variable der ikke er defineret
    omax.FindVariable
    ea.Text = DefList
    For i = 0 To sl.Length - 1
        fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
        If sl.GetVal(i) <> ReplacedVar Then
            DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
            p = InStr(sl.GetName(i), "(")
            If p > 0 Then
                cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
            Else
                cmd = sl.GetName(i) & "=" & fktudtryk
            End If
        Else
            cmd = sl.GetName(i) & "=" & fktudtryk
        End If
        cmd = Replace(Replace(ConvertToGeogebraSyntax(cmd, False), "+", "%2B"), "&", "%26") & ";" ' v.1.26 tilføjet false ved konvertering fordi den vidst allerede er konverteret
        UrlLink = UrlLink & cmd
   Next
End If

    j = 1
    ' indsæt de markerede funktioner
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
        udtryk = ConvertToGeogebraSyntax(udtryk)
        If Gtype <> "CAS" Then
            If Len(udtryk) > 0 Then
                If InStr(udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                    If InStr(udtryk, "=") > 0 Then
                        arr = Split(udtryk, "=")
                        lhs = arr(0)
                        rhs = arr(1)
                        ea.Text = lhs
                        fktnavn = ea.GetNextVar(1)
                        varnavn = ea.GetNextBracketContent(1)
                    
                        If lhs = fktnavn & "(" & varnavn & ")" Then
                            ea.Text = rhs
                            ea.Pos = 1
                            ea.ReplaceVar varnavn, "x"
                            fktudtryk = ea.Text
                            DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        
                            cmd = fktnavn & "(x)=" & fktudtryk
                            cmd = Replace(cmd, "+", "%2B") & ";"
                            UrlLink = UrlLink & cmd

                        Else
                            fktudtryk = ReplaceIndepvarX(rhs, uvar)
                            If Not (uvar = "" Or uvar = "x") Then  'Or uvar = "t"
                                DefList = DefList & uvar & ","
                            End If
                            DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                            If fktnavn = "y" Then
                                cmd = fktnavn & "=" & fktudtryk
                            Else
                                cmd = fktnavn & "(x)=" & fktudtryk
                            End If
                            cmd = Replace(cmd, "+", "%2B") & ";"
                            UrlLink = UrlLink & cmd
                            j = j + 1
                        End If
                    ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                        DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                        cmd = "u" & j & "=" & udtryk
                        cmd = Replace(cmd, "+", "%2B") & ";"
                        UrlLink = UrlLink & cmd
                        '                    geogebrafil.CreateFunction "u" & j, udtryk, True
                    Else
                        udtryk = ReplaceIndepvarX(udtryk)
                        DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                        cmd = "f" & j & "=" & udtryk
                        cmd = Replace(cmd, "+", "%2B") & ";"
                        UrlLink = UrlLink & cmd

                        '                    geogebrafil.CreateFunction "f" & j, udtryk, False
                        j = j + 1
                    End If
                End If
            End If
        Else 'CAS
            If CASfunc <> "" Then
                udtryk = CASfunc & "(" & udtryk & ")"
            End If
            cmd = Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ";"
            UrlLink = UrlLink & cmd
        End If
    Next

    'Punkter
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, setdata As String
        Cregr.GetTableData
        For j = 1 To UBound(Cregr.XValues)
            '            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
            '            cmd = "(" & Replace(Replace(geogebrafil.ConvertToGeoGebraSyntax(Cregr.XValues(j)), "+", "%2B"), ",", ".") & "," & Replace(Replace(geogebrafil.ConvertToGeoGebraSyntax(Cregr.YValues(j)), "+", "%2B"), ",", ".") & ")"
            cmd = "(" & Replace(ConvertToGeogebraSyntax(Cregr.XValues(j)), "+", "%2B") & "," & Replace(ConvertToGeogebraSyntax(Cregr.YValues(j)), "+", "%2B") & ")"
            setdata = setdata & cmd & ","
            UrlLink = UrlLink & cmd & ";"
        Next
        setdata = Left(setdata, Len(setdata) - 1)
        setdata = "{" & setdata & "}"
        UrlLink = UrlLink & setdata & ";"
        '        geogebrafil.CreateList "punkter", setdata
    End If
            
    '    MsgBox UrlLink & cmd
    '    OpenLink UrlLink, True
    
    OpenGeoGebraWeb UrlLink, Gtype, False, False ' v.1.26 false, false tilføjet da definitioner kom med to gange
fejl:

slut:
    CASengine = TempCas
End Sub

Sub OpenGeoGebraWeb(ByVal cmd As String, Gtype As String, Optional ConvertSyntax As Boolean = False, Optional UseDefs As Boolean = True)
    Dim UrlLink As String, arr() As String, c As Variant, ArrDef() As String, ArrCas() As String, i As Integer, AssumeString As String
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
    
    If Gtype = "CAS" Or Gtype = "Classic" Then
        ArrCas = Split(cmd, ";")
        cmd = ""
        For i = 0 To UBound(ArrCas)
            ArrCas(i) = ConvertToGeogebraSyntax(ArrCas(i), ConvertSyntax)
            If AssumeString <> "" Then
                    cmd = cmd & AssumeString & " , " & ArrCas(i) & ");"
            Else
                    cmd = cmd & ArrCas(i) & ";"
            End If
        Next
    End If
    
    If Len(cmd) > 0 Then If right(cmd, 1) = ";" Then cmd = Left(cmd, Len(cmd) - 1)
'    If ConvertSyntax Then Cmd = ConvertToGeogebraSyntax(Cmd, True)
    cmd = DefS & cmd
    cmd = Replace(cmd, "+", "%2B")
        
#If Mac Then
'    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html"
    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra" & Gtype & "Applet.html"
#Else
'    UrlLink = "https://geogebra.org/calculator"
'    UrlLink = "file:///C:/Program%20Files%20(x86)/WordMat/geogebra-math-apps/GeoGebraApplet.html"
    UrlLink = "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebra" & Gtype & "Applet.html"
#End If
    UrlLink = UrlLink & "?command=" & cmd

    OpenLink UrlLink, True

End Sub

Sub FindGeoGebraDefsAndAssumes()
' sætter stregene GeoGebraDefs og GeoGebraAssumes ud fra omax
Dim arr() As String, i As Integer
'    MsgBox omax.DefString
'    MsgBox omax.defstringtext ' alle definitioner og assumes, men i maxima syntaks
    
    GeoGebraAssumes = ""
    GeoGebraDefs = ""
    
    ' Tilføj definitioner og assumes der er skrevet i mat-felter
    arr = Split(omax.DefString, "$")
    For i = 0 To UBound(arr) - 1
        If Left(arr(i), 7) = "assume(" Then
            GeoGebraAssumes = GeoGebraAssumes & Mid(arr(i), 8, Len(arr(i)) - 8) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(arr(i), False) & ";"
        End If
    Next
    
    ' Tilføj midlertidige definitioner og assumes
    arr = Split(omax.TempDefs, ";") ' alle midlertidige definitioner og assumes adskilt ad semikolon
    For i = 0 To UBound(arr)
        If InStr(arr(i), ">") > 0 Or InStr(arr(i), "<") > 0 Then
            GeoGebraAssumes = GeoGebraAssumes & arr(i) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(arr(i)) & ";"
        End If
    Next
    If GeoGebraAssumes <> "" Then GeoGebraAssumes = Left(GeoGebraAssumes, Len(GeoGebraAssumes) - 1)
    If GeoGebraDefs <> "" Then GeoGebraDefs = Left(GeoGebraDefs, Len(GeoGebraDefs) - 1)

End Sub

Public Function GetDefsForGeoGebra(Optional ConvertHTML As Boolean = True) As String
' bruges ikke mere, men måske er metoden bedre, da den gør noget med rækkefølgen
    Dim cmd As String, i As Integer, k As Integer, j As Integer, fktudtryk As String, UrlLink As String, p As Integer
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim ea As New ExpressionAnalyser, arr() As String
    Dim geogebrafil As New CGeoGebraFile
    
    ea.SetNormalBrackets

    ' sæt definitioner i rigtig rækkefølge
    For i = 0 To omax.defindex - 1
        ea.Text = omax.DefValue(i)
        DefList = DefList & "," & omax.DefName(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.Text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next
    
    
       
    ' definer variable der ikke er defineret
    ea.Text = DefList
    For i = 0 To sl.Length - 1
'        If InStr(sl.GetVal(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
                p = InStr(sl.GetName(i), "(")
                If p > 0 Then
                    cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
                Else
                    cmd = sl.GetName(i) & "=" & fktudtryk
                End If
            Else
                cmd = sl.GetName(i) & "=" & fktudtryk
            End If
            cmd = ConvertToGeogebraSyntax(cmd, False)
            If ConvertHTML Then cmd = Replace(cmd, "+", "%2B")
            cmd = cmd & ";"
            UrlLink = UrlLink & cmd
'        End If
    Next
    
    If UrlLink <> "" Then
        If right(UrlLink, 1) = ";" Then UrlLink = Left(UrlLink, Len(UrlLink) - 1)
    End If
    GetDefsForGeoGebra = UrlLink
End Function
Function RunGeoGebraDirect(ByVal cmd As String, Optional UseDefs As Boolean = True) As String
'        If Defliste <> "" Then Defliste = Defliste & ";"
        Dim Res As String, i As Integer, s As String
        Dim UfWait2 As UserFormWaitForMaxima
'        On Error Resume Next

        If UseDefs Then FindGeoGebraDefsAndAssumes
        
        If cmd = "" Then cmd = omax.Kommando
        cmd = ConvertToGeogebraSyntax(cmd)
        omax.StartKommando = cmd
        Res = ExecuteGeoGebraCasCommand(cmd, UseDefs)
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
            Res = ExecuteGeoGebraCasCommand(cmd)
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
            Res = ExecuteGeoGebraCasCommand(cmd)
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

Function ConvertToGeogebraSyntax(ByVal Text As String, Optional ConvertMaxima As Boolean = True, Optional HtmlReady As Boolean = False) As String
'    Dim geogebrafil As New CGeoGebraFile
' definitioner vil allerede være kørt igennem codeforMaxima, så der skal convertmaxima være false

   Dim p As Integer, p2 As Integer, arr() As String, p3 As Integer, sp As Integer, ep As Integer
   Dim ea As ExpressionAnalyser, s As String, t As String, gexpr As String, i As Integer, n As Integer
   Set ea = New ExpressionAnalyser
   ea.SetNormalBrackets
    
    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, vbLf, "")
    Text = Trim(Text)
    

 '  text = Replace(text, "log", "lg")
   If ConvertMaxima Then
      Text = omax.CodeForMaxima(Text, 1) ' CASengine
   End If

      'græske bogstaver
      Text = Replace(Text, "Delta", VBA.ChrW(916))
      Text = Replace(Text, "delta", VBA.ChrW(948))
      Text = Replace(Text, "alpha", VBA.ChrW(945))
      Text = Replace(Text, "beta", VBA.ChrW(946))
      Text = Replace(Text, "gamma_incomplete", "\Gamma_incomplete ")
      Text = Replace(Text, "gamma", VBA.ChrW(915))
      Text = Replace(Text, "gammaLB", VBA.ChrW(947))
      Text = Replace(Text, "theta", VBA.ChrW(952))
      Text = Replace(Text, "Theta", VBA.ChrW(920))
      Text = Replace(Text, "lambda", VBA.ChrW(955))
      Text = Replace(Text, "Lambda", VBA.ChrW(923))
      Text = Replace(Text, "rho", VBA.ChrW(961))
      Text = Replace(Text, "varphi", VBA.ChrW(966))
      Text = Replace(Text, "phi", VBA.ChrW(981))
      Text = Replace(Text, "Phi", VBA.ChrW(934))
      Text = Replace(Text, "varepsilon", VBA.ChrW(949))
      Text = Replace(Text, "epsilon", VBA.ChrW(1013))
      Text = Replace(Text, "psi", VBA.ChrW(968))
      Text = Replace(Text, "Psi", VBA.ChrW(936))
      Text = Replace(Text, "sigma", VBA.ChrW(963))
      Text = Replace(Text, "Sigma", VBA.ChrW(931))
      Text = Replace(Text, "mu", VBA.ChrW(956))
      Text = Replace(Text, "Ohm", VBA.ChrW(937))
      Text = Replace(Text, "Omega", VBA.ChrW(937))
      Text = Replace(Text, "omega", VBA.ChrW(969))
      Text = Replace(Text, "Xi", VBA.ChrW(926))
      Text = Replace(Text, "xi", VBA.ChrW(958))
      Text = Replace(Text, "Chi", VBA.ChrW(935))
      Text = Replace(Text, "chi", VBA.ChrW(967))
      Text = Replace(Text, "tau", VBA.ChrW(964))
      Text = Replace(Text, "Pi", VBA.ChrW(928))
      Text = Replace(Text, "greek-nu", VBA.ChrW(957))
      Text = Replace(Text, "kappa", VBA.ChrW(954))
      Text = Replace(Text, "zeta", VBA.ChrW(950))
      Text = Replace(Text, "eta", VBA.ChrW(951)) ' skal være sidst da eta indgår i andre
      Text = Replace(Text, "increment", VBA.ChrW(8710))  ' specielt delta increment
      Text = Replace(Text, "Symhalf", VBA.ChrW(189)) ' _
      text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
      Text = Replace(Text, "<=", VBA.ChrW(8804))  ' Virker kun med geogebra app
      Text = Replace(Text, ">=", VBA.ChrW(8805)) ' Virker kun med geogebra app
      Text = Replace(Text, "CVinkelO", VBA.ChrW(8736)) '
      Text = Replace(Text, "CVinkel", VBA.ChrW(8736)) '
      Text = Replace(Text, "Symangle", VBA.ChrW(8736))  '
      Text = Replace(Text, "SymVecta", Sprog.A(683))
      
'      text = Replace(text, "Integrate", "Integral")  ' anden variabel end x godtages ikke i NIntegral, men ok i IntegralSymbolic
    
    
      Text = Replace(Text, "diff", "Derivative")  ' anden variabel end x godtages ikke i NIntegral, men ok i IntegralSymbolic
    
      'Else
      '    text = Replace(text, "log", "ln")
   
      ' integrate erstattes med Integral. I GeoGebra oversættes Integral automatisk til NItegral eller IntegralSymbolic afh af omstændigheder.
      ' Der ser dog ud til at være et problem med Integral(f(x),x,0,1)  man kan ikke selv angive en variabel, selvom det står at det er en mulighed
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
    
   
   ea.Text = Text
   ea.ReplaceVar "NIntegrate", "NIntegral"
   ea.ReplaceVar "integrate", "Integral"
   ea.ReplaceVar "Integrate", "Integral"
   ea.ReplaceVar "minf", "-infinity"
   ea.ReplaceVar "inf", "infinity"
   ea.ReplaceVar "log", "ln"
   ea.ReplaceVar "lg10", "log10"
   Text = ea.Text
   Text = Replace(Text, "%e", "exp(1)") 'VBA.ChrW(101)
   Text = Replace(Text, "%pi", VBA.ChrW(960)) '"pi"
   Text = Replace(Text, "%", "")
'   text = Replace(text, "[", "{")
'   text = Replace(text, "]", "}")
   Text = Replace(Text, "##", "*") 'prikprodukt
   Text = Replace(Text, "~", "*") ' vectorprodukt
   Text = Replace(Text, "^^", "^") ' vectorprodukt
'   text = Replace(text, "matrix", "")

'
      p = InStr(Text, "logbase(")
      Do While p > 0
        If p > 0 Then
          ea.Text = Text
          s = ea.GetNextBracketContent(p + 7)
          arr = Split(s, ",")
          If UBound(arr) > 0 Then Text = Left(Text, p - 1) & "log(" & arr(1) & "," & arr(0) & right(Text, Len(Text) - p - Len(s) - 7)
        End If
        p = InStr(Text, "logbase(")
      Loop


      p = InStr(Text, "if")
      p2 = InStr(Text, "then")
      If p > 0 And p2 > 0 Then
         sp = p
         ea.Text = Text
         ea.Pos = p - 1
         s = ea.GetNextBracketContent()
         ep = p + Len(s) + 1
         p3 = 1
         Do
            p = InStr(p3, s, "if")
            If p <= 0 Then Exit Do
            p2 = InStr(p3, s, "then")
            p3 = InStr(p3, s, "else")
            If p3 <= 0 Then p3 = Len(s) + 1
'            gexpr = gexpr & "If[" & Trim(Mid(s, p + 2, p2 - p - 2)) & "," & Trim(Mid(s, p2 + 4, p3 - p2 - 4)) & ","
            gexpr = gexpr & "If(" & Trim(Mid(s, p + 2, p2 - p - 2)) & "," & Trim(Mid(s, p2 + 4, p3 - p2 - 4)) & ","
            n = n + 1
            If p3 = Len(s) + 1 Then Exit Do
            p3 = p3 + 1
         Loop While p3 < Len(s)
         If right(gexpr, 1) = "," Then gexpr = Left(gexpr, Len(gexpr) - 1)
         For i = 1 To n
'            gexpr = gexpr & "]"
            gexpr = gexpr & ")"
         Next
         Text = Left(Text, sp - 1) & gexpr & right(Text, Len(Text) - ep + 2)
         
'         If Left(Text, 1) = "(" Then Text = right(Text, Len(Text) - 1) ' denne gav fejl, så der manglede startparentes. ved ikke hvorfor den er der

         Text = Replace(Text, " and ", " &amp;&amp; ") '&&
         Text = Replace(Text, " or ", " || ") '||
      End If

   ConvertToGeogebraSyntax = Text
       
'    ConvertToGeogebraSyntax = geogebrafil.ConvertToGeogebraSyntax(s, True)
    If HtmlReady Then
        ConvertToGeogebraSyntax = Replace(ConvertToGeogebraSyntax, "+", "%2B")
        ConvertToGeogebraSyntax = Replace(ConvertToGeogebraSyntax, "&", "%26")
    End If
'    Set geogebrafil = Nothing
End Function
Function ConvertGeoGebraSyntaxToWord(ByVal Text As String) As String
    Dim p As Long, p2 As Long, ea As New ExpressionAnalyser, mtext As String, s As String, n As Integer
    ea.SetNormalBrackets
    
    Text = TrimB(Text, """")
    
    ' mellemrum mellem variable skal have *
    ea.Text = Text
    p = InStr(ea.Text, " ")
    Do While p > 0
        If ea.IsLetterPos(p - 1) And ea.IsLetterPos(p + 1) Then
            ea.ReplaceStringAt "*", p
        End If
        p = InStr(p + 1, ea.Text, " ")
    Loop
    
    'potenser kan være skrevet som specialtegn for 0,1,2,3    p = InStr(ea.text, VBA.ChrW(176)) ' hævet 0
    ReplaceSuperScripts ea
    
    Text = ea.Text
    Text = Replace(Text, "\u003C", "<") ' mærkeligt at lige præcis denne står sådan
    Text = Replace(Text, " = ", "=") ' mellemrum omkring =
    Text = Replace(Text, " < ", "<")
    Text = Replace(Text, " > ", ">")


    Text = Replace(Text, "sin^-1", "asin")
    Text = Replace(Text, "cos^-1", "acos")
    Text = Replace(Text, "tan^-1", "atan")

    Text = omax.ConvertToWordSymbols(Text)
    
    ea.Text = Text
    
    ReplaceTrigSuperscript ea, "sin"
    ReplaceTrigSuperscript ea, "cos"
    ReplaceTrigSuperscript ea, "tan"
        
    
    'matricer og vektorer
     ea.SetTuborgBrackets
    p = InStr(ea.Text, "{{")
    Do While p > 0
        mtext = ea.GetNextBracketContent(p)
        mtext = Replace(mtext, " ", "") ' mellem efter ; giver problemer i næste linje
        mtext = Replace(mtext, "}" & ListSeparator & "{", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = omax.matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & omax.matrixendbracket
        ea.Text = Left(ea.Text, p - 1) & mtext & right(ea.Text, Len(ea.Text) - ea.Pos + 1)

        p = InStr(p + 1, ea.Text, "{{")
    Loop
    
    p = InStr(ea.Text, "_{")
    Do While p > 0
        s = ea.GetNextBracketContent(p)
        ea.Text = Left(ea.Text, p) & s & right(ea.Text, Len(ea.Text) - Len(s) - p - 2)
        p = InStr(p + 1, ea.Text, "_{")
    Loop
    Text = ea.Text
    
    ConvertGeoGebraSyntaxToWord = Text
End Function

Sub ReplaceSuperScripts(ByRef ea As ExpressionAnalyser)
    Dim p As Integer, Maxp As Integer
    Dim FirstSup As Boolean
    
    Maxp = ea.Length
    FirstSup = True
    Do
        If ea.IsSuperScriptNoPos(p) Then
            If FirstSup Then
                ReplaceSuperScriptNoAtPos ea, p
                ea.InsertBeforePos "^", p
                Maxp = Maxp + 1
                p = p + 1
                FirstSup = False
            Else
                ReplaceSuperScriptNoAtPos ea, p
            End If
        Else
            FirstSup = True
        End If
        p = p + 1
    Loop While p <= Maxp
    
End Sub
Sub ReplaceSuperScriptNoAtPos(ByRef ea As ExpressionAnalyser, Pos As Integer)
    Dim c As Integer
    c = AscW(ea.ChrByIndex(Pos))
    If c = 8314 Then
        ea.ReplaceStringAt "+", Pos
    ElseIf c = 8315 Then
        ea.ReplaceStringAt "-", Pos
    ElseIf c = 8317 Then
        ea.ReplaceStringAt "(", Pos
    ElseIf c = 8318 Then
        ea.ReplaceStringAt ")", Pos
    ElseIf c = 8304 Then
        ea.ReplaceStringAt "0", Pos
    ElseIf c = 185 Then
        ea.ReplaceStringAt "1", Pos
    ElseIf c = 178 Then
        ea.ReplaceStringAt "2", Pos
    ElseIf c = 179 Then
        ea.ReplaceStringAt "3", Pos
    ElseIf c = 8308 Then
        ea.ReplaceStringAt "4", Pos
    ElseIf c = 8309 Then
        ea.ReplaceStringAt "5", Pos
    ElseIf c = 8310 Then
        ea.ReplaceStringAt "6", Pos
    ElseIf c = 8311 Then
        ea.ReplaceStringAt "7", Pos
    ElseIf c = 8312 Then
        ea.ReplaceStringAt "8", Pos
    ElseIf c = 8313 Then
        ea.ReplaceStringAt "9", Pos
    End If
End Sub


Sub ReplaceTrigSuperscript(ByRef ea As ExpressionAnalyser, Trig As String)
' erstatter fx sin & chrw(8289) & "^2" med "sin" & "^2" & chrw(8289)
' altså retter op på output fra converttowordsymols
Dim p As Integer, p2 As Integer
    p = InStr(ea.Text, Trig & VBA.ChrW(8289) & "^")
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
        p = InStr(ea.Text, Trig & VBA.ChrW(8289) & "^")
    Loop
End Sub

Sub GeoGebra()
' sender den valgte ligning og definitioner over i GeoGebra 5
    On Error GoTo fejl
    Dim geogebrasti As String
    Dim geogebrafilersti As String
    Dim appnr As Long
    Dim UfWait As New UserFormWaitForMaxima
    Dim TempCas As Integer
    
    TempCas = CASengine
    CASengine = 1
    
    PrepareMaxima ' omax bliver brugt
    
    UfWait.Label_tip.Caption = Sprog.A(362)
    UfWait.Label_progress.Caption = "***"
    UfWait.CommandButton_stop.visible = False
    UfWait.Show vbModeless
    
    geogebrasti = GeoGebraPath()
    If geogebrasti = "" Then ' hvis geogebra ikke installeret
        InstallGeoGebra
        GoTo slut
    End If
'    geogebrafilersti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    geogebrafilersti = GetTempDir()
    
    CreateGeoGebraFil geogebrafilersti
    
    geogebrafilersti = geogebrafilersti & "geogebra.ggb"
    
    UfWait.Label_progress.Caption = "******"
    If FileExists(geogebrafilersti) Then ' check om geogebrafilen er lavet
        geogebrasti = geogebrasti & " """ & geogebrafilersti & """"
    Else
        MsgBox "The GeoGebra.ggb file cannot be located", vbOKOnly, Sprog.Error
        GoTo fejl
    End If
    
#If Mac Then
    RunScript "OpenGeoGebra", geogebrafilersti
#Else
    appnr = shell(geogebrasti, vbNormalFocus)   'vbNormalFocus vbMinimizedFocus ' til offline installer
#End If
' til webstart:
'    appnr = Shell("javaws -system -open """ & geogebrafilersti & "geogebra.ggb""" & " http://www.geogebra.org/webstart/geogebra.jnlp", vbNormalFocus)    'vbNormalFocus vbMinimizedFocus
    
    UfWait.Label_progress.Caption = "*********"
    On Error Resume Next
    DoEvents
    Unload UfWait
    
    GoTo slut
fejl:
'    UserFormGeoGebra.Show
slut:
    If Not UfWait Is Nothing Then Unload UfWait
    CASengine = TempCas
End Sub
Sub InstallGeoGebra()
    Dim i As Long, DDir As String, FN As String
    Dim UfWait As UserFormWaitForMaxima
    
    UserFormGeoGebra.Show
#If Mac Then
    MsgBox "The download page will now open. Install GeoGebra classic 5", vbOKOnly, "Download"
    OpenLink "https://www.geogebra.org/download"
#Else ' win

    If UserFormGeoGebra.ReturnVal = 1 Then
        Set UfWait = New UserFormWaitForMaxima
        UfWait.Label_tip.Font.Size = 10
        UfWait.Label_tip.Font.Italic = False
        UfWait.Show vbModeless
        UfWait.Label_tip.Caption = "Downloader GeoGebra 5"
        UfWait.Label_progress.Caption = "*"
        '    DownloadFile "https://download.geogebra.org/package/win"
        OpenLink "https://download.geogebra.org/package/win" ' åbning af dette link starter automatisk download af den rigtige fil
    
        DDir = GetDownloadsFolder
        Do While i < 12
            Sleep2 1
            FN = Dir(DDir & "\GeoGebra-Windows-Installer-5*.exe")
            If FN <> "" Then Exit Do
            UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
            i = i + 1
        Loop
        UfWait.Hide
        If i < 12 Then
            If Not RunApplication(DDir & "\" & FN) Then
                MsgBox "The GeoGebra installation file is now in the downloadsfolder. Go run it", vbOKOnly, "Run installer"
                shell "explorer.exe " & DDir, vbNormalFocus
            End If
        Else ' hvis der ikke er blevet hentet en fil, må brugeren selv hente
            MsgBox "The download page will now be shown. Download and install the 'GeoGebra Classic 5' version", vbOKOnly, "Download page"
            OpenLink "https://www.geogebra.org/download"
        End If
    Else
        GeoGebraWeb
    End If
#End If
    
    GoTo slut
fejl:

slut:
    Unload UfWait
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
    Dim arr As Variant
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
    Wait 1 ' anden gang man åbner geogebra på mac låser word hvis denne ikke er der. Det har noget med finddefinitions at gøre, men har ikke kunnet se hvad
#End If

    PrepareMaxima
'    omax.FindDefinitions ' bliver fundet i preparemaxima
    omax.ReadSelection

'    defs = ConvertToAscii(omax.defstring)
'    arr = Split(defs, ";")
    Dim sl As New CSortList
    Dim var As String, DefList As String
    Dim k As Integer
    
    ' sæt definitioner i rigtig rækkefølge
    For i = 0 To omax.defindex - 1
        ea.Text = omax.DefValue(i)
        DefList = DefList & "," & omax.DefName(i)
        var = ea.GetNextVar
        If var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.Text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next
    
    ' definer variable der ikke er defineret
    ea.Text = DefList
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
    ' indsæt de markerede funktioner
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
                    arr = Split(udtryk, "=")
                    lhs = arr(0)
                    rhs = arr(1)
                    ea.Text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.Text = rhs
                        ea.Pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.Text
                        DefinerKonstanter fktudtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction fktnavn, fktudtryk, False, True
                    Else
                        fktudtryk = ReplaceIndepvarX(rhs)
                        DefinerKonstanter fktudtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction "f" & j, fktudtryk, False
                        j = j + 1
                    End If
                ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                ' kan først bruges med GeoGebra 4.0
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
    
' Problemet med dette er at tabellen skal være lodret
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
    ea.Text = DefList
    ea2.Text = Expr
    ea2.Pos = 0
    Do
        var = ea2.GetNextVar
'        MsgBox AscW(var) & vbCrLf & VBA.ChrW(960)
        ea2.Pos = ea2.Pos + 1
'        If var = "z" Then geogebrafil.Show3D = True  ' ikke implementeret endnu
        If Not (ea2.ChrByIndex(ea2.Pos) = "(") And Not (Left(var, 1) = "_") And Not (ea.IsFunction(var)) And Not (ea.ContainsVar(var)) And var <> "" And var <> "x" And var <> "y" And var <> "z" And var <> "e" And var <> "pi" And var <> "matrix" And var <> "if" And var <> "elseif" And var <> "then" And var <> "and" And var <> "or" And var <> "else" And var <> "amp" And var <> VBA.ChrW(960) Then  ' 960=pi
            varval = InputBox(Sprog.A(363) & " " & var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
            If Len(varval) > 0 Then
                If Not geogebrafil Is Nothing Then
                    geogebrafil.CreateFunction var, varval, False, False
                Else
                    UrlLink = UrlLink & var & "=" & varval & ";"
                End If
                DefList = DefList & "," & var
                ea.Text = DefList
                i = i + 1
            End If
        End If
    Loop While var <> ""
    
'    If i > 1 Then ' kan ikke lige se hvad denne sektion er nÀdvendig for
'        var = Split(DefList, ",")(0)
'            varval = InputBox(Sprog.A(363) & " " & var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
'            If Len(varval) > 0 Then
'                If Not geogebrafil Is Nothing Then
'                    geogebrafil.CreateFunction var, varval, False, False
'                Else
'                    UrlLink = UrlLink & var & "=" & varval & ";"
'                End If
'            End If
'    End If
    
    DefinerKonstanter = var
End Function
Sub TestZipFile()

'    DateiZippen "C:\TEMP\Example.xml", "C:\TEMP\Example.zip"
    CreateZipFile "C:\WordMatGraf.zip", "C:\WordMatGraf.gif", "C:\WordMatGraf.jpg"

End Sub

 

Sub CreateZipFile(zipfilnavn As Variant, FilNavn As Variant, Optional filnavn2 As Variant = "", Optional filnavn3 As Variant = "", Optional filnavn4 As Variant = "")
#If Mac Then
    RunScript "ZipFile", zipfilnavn & ";" & FilNavn
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
    objFolder.CopyHere FilNavn
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



