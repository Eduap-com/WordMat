Attribute VB_Name = "GeoGebraModule"
Option Explicit

Public GeoGebraDefs As String
Public GeoGebraAssumes As String

Sub GeoGebraWeb(Optional Gtype As String = "", Optional CASfunc As String = "")
' Uses Browser to plot with GeoGebra.
' There are no functions that use the parameters. If you just run GeoGebraWeb without parameters, it's plotting, so that's the only thing this function is used for
' The function prepares to send to OpenGeoGebra, which actually opens edge and the right html file
' reads the selected equation/function
' Definitions are also found and inserted

    ' gtype="", "3d", "CAS"
    Dim cmd As String, UrlLink As String, s As String
    Dim sl As New CSortList
    Dim Var As String, DefList As String
    Dim k As Integer, i As Integer, j As Integer, p As Integer
    Dim Arr As Variant, uvar As String
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
    Dim TempCas As Integer
    Dim VektNArr As Variant, VNi As Integer
    VektNArr = Array("a", "b", "c", "v", "w")

    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
    ea2.SetNormalBrackets

    On Error GoTo Fejl

    TempCas = CASengine
    CASengineTempOnly = 1

    PrepareMaxima ' finds definitions
    
    omax.ReadSelection ' reads selected function

    If Gtype <> "CAS" Then
        ' put definitions in the correct order
        For i = 0 To omax.defindex - 1
            DefList = DefList & "," & omax.DefName(i)
            ea.Text = omax.DefValue(i)
            Var = ea.GetNextVar
            If Var = "" Then
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

        ' define variables that are not defined
        omax.FindVariable
        ea.Text = DefList
        For i = 0 To sl.Length - 1
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
                varnavn = sl.GetName(i)
                p = InStr(sl.GetName(i), "(")
                If p > 0 Then
                    cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
                ElseIf InStr(varnavn, "SymVect") > 0 Then
                    varnavn = Replace(varnavn, "SymVecta", "")
                    fktudtryk = Replace(fktudtryk, "{", "(")
                    fktudtryk = Replace(fktudtryk, "}", ")")
                    fktudtryk = Replace(fktudtryk, "((", "(")
                    fktudtryk = Replace(fktudtryk, "))", ")")
                    fktudtryk = "vector((0,0)," & fktudtryk & ")"
                    cmd = varnavn & "=" & fktudtryk
                Else
                    cmd = sl.GetName(i) & "=" & fktudtryk
                End If
            Else
                cmd = sl.GetName(i) & "=" & fktudtryk
            End If
            cmd = Replace(Replace(ConvertToGeogebraSyntax(cmd, False), "+", "%2B"), "&", "%26") & ";" ' v.1.26 added false when converting because it is probably already converted
            UrlLink = UrlLink & cmd
        Next
    End If

    j = 1
    ' insert the marked functions
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        s = Trim(LCase(Udtryk))
        If Not (InStr(s, "definer:") > 0 Or InStr(s, "define:") > 0) Then
'            Udtryk = Replace(Udtryk, "definer:", "")
'            Udtryk = Replace(Udtryk, "Definer:", "")
'            Udtryk = Replace(Udtryk, "define:", "")
'            Udtryk = Replace(Udtryk, "Define:", "")
            Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
            Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
            Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
            '
            Udtryk = Trim(Udtryk)
            Udtryk = ConvertToGeogebraSyntax(Udtryk)
            If Gtype <> "CAS" Then
                If Len(Udtryk) > 0 Then
                    If InStr(Udtryk, "matrix") < 1 Then
                        If InStr(Udtryk, "=") > 0 Then
                            Arr = Split(Udtryk, "=")
                            LHS = Trim(Arr(0))
                            RHS = Trim(Arr(1))
                            ea.Text = LHS
                            fktnavn = ea.GetNextVar(1)
                            varnavn = ea.GetNextBracketContent(1)
                    
                            If LHS = fktnavn & "(" & varnavn & ")" Then
                                ea.Text = RHS
                                ea.Pos = 1
                                ea.ReplaceVar varnavn, "x"
                                fktudtryk = ea.Text
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        
                                cmd = fktnavn & "(x)=" & fktudtryk
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            ElseIf LHS = "y" Then
                                fktudtryk = ReplaceIndepvarX(RHS, uvar, DefList)
                                If Not (uvar = "" Or uvar = "x") Then  'Or uvar = "t"
                                    DefList = DefList & "," & uvar
                                End If
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                If fktnavn = "y" Then
                                    cmd = "f" & j & ":" & fktnavn & "=" & fktudtryk
                                Else
                                    cmd = fktnavn & "(x)=" & fktudtryk
                                End If
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                                j = j + 1
                            ElseIf LHS = "({{x},{y}})" Then 'parametric plot
                                RHS = Replace(RHS, "{", "(")
                                RHS = Replace(RHS, "}", ")")
                                RHS = Replace(RHS, "((", "(")
                                RHS = Replace(RHS, "))", ")")
                                cmd = "Param:X=" & RHS
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            ElseIf Right(LHS, 3) = "pil" Then ' vector
                                LHS = Left(LHS, Len(LHS) - 3)
                                RHS = Replace(RHS, "{", "(")
                                RHS = Replace(RHS, "}", ")")
                                cmd = LHS & "=vector((0,0)," & RHS & ")"
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            ElseIf fktnavn = LHS Then
                                fktudtryk = ReplaceIndepvarX(RHS, uvar, DefList)
                                If fktudtryk <> vbNullString Then
                                    If Not (uvar = "" Or uvar = "x") Then  'Or uvar = "t"
                                        DefList = DefList & "," & uvar
                                    End If
                                    DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                    cmd = fktnavn & j & ":" & "y=" & fktudtryk
                                    cmd = Replace(cmd, "+", "%2B") & ";"
                                    UrlLink = UrlLink & cmd
                                    j = j + 1
                                End If
                            Else ' ligning
                                cmd = LHS & "=" & RHS
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            End If
                        ElseIf Left(Udtryk, 3) = "({{" Then ' vector
                            Udtryk = Replace(Udtryk, "{", "(")
                            Udtryk = Replace(Udtryk, "}", ")")
                            '                        Udtryk = Replace(Udtryk, "((", "(")
                            '                        Udtryk = Replace(Udtryk, "))", ")")
                            If VNi < 5 Then
                                cmd = VektNArr(VNi) & "=vector((0,0)," & Udtryk & ")"
                            Else
                                cmd = "v" & VNi - 4 & "=vector((0,0)," & Udtryk & ")"
                            End If
                            VNi = VNi + 1
                            cmd = Replace(cmd, "+", "%2B") & ";"
                            UrlLink = UrlLink & cmd
                        ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                            DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                            cmd = "u" & j & "=" & Udtryk
                            cmd = Replace(cmd, "+", "%2B") & ";"
                            UrlLink = UrlLink & cmd
                            '                    geogebrafil.CreateFunction "u" & j, udtryk, True
                        ElseIf LHS = "" Then
                            Udtryk = ReplaceIndepvarX(Udtryk, , DefList)
                            If Udtryk <> vbNullString Then
                                DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                                cmd = "f" & j & "=" & Udtryk
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                                j = j + 1
                            End If
                        Else
                            Udtryk = ReplaceIndepvarX(Udtryk, , DefList)
                            If Udtryk <> vbNullString Then
                                DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                                cmd = "f" & j & "=" & Udtryk
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                                '                    geogebrafil.CreateFunction "f" & j, udtryk, False
                                j = j + 1
                            End If
                        End If
                    End If
                End If
            Else 'CAS
                If CASfunc <> "" Then
                    Udtryk = CASfunc & "(" & Udtryk & ")"
                End If
                cmd = Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ";"
                UrlLink = UrlLink & cmd
            End If
        End If
    Next

    'Points
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
    
    OpenGeoGebraWeb UrlLink, Gtype, False, False
Fejl:

slut:
    CASengineTempOnly = TempCas
End Sub

Sub OpenGeoGebraWeb(ByVal cmd As String, Gtype As String, Optional ConvertSyntax As Boolean = False, Optional UseDefs As Boolean = True)
' Opens GeoGebra in Edge.
' Gtype="" for plotting. Opens in Calculator suite
' Gtype="classic" for plotting. Used by haeldningsfelt
' Gtype="CAS"  opens GeoGebraCASapplet.html
' Function does not read in the document. Preparemaxima must be run prior, to find definitions, when UseDefs=true
' cmd added to the end of url'en with ?command=       Definitions are also added to command

    Dim UrlLink As String, ArrDef() As String, ArrCas() As String, i As Integer, AssumeString As String
    Dim DefS As String, DN As String
       
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
    
    If Len(cmd) > 0 Then If Right(cmd, 1) = ";" Then cmd = Left(cmd, Len(cmd) - 1)
    '    If ConvertSyntax Then Cmd = ConvertToGeogebraSyntax(Cmd, True)
    cmd = DefS & cmd
    cmd = Replace(cmd, "+", "%2B")
        
#If Mac Then
    '    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra" & Gtype & "Applet.html"
    If Gtype = "" Then
'        UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra" & Gtype & "Applet.html"
        UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra/HTML5/5.0/GeoGebra.html"
    Else
        UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra" & Gtype & "Applet.html"
    End If
#Else
    '    UrlLink = "https://geogebra.org/calculator"
    DN = GetGeoGebraMathAppsFolder()
    If DN <> vbNullString Then
        If Gtype = "" Then
            UrlLink = "file://" & DN & "GeoGebra/HTML5/5.0/GeoGebra.html"
        Else
            UrlLink = "file://" & DN & "GeoGebra" & Gtype & "Applet.html"
        End If
    Else: GoTo slut
    End If
#End If
    UrlLink = UrlLink & "?command=" & cmd

    OpenLink UrlLink, True
slut:
End Sub

Function GetGeoGebraMathAppsFolder() As String
' Finds GeoGebramath apps in Appdata or program files. It first searches for what was last installed
#If Mac Then
    GetGeoGebraMathAppsFolder = "/Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/"
#Else
    Dim DN As String
    If InstallLocation = "All" Then
        DN = GetProgramFilesDir & "/WordMat/geogebra-math-apps/"
        If Dir(DN, vbDirectory) = vbNullString Then
            DN = Environ("AppData") & "/WordMat/geogebra-math-apps/"
            If Dir(DN, vbDirectory) = vbNullString Then
                MsgBox "geogebra-math-apps could not be found", vbOKOnly, "Error"
            End If
        End If
    Else
        DN = Environ("AppData") & "/WordMat/geogebra-math-apps/"
        If Dir(DN, vbDirectory) = vbNullString Then
            DN = GetProgramFilesDir & "/WordMat/geogebra-math-apps/"
            If Dir(DN, vbDirectory) = vbNullString Then
                MsgBox "geogebra-math-apps could not be found", vbOKOnly, "Error"
            End If
        End If
    End If
    GetGeoGebraMathAppsFolder = DN
#End If
End Function

Sub FindGeoGebraDefsAndAssumes()
' sets the lines GeoGebraDefs and GeoGebraAssumes from omax
Dim Arr() As String, i As Integer
'    MsgBox omax.DefString
'    MsgBox omax.defstringtext
    
    GeoGebraAssumes = ""
    GeoGebraDefs = ""
    
' Add definitions and assumes written in math fields
    Arr = Split(omax.DefString, ";")
    For i = 0 To UBound(Arr) - 1
        If Left(Arr(i), 7) = "assume(" Then
            GeoGebraAssumes = GeoGebraAssumes & Mid(Arr(i), 8, Len(Arr(i)) - 8) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(Arr(i), False) & ";"
        End If
    Next
    
' Add temporary definitions and assumes
    Arr = Split(omax.TempDefs, ";") ' all temporary definitions and assumptions separated by semicolons
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

Function RunGeoGebraDirect(ByVal cmd As String, Optional UseDefs As Boolean = True) As String
        Dim res As String, i As Integer, s As String
        Dim UFwait2 As UserFormWaitForMaxima
'        On Error Resume Next

        If UseDefs Then FindGeoGebraDefsAndAssumes
        
        If cmd = "" Then cmd = omax.Kommando
        cmd = ConvertToGeogebraSyntax(cmd)
        omax.StartKommando = cmd
        res = ExecuteGeoGebraCasCommand(cmd, UseDefs)
        i = 0
        If Left(res, 9) = "reloading" Then
            Set UFwait2 = New UserFormWaitForMaxima
            UFwait2.Label_tip.Font.Size = 10
            UFwait2.Label_tip.Font.Italic = False
            UFwait2.Show vbModeless
            UFwait2.Label_tip.Caption = TT.A(684)
            UFwait2.Label_progress.Caption = "*"
#If Mac Then
            Do
                Wait (0.2)
                UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
                If UFwait2.StopNow Then GoTo slut
                s = RunScript("IsGeoGebraAppReady", "")
                If Left(s, 3) = "yes" Then Exit Do
                i = i + 1
            Loop Until i > 50
            Wait (1)
#End If
            UFwait2.Label_tip.Caption = "Executing GeoGebra command"
            res = ExecuteGeoGebraCasCommand(cmd)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
        ElseIf Left(res, 7) = "opening" Or res = "null" Then
            Set UFwait2 = New UserFormWaitForMaxima
            UFwait2.Label_tip.Font.Size = 10
            UFwait2.Label_tip.Font.Italic = False
            UFwait2.Show vbModeless
            UFwait2.Label_tip.Caption = TT.A(684)
            UFwait2.Label_progress.Caption = "*"
#If Mac Then
            Do
                Wait (0.2)
                UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
                If UFwait2.StopNow Then GoTo slut
                s = RunScript("IsGeoGebraAppReady", "")
                If Left(s, 3) = "yes" Then Exit Do
                i = i + 1
            Loop Until i > 50
            Wait (1)
#Else
'            Wait (1)
#End If
            UFwait2.Label_tip.Caption = "Executing GeoGebra command"
            res = ExecuteGeoGebraCasCommand(cmd)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
        ElseIf Left(res, 5) = "error" Then
'            Wait (1)
            GoTo slut
        End If
'        Res = Replace(Res, " ", "")
        res = ConvertGeoGebraSyntaxToWord(res)
        omax.MaximaOutput = res
        RunGeoGebraDirect = res
slut:
     If Not UFwait2 Is Nothing Then Unload UFwait2
End Function

Function ConvertToGeogebraSyntax(ByVal Text As String, Optional ConvertMaxima As Boolean = True, Optional HtmlReady As Boolean = False) As String
' definitions will have already been run through codeforMaxima, so convertmaxima should be false

   Dim p As Integer, p2 As Integer, Arr() As String, p3 As Integer, sp As Integer, ep As Integer
   Dim ea As ExpressionAnalyser, s As String, gexpr As String, i As Integer, n As Integer
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

      'greek letters
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
      Text = Replace(Text, "eta", VBA.ChrW(951)) ' must be last as eta is included in others
      Text = Replace(Text, "increment", VBA.ChrW(8710))  ' speciel delta increment
      Text = Replace(Text, "Symhalf", VBA.ChrW(189)) ' _
      text = Replace(text, "degC", VBA.ChrW(8451))   ' speciel oC symbol
      Text = Replace(Text, "<=", VBA.ChrW(8804))  ' Only works with geogebra app
      Text = Replace(Text, ">=", VBA.ChrW(8805)) ' Only works with geogebra app
      Text = Replace(Text, "CVinkelO", VBA.ChrW(8736))
      Text = Replace(Text, "CVinkel", VBA.ChrW(8736))
      Text = Replace(Text, "Symangle", VBA.ChrW(8736))
      Text = Replace(Text, "SymVecta", TT.A(683))
    
      Text = Replace(Text, "diff", "Derivative")  ' variable other than x is not accepted in NIntegral, but ok in IntegralSymbolic
    
      'Else
      '    text = Replace(text, "log", "ln")
   
' integrate is replaced with Integral. In GeoGebra, Integral is automatically translated to NItegral or IntegralSymbolic depending on the circumstances.
' However, there seems to be a problem with Integral(f(x),x,0,1) you cannot specify a variable yourself, even though it says it is a possibility
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
   ea.ReplaceVar "log", "ln" ' The text will have been run through codeformaxima, so log will be ln. It needs to be changed back
   ea.ReplaceVar "lg10", "log10"
   Text = ea.Text
   Text = Replace(Text, "%e", "exp(1)") 'VBA.ChrW(101)
   Text = Replace(Text, "%pi", VBA.ChrW(960)) '"pi"
   Text = Replace(Text, "%", "")
   Text = Replace(Text, "##", "*") 'dot product
   Text = Replace(Text, "~", "*") ' vectorprodukt
   Text = Replace(Text, "^^", "^") ' vectorprodukt

'
      p = InStr(Text, "logbase(")
      Do While p > 0
        If p > 0 Then
          ea.Text = Text
          s = ea.GetNextBracketContent(p + 7)
          Arr = Split(s, ",")
          If UBound(Arr) > 0 Then Text = Left(Text, p - 1) & "log(" & Arr(1) & "," & Arr(0) & Right(Text, Len(Text) - p - Len(s) - 7)
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
         If Right(gexpr, 1) = "," Then gexpr = Left(gexpr, Len(gexpr) - 1)
         For i = 1 To n
'            gexpr = gexpr & "]"
            gexpr = gexpr & ")"
         Next
         Text = Left(Text, sp - 1) & gexpr & Right(Text, Len(Text) - ep + 2)
         
         Text = Replace(Text, " and ", " &amp;&amp; ") '&&
         Text = Replace(Text, " or ", " || ") '||
      End If

   ConvertToGeogebraSyntax = Text
       
'    ConvertToGeogebraSyntax = geogebrafil.ConvertToGeogebraSyntax(s, True)
    If HtmlReady Then
        ConvertToGeogebraSyntax = Replace(ConvertToGeogebraSyntax, "+", "%2B")
        ConvertToGeogebraSyntax = Replace(ConvertToGeogebraSyntax, "&", "%26")
    End If
End Function
Function ConvertGeoGebraSyntaxToWord(ByVal Text As String) As String
    Dim p As Long, ea As New ExpressionAnalyser, mtext As String, s As String
    ea.SetNormalBrackets
    
    Text = TrimB(Text, """")
    
    ' spaces between variables must have *
    ea.Text = Text
    p = InStr(ea.Text, " ")
    Do While p > 0
        If ea.IsLetterPos(p - 1) And ea.IsLetterPos(p + 1) Then
            ea.ReplaceStringAt "*", p
        End If
        p = InStr(p + 1, ea.Text, " ")
    Loop
    
    'powers can be written as special characters for 0,1,2,3 p = InStr(ea.text, VBA.ChrW(176)) ' superscript 0
    ReplaceSuperScripts ea
    
    Text = ea.Text
    Text = Replace(Text, "\u003C", "<") ' strange that this one stands exactly like this
    Text = Replace(Text, " = ", "=") ' space around =
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
        
    
    'matrices and vectors
     ea.SetTuborgBrackets
    p = InStr(ea.Text, "{{")
    Do While p > 0
        mtext = ea.GetNextBracketContent(p)
        mtext = Replace(mtext, " ", "") ' between after ; causes problems in the next line
        mtext = Replace(mtext, "}" & ListSeparator & "{", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = omax.matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & omax.matrixendbracket
        ea.Text = Left(ea.Text, p - 1) & mtext & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)

        p = InStr(p + 1, ea.Text, "{{")
    Loop
    
    p = InStr(ea.Text, "_{")
    Do While p > 0
        s = ea.GetNextBracketContent(p)
        ea.Text = Left(ea.Text, p) & s & Right(ea.Text, Len(ea.Text) - Len(s) - p - 2)
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
' replaces e.g. sin & chrw(8289) & "^2" with "sin" & "^2" & chrw(8289)
' thus corrects the output from converttowordsymols
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
' sends the selected equation and definitions to GeoGebra 5
    On Error GoTo Fejl
    Dim geogebrasti As String
    Dim geogebrafilersti As String
    Dim geogebracmd As String
    Dim UfWait As New UserFormWaitForMaxima
    Dim TempCas As Integer
    
    TempCas = CASengine
    CASengineTempOnly = 1
    
    PrepareMaxima ' omax must be prepared
    
    UfWait.Label_tip.Caption = TT.A(362)
    UfWait.Label_progress.Caption = "***"
    UfWait.Label_stop.visible = False
    UfWait.Show vbModeless
    
    geogebrasti = GeoGebraPath()
    If geogebrasti = "" Then ' if geogebra not installed
        Unload UfWait
        InstallGeoGebra
        GoTo slut
    End If
'    geogebrafilersti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    geogebrafilersti = GetTempDir()
    
    CreateGeoGebraFil geogebrafilersti
    
    geogebrafilersti = geogebrafilersti & "geogebra.ggb"
    
    UfWait.Label_progress.Caption = "******"
    If fileExists(geogebrafilersti) Then ' check if the geogebra file has been created
        geogebracmd = geogebrasti & " """ & geogebrafilersti & """"
    Else
        MsgBox "The GeoGebra.ggb file cannot be located", vbOKOnly, TT.Error
        GoTo Fejl
    End If
    
#If Mac Then
    RunScript "OpenGeoGebra", geogebrafilersti
#Else
    MaxProc.RunFile geogebrasti, geogebrafilersti
#End If
    
    UfWait.Label_progress.Caption = "*********"
    On Error Resume Next
    DoEvents
    Unload UfWait
    Set UfWait = Nothing
    
    GoTo slut
Fejl:
'    UserFormGeoGebra.Show
slut:
    CASengineTempOnly = TempCas
    If Not UfWait Is Nothing Then
        Unload UfWait
        Set UfWait = Nothing
    End If
End Sub

Sub InstallGeoGebra(Optional ConfirmPrompt As Boolean = True)

    If ConfirmPrompt Then
        UserFormGeoGebra.Show
    Else
        UserFormGeoGebra.ReturnVal = 1
        
    End If
#If Mac Then
    If UserFormGeoGebra.ReturnVal = 1 Then
        UserFormGeoGebraMacInstall.Show
        UserFormGeoGebraMacInstall.DownloadNow
    Else
        GeoGebraWeb
    End If
#Else
    If UserFormGeoGebra.ReturnVal = 1 Then
    Else
        GeoGebraWeb
    End If
#End If
    
    GoTo slut
Fejl:

slut:
End Sub
Function GeoGebraPath() As String
' path to the geogebra executable. Returns "" if not found. NOTE: On Mac the path is not used. There is an applescript for that, but the function is used to determine if there is a GeoGebra installation.
' Downloading to the latest version if possible
Dim DN As String
On Error GoTo Fejl
#If Mac Then
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra.app"
    If fileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = ""
#Else
    GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 5*", vbDirectory)
    If GeoGebraPath <> "" Then
        DN = GeoGebraPath
        Do While GeoGebraPath <> "" ' we download the GeoGebra 5 with the highest version number. It will be the last in the list
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = """" & GetProgramFilesDir & "\" & DN & "\GeoGebra.exe"""
            GoTo slut
        End If
    End If
    
    If GeoGebraPath = "" Then 'calc suite
        GeoGebraPath = Dir(Environ("USERPROFILE") & "\AppData\Local\GeoGebra_Calculator\app-*", vbDirectory)
        DN = GeoGebraPath
        Do While GeoGebraPath <> ""
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = Environ("USERPROFILE") & "\AppData\Local\GeoGebra_Calculator\" & DN & "\GeoGebraCalculator.exe"
            GeoGebraPath = """" & GeoGebraPath & """"
            GoTo slut
        End If
    End If
    
    If GeoGebraPath = "" Then ' GeoGebra classic 6
        GeoGebraPath = Dir(Environ("USERPROFILE") & "\AppData\Local\GeoGebra_6\app-6*", vbDirectory)
        DN = GeoGebraPath
        Do While GeoGebraPath <> ""
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = Environ("USERPROFILE") & "\AppData\Local\GeoGebra_6\" & DN & "\GeoGebra.exe"
            GeoGebraPath = """" & GeoGebraPath & """"
            GoTo slut
        End If
    End If
    
    If GeoGebraPath = "" Then ' Graf plotter
        GeoGebraPath = Dir(Environ("USERPROFILE") & "\AppData\Local\GeoGebra_Graphing\app-*", vbDirectory)
        DN = GeoGebraPath
        Do While GeoGebraPath <> ""
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = Environ("USERPROFILE") & "\AppData\Local\GeoGebra_Graphing\" & DN & "\GeoGebraGraphing.exe"
            GeoGebraPath = """" & GeoGebraPath & """"
            GoTo slut
        End If
    End If
        
    If GeoGebraPath = "" Then ' CAS calculator can also plot grahs
        GeoGebraPath = Dir(Environ("USERPROFILE") & "\AppData\Local\GeoGebra_CAS\app-*", vbDirectory)
        DN = GeoGebraPath
        Do While GeoGebraPath <> ""
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = Environ("USERPROFILE") & "\AppData\Local\GeoGebra_CAS\" & DN & "\GeoGebraCAS.exe"
            GeoGebraPath = """" & GeoGebraPath & """"
            GoTo slut
        End If
    End If
    
    ' look in program files for older installations
    
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 6*", vbDirectory)
    End If
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra Classic*", vbDirectory)
    End If
    If GeoGebraPath = "" Then
        GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra*", vbDirectory)
    End If
    If Not GeoGebraPath = "" Then
        GeoGebraPath = GetProgramFilesDir & "\" & GeoGebraPath & "\GeoGebra.exe"
        GeoGebraPath = """" & GeoGebraPath & """"
        GoTo slut
    End If
    
#End If
    GoTo slut
Fejl:
    GeoGebraPath = ""
slut:
End Function
Sub CreateGeoGebraFil(geogebrasti As String)
    Dim geogebrafil As New CGeoGebraFile
    Dim i As Integer, j As Integer
    Dim Arr As Variant, s As String, p As Long, cmd As String
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    On Error GoTo Fejl
    ea.SetNormalBrackets
    ea2.SetNormalBrackets
    geogebrafil.Show3D = False
    PrepareMaxima
    'geogebrasti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    omax.ConvertLnLog = False
#If Mac Then
    Wait 1 ' the second time you open geogebra on mac word locks up if it is not there. It has something to do with finddefinitions, but I haven't been able to see what
#End If

    PrepareMaxima
    omax.ReadSelection

    Dim sl As New CSortList
    Dim Var As String, DefList As String
    Dim k As Integer
    
    ' put definitions in the correct order
    For i = 0 To omax.defindex - 1
        ea.Text = omax.DefValue(i)
        DefList = DefList & "," & omax.DefName(i)
        Var = ea.GetNextVar
        If Var = "" Then
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
    
    ' define variables that are not defined
    ea.Text = DefList
    For i = 0 To sl.Length - 1
        varnavn = sl.GetName(i)
        fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
        p = InStr(sl.GetName(i), "(")
        If p > 0 Then
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), DefList, geogebrafil
            End If
            geogebrafil.CreateFunction sl.GetName(i), fktudtryk, False, False
        ElseIf InStr(varnavn, "SymVect") > 0 Then
            varnavn = Replace(varnavn, "SymVecta", "")
            fktudtryk = Replace(fktudtryk, "{", "(")
            fktudtryk = Replace(fktudtryk, "}", ")")
            fktudtryk = Replace(fktudtryk, "((", "(")
            fktudtryk = Replace(fktudtryk, "))", ")")
            cmd = varnavn & "=" & fktudtryk
            geogebrafil.CreateVector varnavn, fktudtryk, False, False
        Else
            geogebrafil.CreateEquation sl.GetName(i), fktudtryk, False, False
        End If
    Next
    
    j = 1
    ' insert the selected functions
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        s = Trim(LCase(Udtryk))
        If Not (InStr(s, "definer:") > 0 Or InStr(s, "define:") > 0) Then
            '            Udtryk = Replace(Udtryk, "definer:", "")
            '            Udtryk = Replace(Udtryk, "Definer:", "")
            '            Udtryk = Replace(Udtryk, "define:", "")
            '            Udtryk = Replace(Udtryk, "Define:", "")
            Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
            Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
            Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
            Udtryk = Trim(Udtryk)
            If Len(Udtryk) > 0 Then
                If InStr(Udtryk, "matrix") < 1 Then
                    If InStr(Udtryk, "=") > 0 Then
                        Arr = Split(Udtryk, "=")
                        LHS = Arr(0)
                        RHS = Arr(1)
                        ea.Text = LHS
                        fktnavn = ea.GetNextVar(1)
                        varnavn = ea.GetNextBracketContent(1)
                    
                        If LHS = fktnavn & "(" & varnavn & ")" Then
                            ea.Text = RHS
                            ea.Pos = 1
                            ea.ReplaceVar varnavn, "x"
                            fktudtryk = ea.Text
                            DefinerKonstanter fktudtryk, DefList, geogebrafil
                            geogebrafil.CreateFunction fktnavn, fktudtryk, False, True
                        ElseIf LHS = "y" Then
                            fktudtryk = ReplaceIndepvarX(RHS, , DefList)
                            If fktudtryk <> vbNullString Then
                                DefinerKonstanter fktudtryk, DefList, geogebrafil
                                geogebrafil.CreateFunction "f" & j, fktudtryk, False
                                j = j + 1
                            End If
                        ElseIf InStr(LHS, VBA.ChrW(8407)) > 0 Then ' arrow -> vector
                            If InStr(RHS, "¦") > 0 Then ' vector inserted using template from equation menu
                                RHS = Replace(RHS, "¦", ";")
                                geogebrafil.CreateVector fktnavn, RHS, False, True
                            ElseIf Right(LHS, 1) = VBA.ChrW(8407) Then ' vector
                                RHS = Replace(RHS, VBA.ChrW(9608), "")
                                RHS = Replace(RHS, VBA.ChrW(183), "*")
                                RHS = Replace(RHS, ",", ".")
                                RHS = Replace(RHS, "@", ";")
                                RHS = Replace(RHS, "((", "(")
                                RHS = Replace(RHS, "))", ")")
                                geogebrafil.CreateVector fktnavn, RHS, False, True
                            End If
                        ElseIf LHS = "(x¦y)" Then 'paramettric plot inserted using template
                            RHS = Replace(RHS, "¦", ";")
                            RHS = Replace(RHS, VBA.ChrW(183), "*")
                            RHS = Replace(RHS, ",", ".")
                            fktudtryk = "param1: X = " & RHS
                            geogebrafil.CreateEquation "param" & j, fktudtryk, False, True
                            j = j + 1
                        ElseIf LHS = "(" & VBA.ChrW(9608) & "(x@y))" Then 'parametric plot
                            RHS = Replace(RHS, VBA.ChrW(9608), "")
                            RHS = Replace(RHS, VBA.ChrW(183), "*")
                            RHS = Replace(RHS, ",", ".")
                            RHS = Replace(RHS, "@", ";")
                            RHS = Replace(RHS, "((", "(")
                            RHS = Replace(RHS, "))", ")")
                            fktudtryk = "param1: X = " & RHS
                            geogebrafil.CreateEquation "param" & j, fktudtryk, False, True
                            j = j + 1
                        ElseIf fktnavn = LHS Then
                            fktudtryk = ReplaceIndepvarX(RHS, , DefList)
                            If fktudtryk <> vbNullString Then
                                DefinerKonstanter fktudtryk, DefList, geogebrafil
                                geogebrafil.CreateEquation LHS & j, RHS, False, True
                                j = j + 1
                            End If
                        Else 'ligning
                            DefinerKonstanter fktudtryk, DefList, geogebrafil
                            geogebrafil.CreateEquation "eq" & j, LHS & "=" & RHS, False, True
                            j = j + 1
                        End If
                    ElseIf InStr(Udtryk, "¦") > 0 Then ' vector inserted using template from equation menu
                        Udtryk = Replace(Udtryk, "¦", ";")
                        geogebrafil.CreateVector "v", Udtryk, False, True
                    ElseIf Left(Udtryk, 3) = "(" & VBA.ChrW(9608) & "(" Then ' vector
                        Udtryk = Replace(Udtryk, VBA.ChrW(9608), "")
                        Udtryk = Replace(Udtryk, VBA.ChrW(183), "*")
                        Udtryk = Replace(Udtryk, ",", ".")
                        Udtryk = Replace(Udtryk, "@", ";")
                        Udtryk = Replace(Udtryk, "((", "(")
                        Udtryk = Replace(Udtryk, "))", ")")
                        geogebrafil.CreateVector "v", Udtryk, False, True
                    ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                        ' can only be used with GeoGebra 4.0
                        DefinerKonstanter Udtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction "u" & j, Udtryk, True
                    Else
                        Udtryk = ReplaceIndepvarX(Udtryk, , DefList)
                        If Udtryk <> vbNullString Then
                            DefinerKonstanter Udtryk, DefList, geogebrafil
                            geogebrafil.CreateFunction "f" & j, Udtryk, False
                            j = j + 1
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    'datapoints
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
    
    If fileExists(geogebrasti & "geogebra.xml") Then
        Kill geogebrasti & "geogebra.xml"
    End If
    geogebrafil.Save geogebrasti & "geogebra.xml"
    
    DoEvents
    i = 0
    Do While i < 10 And Not (fileExists(geogebrasti & "geogebra.xml"))
        Wait 1
        i = i + 1
    Loop
    '    CreateZipFile geogebrasti & "geogebra.zip", geogebrasti & "geogebra_thumbnail.png", geogebrasti & "geogebra.xml"
    If fileExists(geogebrasti & "geogebra.ggb") Then
        Kill geogebrasti & "geogebra.ggb"
    End If
#If Mac Then
    CreateZipFile geogebrasti & "geogebra.ggb", geogebrasti & "geogebra.xml"
#Else
    CreateZipFile geogebrasti & "geogebra.zip", geogebrasti & "geogebra.xml"
    Name geogebrasti & "geogebra.zip" As geogebrasti & "geogebra.ggb"
#End If
    GoTo slut
Fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
    On Error Resume Next
    omax.ConvertLnLog = True
End Sub
Function DefinerKonstanter(Expr As String, ByRef DefList As String, Optional ByRef geogebrafil As CGeoGebraFile = Nothing, Optional ByRef UrlLink As String) As String
' define variables not defined in expr
' deflist is a list of variables that are defined
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim Var As String, i As Integer
Dim varval As String
    ea.Text = DefList
    ea2.Text = Expr
    ea2.Pos = 0
    Do
        Var = ea2.GetNextVar
'        MsgBox AscW(var) & vbCrLf & VBA.ChrW(960)
        ea2.Pos = ea2.Pos + 1
'        If var = "z" Then geogebrafil.Show3D = True  ' not yet implemented
        If Not (ea2.ChrByIndex(ea2.Pos) = "(") And Not (Left(Var, 1) = "_") And Not (ea.IsFunction(Var)) And Not (ea.ContainsVar(Var)) And Var <> "" And Var <> "x" And Var <> "y" And Var <> "z" And Var <> "e" And Var <> "pi" And Var <> "matrix" And Var <> "if" And Var <> "elseif" And Var <> "then" And Var <> "and" And Var <> "or" And Var <> "else" And Var <> "amp" And Var <> "infinity" And Var <> VBA.ChrW(960) Then  ' 960=pi
            varval = "1"
            If Len(varval) > 0 Then
                If Not geogebrafil Is Nothing Then
                    geogebrafil.CreateFunction Var, varval, False, False
                Else
'                    UrlLink = UrlLink & Var & "=" & varval & ";"
                    UrlLink = UrlLink & Var & "=1;" & Var & "=slider(-5,5,0.1,1,100,false,true,true,false);"
                End If
                DefList = DefList & "," & Var
                ea.Text = DefList
                i = i + 1
            End If
        End If
    Loop While Var <> ""
        
    DefinerKonstanter = Var
End Function

Sub CreateZipFile(zipfilnavn As Variant, FilNavn As Variant, Optional filnavn2 As Variant = "", Optional filnavn3 As Variant = "", Optional filnavn4 As Variant = "")
#If Mac Then
    RunScript "ZipFile", zipfilnavn & ";" & FilNavn
#Else
    PrepareMaxima False
    MaxProc.CreateZipFile zipfilnavn, FilNavn
#End If
GoTo slut
Fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Sub
 


