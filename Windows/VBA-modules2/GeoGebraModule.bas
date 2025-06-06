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
    ' Bruger Browser til at plotte med GeoGebra.
    ' Der er ingen funktioner der anvender parametrene. Hvis der bare k�res GeoGebraWeb uden parametre, s� er det plotning, s� det er det eneste denne funktion bruges til
    ' Funktionen klarg�r til at sende til OpenGeoGebra, der rent faktisk �bner edge og den rigtige html-fil
    ' l�ser valgte ligning/funktion
    ' Definitioner findes og inds�ttes ogs�
    

    ' gtype="", "3d", "CAS"
    Dim cmd As String, UrlLink As String, s As String
    Dim sl As New CSortList
    Dim Var As String, DefList As String
    Dim k As Integer, i As Integer, j As Integer, p As Integer
    Dim Arr As Variant, uvar As String
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
    Dim TempCas As Integer
    Dim VektNArr As Variant, VNi As Integer
    Dim UdtArr() As String
    VektNArr = Array("a", "b", "c", "v", "w")

    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
    ea2.SetNormalBrackets

    On Error GoTo Fejl

    TempCas = CASengine
    CASengineTempOnly = 1

    PrepareMaxima ' finder bla. definitioner
    
    omax.ReadSelection ' l�ser valgte funktion

    If Gtype <> "CAS" Then
        ' s�t definitioner i rigtig r�kkef�lge
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

        '    ' definer variable der ikke er defineret
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
            cmd = Replace(Replace(ConvertToGeogebraSyntax(cmd, False), "+", "%2B"), "&", "%26") & ";" ' v.1.26 tilf�jet false ved konvertering fordi den vidst allerede er konverteret
            UrlLink = UrlLink & cmd
        Next
    End If

    j = 1
    ' inds�t de markerede funktioner
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
                    If InStr(Udtryk, "matrix") < 1 Then ' matricer er ikke implementeret endnu
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
                                fktudtryk = ReplaceIndepvarX(RHS, uvar)
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
                            ElseIf LHS = "({{x},{y}})" Then 'parameterfremstilling
                                RHS = Replace(RHS, "{", "(")
                                RHS = Replace(RHS, "}", ")")
                                RHS = Replace(RHS, "((", "(")
                                RHS = Replace(RHS, "))", ")")
                                cmd = "Param:X=" & RHS
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            ElseIf right(LHS, 3) = "pil" Then ' vektor
                                LHS = Left(LHS, Len(LHS) - 3)
                                RHS = Replace(RHS, "{", "(")
                                RHS = Replace(RHS, "}", ")")
                                cmd = LHS & "=vector((0,0)," & RHS & ")"
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            Else ' ligning
                                cmd = LHS & "=" & RHS
                                DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                                cmd = Replace(cmd, "+", "%2B") & ";"
                                UrlLink = UrlLink & cmd
                            End If
                        ElseIf Left(Udtryk, 3) = "({{" Then ' vektor
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
                        Else
                            Udtryk = ReplaceIndepvarX(Udtryk)
                            DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                            cmd = "f" & j & "=" & Udtryk
                            cmd = Replace(cmd, "+", "%2B") & ";"
                            UrlLink = UrlLink & cmd
                            '                    geogebrafil.CreateFunction "f" & j, udtryk, False
                            j = j + 1
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
    
    OpenGeoGebraWeb UrlLink, Gtype, False, False ' v.1.26 false, false tilf�jet da definitioner kom med to gange
Fejl:

slut:
    CASengineTempOnly = TempCas
End Sub

Sub OpenGeoGebraWeb(ByVal cmd As String, Gtype As String, Optional ConvertSyntax As Boolean = False, Optional UseDefs As Boolean = True)
' �bner GeoGebra i Edge.
' Gtype="" til plotning. �bner i Calculator suite
' Gtype="classic" til plotning. bruges af h�ldningsfelt
' Gtype="CAS"  �bner GeoGebraCASapplet.html
' Funktionen l�ser ikke noget i dokumentet. Preparemaxima skal v�re k�rt inden, hvor definitioner findes, n�r UseDefs=true anvendes
' cmd tilf�jes i slutningen af url'en med ?command=       Definitioner tilf�jes ogs� til command

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
    
    If Len(cmd) > 0 Then If right(cmd, 1) = ";" Then cmd = Left(cmd, Len(cmd) - 1)
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
' Finder GeoGebramath apps i Appdata eller programfiles. Der s�ges f�rst efter hvad sidste installation
#If Mac Then
    GetGeoGebraMathAppsFolder = "/Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/"
#Else
    Dim DN As String, RK As String
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
' s�tter stregene GeoGebraDefs og GeoGebraAssumes ud fra omax
Dim Arr() As String, i As Integer
'    MsgBox omax.DefString
'    MsgBox omax.defstringtext ' alle definitioner og assumes, men i maxima syntaks
    
    GeoGebraAssumes = ""
    GeoGebraDefs = ""
    
    ' Tilf�j definitioner og assumes der er skrevet i mat-felter
    Arr = Split(omax.DefString, ";")
    For i = 0 To UBound(Arr) - 1
        If Left(Arr(i), 7) = "assume(" Then
            GeoGebraAssumes = GeoGebraAssumes & Mid(Arr(i), 8, Len(Arr(i)) - 8) & ChrW(8743)
        Else
            GeoGebraDefs = GeoGebraDefs & ConvertToGeogebraSyntax(Arr(i), False) & ";"
        End If
    Next
    
    ' Tilf�j midlertidige definitioner og assumes
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
' bruges ikke mere, men m�ske er metoden bedre, da den g�r noget med r�kkef�lgen
    Dim cmd As String, i As Integer, k As Integer, j As Integer, fktudtryk As String, UrlLink As String, p As Integer
    Dim sl As New CSortList
    Dim Var As String, DefList As String
    Dim ea As New ExpressionAnalyser
    
    ea.SetNormalBrackets

    ' s�t definitioner i rigtig r�kkef�lge
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
        Dim res As String, i As Integer, s As String
        Dim UFwait2 As UserFormWaitForMaxima
'        On Error Resume Next

        If UseDefs Then FindGeoGebraDefsAndAssumes
        
        If cmd = "" Then cmd = omax.Kommando
        cmd = ConvertToGeogebraSyntax(cmd)
        omax.StartKommando = cmd
        res = ExecuteGeoGebraCasCommand(cmd, UseDefs)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd & "#?" & Defliste)
'            Res = RunScript("ExecuteGeoGebraCASCommand", Cmd)
        i = 0
        If Left(res, 9) = "reloading" Then
            Set UFwait2 = New UserFormWaitForMaxima
            UFwait2.Label_tip.Font.Size = 10
            UFwait2.Label_tip.Font.Italic = False
            UFwait2.Show vbModeless
            UFwait2.Label_tip.Caption = Sprog.A(684)
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
            UFwait2.Label_tip.Caption = Sprog.A(684)
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
'    Dim geogebrafil As New CGeoGebraFile
' definitioner vil allerede v�re k�rt igennem codeforMaxima, s� der skal convertmaxima v�re false

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

      'gr�ske bogstaver
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
      Text = Replace(Text, "eta", VBA.ChrW(951)) ' skal v�re sidst da eta indg�r i andre
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
   
      ' integrate erstattes med Integral. I GeoGebra overs�ttes Integral automatisk til NItegral eller IntegralSymbolic afh af omst�ndigheder.
      ' Der ser dog ud til at v�re et problem med Integral(f(x),x,0,1)  man kan ikke selv angive en variabel, selvom det st�r at det er en mulighed
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
   ea.ReplaceVar "log", "ln" ' Texten vil v�re k�rt gennem codeformaxima, s� log vil v�re ln. Det skal skiftes tilbage
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
          Arr = Split(s, ",")
          If UBound(Arr) > 0 Then Text = Left(Text, p - 1) & "log(" & Arr(1) & "," & Arr(0) & right(Text, Len(Text) - p - Len(s) - 7)
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
         
'         If Left(Text, 1) = "(" Then Text = right(Text, Len(Text) - 1) ' denne gav fejl, s� der manglede startparentes. ved ikke hvorfor den er der

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
    Dim p As Long, ea As New ExpressionAnalyser, mtext As String, s As String
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
    
    'potenser kan v�re skrevet som specialtegn for 0,1,2,3    p = InStr(ea.text, VBA.ChrW(176)) ' h�vet 0
    ReplaceSuperScripts ea
    
    Text = ea.Text
    Text = Replace(Text, "\u003C", "<") ' m�rkeligt at lige pr�cis denne st�r s�dan
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
        mtext = Replace(mtext, " ", "") ' mellem efter ; giver problemer i n�ste linje
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
' alts� retter op p� output fra converttowordsymols
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
    On Error GoTo Fejl
    Dim geogebrasti As String
    Dim geogebrafilersti As String
    Dim geogebracmd As String
    Dim appnr As Long
    Dim UfWait As New UserFormWaitForMaxima
    Dim TempCas As Integer
    
    TempCas = CASengine
    CASengineTempOnly = 1
    
    PrepareMaxima ' omax bliver brugt
    
    UfWait.Label_tip.Caption = Sprog.A(362)
    UfWait.Label_progress.Caption = "***"
    UfWait.Label_stop.visible = False
    UfWait.Show vbModeless
    
    geogebrasti = GeoGebraPath()
    If geogebrasti = "" Then ' hvis geogebra ikke installeret
        Unload UfWait
        InstallGeoGebra
        GoTo slut
    End If
'    geogebrafilersti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    geogebrafilersti = GetTempDir()
    
    CreateGeoGebraFil geogebrafilersti
    
    geogebrafilersti = geogebrafilersti & "geogebra.ggb"
    
    UfWait.Label_progress.Caption = "******"
    If fileExists(geogebrafilersti) Then ' check om geogebrafilen er lavet
        geogebracmd = geogebrasti & " """ & geogebrafilersti & """"
    Else
        MsgBox "The GeoGebra.ggb file cannot be located", vbOKOnly, Sprog.Error
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

'    Dim UfWait As UserFormWaitForMaxima
    
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
'    MsgBox "The download page will now open. Install GeoGebra classic 5", vbOKOnly, "Download"
'    OpenLink "https://www.geogebra.org/download"
#Else ' win

    If UserFormGeoGebra.ReturnVal = 1 Then
    Else
        GeoGebraWeb
    End If
#End If
    
    GoTo slut
Fejl:

slut:
'    If Not UfWait Is Nothing Then Unload UfWait
End Sub
Function GeoGebraPath() As String
' path to the geogebra executable. Returns "" if not found. OBS: P� mac bruges stien ikke. Der er applescript til det, men funktionen bruges til at afg�re om der er en GeoGebra installation.
' Der hentes til den nyeste version hvis muligt
Dim DN As String
On Error GoTo Fejl
#If Mac Then
    GeoGebraPath = GetProgramFilesDir() & "GeoGebra.app"
    If fileExists(GeoGebraPath) Then Exit Function
'    GeoGebraPath = GetProgramFilesDir() & "GeoGebra 5.app"
'    If FileExists(GeoGebraPath) Then Exit Function
' Disse kan godt startes p� Mac, men de kan ikke �bne en ggb-fil som argument
'    GeoGebraPath = GetProgramFilesDir() & "GeoGebra 6.app"
'    If FileExists(GeoGebraPath) Then Exit Function
'    GeoGebraPath = GetProgramFilesDir() & "GeoGebra Graphing Calculator.app"
'    If FileExists(GeoGebraPath) Then Exit Function
    GeoGebraPath = ""
#Else ' Windows
    
    GeoGebraPath = Dir(GetProgramFilesDir & "\GeoGebra 5*", vbDirectory)
    If GeoGebraPath <> "" Then
        DN = GeoGebraPath
        Do While GeoGebraPath <> "" ' vi henter den GeoGebra 5 med h�jst versions nr. Den vil v�re sidst p� listen
            GeoGebraPath = Dir()
            If GeoGebraPath <> "" Then DN = GeoGebraPath
        Loop
        If DN <> "" Then
            GeoGebraPath = """" & GetProgramFilesDir & "\" & DN & "\GeoGebra.exe"""
            GoTo slut
        End If
    End If
    
    If GeoGebraPath = "" Then 'Matematikv�rkt�jssuite
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
    
    If GeoGebraPath = "" Then ' Graftegner
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
    Wait 1 ' anden gang man �bner geogebra p� mac l�ser word hvis denne ikke er der. Det har noget med finddefinitions at g�re, men har ikke kunnet se hvad
#End If

    PrepareMaxima
    '    omax.FindDefinitions ' bliver fundet i preparemaxima
    omax.ReadSelection

    '    defs = ConvertToAscii(omax.defstring)
    '    arr = Split(defs, ";")
    Dim sl As New CSortList
    Dim Var As String, DefList As String
    Dim k As Integer
    
    ' s�t definitioner i rigtig r�kkef�lge
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
    
    ' definer variable der ikke er defineret
    ea.Text = DefList
    For i = 0 To sl.Length - 1
        varnavn = sl.GetName(i)
        fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
        p = InStr(sl.GetName(i), "(")
        '        If InStr(sl.GetVal(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
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
    ' definer de definerede funktioner/variable
    '    For i = 0 To sl.Length - 1
    '    Next
    
    j = 1
    ' inds�t de markerede funktioner
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
                If InStr(Udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
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
                            fktudtryk = ReplaceIndepvarX(RHS)
                            DefinerKonstanter fktudtryk, DefList, geogebrafil
                            geogebrafil.CreateFunction "f" & j, fktudtryk, False
                            j = j + 1
                        ElseIf InStr(LHS, VBA.ChrW(8407)) > 0 Then ' pil -> vektor
                            If InStr(RHS, "�") > 0 Then ' vektor med skabelon
                                RHS = Replace(RHS, "�", ";")
                                geogebrafil.CreateVector fktnavn, RHS, False, True
                            ElseIf right(LHS, 1) = VBA.ChrW(8407) Then ' vektor
                                RHS = Replace(RHS, VBA.ChrW(9608), "")
                                RHS = Replace(RHS, VBA.ChrW(183), "*")
                                RHS = Replace(RHS, ",", ".")
                                RHS = Replace(RHS, "@", ";")
                                RHS = Replace(RHS, "((", "(")
                                RHS = Replace(RHS, "))", ")")
                                geogebrafil.CreateVector fktnavn, RHS, False, True
                            End If
                        ElseIf LHS = "(x�y)" Then 'parameterfremstilling med skabelon
                            RHS = Replace(RHS, "�", ";")
                            RHS = Replace(RHS, VBA.ChrW(183), "*")
                            RHS = Replace(RHS, ",", ".")
                            fktudtryk = "param1: X = " & RHS
                            geogebrafil.CreateEquation "param" & j, fktudtryk, False, True
                            j = j + 1
                        ElseIf LHS = "(" & VBA.ChrW(9608) & "(x@y))" Then 'parameterfremstilling
                            RHS = Replace(RHS, VBA.ChrW(9608), "")
                            RHS = Replace(RHS, VBA.ChrW(183), "*")
                            RHS = Replace(RHS, ",", ".")
                            RHS = Replace(RHS, "@", ";")
                            RHS = Replace(RHS, "((", "(")
                            RHS = Replace(RHS, "))", ")")
                            fktudtryk = "param1: X = " & RHS
                            geogebrafil.CreateEquation "param" & j, fktudtryk, False, True
                            j = j + 1
                        Else 'ligning
                            DefinerKonstanter fktudtryk, DefList, geogebrafil
                            geogebrafil.CreateEquation "eq" & j, LHS & "=" & RHS, False, True
                            j = j + 1
                        End If
                    ElseIf InStr(Udtryk, "�") > 0 Then ' vektor med skabelon
                        Udtryk = Replace(Udtryk, "�", ";")
                        geogebrafil.CreateVector "v", Udtryk, False, True
                    ElseIf Left(Udtryk, 3) = "(" & VBA.ChrW(9608) & "(" Then ' vektor
                        Udtryk = Replace(Udtryk, VBA.ChrW(9608), "")
                        Udtryk = Replace(Udtryk, VBA.ChrW(183), "*")
                        Udtryk = Replace(Udtryk, ",", ".")
                        Udtryk = Replace(Udtryk, "@", ";")
                        Udtryk = Replace(Udtryk, "((", "(")
                        Udtryk = Replace(Udtryk, "))", ")")
                        geogebrafil.CreateVector "v", Udtryk, False, True
                    ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                        ' kan f�rst bruges med GeoGebra 4.0
                        DefinerKonstanter Udtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction "u" & j, Udtryk, True
                    Else
                        Udtryk = ReplaceIndepvarX(Udtryk)
                        DefinerKonstanter Udtryk, DefList, geogebrafil
                        geogebrafil.CreateFunction "f" & j, Udtryk, False
                        j = j + 1
                    End If
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
    
    ' Problemet med dette er at tabellen skal v�re lodret
    '    dd.ReadSelection
    '    If Len(dd.GetSetForm) > 4 Then
    '        geogebrafil.CreateList "punkter", dd.GetSetForm
    '    End If

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
Dim Var As String, i As Integer
Dim varval As String
    ea.Text = DefList
    ea2.Text = Expr
    ea2.Pos = 0
    Do
        Var = ea2.GetNextVar
'        MsgBox AscW(var) & vbCrLf & VBA.ChrW(960)
        ea2.Pos = ea2.Pos + 1
'        If var = "z" Then geogebrafil.Show3D = True  ' ikke implementeret endnu
        If Not (ea2.ChrByIndex(ea2.Pos) = "(") And Not (Left(Var, 1) = "_") And Not (ea.IsFunction(Var)) And Not (ea.ContainsVar(Var)) And Var <> "" And Var <> "x" And Var <> "y" And Var <> "z" And Var <> "e" And Var <> "pi" And Var <> "matrix" And Var <> "if" And Var <> "elseif" And Var <> "then" And Var <> "and" And Var <> "or" And Var <> "else" And Var <> "amp" And Var <> "infinity" And Var <> VBA.ChrW(960) Then  ' 960=pi
'            varval = InputBox(Sprog.A(363) & " " & Var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
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
    
'    If i > 1 Then ' kan ikke lige se hvad denne sektion er n�dvendig for
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
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:

End Sub
 


