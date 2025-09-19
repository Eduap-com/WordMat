Attribute VB_Name = "Grafer"
Option Explicit
Public UF2Dgraph As UserForm2DGraph
Public ReplacedVar As String  ' contains var which was replaced by x by replaceindepvarx
Public Sub StandardPlot()
    If GraphApp = 0 Then
        Plot2DGraph
    ElseIf GraphApp = 1 Then
        InsertGraphOleObject
    ElseIf GraphApp = 2 Then
        GeoGebra
    ElseIf GraphApp = 3 Then
        InsertChart
    ElseIf GraphApp = 4 Then
        GeoGebraWeb ""
    End If
    
End Sub
Public Sub Plot2DGraph()
' gnuplot
    Dim forskrifter As String
    Dim arr As Variant
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    On Error GoTo fejl
    Dim sstart As Long, sslut As Long
    Dim TempCas As Integer
    
#If Mac Then
    If MsgBox("Support for gnuplot on Mac has ended. You will be redirected to GeoGebra", vbOKCancel, "No GnuPlot") = vbOK Then
        GraphApp = 2
        GeoGebra
    End If
    Exit Sub
#End If

    sstart = Selection.start
    sslut = Selection.End
    
    TempCas = CASengine
    CASengineTempOnly = 0
    Set UF2Dgraph = New UserForm2DGraph
    PrepareMaxima
    Dim Sel As Selection
    Set Sel = Selection
    If Sel.InlineShapes.Count > 0 Then
        If Sel.InlineShapes(1).Type = wdInlineShapePicture Then
            s = Sel.InlineShapes(1).AlternativeText
            arr = Split(s, "|")
            If arr(0) = "WordMat" Then
                On Error Resume Next
                PicOpen = False
                UF2Dgraph.TextBox_definitioner.text = arr(2)
                UF2Dgraph.TextBox_titel.text = arr(3)
                UF2Dgraph.TextBox_xaksetitel.text = arr(4)
                UF2Dgraph.TextBox_yaksetitel.text = arr(5)
        
                UF2Dgraph.TextBox_xmin.text = arr(6)
                UF2Dgraph.TextBox_xmax.text = arr(7)
                UF2Dgraph.TextBox_ymin.text = arr(8)
                UF2Dgraph.TextBox_ymax.text = arr(9)
        
                UF2Dgraph.TextBox_ligning1.text = arr(10)
                UF2Dgraph.TextBox_var1.text = arr(11)
                UF2Dgraph.TextBox_xmin1.text = arr(12)
                UF2Dgraph.TextBox_xmax1.text = arr(13)
                UF2Dgraph.ComboBox_ligning1.ListIndex = CInt(arr(14))
        
                UF2Dgraph.TextBox_ligning2.text = arr(15)
                UF2Dgraph.TextBox_var2.text = arr(16)
                UF2Dgraph.TextBox_xmin2.text = arr(17)
                UF2Dgraph.TextBox_xmax2.text = arr(18)
                UF2Dgraph.ComboBox_ligning2.ListIndex = CInt(arr(19))
        
                UF2Dgraph.TextBox_ligning3.text = arr(20)
                UF2Dgraph.TextBox_var3.text = arr(21)
                UF2Dgraph.TextBox_xmin3.text = arr(22)
                UF2Dgraph.TextBox_xmax3.text = arr(23)
                UF2Dgraph.ComboBox_ligning3.ListIndex = CInt(arr(24))
        
                UF2Dgraph.TextBox_ligning4.text = arr(25)
                UF2Dgraph.TextBox_var4.text = arr(26)
                UF2Dgraph.TextBox_xmin4.text = arr(27)
                UF2Dgraph.TextBox_xmax4.text = arr(28)
                UF2Dgraph.ComboBox_ligning4.ListIndex = CInt(arr(29))
        
                UF2Dgraph.TextBox_ligning5.text = arr(30)
                UF2Dgraph.TextBox_var5.text = arr(31)
                UF2Dgraph.TextBox_xmin5.text = arr(32)
                UF2Dgraph.TextBox_xmax5.text = arr(33)
                UF2Dgraph.ComboBox_ligning5.ListIndex = CInt(arr(34))
        
                UF2Dgraph.TextBox_ligning6.text = arr(35)
                UF2Dgraph.TextBox_var6.text = arr(36)
                UF2Dgraph.TextBox_xmin6.text = arr(37)
                UF2Dgraph.TextBox_xmax6.text = arr(38)
                UF2Dgraph.ComboBox_ligning6.ListIndex = CInt(arr(39))
        
                UF2Dgraph.TextBox_lig1.text = arr(40)
                UF2Dgraph.TextBox_lig2.text = arr(41)
                UF2Dgraph.TextBox_Lig3.text = arr(42)
        
                UF2Dgraph.TextBox_parametric1x.text = arr(43)
                UF2Dgraph.TextBox_parametric1y.text = arr(44)
                UF2Dgraph.TextBox_tmin1.text = arr(45)
                UF2Dgraph.TextBox_tmax1.text = arr(46)
                UF2Dgraph.TextBox_parametric2x.text = arr(47)
                UF2Dgraph.TextBox_parametric2y.text = arr(48)
                UF2Dgraph.TextBox_tmin2.text = arr(49)
                UF2Dgraph.TextBox_tmax2.text = arr(50)
                UF2Dgraph.TextBox_parametric3x.text = arr(51)
                UF2Dgraph.TextBox_parametric3y.text = arr(51)
                UF2Dgraph.TextBox_tmin3.text = arr(53)
                UF2Dgraph.TextBox_tmax3.text = arr(54)
        
                UF2Dgraph.TextBox_punkter.text = arr(55)
                UF2Dgraph.TextBox_punkter2.text = arr(56)
                UF2Dgraph.TextBox_markerpunkter.text = arr(57)
                UF2Dgraph.CheckBox_pointsjoined.Value = CBool(arr(58))
                UF2Dgraph.CheckBox_pointsjoined2.Value = CBool(arr(59))
                UF2Dgraph.TextBox_pointsize.text = arr(60)
                UF2Dgraph.TextBox_pointsize2.text = arr(61)
        
                UF2Dgraph.TextBox_vektorer.text = arr(62)
                UF2Dgraph.TextBox_labels.text = arr(63)
        
                UF2Dgraph.CheckBox_gitter.Value = CBool(arr(64))
                UF2Dgraph.CheckBox_logx.Value = CBool(arr(65))
                UF2Dgraph.CheckBox_logy.Value = CBool(arr(66))
                UF2Dgraph.CheckBox_visforklaring.Value = CBool(arr(67))
                
                Sel.InlineShapes(1).Select
                UF2Dgraph.Show vbModeless
                PicOpen = True
                GoTo slut
            End If
        End If
    End If
    
    
    omax.ReadSelection
        
    forskrifter = omax.KommandoerStreng
    
    If Len(forskrifter) > 1 Then
        arr = Split(forskrifter, ListSeparator)
        For i = 0 To UBound(arr)
            arr(i) = Trim$(Replace(arr(i), "  ", " ")) ' do not remove all spaces, for example 1/x 3 becomes 1/x3 where x3 is variable
            If arr(i) <> "" Then InsertNextEquation (arr(i))
        Next
    End If
    
    'datapoints
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, xmin As Double, xmax As Double
        Cregr.GetTableData
        If UF2Dgraph.TextBox_punkter.text <> "" Then UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & VbCrLfMac
        xmin = Cregr.XValues(1)
        xmax = Cregr.XValues(1)
        For j = 1 To UBound(Cregr.XValues)
            '            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & ConvertNumber(Cregr.XValues(j)) & ListSeparator & ConvertNumber(Cregr.YValues(j)) & VbCrLfMac
            If Cregr.XValues(j) > xmax Then xmax = Cregr.XValues(j)
            If Cregr.XValues(j) < xmin Then xmin = Cregr.XValues(j)
        Next
        UF2Dgraph.TextBox_xmin = xmin
        UF2Dgraph.TextBox_xmax = xmax
    End If
    
    '    UserForm2DGraph.Show vbModeless
    Selection.End = sslut ' The end must come first, otherwise things will go wrong.
    Selection.start = sstart
    UF2Dgraph.Show vbModeless
    
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
    CASengineTempOnly = TempCas
End Sub
Sub InsertNextEquation(Ligning As String)
    Dim arr As Variant
    On Error GoTo fejl
    Ligning = Replace(Ligning, VBA.ChrW$(8788), "=") ' :=
    Ligning = Replace(Ligning, VBA.ChrW$(8797), "=") ' tripel =
    Ligning = Replace(Ligning, VBA.ChrW$(8801), "=") ' def =

    arr = Split(Ligning, "=")

    'If Not (InStr(Ligning, VBA.chrw$(9608)) > 0 And InStr(Ligning, VBA.chrw$(9508)) > 0) Then ' tuborg
    arr = Split(arr(UBound(arr)), VBA.ChrW$(8776)) ' for inserting the rule itself instead of f(x)
    Ligning = omax.ConvertToAscii(arr(UBound(arr)))
    'End If
    Ligning = omax.ConvertToAscii(Trim$(Replace(Replace(Replace(Replace(arr(0), "Definer:", ""), "Define:", ""), "definer:", ""), "define:", "")))

    If UF2Dgraph.TextBox_ligning1.text = Ligning Then
        Exit Sub
    ElseIf UF2Dgraph.TextBox_ligning2.text = Ligning Then
        Exit Sub
    ElseIf UF2Dgraph.TextBox_ligning3.text = Ligning Then
        Exit Sub
    ElseIf UF2Dgraph.TextBox_ligning4.text = Ligning Then
        Exit Sub
    ElseIf UF2Dgraph.TextBox_ligning5.text = Ligning Then
        Exit Sub
    ElseIf UF2Dgraph.TextBox_ligning6.text = Ligning Then
        Exit Sub
    End If

    If UF2Dgraph.TextBox_ligning1.text = "" Then
        UF2Dgraph.TextBox_ligning1.text = Ligning
    ElseIf UF2Dgraph.TextBox_ligning2.text = "" Then
        UF2Dgraph.TextBox_ligning2.text = Ligning
    ElseIf UF2Dgraph.TextBox_ligning3.text = "" Then
        UF2Dgraph.TextBox_ligning3.text = Ligning
    ElseIf UF2Dgraph.TextBox_ligning4.text = "" Then
        UF2Dgraph.TextBox_ligning4.text = Ligning
    ElseIf UF2Dgraph.TextBox_ligning5.text = "" Then
        UF2Dgraph.TextBox_ligning5.text = Ligning
    ElseIf UF2Dgraph.TextBox_ligning6.text = "" Then
        UF2Dgraph.TextBox_ligning6.text = Ligning
    End If
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub PlotDF(Optional DE As String, Optional IndepVar As String = "x", Optional DepVar As String = "y", Optional DePoint As String = "(1, 2)")
' plot direction field
    Dim s As String, v As String
    Dim arr As Variant
    Dim ea As New ExpressionAnalyser
  '  On Error GoTo fejl
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
        
    If DE = vbNullString Then
        PrepareMaxima
        omax.ReadSelection
    Else
        omax.Kommando = DE
    End If
#If Mac Then
#Else
   If GraphApp > 0 Then
#End If
                
        s = Trim$(omax.Kommando)
        s = GetCmdAfterEqualSign(s)
        If s = "" Then
            MsgBox2 TT.A(804), vbOKOnly, "Error"
            Exit Sub
        End If
        ea.text = s
        v = ea.GetNextVar
        If v <> "x" And v <> "y" Then
            If v = "t" Then
                ea.ReplaceVar "t", "x"
            ElseIf v = "N" Then
                ea.ReplaceVar v, "y"
            Else
                ea.ReplaceVar v, "y"
            End If
        End If
        s = ea.text
        s = "SlopeField(" & s & ");"
        s = s & "A=" & DePoint & ";Xmin=-100;Xmax=100;Tic=0.1;"
        s = s & "SolveODE(" & ea.text & ", x(A), y(A), Xmin, Tic);" ' y(A) does not work
        s = s & "SolveODE(" & ea.text & ", x(A), y(A), Xmax, Tic)"
        OpenGeoGebraWeb s, "", True, True
        GoTo slut
#If Mac Then
#Else
    End If
#End If
    Set UF2Dgraph = New UserForm2DGraph
           
    If Len(omax.Kommando) > 0 Then
    arr = Split(omax.Kommando, "=")
    omax.Kommando = arr(UBound(arr))
    End If
    UF2Dgraph.TextBox_dfligning.text = omax.ConvertToAscii(omax.Kommando)
    
    omax.FindVariable
    If InStr(omax.Vars, "x") > 0 Then
        UF2Dgraph.TextBox_dfx.text = "x"
    ElseIf InStr(omax.Vars, "t") > 0 Then
        UF2Dgraph.TextBox_dfx.text = "t"
    Else
        UF2Dgraph.TextBox_dfx.text = "x"
    End If
    If InStr(omax.Vars, "y") > 0 Then
        UF2Dgraph.TextBox_dfy.text = "y"
    ElseIf InStr(omax.Vars, "N") > 0 Then
        UF2Dgraph.TextBox_dfy.text = "N"
    Else
        ea.text = omax.Vars
        UF2Dgraph.TextBox_dfy.text = ea.GetNextVar
        If UF2Dgraph.TextBox_dfy.text = "" Then UF2Dgraph.TextBox_dfy.text = "y"
    End If

    Selection.End = sslut ' The end must come first, otherwise things will go wrong.
    Selection.start = sstart
    UF2Dgraph.MultiPage1.Value = 5
    UF2Dgraph.MultiPage1.SetFocus
    UF2Dgraph.Show vbModeless

    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub InsertEmptyGraphOleObject()
' inserts graph object padowan
Dim ils As InlineShape
Application.ScreenUpdating = False

If Not fileExists(GetProgramFilesDir & "\Graph\graph.exe") Then
    Dim result As VbMsgBoxResult
    result = MsgBox(TT.A(366), vbOKCancel, TT.Error)
    If result = vbOK Then
        OpenLink ("https://www.google.dk/search?q=padowan+graph")
    End If
    Exit Sub
End If

' insert using classname
Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="GraphFile", fileName:="", Range:=Selection.Range, LinkToFile:=False, DisplayAsIcon:=False)

Application.ScreenUpdating = True

End Sub
Sub InsertGraphOleObject()
    ' inserts graph object padowan
#If Mac Then
    MsgBox "Sorry. Graph is not supported on Mac.", vbOKOnly, TT.Error
#Else
    Dim path As String
    Dim ils As InlineShape
    Dim arr As Variant
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
    Dim ea As New ExpressionAnalyser
    Dim p As Integer
    Dim sslut As Long
    sslut = Selection.End

    ea.SetNormalBrackets
    Dim UfWait As New UserFormWaitForMaxima
    UfWait.Label_tip.Caption = TT.A(371)
    UfWait.Label_progress.Caption = "***"
    UfWait.Label_stop.visible = False
    UfWait.Show vbModeless
    On Error GoTo fejl
    Application.ScreenUpdating = False

    If Not fileExists(GetProgramFilesDir & "\Graph\graph.exe") Then
        Dim result As VbMsgBoxResult
        result = MsgBox(TT.A(366), vbOKCancel, TT.Error)
        Exit Sub
    End If

    'path = """" & GetProgramFilesDir & "\WordMat\graphtemplate.grf"""
    path = Environ("TEMP") & "\" & "wordmatgraph.grf"
    'path = "c:\wordmatgraph.grf" ' til test

    Dim graphfil As New CGraphFile
    Dim DefList As String, deflist2 As String
    Dim i As Integer
    PrepareMaxima
    omax.ConvertLnLog = False
    omax.FindDefinitions
    omax.ReadSelection
    omax.ConvertLnLog = True
        
    For i = omax.defindex - 1 To 0 Step -1
        DefList = DefList & "," & omax.DefName(i)
    Next
    
    For i = omax.defindex - 1 To 0 Step -1
        '        graphfil.InsertFunction omax.DefValue(i)
        If InStr(omax.DefValue(i), "matrix") < 1 Then
            If Not (InStr(deflist2, omax.DefName(i)) > 0) Then
                deflist2 = deflist2 & "," & omax.DefName(i)
                graphfil.AddCustomFunction omax.DefName(i) & "=" & omax.DefValue(i)
                p = InStr(omax.DefName(i), "(")
                If p > 0 Then
                    graphfil.InsertFunction Left$(omax.DefName(i), p - 1) & "(x)", 0
                Else
                    graphfil.InsertFunction omax.DefName(i), 0
                End If
                DefinerKonstanterGraph omax.DefValue(i), DefList, graphfil
            End If
        End If
    Next
    
    ' selected functions
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim$(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW$(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW$(8797), "=") ' triple =
        Udtryk = Replace(Udtryk, VBA.ChrW$(8801), "=") ' def =
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then
                If InStr(Udtryk, "=") > 0 Then
                    arr = Split(Udtryk, "=")
                    LHS = arr(0)
                    RHS = arr(1)
                    ea.text = LHS
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    If LHS = fktnavn & "(" & varnavn & ")" Then
                        ea.text = RHS
                        ea.pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.text
                        DefinerKonstanterGraph fktudtryk, DefList, graphfil
                        graphfil.InsertFunction fktudtryk
                    Else
                        DefinerKonstanterGraph Udtryk, DefList, graphfil, True
                        graphfil.InsertRelation Udtryk
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW$(8804)) > 0 Or InStr(Udtryk, VBA.ChrW$(8805)) > 0 Then
                    DefinerKonstanterGraph Udtryk, DefList, graphfil, True
                    graphfil.InsertRelation Udtryk
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk, , DefList)
                    If Udtryk <> vbNullString Then
                        DefinerKonstanterGraph Udtryk, DefList, graphfil
                        graphfil.InsertFunction Udtryk
                    End If
                End If
            End If
        End If
    Next
    
    'datapoints
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, setdata As String, j As Integer
        Cregr.GetTableData
        For j = 1 To UBound(Cregr.XValues)
            setdata = setdata & Replace(Cregr.XValues(j), ",", ".") & "," & Replace(Cregr.YValues(j), ",", ".") & ";"
        Next
        If Len(setdata) > 0 Then
            setdata = Left$(setdata, Len(setdata) - 1)
            graphfil.InsertPointSeries setdata
        End If
    End If
    
    Selection.start = sslut
    Selection.End = sslut

    If Selection.OMaths.Count > 0 Then
        omax.GoToEndOfSelectedMaths
    End If
    If Selection.Tables.Count > 0 Then
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse wdCollapseEnd
    End If
    Selection.MoveRight wdCharacter, 1
    Selection.TypeParagraph

    UfWait.Label_progress.Caption = "******"

    If graphfil.funkno > 0 Or Len(graphfil.CustomFunctions) > 0 Or graphfil.relationno > 0 Or graphfil.pointno > 0 Then
        graphfil.Save path

'insert using empty graph file. A little slower, but could be used for communication at some point
        On Error GoTo hop
        Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
        On Error GoTo fejl
        ils.OLEFormat.DoVerb (wdOLEVerbShow)

    Else
        ' insert using classname
hop:
        Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="GraphFile", fileName:="", Range:=Selection.Range, LinkToFile:=False, DisplayAsIcon:=False)
    End If

    DoEvents
    Unload UfWait

    Application.ScreenUpdating = True
    GoTo slut
fejl:
    MsgBox TT.A(97), vbOKOnly, TT.Error
    omax.ConvertLnLog = True
    Unload UfWait
slut:
    omax.ConvertLnLog = True
#End If
End Sub
#If Mac Then
#Else
Sub DefinerKonstanterGraph(Expr As String, DefList As String, ByRef graphfil As CGraphFile, Optional noty As Boolean = False)
' define variables not defined in expr
' deflist is a list of variables that are defined
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim Var As String
    ea.text = DefList
    If noty Then ea.text = ea.text & ",y"
    ea2.text = Expr
    ea2.pos = 0
    Do
        Var = ea2.GetNextVar
        ea2.pos = ea2.pos + 1
        If Not (ea2.ChrByIndex(ea2.pos) = "(") And Not (ea.IsFunction(Var)) And Not (ea.ContainsVar(Var)) And Var <> "" And Var <> "x" And Var <> "y" And Var <> "e" And Var <> "pi" And Var <> "matrix" Then ' maybe not y? copied from geogebra
            graphfil.AddCustomFunction Var & "=" & InputBox(TT.A(363) & " " & Var & vbCrLf & vbCrLf & TT.A(367), TT.A(365), "1")
            DefList = DefList & "," & Var
        End If
    Loop While Var <> ""

End Sub
#End If
Function ReplaceIndepvarX(fkt As String, Optional ByRef uvar = "", Optional DefList As String) As String
    ' makes sure to insert x as independent variable
    ' uvar and global var "ReplacedVar" will after hold the variable which was replaced
    ' Deflist is a list of variables that are defined, and may not be the independent variable

    Dim ea As New ExpressionAnalyser
    Dim Var As String, VarColl As Collection, v As String
    Dim DefArr() As String, i As Integer, DefFound As Boolean
    
    ea.text = fkt
    If ea.ContainsVar("x") Then
        ReplacedVar = "x"
        uvar = "x"
        GoTo slut
    End If

    DefArr = Split(DefList, ",")
    
    If ea.ContainsVar("t") Then
        ReplacedVar = "t"
        uvar = "t"
        ea.ReplaceVar uvar, "x"
        GoTo slut
    End If

    ' create a collection of all variables in expression, that are not already defined
    Set VarColl = New Collection
    ea.pos = -1
    Do
        v = ea.GetNextVar(ea.pos + 1)
        If v <> vbNullString And v <> "matrix" Then
            DefFound = False
            For i = 0 To UBound(DefArr)
                If DefArr(i) = v Then
                    DefFound = True
                    Exit For
                End If
            Next
            If Not DefFound Then VarColl.Add v
        End If
    Loop While v <> vbNullString

    ' if only 1 variable, that variable is assumed to be the independent variable
    If VarColl.Count = 1 Then
        ReplacedVar = VarColl(1)
        uvar = VarColl(1)
        ea.ReplaceVar uvar, "x"
    ElseIf VarColl.Count > 1 Then
        MsgBox2 fkt & " " & TT.A(115), vbOKOnly, TT.Error
        ReplaceIndepvarX = ""
        Exit Function
    End If

slut:
    ReplaceIndepvarX = ea.text
End Function
#If Mac Then
Sub InsertChart()
    Dim dd As New DocData
    Dim ea As New ExpressionAnalyser
    Dim xmin As Double, xmax As Double
    Dim arr As Variant
    Dim i As Integer
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String
    Dim srange As Range
    Dim ScriptDataPoints As String ' "1:2#3:4"
    Dim ScriptFunctions As String ' "2*x+1#3*t-1:t"
    ea.SetNormalBrackets

    Set srange = Selection.Range

    dd.ReadSelection

    DoEvents
    Application.ScreenUpdating = False
    
    Dim UFwait2 As New UserFormWaitForMaxima
    UFwait2.Show vbModeless
    DoEvents
    UFwait2.Label_progress = "***"
    
    ' selected functions
    For i = 0 To dd.AntalMathBoxes - 1
        Udtryk = dd.MathBoxes(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim$(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW$(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW$(8797), "=") ' triple =
        Udtryk = Replace(Udtryk, VBA.ChrW$(8801), "=") ' def =
        Udtryk = Replace(Udtryk, vbCrLf, "") '
        Udtryk = Replace(Udtryk, vbCr, "") '
        Udtryk = Replace(Udtryk, vbLf, "") '
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then
                If InStr(Udtryk, "=") > 0 Then
                    arr = Split(Udtryk, "=")
                    LHS = arr(0)
                    RHS = arr(1)
                    ea.text = LHS
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    If LHS = fktnavn & "(" & varnavn & ")" Then
                        ScriptFunctions = ScriptFunctions & RHS & ":" & varnavn & "#"
                    Else
                        ScriptFunctions = ScriptFunctions & RHS & ":x" & "#"
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW$(8804)) > 0 Or InStr(Udtryk, VBA.ChrW$(8805)) > 0 Then
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
                    ScriptFunctions = ScriptFunctions & Udtryk & ":x" & "#"
                End If
            End If
        End If
    Next
    If Right$(ScriptFunctions, 1) = "#" Then ScriptFunctions = Left$(ScriptFunctions, Len(ScriptFunctions) - 1)
    
    srange.Select
    'datapoints
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression
        Cregr.GetTableData
        xmin = Cregr.XValues(1)
        xmax = Cregr.XValues(1)
        For i = 1 To UBound(Cregr.XValues)
            ScriptDataPoints = ScriptDataPoints & val(Replace(Cregr.XValues(i), ",", ".")) & ":"
            ScriptDataPoints = ScriptDataPoints & val(Replace(Cregr.YValues(i), ",", ".")) & "#"
            If Cregr.XValues(i) > xmax Then xmax = Cregr.XValues(i)
            If Cregr.XValues(i) < xmin Then xmin = Cregr.XValues(i)
        Next
        If Right$(ScriptDataPoints, 1) = "#" Then ScriptDataPoints = Left$(ScriptDataPoints, Len(ScriptDataPoints) - 1)
        ScriptDataPoints = ScriptDataPoints & ";" & xmin & ":" & xmax
    End If
    
    OpenExcelMac "Graphs.xltm", ";" & ScriptFunctions & ";" & ScriptDataPoints

slut2:
    On Error Resume Next
    Unload UFwait2

End Sub
#Else
Sub InsertChart()
Dim WB As Object
Dim ws As Object
Dim xlap As Object 'Excel.Application
Dim xmin As Double, xmax As Double
Dim i As Integer
Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String
Dim arr As Variant
Dim dd As New DocData
Dim ea As New ExpressionAnalyser
Dim srange As Range
On Error GoTo fejl
ea.SetNormalBrackets
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    Set srange = Selection.Range

    dd.ReadSelection

    DoEvents

    Application.ScreenUpdating = False
    Dim UFwait2 As New UserFormWaitForMaxima
    UFwait2.Show vbModeless
    DoEvents
    UFwait2.Label_progress = "***"

If Not ExcelIndlejret Then ' open in Excel
    If cxl Is Nothing Then Set cxl = New CExcel
    cxl.LoadFile ("Graphs.xltm")
    
    UFwait2.Label_progress = UFwait2.Label_progress & "***"
    Set WB = cxl.xlwb
'    Set ws = cxl.xlwb.worksheets(1)
    Set ws = cxl.XLapp.ActiveSheet

    Set xlap = cxl.XLapp
Else ' indlejret
    GoToInsertPoint
    Selection.TypeParagraph
'    Set xlap = New Excel.Application
    Set WB = InsertIndlejret("Graphs.xltm", TT.A(633))
    Set ws = WB.Sheets(1)
    Set xlap = WB.Application
End If
    
    UFwait2.Label_progress = UFwait2.Label_progress & "*****"

' settings
If Radians Then
    ws.Range("A4").Value = "rad"
Else
    ws.Range("A4").Value = "grad"
End If

    ' selected functions
    For i = 0 To dd.AntalMathBoxes - 1
        Udtryk = dd.MathBoxes(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim$(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW$(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW$(8797), "=") ' triple =
        Udtryk = Replace(Udtryk, VBA.ChrW$(8801), "=") ' def =
        Udtryk = Replace(Udtryk, vbCrLf, "") '
        Udtryk = Replace(Udtryk, vbCr, "") '
        Udtryk = Replace(Udtryk, vbLf, "") '
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then
                If InStr(Udtryk, "=") > 0 Then
                    arr = Split(Udtryk, "=")
                    LHS = arr(0)
                    RHS = arr(1)
                    ea.text = LHS
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    If LHS = fktnavn & "(" & varnavn & ")" Then
                        ws.Range("B4").Offset(0, i).Value = RHS
                        ws.Range("B1").Offset(0, i).Value = varnavn
                    Else
                        ws.Range("B4").Offset(0, i).Value = RHS
                        ws.Range("B1").Offset(0, i).Value = "x"
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW$(8804)) > 0 Or InStr(Udtryk, VBA.ChrW$(8805)) > 0 Then
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
                    ws.Range("B4").Offset(0, i).Value = Udtryk
                    ws.Range("B1").Offset(0, i).Value = "x"
               End If
            End If
        End If
    Next

    srange.Select

    'datapoints
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression
        Cregr.GetTableData
        xmin = Cregr.XValues(1)
        xmax = Cregr.XValues(1)
        For i = 1 To UBound(Cregr.XValues)
            ws.Range("Q6").Offset(i, 0).Value = val(Replace(Cregr.XValues(i), ",", "."))
            ws.Range("R6").Offset(i, 0).Value = val(Replace(Cregr.YValues(i), ",", "."))
'            ws.Range("H6").Offset(i, 0).Value = ConvertNumber(Cregr.XValues(i))
'            ws.Range("I6").Offset(i, 0).Value = ConvertNumber(Cregr.YValues(i))
            If Cregr.XValues(i) > xmax Then xmax = Cregr.XValues(i)
            If Cregr.XValues(i) < xmin Then xmin = Cregr.XValues(i)
        Next
        ws.Range("W3").Value = xmin
        ws.Range("X3").Value = xmax
    Else
        If Len(Udtryk) > 0 Then ' if only function, then xmin and xmax must be set
            ws.Range("W3").Value = -2
            ws.Range("X3").Value = 5
        End If
    End If
    
    srange.Select
    Selection.Collapse wdCollapseEnd

GoTo slut:
fejl:
    MsgBox TT.A(98), vbOKOnly, TT.Error
slut:
On Error GoTo slut2
    UFwait2.Label_progress = UFwait2.Label_progress & "**"
    xlap.Run ("Auto_open")
    xlap.Run ("UpDateAll")

    UFwait2.Label_progress = UFwait2.Label_progress & "***"

slut2:
    On Error Resume Next
    Unload UFwait2
    xlap.EnableEvents = True
    xlap.ScreenUpdating = True

End Sub

#End If

Sub InsertGeoGeobraObject()
    Dim ils As InlineShape
#If Mac Then
    MsgBox2 "This function is not supported on Mac", vbOKOnly, "Mac"
#Else
    If InStr(GeoGebraPath, AppNavn) > 0 Then
        UserFormGeoGebra.Show
    Else
        CreateGeoGebraFil GetTempDir()
        Application.ScreenUpdating = False
        If Selection.Range.Tables.Count > 0 Then
            Selection.Tables(Selection.Tables.Count).Select
            Selection.Collapse (wdCollapseEnd)
            Selection.TypeParagraph
        End If
        If Selection.OMaths.Count > 0 Then
            Selection.OMaths(Selection.OMaths.Count).Range.Select
            Selection.Collapse (wdCollapseEnd)
            Selection.TypeParagraph
        End If
        Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="Package", fileName:=GetTempDir() & "geogebra.ggb", LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
        '        ils.OLEFormat.DoVerb (wdOLEVerbOpen)
    End If
#End If
End Sub

Function ReadTextFile(fil As String) As String
Dim filno As Integer
Dim linje, text As String
filno = FreeFile

Open fil For Input As filno
Do While Not EOF(filno) ' Loop until end of file.
  Line Input #filno, linje
   text = text & vbCrLf & linje
Loop
Close filno

ReadTextFile = text
End Function
Sub TestEmbed()
Dim path As String
Dim ils As InlineShape
path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""

Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
End Sub
Function InsertIndlejret(FilNavn As String, Optional startark As String) As Object
' insert excel document as embedded document
' note fails if google cloud connect installed
    Dim path As String
    Dim ils As InlineShape
    Dim vers As String
    On Error GoTo fejl
    Application.ScreenUpdating = False
    
    Dim UFwait2 As New UserFormWaitForMaxima
    UFwait2.Label_stop.visible = False
    UFwait2.Label_tip.Caption = "      " & TT.A(372) & "..."
    UFwait2.Label_progress.Caption = TT.A(373)
    UFwait2.Show vbModeless
    DoEvents
    UFwait2.Label_progress = "***"
#If Mac Then
    path = GetWordMatDir() & "Excelfiles/" & FilNavn
#Else
    path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\" & FilNavn & """"
#End If
    If Selection.Range.Tables.Count > 0 Then
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
    End If
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
    End If
    If val(Application.Version) = 12 Then
        vers = ".12"
    Else
        vers = ""
    End If

    Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="Excel.SheetMacroEnabled" & vers & Application.Version, _
        fileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
        
    UFwait2.Label_progress = "***************************************"

    If startark <> "" Then
        ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
        DoEvents
        Set InsertIndlejret = ils.OLEFormat.Object
    End If
    Unload UFwait2

    GoTo slut
fejl:
    On Error Resume Next
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
    Unload UFwait2
slut:
End Function

Sub InsertPindeDiagram()
    InsertOpenExcel "Pindediagram.xltm"
End Sub
Sub InsertBoksplot()
    InsertOpenExcel "Boksplot.xltm"
End Sub
Sub InsertHistogram()
    InsertOpenExcel "Histogram.xltm"
End Sub
Sub InsertSumkurve()
    InsertOpenExcel "Sumkurve.xltm"
End Sub
Sub InsertUGrupObs()
Dim s As String
If TT.LangNo = 1 Then
    s = "Ugrup"
Else
    s = "Ungroup"
End If
    InsertOpenExcel "statistik.xltm", s
End Sub
Sub InsertGrupObs()
Dim s As String
If TT.LangNo = 1 Then
    s = "Grup"
Else
    s = "Group"
End If
    InsertOpenExcel "statistik.xltm", s
End Sub
Sub InsertTrappediagram()
    InsertOpenExcel "TrappeDiagram.xltm"
End Sub
Function ConvertDrawLabel(text As String) As String
' converts characters to draw2d plot

    text = Replace(text, VBA.ChrW$(916), "{/Symbol D}")
    text = Replace(text, VBA.ChrW$(948), "{/Symbol d}")
    text = Replace(text, VBA.ChrW$(945), "{/Symbol a}")
    text = Replace(text, VBA.ChrW$(946), "{/Symbol b}")
    text = Replace(text, VBA.ChrW$(947), "{/Symbol g}")
    text = Replace(text, VBA.ChrW$(952), "{/Symbol t}") 'theta
    text = Replace(text, VBA.ChrW$(920), "{/Symbol T}")
    text = Replace(text, VBA.ChrW$(955), "{/Symbol l}")
    text = Replace(text, VBA.ChrW$(923), "{/Symbol L}")
    text = Replace(text, VBA.ChrW$(956), "{/Symbol m}")
    text = Replace(text, VBA.ChrW$(961), "{/Symbol r}") ' rho
    text = Replace(text, VBA.ChrW$(963), "{/Symbol s}")
    text = Replace(text, VBA.ChrW$(931), "{/Symbol S}")
    text = Replace(text, VBA.ChrW$(981), "{/Symbol p}") ' phi
    text = Replace(text, VBA.ChrW$(934), "{/Symbol P}")
    text = Replace(text, VBA.ChrW$(949), "{/Symbol v}") 'varepsilon
    text = Replace(text, VBA.ChrW$(1013), "{/Symbol e}") 'epsilon
    text = Replace(text, VBA.ChrW$(968), "{/Symbol p}") 'psi
    text = Replace(text, VBA.ChrW$(936), "{/Symbol P}")
    text = Replace(text, VBA.ChrW$(926), "{/Symbol X}") 'xi
    text = Replace(text, VBA.ChrW$(958), "{/Symbol x}")
    text = Replace(text, VBA.ChrW$(935), "{/Symbol C}") 'chi
    text = Replace(text, VBA.ChrW$(967), "{/Symbol c}")
    text = Replace(text, VBA.ChrW$(928), "{/Symbol Pi}")
    text = Replace(text, VBA.ChrW$(964), "{/Symbol t}") 'tau
    text = Replace(text, VBA.ChrW$(957), "{/Symbol n}") 'greeknu
    text = Replace(text, VBA.ChrW$(954), "{/Symbol k}") 'kappa
    text = Replace(text, VBA.ChrW$(951), "{/Symbol e}") 'eta
    text = Replace(text, VBA.ChrW$(950), "{/Symbol z}") 'zeta

ConvertDrawLabel = text
End Function

