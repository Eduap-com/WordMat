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
        GeoGebraWeb
    End If
    
End Sub
Public Sub Plot2DGraph()
' gnuplot
'    Dim omax As New CMaxima
    Dim forskrifter As String
    Dim Arr As Variant
    Dim i As Integer
    Dim j As Integer
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
    CASengine = 0
    PrepareMaxima
    omax.ReadSelection
'    If UF2Dgraph Is Nothing Then
       Set UF2Dgraph = New UserForm2DGraph
'    End If
    
'    forskrifter = omax.FindDefinitions
'    If Len(forskrifter) > 3 Then
'    forskrifter = Mid(forskrifter, 2, Len(forskrifter) - 3)
'    arr = Split(forskrifter, ",")
'    forskrifter = ""
    
'    For i = 0 To UBound(arr)
'        If InStr(arr(i), "):") > 0 Then
'            forskrifter = forskrifter & omax.ConvertToWordSymbols(arr(i)) & ListSeparator
'        End If
'    Next
'    End If
    
'    If forskrifter <> "" Then
'        forskrifter = Left(forskrifter, Len(forskrifter) - 1)
'    End If
'    forskrifter = omax.KommandoerStreng & ListSeparator & forskrifter
    
    forskrifter = omax.KommandoerStreng
    
    If Len(forskrifter) > 1 Then
    Arr = Split(forskrifter, ListSeparator)
    For i = 0 To UBound(Arr)
        Arr(i) = Trim(Replace(Arr(i), "  ", " ")) ' må ikke fjerne alle mellemrum da f.eks 1/x 3 så bliver 1/x3 hvor x3 er variabel
        If Arr(i) <> "" Then InsertNextEquation (Arr(i))
    Next
    End If
    
    'datapunkter
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, xmin As Double, xmax As Double
        Cregr.GetTableData
        If UF2Dgraph.TextBox_punkter.Text <> "" Then UF2Dgraph.TextBox_punkter.Text = UF2Dgraph.TextBox_punkter.Text & VbCrLfMac
        xmin = Cregr.XValues(1)
        xmax = Cregr.XValues(1)
        For j = 1 To UBound(Cregr.XValues)
'            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
            UF2Dgraph.TextBox_punkter.Text = UF2Dgraph.TextBox_punkter.Text & ConvertNumber(Cregr.XValues(j)) & ListSeparator & ConvertNumber(Cregr.YValues(j)) & VbCrLfMac
            If Cregr.XValues(j) > xmax Then xmax = Cregr.XValues(j)
            If Cregr.XValues(j) < xmin Then xmin = Cregr.XValues(j)
        Next
        UF2Dgraph.TextBox_xmin = xmin
        UF2Dgraph.TextBox_xmax = xmax
    End If
    
'    UserForm2DGraph.Show vbModeless
    Selection.End = sslut ' slut skal være først ellers går det galt
    Selection.start = sstart
    UF2Dgraph.Show vbModeless
    
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    CASengine = TempCas
End Sub
Sub InsertNextEquation(Ligning As String)
Dim Arr As Variant
On Error GoTo fejl
Ligning = Replace(Ligning, VBA.ChrW(8788), "=") ' :=
Ligning = Replace(Ligning, VBA.ChrW(8797), "=") ' tripel =
Ligning = Replace(Ligning, VBA.ChrW(8801), "=") ' def =

Arr = Split(Ligning, "=")

'If Not (InStr(Ligning, VBA.ChrW(9608)) > 0 And InStr(Ligning, VBA.ChrW(9508)) > 0) Then ' tuborg
   Arr = Split(Arr(UBound(Arr)), VBA.ChrW(8776)) ' til indsættelse af selve forskrift i stedet for f(x)
   Ligning = omax.ConvertToAscii(Arr(UBound(Arr)))
'End If
Ligning = omax.ConvertToAscii(Trim(Replace(Replace(Replace(Replace(Arr(0), "Definer:", ""), "Define:", ""), "definer:", ""), "define:", "")))

If UF2Dgraph.TextBox_ligning1.Text = Ligning Then
    Exit Sub
ElseIf UF2Dgraph.TextBox_ligning2.Text = Ligning Then
    Exit Sub
ElseIf UF2Dgraph.TextBox_ligning3.Text = Ligning Then
    Exit Sub
ElseIf UF2Dgraph.TextBox_ligning4.Text = Ligning Then
    Exit Sub
ElseIf UF2Dgraph.TextBox_ligning5.Text = Ligning Then
    Exit Sub
ElseIf UF2Dgraph.TextBox_ligning6.Text = Ligning Then
    Exit Sub
End If

If UF2Dgraph.TextBox_ligning1.Text = "" Then
    UF2Dgraph.TextBox_ligning1.Text = Ligning
ElseIf UF2Dgraph.TextBox_ligning2.Text = "" Then
    UF2Dgraph.TextBox_ligning2.Text = Ligning
ElseIf UF2Dgraph.TextBox_ligning3.Text = "" Then
    UF2Dgraph.TextBox_ligning3.Text = Ligning
ElseIf UF2Dgraph.TextBox_ligning4.Text = "" Then
    UF2Dgraph.TextBox_ligning4.Text = Ligning
ElseIf UF2Dgraph.TextBox_ligning5.Text = "" Then
    UF2Dgraph.TextBox_ligning5.Text = Ligning
ElseIf UF2Dgraph.TextBox_ligning6.Text = "" Then
    UF2Dgraph.TextBox_ligning6.Text = Ligning
End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub PlotDF()
' plot retningsfelt
    Dim s As String, V As String
    Dim Arr As Variant
    Dim ea As New ExpressionAnalyser
  '  On Error GoTo fejl
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
        
    PrepareMaxima
    omax.ReadSelection
#If Mac Then
#Else
   If GraphApp > 0 Then
'    If CASengine > 0 Then ' fra 1.26 er der skiftet til, at det er det valgte grafprogram der anvendes
#End If
        s = Trim(omax.Kommando)
        s = GetCmdAfterEqualSign(s)
        If s = "" Then
            MsgBox2 Sprog.EquationMissingError2, vbOKOnly, "Error"
            Exit Sub
        End If
        ea.Text = s
        V = ea.GetNextVar
        If V <> "x" And V <> "y" Then
            If V = "t" Then
                ea.ReplaceVar "t", "x"
            ElseIf V = "N" Then
                ea.ReplaceVar V, "y"
            Else
                ea.ReplaceVar V, "y"
            End If
        End If
        s = ea.Text
        s = "SlopeField(" & s & ");"
        s = s & "A=(1, 2);Xmin=-100;Xmax=100;Tic=0.1;"
        s = s & "SolveODE(" & ea.Text & ", x(A), y(A), Xmin, Tic);" ' y(A) virker ikke
        s = s & "SolveODE(" & ea.Text & ", x(A), y(A), Xmax, Tic)"
        OpenGeoGebraWeb s, "Classic", True, True
        GoTo slut
#If Mac Then
#Else
    End If
#End If
    Set UF2Dgraph = New UserForm2DGraph
       
'    forskrifter = omax.KommandoerStreng
    
        
    If Len(omax.Kommando) > 0 Then
    Arr = Split(omax.Kommando, "=")
    omax.Kommando = Arr(UBound(Arr))
    End If
    UF2Dgraph.TextBox_dfligning.Text = omax.ConvertToAscii(omax.Kommando)
    
    omax.FindVariable
    If InStr(omax.vars, "x") > 0 Then
        UF2Dgraph.TextBox_dfx.Text = "x"
    ElseIf InStr(omax.vars, "t") > 0 Then
        UF2Dgraph.TextBox_dfx.Text = "t"
    Else
        UF2Dgraph.TextBox_dfx.Text = "x"
    End If
    If InStr(omax.vars, "y") > 0 Then
        UF2Dgraph.TextBox_dfy.Text = "y"
    ElseIf InStr(omax.vars, "N") > 0 Then
        UF2Dgraph.TextBox_dfy.Text = "N"
    Else
        ea.Text = omax.vars
        UF2Dgraph.TextBox_dfy.Text = ea.GetNextVar
        If UF2Dgraph.TextBox_dfy.Text = "" Then UF2Dgraph.TextBox_dfy.Text = "y"
    End If



    Selection.End = sslut ' slut skal være først ellers går det galt
    Selection.start = sstart
    UF2Dgraph.MultiPage1.Value = 5
    UF2Dgraph.MultiPage1.SetFocus
    UF2Dgraph.Show vbModeless

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEmptyGraphOleObject()
' indsætter graph object www.padowan.dk
Dim ils As InlineShape
Application.ScreenUpdating = False

If Not FileExists(GetProgramFilesDir & "\Graph\graph.exe") Then
    Dim result As VbMsgBoxResult
    result = MsgBox(Sprog.A(366), vbOKCancel, Sprog.Error)
    If result = vbOK Then
        OpenLink ("http://www.padowan.dk/graph/Download.php")
    End If
    Exit Sub
End If

'path = """" & GetProgramFilesDir & "\WordMat\graphtemplate.grf"""

' indsæt vha. classname
Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="GraphFile", fileName:="", Range:=Selection.Range, LinkToFile:=False, DisplayAsIcon:=False)

'indsæt vha. tom graphfil. Nok lidt langsommere, men kan på et tidspunkt måske bruges til kommunikation
'Set ils = ActiveDocument.InlineShapes.AddOLEObject(FileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
'ils.OLEFormat.DoVerb (wdOLEVerbShow)

Application.ScreenUpdating = True

End Sub
Sub InsertGraphOleObject()
' indsætter graph object www.padowan.dk
#If Mac Then
    MsgBox "Sorry. Graph is not supported on Mac. There is a beta version you could try though. You will now be forwarded to the download page", vbOKOnly, Sprog.Error
    OpenLink "http://www.padowan.dk/mac/"
#Else
Dim Path As String
Dim ils As InlineShape
Dim Arr As Variant
Dim fktnavn As String, Udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String
Dim ea As New ExpressionAnalyser
Dim p As Integer
    Dim sslut As Long
    sslut = Selection.End

ea.SetNormalBrackets
    Dim UfWait As New UserFormWaitForMaxima
    UfWait.Label_tip.Caption = Sprog.A(371)
    UfWait.Label_progress.Caption = "***"
    UfWait.Label_stop.visible = False
    UfWait.Show vbModeless
On Error GoTo fejl
Application.ScreenUpdating = False

If Not FileExists(GetProgramFilesDir & "\Graph\graph.exe") Then
    Dim result As VbMsgBoxResult
    result = MsgBox(Sprog.A(366), vbOKCancel, Sprog.Error)
    If result = vbOK Then
        OpenLink ("http://www.padowan.dk/graph/Download.php")
    End If
    Exit Sub
End If

'path = """" & GetProgramFilesDir & "\WordMat\graphtemplate.grf"""
Path = Environ("TEMP") & "\" & "wordmatgraph.grf"
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
        If InStr(omax.DefValue(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
            If Not (InStr(deflist2, omax.DefName(i)) > 0) Then ' hvis ikke allerede defineret
               deflist2 = deflist2 & "," & omax.DefName(i)
               graphfil.AddCustomFunction omax.DefName(i) & "=" & omax.DefValue(i)
                p = InStr(omax.DefName(i), "(")
                If p > 0 Then
                    graphfil.InsertFunction Left(omax.DefName(i), p - 1) & "(x)", 0
                Else
                    graphfil.InsertFunction omax.DefName(i), 0
                End If
                DefinerKonstanterGraph omax.DefValue(i), DefList, graphfil
            End If
        End If
    Next
    
    ' funktioner der markeres
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
        Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                If InStr(Udtryk, "=") > 0 Then
                    Arr = Split(Udtryk, "=")
                    lhs = Arr(0)
                    rhs = Arr(1)
                    ea.Text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.Text = rhs
                        ea.Pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.Text
                        DefinerKonstanterGraph fktudtryk, DefList, graphfil
                        graphfil.InsertFunction fktudtryk
                    Else
                        DefinerKonstanterGraph Udtryk, DefList, graphfil, True
                        graphfil.InsertRelation Udtryk
                        ' blev brugt før relation
'                        fktudtryk = ReplaceIndepvarX(rhs)
'                        DefinerKonstanterGraph fktudtryk, deflist, graphfil
'                        graphfil.InsertFunction fktudtryk
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                    DefinerKonstanterGraph Udtryk, DefList, graphfil, True
                    graphfil.InsertRelation Udtryk
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
                    DefinerKonstanterGraph Udtryk, DefList, graphfil
                    graphfil.InsertFunction Udtryk
               End If
            End If
        End If
    Next
    
    'datapunkter
    If Selection.Tables.Count > 0 Then
        Dim Cregr As New CRegression, setdata As String, j As Integer
        Cregr.GetTableData
        For j = 1 To UBound(Cregr.XValues)
'            UF2Dgraph.TextBox_punkter.text = UF2Dgraph.TextBox_punkter.text & CStr(Cregr.XValues(j)) & ListSeparator & CStr(Cregr.YValues(j)) & vbCrLf
'            setdata = setdata & ConvertNumber(Cregr.XValues(j)) & "," & ConvertNumber(Cregr.YValues(j)) & ";" '
            setdata = setdata & Replace(Cregr.XValues(j), ",", ".") & "," & Replace(Cregr.YValues(j), ",", ".") & ";"
        Next
        If Len(setdata) > 0 Then
            setdata = Left(setdata, Len(setdata) - 1)
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
        graphfil.Save Path

        'indsæt vha. tom graphfil. Nok lidt langsommere, men kan på et tidspunkt måske bruges til kommunikation
        Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=Path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
        ils.OLEFormat.DoVerb (wdOLEVerbShow)

    Else
        ' indsæt vha. classname
        Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="GraphFile", fileName:="", Range:=Selection.Range, LinkToFile:=False, DisplayAsIcon:=False)
    End If

    DoEvents
    Unload UfWait

Application.ScreenUpdating = True
GoTo slut
fejl:
    MsgBox Sprog.A(97), vbOKOnly, Sprog.Error
    omax.ConvertLnLog = True
    Unload UfWait
slut:
    omax.ConvertLnLog = True
#End If
End Sub
#If Mac Then
#Else
Sub DefinerKonstanterGraph(Expr As String, DefList As String, ByRef graphfil As CGraphFile, Optional noty As Boolean = False)
' definer variable der ikke er defineret i expr
' deflist er en liste af variable der er defineret
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim Var As String
    ea.Text = DefList
    If noty Then ea.Text = ea.Text & ",y"
    ea2.Text = Expr
    ea2.Pos = 0
    Do
        Var = ea2.GetNextVar
        ea2.Pos = ea2.Pos + 1
        If Not (ea2.ChrByIndex(ea2.Pos) = "(") And Not (ea.IsFunction(Var)) And Not (ea.ContainsVar(Var)) And Var <> "" And Var <> "x" And Var <> "y" And Var <> "e" And Var <> "pi" And Var <> "matrix" Then ' måske ikke y? kopieret fra geogebra
            graphfil.AddCustomFunction Var & "=" & InputBox(Sprog.A(363) & " " & Var & vbCrLf & vbCrLf & Sprog.A(367), Sprog.A(365), "1")
            DefList = DefList & "," & Var
        End If
    Loop While Var <> ""

End Sub
#End If
Function ReplaceIndepvarX(fkt As String, Optional ByRef uvar = "") As String
' sørger for at indsætte x som uafh variabel
' hvis den ikke er i udtrykket spørges
Dim ea As New ExpressionAnalyser
Dim Var As String
'Dim uvar As String
ea.Text = fkt
Var = ea.GetNextVar
ReplacedVar = "x"
If Not (ea.ContainsVar("x")) And Var <> "" And Var <> "matrix" Then
    If ea.ContainsVar("t") Then
        uvar = "t"
    Else
        uvar = Var
    End If
    uvar = InputBox(Sprog.A(368) & vbCrLf & vbCrLf & "   " & fkt & vbCrLf & vbCrLf, Sprog.A(369), uvar)
    If uvar = "" Then uvar = "x"
    If uvar <> "x" Then
        ea.ReplaceVar uvar, "x"
    End If
    ReplacedVar = uvar
End If

ReplaceIndepvarX = ea.Text
End Function
Sub InsertChart()
Dim WB As Object
Dim ws As Object
'Dim xlap As Excel.Application
Dim xlap As Object 'Excel.Application
Dim xmin As Double, xmax As Double
Dim i As Integer
Dim fktnavn As String, Udtryk As String, lhs As String, rhs As String, varnavn As String
Dim Arr As Variant
Dim dd As New DocData
Dim ea As New ExpressionAnalyser
Dim srange As Range
On Error GoTo fejl
ea.SetNormalBrackets
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    Set srange = Selection.Range

Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone

dd.ReadSelection

'cxl.PrePareExcel
    DoEvents

Application.ScreenUpdating = False
    Dim UfWait2 As New UserFormWaitForMaxima
    UfWait2.Show vbModeless
    DoEvents
    UfWait2.Label_progress = "***"

If Not ExcelIndlejret Then ' åben i excel
    If cxl Is Nothing Then Set cxl = New CExcel
    cxl.LoadFile ("Graphs.xltm")
    UfWait2.Label_progress = UfWait2.Label_progress & "***"
    Set WB = cxl.xlwb
'    Set ws = cxl.xlwb.worksheets(1)
    Set ws = cxl.XLapp.ActiveSheet

    Set xlap = cxl.XLapp
Else ' indlejret
    GoToInsertPoint
    Selection.TypeParagraph
'    Set xlap = New Excel.Application
    Set WB = InsertIndlejret("Graphs.xltm", Sprog.A(633)) ' "tabel"
    Set ws = WB.Sheets(1)
    Set xlap = WB.Application
End If
    
    UfWait2.Label_progress = UfWait2.Label_progress & "*****"

' indstillinger
If Radians Then
    ws.Range("A4").Value = "rad"
Else
    ws.Range("A4").Value = "grad"
End If

    ' funktioner der markeres
    For i = 0 To dd.AntalMathBoxes - 1
        Udtryk = dd.MathBoxes(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
        Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
        Udtryk = Replace(Udtryk, vbCrLf, "") '
        Udtryk = Replace(Udtryk, vbCr, "") '
        Udtryk = Replace(Udtryk, vbLf, "") '
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                If InStr(Udtryk, "=") > 0 Then
                    Arr = Split(Udtryk, "=")
                    lhs = Arr(0)
                    rhs = Arr(1)
                    ea.Text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
'                    If varnavn = "" And fktnavn = Y Then varnavn = X
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ws.Range("B4").Offset(0, i).Value = rhs
                        ws.Range("B1").Offset(0, i).Value = varnavn
                    Else
'                        DefinerKonstanterGraph udtryk, deflist, graphfil, True
                        ws.Range("B4").Offset(0, i).Value = rhs
                        ws.Range("B1").Offset(0, i).Value = "x"
                        ' blev brugt før relation
'                        fktudtryk = ReplaceIndepvarX(rhs)
'                        DefinerKonstanterGraph fktudtryk, deflist, graphfil
'                        graphfil.InsertFunction fktudtryk
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
'                    DefinerKonstanterGraph udtryk, deflist, graphfil
                    ws.Range("b4").Offset(0, i).Value = Udtryk
                    ws.Range("B1").Offset(0, i).Value = "x"
               End If
            End If
        End If
    Next

'    Selection.start = sstart
'    Selection.End = sslut
    srange.Select

    'datapunkter
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
    End If

' virker kun med lodret tabel
'    If dd.nrows > 1 And dd.ncolumns > 1 Then
'        For i = 1 To dd.nrows
'            ws.Range("H6").Offset(i, 0).Value = dd.TabelCelle(i, 1)
'            ws.Range("I6").Offset(i, 0).Value = dd.TabelCelle(i, 2)
'        Next
'    End If

'    Selection.start = sslut
'    Selection.End = sslut
    srange.Select
    Selection.Collapse wdCollapseEnd

GoTo slut:
fejl:
    MsgBox Sprog.A(98), vbOKOnly, Sprog.Error
slut:
On Error GoTo slut2
    UfWait2.Label_progress = UfWait2.Label_progress & "**"
    xlap.Run ("Auto_open")
xlap.Run ("UpDateAll")
'excel.Run ("UpDateAll")
UfWait2.Label_progress = UfWait2.Label_progress & "***"

'If Not wb Is Nothing Then ' start på tabel sheet, ikke graph
'    wb.Charts(1).Activate
'End If

slut2:
    On Error Resume Next
'    excel.Application.EnableEvents = True
'    excel.Application.ScreenUpdating = True
'    excel.Application.DisplayAlerts = True
    Unload UfWait2
    xlap.EnableEvents = True
    xlap.ScreenUpdating = True

'Excel.Application.ScreenUpdating = True

'excel.ActiveSheet.ChartObjects(1).Copy
'Selection.Collapse Direction:=wdCollapseStart
'Selection.Paste
'Selection.PasteSpecial DataType:=wdPasteOLEObject
'Selection.PasteSpecial DataType:=wdPasteShape
End Sub
Sub InsertChartG()
'indsætter exceldokument som indlejret dokument
'Dim wb As Workbook
'Dim ws As Worksheet
Dim WB As Object 'Workbook
Dim ws As Object
Dim fktnavn As String, Udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String
Dim dd As New DocData
Dim i As Integer
Dim Arr As Variant
Dim ea As New ExpressionAnalyser
Dim Path As String
Dim ils As InlineShape
Application.ScreenUpdating = False
ea.SetNormalBrackets

PrepareMaxima
omax.ReadSelection
'    dd.ReadSelection

    MsgBox Sprog.Wait

If ExcelIndlejret Then
    GoToInsertPoint
    Selection.TypeParagraph
Path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""

EnableExcelMacros

Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=Path, LinkToFile:=False, _
DisplayAsIcon:=False, Range:=Selection.Range)

ils.OLEFormat.DoVerb (wdOLEVerbShow)
Set WB = ils.OLEFormat.Object
DisableExcelMacros



Else
    Set WB = InsertOpenExcel("Graphs.xltm", "Tabel")
End If
Set ws = WB.Sheets("Tabel")
'Excel.Application.EnableEvents = False
'Excel.Application.ScreenUpdating = False
XLapp.Application.EnableEvents = False
XLapp.Application.ScreenUpdating = False

GoTo hop

If Radians Then
    ws.Range("A4").Value = "rad"
Else
    ws.Range("A4").Value = "grad"
End If


    ' funktioner der markeres
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Trim(Udtryk)
        Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
        Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                If InStr(Udtryk, "=") > 0 Then
                    Arr = Split(Udtryk, "=")
                    lhs = Arr(0)
                    rhs = Arr(1)
                    ea.Text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.Text = rhs
                        ea.Pos = 1
'                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.Text
'                        DefinerKonstanterGraph fktudtryk, deflist, graphfil
                        ws.Range("b4").Offset(0, i).Value = fktudtryk
                        ws.Range("B1").Offset(0, i).Value = varnavn
                    Else
'                        DefinerKonstanterGraph udtryk, deflist, graphfil, True
                        ws.Range("b4").Offset(0, i).Value = Udtryk
                        ws.Range("B1").Offset(0, i).Value = "x"
                        ' blev brugt før relation
'                        fktudtryk = ReplaceIndepvarX(rhs)
'                        DefinerKonstanterGraph fktudtryk, deflist, graphfil
'                        graphfil.InsertFunction fktudtryk
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
'                    DefinerKonstanterGraph udtryk, deflist, graphfil
                    ws.Range("b4").Offset(0, i).Value = Udtryk
                    ws.Range("B1").Offset(0, i).Value = "x"
               End If
            End If
        End If
    Next
    
    If dd.nrows > 1 And dd.ncolumns > 1 Then
        For i = 1 To dd.nrows
            ws.Range("H6").Offset(i, 0).Value = dd.TabelCelle(i, 1)
            ws.Range("I6").Offset(i, 0).Value = dd.TabelCelle(i, 2)
        Next
    End If


hop:
' Opdater Excel med ændringer
On Error Resume Next
'wb.Application.Run ("UpDateAll")
WB.Charts(1).Activate
'Excel.Application.EnableEvents = True
'Excel.Application.ScreenUpdating = True
'Excel.Run ("UpDateAll")
XLapp.Application.EnableEvents = True
XLapp.Application.ScreenUpdating = True
XLapp.Run ("UpDateAll")

Exit Sub

'On Error GoTo slut
'Dim path As String
'Dim ils As InlineShape
'Dim wb As Variant
'Dim ws As Variant
EnableExcelMacros
Application.ScreenUpdating = False
Path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""

'Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
Set ils = ActiveDocument.InlineShapes.AddOLEObject( _
fileName:=Path, LinkToFile:=False, _
DisplayAsIcon:=False, Range:=Selection.Range)
'Ils.Height = 300
'Ils.Width = 500

'ils.OLEFormat.DoVerb (wdOLEVerbShow)
ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)

Set WB = ils.OLEFormat.Object
Set ws = WB.Sheets("Tabel")
ws.Activate


'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
'Ils.OLEFormat.DoVerb (wdOLEVerbShow)
'Ils.OLEFormat.DoVerb (wdOLEVerbUIActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbHide)

'DoEvents
'Application.ScreenUpdating = True
'Dim wb As excel.Workbook
'Dim excel As excel.Application
'Set excel = CreateObject("Excel.Application")
'Set wb = excel.Workbooks(excel.Workbooks.Count)
slut:
DisableExcelMacros
End Sub

Sub InsertGeoGeobraObject()
Dim ils As InlineShape
#If Mac Then
    MsgBox "This function is not supported on Mac", vbOKOnly, "Mac"
#Else
    If InStr(GeoGebraPath, AppNavn) > 0 Then
        MsgBox "Denne funktion kræver at GeoGebra installeres separat", vbOKOnly, Sprog.Error
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
Dim linje, Text As String
filno = FreeFile

Open fil For Input As filno
Do While Not EOF(filno) ' Loop until end of file.
  Line Input #filno, linje
   Text = Text & vbCrLf & linje
Loop
Close filno

ReadTextFile = Text
End Function
Sub TestEmbed()
Dim Path As String
Dim ils As InlineShape
Path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""

Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=Path, LinkToFile:=False, _
DisplayAsIcon:=False, Range:=Selection.Range)

End Sub
Function InsertIndlejret(FilNavn As String, Optional startark As String) As Object
'indsætter exceldokument som indlejret dokument
' bemærk fejler hvis google cloud connect installeret
Dim Path As String
Dim ils As InlineShape
Dim vers As String
On Error GoTo fejl
Application.ScreenUpdating = False
EnableExcelMacros
    
    Dim UfWait2 As New UserFormWaitForMaxima
    UfWait2.Label_stop.visible = False
    UfWait2.Label_tip.Caption = "      " & Sprog.A(372) & "..."
    UfWait2.Label_progress.Caption = Sprog.A(373) ' "Indsætning af indlejrede objekter kan tage tid. Dobbeltklik på objektet for at redigere det."
    UfWait2.Show vbModeless
        DoEvents
        UfWait2.Label_progress = "***"
#If Mac Then
Path = GetWordMatDir() & "Excelfiles/" & FilNavn
#Else
Path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\" & FilNavn & """"
#End If
    If Selection.Range.Tables.Count > 0 Then
'        Selection.Copy
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
'        startark = "Data"
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

'Set ils = ActiveDocument.InlineShapes.AddOLEObject(FileName:=path, LinkToFile:=False, _
'DisplayAsIcon:=False, Range:=Selection.Range)
' forsøg uden classtype. prøv evt med classtype:=Excel.SheetMacroEnabled.12
'Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="Excel.Sheet" & vers & Application.Version, _
'FileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="Excel.SheetMacroEnabled" & vers & Application.Version, _
fileName:=Path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
        
        UfWait2.Label_progress = "***************************************"
'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
'ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
'ils.OLEFormat.DoVerb (wdOLEVerbShow)
If startark <> "" Then
    ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
    DoEvents
    Set InsertIndlejret = ils.OLEFormat.Object

'    Dim oWS As Object ' Worksheet Object
'    Set oWS = ils.OLEFormat.Object
'    ils.OLEFormat.Object.Sheets(startark).Activate
'    oWS.ActiveSheet.Cells(4, 1).Activate
'    oWS.ActiveSheet.Application.Selection.Paste ' virker ikke
'    oWS.Selection.Paste
End If
Unload UfWait2
'Ils.OLEFormat.DoVerb (wdOLEVerbUIActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbHide)
DisableExcelMacros
GoTo slut
fejl:
    On Error Resume Next
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    Unload UfWait2
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
'    InsertOpenExcel "UGrupperedeObservationer.xltm"
Dim s As String
If Sprog.SprogNr = 1 Then
    s = "Ugrup"
Else
    s = "Ungroup"
End If
    InsertOpenExcel "statistik.xltm", s
End Sub
Sub InsertGrupObs()
'    InsertOpenExcel "GrupperedeObservationer.xltm"
Dim s As String
If Sprog.SprogNr = 1 Then
    s = "Grup"
Else
    s = "Group"
End If
    InsertOpenExcel "statistik.xltm", s
End Sub
Sub InsertTrappediagram()
    InsertOpenExcel "TrappeDiagram.xltm"
End Sub
Function ConvertDrawLabel(Text As String) As String
' konverterer tegn til draw2d plot
'text = Replace(text, "", "")
'text = Replace(text, "", "")
'text = Replace(text, "", "")

    Text = Replace(Text, VBA.ChrW(916), "{/Symbol D}")
    Text = Replace(Text, VBA.ChrW(948), "{/Symbol d}")
    Text = Replace(Text, VBA.ChrW(945), "{/Symbol a}")
    Text = Replace(Text, VBA.ChrW(946), "{/Symbol b}")
    Text = Replace(Text, VBA.ChrW(947), "{/Symbol g}")
    Text = Replace(Text, VBA.ChrW(952), "{/Symbol t}") 'theta
    Text = Replace(Text, VBA.ChrW(920), "{/Symbol T}")
    Text = Replace(Text, VBA.ChrW(955), "{/Symbol l}")
    Text = Replace(Text, VBA.ChrW(923), "{/Symbol L}")
    Text = Replace(Text, VBA.ChrW(956), "{/Symbol m}")
    Text = Replace(Text, VBA.ChrW(961), "{/Symbol r}") ' rho
    Text = Replace(Text, VBA.ChrW(963), "{/Symbol s}")
    Text = Replace(Text, VBA.ChrW(931), "{/Symbol S}")
    Text = Replace(Text, VBA.ChrW(981), "{/Symbol p}") ' phi
    Text = Replace(Text, VBA.ChrW(934), "{/Symbol P}")
    Text = Replace(Text, VBA.ChrW(949), "{/Symbol v}") 'varepsilon
    Text = Replace(Text, VBA.ChrW(1013), "{/Symbol e}") 'epsilon
    Text = Replace(Text, VBA.ChrW(968), "{/Symbol p}") 'psi
    Text = Replace(Text, VBA.ChrW(936), "{/Symbol P}")
    Text = Replace(Text, VBA.ChrW(926), "{/Symbol X}") 'xi
    Text = Replace(Text, VBA.ChrW(958), "{/Symbol x}")
    Text = Replace(Text, VBA.ChrW(935), "{/Symbol C}") 'chi
    Text = Replace(Text, VBA.ChrW(967), "{/Symbol c}")
    Text = Replace(Text, VBA.ChrW(928), "{/Symbol Pi}")
    Text = Replace(Text, VBA.ChrW(964), "{/Symbol t}") 'tau
    Text = Replace(Text, VBA.ChrW(957), "{/Symbol n}") 'greeknu
    Text = Replace(Text, VBA.ChrW(954), "{/Symbol k}") 'kappa
    Text = Replace(Text, VBA.ChrW(951), "{/Symbol e}") 'eta
    Text = Replace(Text, VBA.ChrW(950), "{/Symbol z}") 'zeta


ConvertDrawLabel = Text
End Function




