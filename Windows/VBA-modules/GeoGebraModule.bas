Attribute VB_Name = "GeoGebraModule"
Option Explicit
' duer ikke med 64bit
'#If VBA7 Then
'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
'#Else
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#End If


Sub GeoGebra()
    On Error GoTo fejl
    Dim geogebrasti As String
    Dim geogebrafilersti As String
    Dim appnr As Integer
    Dim ufwait As New UserFormWaitForMaxima
    PrepareMaxima ' omax bliver brugt
    
    ufwait.Label_tip.Caption = Sprog.A(362)
    ufwait.Label_progress.Caption = "***"
    ufwait.CommandButton_stop.visible = False
    ufwait.Show vbModeless
    
    geogebrasti = GeoGebraPath()
    If geogebrasti = "" Then GoTo fejl ' hvis geogebra ikke installeret så fejl
'    geogebrafilersti = GetProgramFilesDir & "\WordMat\GeoGebraFiler\"
    geogebrafilersti = GetTempDir()
    
    CreateGeoGebraFil geogebrafilersti
    
    geogebrafilersti = geogebrafilersti & "geogebra.ggb"
    
    ufwait.Label_progress.Caption = "******"
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
    
    ufwait.Label_progress.Caption = "*********"
    On Error Resume Next
    DoEvents
    Unload ufwait
    
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
    Dim defs As String
    Dim arr As Variant
    Dim fktnavn As String, udtryk As String, lhs As String, RHS As String, varnavn As String, fktudtryk As String
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
    Dim var As String, deflist As String
    Dim k As Integer
    
    ' sæt definitioner i rigtig rækkefølge
    For i = 0 To omax.defindex - 1
        ea.text = omax.DefValue(i)
        deflist = deflist & "," & omax.DefName(i)
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
    ea.text = deflist
    For i = 0 To sl.Length - 1
        If InStr(sl.GetVal(i), "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
            fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
            If sl.GetVal(i) <> ReplacedVar Then
                DefinerKonstanter sl.GetVal(i), deflist, geogebrafil
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
                    RHS = arr(1)
                    ea.text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.text = RHS
                        ea.pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.text
                        DefinerKonstanter fktudtryk, deflist, geogebrafil
                        geogebrafil.CreateFunction fktnavn, fktudtryk, False, True
                    Else
                        fktudtryk = ReplaceIndepvarX(RHS)
                        DefinerKonstanter udtryk, deflist, geogebrafil
                        geogebrafil.CreateFunction "f" & j, udtryk, False
                        j = j + 1
                    End If
                ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                ' kan først bruges med GeoGebra 4.0
                    DefinerKonstanter udtryk, deflist, geogebrafil
                    geogebrafil.CreateFunction "u" & j, udtryk, True
                Else
                    udtryk = ReplaceIndepvarX(udtryk)
                    DefinerKonstanter udtryk, deflist, geogebrafil
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
Sub DefinerKonstanter(Expr As String, ByRef deflist As String, ByRef geogebrafil As CGeoGebraFile)
' definer variable der ikke er defineret i expr
' deflist er en liste af variable der er defineret
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim var As String
Dim varval As String
    ea.text = deflist
    ea2.text = Expr
    ea2.pos = 0
    Do
        var = ea2.GetNextVar
'        MsgBox AscW(var) & vbCrLf & VBA.ChrW(960)
        ea2.pos = ea2.pos + 1
'        If var = "z" Then geogebrafil.Show3D = True  ' ikke implementeret endnu
        If Not (ea2.ChrByIndex(ea2.pos) = "(") And Not (ea.IsFunction(var)) And Not (ea.ContainsVar(var)) And var <> "" And var <> "x" And var <> "y" And var <> "z" And var <> "e" And var <> "pi" And var <> "matrix" And var <> "if" And var <> "elseif" And var <> "then" And var <> "and" And var <> "or" And var <> "else" And var <> VBA.ChrW(960) Then  ' 960=pi
            varval = InputBox(Sprog.A(363) & " " & var & vbCrLf & vbCrLf & Sprog.A(364), Sprog.A(365), "1")
            If Len(varval) > 0 Then
                geogebrafil.CreateFunction var, varval, False, False
                deflist = deflist & "," & var
            End If
        End If
    Loop While var <> ""

End Sub
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



