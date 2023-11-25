Attribute VB_Name = "OpretMathMenu"
Option Explicit
'Sub IndsaetFraScanner()
'    Application.Run MacroName:="IndsætImagerScan"
'End Sub

Sub FjernRetteMenu()
    On Error Resume Next
    Application.CommandBars("MathMenu").Delete
End Sub
Sub SkjulMathMenu()
    Application.CommandBars("MathMenu").visible = False
End Sub
Sub OpretMathMenu()
    Dim myCB As CommandBar
    Dim CBB As CommandBarButton
    Dim CBP As CommandBarPopup
    Dim myCPup1 As CommandBarPopup
    Dim myCPup2 As CommandBarPopup
    Dim myCPup3 As CommandBarPopup
    Dim myCP1Btn1 As CommandBarButton

    On Error Resume Next

    ' Delete the commandbar if it exists already
    Application.CommandBars("MathMenu").Delete
   
    ' Create a new Command Bar 'The ampersand (&) in the name of the menu underlines the letter that follows it to give it a keyboard command (Alt-m) as many menus have.
    Set myCB = CommandBars.Add(Name:="MathMenu", position:=msoBarFloating)
    CommandBars("Menu Bar").Controls("MathMenu").Caption = "MathMenu"
   
    'Indsæt formelsamling menu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Formelsamling"
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Procentregning"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "S=B(1+r)"
     .DescriptionText = ""
     .Tag = "S=B" & VBA.ChrW(183) & "(1+r)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Kapitalfremskrivningsformel: Kn=Ko" & VBA.ChrW(183) & "(1+r)" & VBA.ChrW(&H207F)
     .Tag = "K_n=K_0" & VBA.ChrW(183) & "(1+r)^n"
     .DescriptionText = "Kapitalfremskrivningsformel"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Funktioner"
    Set myCPup3 = myCPup2.Controls.Add(Type:=msoControlPopup)
    myCPup3.Caption = "Lineær"
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "ligning: y=a" & VBA.ChrW(183) & "x+b"
     .DescriptionText = ""
     .Tag = "y=a" & VBA.ChrW(183) & "x+b"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "hældningskoefficient a=(y" & VBA.ChrW(&H2082) & "-y" & VBA.ChrW(&H2081) & ")/(x" & VBA.ChrW(&H2082) & "-x" & VBA.ChrW(&H2081) & ")"
     .DescriptionText = ""
     .Tag = "a=(y_2-y_1)/(x_2-x_1)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Ligning ud fra punkt (x" & VBA.ChrW(&H2081) & ",y" & VBA.ChrW(&H2081) & ") og a  y=a(x-x" & VBA.ChrW(&H2081) & ")+y" & VBA.ChrW(&H2081)
     .DescriptionText = "Ligning til Bestemmelse af ligning for ret linje ud fra kendt punkt (x1,y1) og hældningskoefficient a."
     .Tag = "y=a" & VBA.ChrW(183) & "(x-x_1)+y_1"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    Set myCPup3 = myCPup2.Controls.Add(Type:=msoControlPopup)
    myCPup3.Caption = "Ekponentiel"
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Ligning: y=b" & VBA.ChrW(183) & "a^x"
     .DescriptionText = ""
     .Tag = "y=b" & VBA.ChrW(183) & "a^x"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Ligning: y=b" & VBA.ChrW(183) & "e^kx"
     .DescriptionText = ""
     .Tag = "y=b" & VBA.ChrW(183) & "e^kx"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Bestemmelse af a=(x" & VBA.ChrW(&H2082) & "-x" & VBA.ChrW(&H2081) & ")" & VBA.ChrW(&H221A) & "(y" & VBA.ChrW(&H2082) & "/y" & VBA.ChrW(&H2081) & ")"
     .DescriptionText = ""
     .Tag = "a=" & VBA.ChrW(&H221A) & "(x_2-x_1&y_2/y_1)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Fordoblingskonstant: T" & VBA.ChrW(&H2082) & "=log(2)/log(a)"
     .DescriptionText = "Fordoblingskonstant"
     .Tag = "T_2=log" & VBA.ChrW(8289) & "(2)/log" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(2)/ln" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(2)/k"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Halveringskonstant: T" & VBA.ChrW(&H2081) & "," & VBA.ChrW(&H2082) & "=log" & VBA.ChrW(&HBD) & "/loga"
     .DescriptionText = "Halveringskonstant"
     .Tag = "T_(1/2)=log" & VBA.ChrW(8289) & "(1/2)/log" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(1/2)/ln" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(2)/k"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    Set myCPup3 = myCPup2.Controls.Add(Type:=msoControlPopup)
    myCPup3.Caption = "Potens"
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Ligning: y=b" & VBA.ChrW(183) & "x^a"
     .DescriptionText = ""
     .Tag = "y=b" & VBA.ChrW(183) & "x^a"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "a=log" & VBA.ChrW(8289) & "(y" & VBA.ChrW(&H2082) & "/y" & VBA.ChrW(&H2081) & ")/log" & VBA.ChrW(8289) & "(x" & VBA.ChrW(&H2082) & "/x" & VBA.ChrW(&H2081) & ")"
     .DescriptionText = ""
     .Tag = "a=log" & VBA.ChrW(8289) & "(y_2/y_1)/log" & VBA.ChrW(8289) & "(x_2/x_1)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "1+ry=(1+rx)" & VBA.ChrW(&H207F)
     .DescriptionText = ""
     .Tag = "1+r_y=(1+r_x)^a"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    Set myCPup3 = myCPup2.Controls.Add(Type:=msoControlPopup)
    myCPup3.Caption = "Polynomier"
    With myCPup3.Controls.Add(Type:=msoControlButton)
     .Caption = "Toppunkt af 2. grads polynomium  x=-b/2a"
     .DescriptionText = ""
     .Tag = "x_t=-b/(2a)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Geometri"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Sinusrelation " & "sin(A)/a=sin(B)/b"
     .DescriptionText = ""
     .Tag = "sin" & VBA.ChrW(8289) & "(A)/a=sin" & VBA.ChrW(8289) & "(B)/b"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Cosinusrelation " & "c" & VBA.ChrW(&HB2) & "=a" & VBA.ChrW(&HB2) & "+b" & VBA.ChrW(&HB2) & "-2a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "cos(C)"
     .DescriptionText = ""
     .Tag = "c^2=a^2+b^2-2a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "cos" & VBA.ChrW(8289) & "(C)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Areal af trekant " & "T=" & VBA.ChrW(&HBD) & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "sin(C)"
     .DescriptionText = ""
     .Tag = "T=1/2" & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "sin" & VBA.ChrW(8289) & "(C)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
   
    With myCB.Controls.Add(Type:=msoControlButton)
     .Caption = "|"
     .Style = msoButtonCaption
    End With
'    Set CBP = myCB.Controls.Add(Type:=msoControlButton)
'    CBP.Caption = "|"
'    CBP.Tag = ""
'    CBP.Visible = True
   
    'MaximaMenu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Maxima"
    myCPup1.Tag = ""
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Beregn"
     .Style = msoButtonCaption
     .OnAction = "beregn"
     .ShortcutText = "Alt + b"
    End With
    
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Omskriv"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Reducer"
     .Style = msoButtonCaption
     .OnAction = "reducer"
     .ShortcutText = "Alt + r"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Faktoriser (sæt udenfor parantes)"
     .Style = msoButtonCaption
     .OnAction = "Faktoriser"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Udvid (Gang ind i paranteser)"
     .Style = msoButtonCaption
     .OnAction = "Udvid"
    End With
    
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Ligninger"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Løs ligning(er)"
     .Style = msoButtonCaption
     .ShortcutText = "Alt + L"
     .OnAction = "MaximaSolve"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Løs ligning(er) numerisk"
     .Style = msoButtonCaption
     .OnAction = "MaximaSolveNumeric"
    End With
    
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Infinitesimalregning"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Differentier"
     .Style = msoButtonCaption
     .OnAction = "Differentier"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Integrer (Find stamfunktioner)"
     .Style = msoButtonIconAndCaption
     .FaceId = 477 ' integrale
     .OnAction = "Integrer"
    End With
    
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "----------------------"
     .Style = msoButtonCaption
    End With
        
    Set myCPup2 = myCPup1.Controls.Add(Type:=msoControlPopup)
    myCPup2.Caption = "Definitioner"
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Vis definitioner"
     .Style = msoButtonCaption
'     .FaceId = 385 ' funktion
     .OnAction = "VisDef"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Definer Variabel"
     .Style = msoButtonCaption
     .OnAction = "DefinerVar"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Definer Funktion"
     .Style = msoButtonIconAndCaption
     .FaceId = 385 ' funktion
     .OnAction = "DefinerFunktion"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Definer Ligning"
     .Style = msoButtonCaption
     .OnAction = "DefinerLigning"
    End With
    With myCPup2.Controls.Add(Type:=msoControlButton)
     .Caption = "Nulstil definitioner"
     .Style = msoButtonCaption
     .Tag = "slet definitioner:"
     .OnAction = "indsaetformel"
    End With
    
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Udfør Maxima Kommando"
     .Style = msoButtonCaption
     .OnAction = "MaximaCommand"
    End With

    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Maxima Indstillinger"
     .Style = msoButtonCaption
     .Tag = "false;false;both;false;7;false;prik;false;false;false;false;auto" ' forklaringer;maximakommando;exact;radianer;cifre;separator;gangetegn;kompleks;Løsningsmængde;enheder;vidnotation;logoutput
     .OnAction = "MaximaSettings"
     .ShortcutText = "Alt + i"
    End With
    
    
        'GrafMenu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Grafer"
    myCPup1.Tag = "Grafer"
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Plot ligninger og punkter i planen"
     .Style = msoButtonIconAndCaption
     .FaceId = 422 ' graf
     .OnAction = "Plot2DGraph"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Lineær regression"
     .Style = msoButtonCaption
     .OnAction = "linregression"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Eksponentiel regression"
     .Style = msoButtonCaption
     .OnAction = "ekspregression"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Potens regression"
     .Style = msoButtonCaption
     .OnAction = "potregression"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Kvadratisk regression"
     .Style = msoButtonCaption
     .OnAction = "polregression"
    End With

    'ReducerMenu
'    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
'    myCPup1.Caption = "Reducer"
'    myCPup1.Tag = ""
'    With myCPup1.Controls.Add(Type:=msoControlButton)
'     .Caption = "Simpel Udregning (Alt+b)"
'     .Style = msoButtonCaption
'     .OnAction = "SimpelUdregning"
'    End With
'    With myCPup1.Controls.Add(Type:=msoControlButton)
'     .Caption = "Omregn decimal til brøk   tofrac()"
'     .Style = msoButtonCaption
'     .OnAction = "tofrac"
'    End With

    ' Ligninger
'    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
'    myCPup1.Caption = "Ligninger"
'    myCPup1.Tag = ""
' Add buttons to popup
'    With myCPup1.Controls.Add(Type:=msoControlButton)
'     .Caption = "Løs ligning numerisk"
'     .Style = msoButtonCaption
'     .OnAction = "nsolve"
'    End With

    'Indsæt statistik menu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Statistik"
    myCPup1.Tag = ""
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Beregn hvor mange gange man kan udtage k elementer fra n. Rækkefølgen underordnet. Combination(n,k)"
     .DescriptionText = "Beregn hvor mange gange man kan udtage k elementer fra n. Rækkefølgen underordnet."
     .Tag = "Combination(n,k)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Beregn hvor mange gange man kan udtage k elementer fra n. Rækkefølgen tæller med. permutation(n,k)"
     .DescriptionText = "Beregn hvor mange gange man kan udtage k elementer fra n. Rækkefølgen af de udtagne elementer regnes med."
     .Tag = "permutation(n,k)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Tilfældigt tal mellem 0 og n. random(n)"
     .DescriptionText = ""
     .Tag = "Random(n)"
     .Style = msoButtonCaption
     .OnAction = "indsaetformel"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & " - Test for sammenhæng"
     .DescriptionText = ""
     .Tag = ""
     .Style = msoButtonCaption
     .OnAction = "Chi2Test"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & " - frekvensfordeling"
     .DescriptionText = ""
     .Tag = ""
     .Style = msoButtonCaption
     .OnAction = "Chi2fordeling"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Statistik regneark"
     .DescriptionText = ""
     .Tag = ""
     .Style = msoButtonCaption
     .OnAction = "OpenSpreadsheet"
    End With
    
    'Indsæt Geometri menu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Geometri"
    myCPup1.Tag = ""
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "GeoGebra"
     .Style = msoButtonIconAndCaption
     .FaceId = 212 ' Geometri
     .Tag = ""
     .OnAction = "GeoGebra"
    End With
    
    ' |
    With myCB.Controls.Add(Type:=msoControlButton)
     .Caption = "|"
     .Style = msoButtonCaption
    End With
        
    'Symbol og Erstat menu
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Symboler"
    myCPup1.Tag = ""
    ' Add buttons to popup
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Indsæt prik gangetegn"
     .Style = msoButtonCaption
     .OnAction = "Gange"
     .ShortcutText = "Alt+G"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "I matematik . " & VBA.ChrW(&H2192) & " , (og , " & VBA.ChrW(&H2192) & " ;)"
     .Style = msoButtonCaption
     .OnAction = "ErstatPunktum"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "I matematik  ," & VBA.ChrW(&H2192) & " . (og ; " & VBA.ChrW(&H2192) & " ,)"
     .Style = msoButtonCaption
     .OnAction = "ErstatKomma"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "* " & VBA.ChrW(&H2192) & " " & VBA.ChrW(183)
     .Style = msoButtonCaption
     .OnAction = "ReplaceStarMult"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = VBA.ChrW(183) & " " & VBA.ChrW(&H2192) & " *"
     .Style = msoButtonCaption
     .OnAction = "ReplaceStarMultBack"
    End With
    
    ' hjælp
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    myCPup1.Caption = "Hjælp"
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Hjælp"
     .Style = msoButtonCaption
     .OnAction = "HjælpeMenu"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "GeoGebra videovejledninger"
     .Tag = "http://www.laerit.dk/geogebra/"
     .Style = msoButtonCaption
     .OnAction = "GoToLink"
    End With
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Om MathMenu"
     .Style = msoButtonCaption
     .OnAction = "OmMathMenu"
    End With
    
    With myCPup1.Controls.Add(Type:=msoControlButton)
     .Caption = "Hjælp jeg kan ikke gemme."
     .Style = msoButtonCaption
     .OnAction = "SolveCantSaveProblem"
    End With
    
'     Application.MacroOptions Macro:="beregn", HasShortcutKey:=True, ShortcutKey:="%b"
'    Application.OnKey "%b", "beregn"

    ' Show the command bar
    myCB.visible = True
End Sub
Sub OmMathMenu()
    Dim v As String
    v = AppVersion
    If PatchVersion <> "" Then
        v = v & PatchVersion
    End If
    MsgBox Sprog.A(20), vbOKOnly, AppNavn & " version " & v
End Sub
Sub hjaelpeMenu()
Dim FilNavn As String
On Error GoTo fejl
'filnavn = """" & Environ("ProgramFiles") & "\MathMenu\MathMenuManual.docx"""
FilNavn = GetProgramFilesDir & "\WordMat\WordMatManual.docx"
If Dir(FilNavn) <> "" Then
    Documents.Open FileName:=FilNavn, ReadOnly:=True
Else
    MsgBox "Cant locate the help-file", vbOKOnly, Sprog.Error
End If

GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub indsaetformel()
    On Error GoTo fejl
'    MsgBox CommandBars.ActionControl.Caption
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If

    Application.ScreenUpdating = False
    If CommandBars.ActionControl.DescriptionText <> "" Then
    Selection.InsertAfter (CommandBars.ActionControl.DescriptionText)
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    End If
    Selection.InsertAfter (CommandBars.ActionControl.Tag)
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).BuildUp
    Selection.MoveRight Unit:=wdCharacter, Count:=2
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Function VisDef() As String
'Dim omax As New CMaxima
Dim deftext As String
    On Error GoTo fejl
    PrepareMaxima
    deftext = omax.DefString
    If Len(deftext) > 3 Then
'    deftext = Mid(deftext, 2, Len(deftext) - 3)
    deftext = Replace(deftext, "$", vbCrLf)
    deftext = Replace(deftext, ":=", " = ")
    deftext = Replace(deftext, ":", " = ")
    If DecSeparator = "," Then
        deftext = Replace(deftext, ",", ";")
        deftext = Replace(deftext, ".", ",")
    End If
    deftext = Sprog.A(113) & vbCrLf & vbCrLf & deftext
    Else
        deftext = Sprog.A(114)
    End If
    VisDef = deftext
'    MsgBox deftext, vbOKOnly, "Definitioner"
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Function
Sub DefinerVar()
    Dim var As String
    On Error GoTo fejl
'    var = InputBox("Indtast definitionen på den nye variabel" & vbCrLf & vbCrLf & "Definitionen kan benyttes i resten af dokumentet, men ikke før. Hvis der indsættes en clearvars: kommando længere nede i dokumentet kan den ikke benyttes derefter." & vbCrLf & vbCrLf & "Definitionen kan indtastes på 4 forskellige måder" & vbCrLf & vbCrLf & "definer: variabel=værdi" & vbCrLf & "variabel:værdi" & vbCrLf & "variabel:=værdi" & vbCrLf & "variabel" & VBA.ChrW(&H2261) & "værdi  (Definitions ligmed)" & vbCrLf & "Der kan defineres flere variable i en ligningsboks ved at adskille definitionerne med semikolon. f.eks. a:1 ; b:2", "Ny variabel", "a=1")
    var = InputBox(Sprog.A(120), Sprog.A(121), "a=1")
    var = Replace(var, ":=", "=")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    If var <> "" Then
        var = Sprog.A(126) & ": " & var
        Selection.InsertAfter (var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
    
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub DefinerFunktion()
    Dim var As String
On Error GoTo fejl
'    var = InputBox("Indtast definitionen på den nye funktion" & vbCrLf & vbCrLf & "Definitionen kan benyttes i resten af dokumentet, men ikke før. Hvis der indsættes en clearvars: kommando længere nede i dokumentet kan den ikke benyttes derefter." & vbCrLf & vbCrLf & "Definitionen kan indtastes på 3 forskellige måder" & vbCrLf & vbCrLf & "f(x):forskrift" & vbCrLf & "f(x):=forskrift" & vbCrLf & "f(x)" & VBA.ChrW(&H2261) & "forskrift  (Definitions ligmed)" & vbCrLf & "Der kan defineres flere funktioner i en ligningsboks ved at adskille definitionerne med semikolon. f.eks. f(x)=x ; g(x)=2x+1", "Ny funktion", "f(x)=x+1")
    var = InputBox(Sprog.A(122), Sprog.A(123), "f(x)=x+1")
    var = Replace(var, ":=", "=")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    
    If var <> "" Then
        var = Sprog.A(126) & ": " & var
        Selection.InsertAfter (var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub DefinerLigning()
    Dim var As String
On Error GoTo fejl
    var = InputBox(Sprog.A(115), Sprog.A(124), Sprog.A(125) & ":     Area:A=1/2*h*b")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    
    If var <> "" Then
        Selection.InsertAfter (var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ErstatPunktum()
'
' ErstatPunktum Makro
'
'
On Error GoTo fejl
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "."
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ErstatKomma()
'
' ErstatPunktum Makro
'
'
On Error GoTo fejl
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ";"
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub Nsolve()
    Dim Ligning As String
    Dim objEq As OMath
    Dim objRange As Range

    Set objRange = Selection.Range
    Ligning = InputBox("Indtast ligning", "Numerisk Ligningsløsning")
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:="nsolve(" & Ligning & ")"
    Set objEq = objRange.OMaths(1)
    objEq.BuildUp

End Sub

Sub FlytLigningerNed(antal As Integer)
    Dim i As Integer
    
    For i = 1 To antal
    UserForm2DGraph.TextBox_ligning6.Text = UserForm2DGraph.TextBox_ligning5.Text
    UserForm2DGraph.TextBox_ligning5.Text = UserForm2DGraph.TextBox_ligning4.Text
    UserForm2DGraph.TextBox_ligning4.Text = UserForm2DGraph.TextBox_ligning3.Text
    UserForm2DGraph.TextBox_ligning3.Text = UserForm2DGraph.TextBox_ligning2.Text
    UserForm2DGraph.TextBox_ligning2.Text = UserForm2DGraph.TextBox_ligning1.Text
    Next
End Sub
Function GetMathText(om As OMath) As String
On Error GoTo fejl
    om.ConvertToNormalText
    GetMathText = om.Range.Text
    om.ConvertToMathText
    om.BuildUp
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Function
Sub Gange()
' prik Gangetegn
On Error Resume Next
'    Selection.InsertSymbol Font:="+Brødtekst", CharacterNumber:=183, Unicode:=True

'    Selection.InsertSymbol Font:="+Brødtekst", CharacterNumber:=AscW(MaximaGangeTegn), Unicode:=True
    Selection.InsertSymbol CharacterNumber:=AscW(MaximaGangeTegn), Unicode:=True 'font brødtekst fjernet for at understøtte international

End Sub
Sub SimpelUdregning()
' laver simpel udregning med 4 regningsarter og ^
    
    On Error GoTo slut
    Dim crange As Range
    Dim r As Range
    Dim sindex As Integer
    Dim resultat As String
    
    Application.ScreenUpdating = False
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        If Selection.Range.Font.Bold = True Then
            Selection.Range.Font.Bold = False
        End If
        Selection.OMaths(1).Range.Select
        Selection.OMaths(1).Linearize
    End If
    If Len(Selection.Text) < 2 Then
'        MsgBox "Marker det udtryk der skal beregnes. Udtrykket må kun indeholde tal, de fire regningsarter og ^ ."
        Set r = Selection.Range

        sindex = 0
        Call r.MoveStart(wdCharacter, -1)
        Do
        sindex = sindex + 1
        Call r.MoveStart(wdCharacter, -1)
        Loop While sindex < 20 And (AscW(r.Characters(1)) > 39 And AscW(r.Characters(1)) < 58 Or AscW(r.Characters(1)) = 94 Or AscW(r.Characters(1)) = 183)
        If sindex < 20 Then Call r.MoveStart(wdCharacter, 1)
        Selection.start = r.start
'        Selection.End = r.End
    End If
    
    Call ActiveDocument.Range.Find.Execute(VBA.ChrW(8727), , , , , , , , , "*", wdReplaceAll) ' nødvendig til mathboxes
    Call Selection.Range.Find.Execute(VBA.ChrW(183), , , , , , , , , "*", wdReplaceAll)
    Call Selection.Range.Find.Execute(".", , , , , , , , , ",", wdReplaceAll)
    resultat = Selection.Range.Calculate()
    Selection.Range.InsertAfter ("=" & resultat)
    Call Selection.Range.Find.Execute(VBA.ChrW(42), , , , , , , , , VBA.ChrW(183), wdReplaceAll)
    Selection.MoveEnd Unit:=wdCharacter, Count:=Len(resultat) + 1
    Call Selection.Range.Find.Execute(",", , , , , , , , , ".", wdReplaceAll)
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths.BuildUp
    Selection.Collapse (wdCollapseEnd)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
'    Selection.TypeText (" ")

slut:
End Sub
Sub tofrac()
Dim udtryk As String
    If Len(Selection.Text) > 1 Then
        udtryk = Selection.Text
    Else
        udtryk = InputBox("Indtast decimaltal", "Fra decimaltal til brøk")
    End If
'    Selection
    Selection.InsertAfter ("tofrac(" & udtryk & ")")
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).BuildUp
    Selection.MoveRight Unit:=wdCharacter, Count:=2

End Sub
Sub ReplaceStarMult()
' fjerner stjerner og indsætter alm. gangetegn
Application.ScreenUpdating = False
On Error GoTo fejl

'    Call ActiveDocument.Range.Find.Execute(chr(42), , , , , , , , , VBA.ChrW(183), wdReplaceAll)
    Call ActiveDocument.Range.Find.Execute(VBA.ChrW(8727), , , , , , , , , VBA.ChrW(183), wdReplaceAll) ' nødvendig til mathboxes
    Call ActiveDocument.Range.Find.Execute("*", , , , , , , , , VBA.ChrW(183), wdReplaceAll)

'    MsgBox "Alle * er nu lavet om til " & VBA.ChrW(183)
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ReplaceStarMultBack()
' fjerner alm gangetegn og indsætter *
Application.ScreenUpdating = False
On Error GoTo fejl
    Call Selection.Range.Find.Execute(VBA.ChrW(183), , , , , , , , , "*", wdReplaceAll)
    Call Selection.Range.Find.Execute(VBA.ChrW(8901), , , , , , , , , "*", wdReplaceAll) '\cdot
    Call Selection.Range.Find.Execute(VBA.ChrW(8729), , , , , , , , , "*", wdReplaceAll) ' \cdot
    Call Selection.Range.Find.Execute(VBA.ChrW(8226), , , , , , , , , "*", wdReplaceAll) ' tyk prik
    
'    MsgBox "Alle " & VBA.ChrW(183) & " er nu lavet om til *"
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub MaximaSettings()
On Error GoTo fejl
    If UFMSettings Is Nothing Then Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
    GoTo slut
fejl:
    Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
slut:
End Sub
Sub GoToLink()
    Dim Link As String
    Dim explorersti As String
    Dim appnr As Integer
    On Error GoTo fejl
    Link = CommandBars.ActionControl.Tag

    explorersti = """" & GetProgramFilesDir & "\Internet Explorer\iexplore.exe"" " & Link
'    On Error GoTo fejl
    appnr = Shell(explorersti, vbNormalFocus) 'vbNormalFocus vbMinimizedFocus
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

