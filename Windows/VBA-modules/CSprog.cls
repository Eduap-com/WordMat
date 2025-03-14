VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SprogNr As Integer
Private SA(800) As String
'0=WordLanguage, 1=Dansk, 2=English, 3=Spanish

Private Sub Class_Initialize()
    CheckSetting
    If Sprog.SprogNr = 1 Then
        AppNavn = "WordMat"
    Else
        AppNavn = "WordMath"
    End If
    LoadSprogArray
End Sub
Public Sub CheckSetting()
Dim b As Boolean
    If SprogNr < 1 Then b = True
    If LanguageSetting = 0 Then
        SprogNr = GetOfficeLanguage
    Else
        SprogNr = LanguageSetting
    End If
    If b = True Then
    End If
End Sub
Function GetOfficeLanguage() As Integer
'Reads the UI-language for Word
Select Case Application.International(wdProductLanguageID)
Case 1030 ' Dansk
    GetOfficeLanguage = 1
Case 1034 ' Spanish
    GetOfficeLanguage = 3
Case Else ' 1033 is english
    GetOfficeLanguage = 2
End Select

End Function
Public Function A(i As Long) As String
    On Error Resume Next
    If SA(0) = vbNullString Then LoadSprogArray
    A = SA(i)
End Function
Public Sub LoadSprogArray()
' method from ReadTextFile
' also check ReadTextfileToString
    Dim FileNo As Integer
    Dim linje As String
    Dim Arr() As String
    Dim FilePath As String
    
    FilePath = GetWordMatDir
    If SprogNr = 1 Then
        FilePath = FilePath & "languages\da.csv"
    ElseIf SprogNr = 2 Then
        FilePath = FilePath & "languages\en.csv"
    ElseIf SprogNr = 3 Then
        FilePath = FilePath & "languages\sp.csv"
    End If
    FileNo = FreeFile

    Open FilePath For Input As FileNo
    Do While Not EOF(FileNo) ' Loop until end of file.
        Line Input #FileNo, linje
        Arr = Split(linje, "#£")
        If UBound(Arr) > 0 Then
            If IsNumeric(Arr(0)) Then
                SA(CInt(Arr(0))) = FormatText(Arr(1))
            End If
        End If
    Loop
    Close FileNo
End Sub

Function FormatText(ByVal t As String) As String
    t = Replace(t, "\r\n", VbCrLfMac)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "$Tab", vbTab)
    t = Replace(t, "$*", ChrW(183))
    t = Replace(t, "$C", ChrW(8450))
    t = Replace(t, "$R", ChrW(8477))
    t = Replace(t, "$X^2", ChrW(&H3C7) & ChrW(&HB2))
    t = Replace(t, "$Degree", ChrW(176))
    
    FormatText = t
End Function

' i xml må der ikke være label attribut når der er  callback.
' prepare for new language by find/replace Case 3: s="" ->getLabel Case 3: s=""   vbrlf   Case 4: s=""

Public Property Get RibFormulae() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Formler"
Case 3: s = "Formulario"
Case 4: s = ""
Case Else: s = "Formulae"
End Select
RibFormulae = s
End Property
Public Property Get RibMath() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Mat"
Case 3: s = "Mat"
Case 4: s = ""
Case Else: s = "Math"
End Select
RibMath = s
End Property
Public Property Get RibPhysics() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fys"
Case 3: s = "Física"
Case 4: s = ""
Case Else: s = "Phys"
End Select
RibPhysics = s
End Property
Public Property Get RibChemistry() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Kemi"
Case 3: s = "Química"
Case 4: s = ""
Case Else: s = "Chem"
End Select
RibChemistry = s
End Property
Public Property Get RibPercentage() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Procent"
Case 3: s = "Porcentaje"
Case 4: s = ""
Case Else: s = "Percentage"
End Select
RibPercentage = s
End Property
Public Property Get RibFunctions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Funktioner"
Case 3: s = "Funciones"
Case 4: s = ""
Case Else: s = "Functions"
End Select
RibFunctions = s
End Property
Public Property Get RibGeometry() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Geometri"
Case 3: s = "Geometría"
Case 4: s = ""
Case Else: s = "Geometry"
End Select
RibGeometry = s
End Property
Public Property Get RibProbabilityShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sands."
Case 3: s = "Probabilidades"
Case 4: s = ""
Case Else: s = "Prob."
End Select
RibProbabilityShort = s
End Property
Public Property Get RibInfinitesimalShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "inf."
Case 3: s = "Infinitesimal"
Case 4: s = ""
Case Else: s = "inf."
End Select
RibInfinitesimalShort = s
End Property
Public Property Get RibVector() As String
Dim s As String
Select Case SprogNr
Case 1: s = "vektor"
Case 3: s = "Vector"
Case 4: s = ""
Case Else: s = "vector"
End Select
RibVector = s
End Property

Public Property Get RibSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indstillinger"
Case 3: s = "Configuración"
Case 4: s = ""
Case Else: s = "Settings"
End Select
RibSettings = s
End Property
Public Property Get RibSettingsShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indst."
Case 3: s = "Configuración"
Case 4: s = ""
Case Else: s = "Settings"
End Select
RibSettingsShort = s
End Property
Public Property Get RibSciNot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "vid"
Case 3: s = "NCi" 'Notación Científica
Case 4: s = ""
Case Else: s = "sci"
End Select
RibSciNot = s
End Property

Public Property Get RibBeregn() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Beregn"
Case 3: s = "Calcular"
Case 4: s = ""
Case Else: s = "Calculate"
End Select
RibBeregn = s
End Property

Public Property Get RibSolve() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs"
Case 3: s = "Resolver"
Case 4: s = ""
Case Else: s = "Solve"
End Select
RibSolve = s
End Property

Public Property Get RibMaximaCommand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Maxima-kommando"
Case 3: s = "Maxima command"
Case 4: s = ""
Case Else: s = "Maxima command"
End Select
RibMaximaCommand = s
End Property
Public Property Get RibSolveNum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs ligning(er) vha. numerisk metode"
Case 3: s = "Resolver ecuación(es) usando métodos numéricos"
Case 4: s = ""
Case Else: s = "Solve equation(s) using numerical method"
End Select
RibSolveNum = s
End Property
Public Property Get RibEliminate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eliminer variable"
Case 3: s = "Eliminar variables"
Case 4: s = ""
Case Else: s = "Eliminate variables"
End Select
RibEliminate = s
End Property
Public Property Get TestTF() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Test sand/falsk"
Case 3: s = "Prueba Verdadero/falso"
Case 4: s = ""
Case Else: s = "Test true/false"
End Select
TestTF = s
End Property
Public Property Get RibSolveDE() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs differentialligning"
Case 3: s = "Resolver ecuación diferencial"
Case 4: s = ""
Case Else: s = "Solve differential equation"
End Select
RibSolveDE = s
End Property
Public Property Get RibDefinitions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
RibDefinitions = s
End Property
Public Property Get RibDefDelete() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Slet definitioner"
Case 3: s = "Borrar definiciones"
Case 4: s = ""
Case Else: s = "Delete definitions"
End Select
RibDefDelete = s
End Property
Public Property Get RibDefVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer variabel"
Case 3: s = "Definir variable"
Case 4: s = ""
Case Else: s = "Define variable"
End Select
RibDefVariable = s
End Property
Public Property Get RibDefFunction() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer funktion"
Case 3: s = "Definir función"
Case 4: s = ""
Case Else: s = "Define function"
End Select
RibDefFunction = s
End Property
Public Property Get RibDefEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer ligning"
Case 3: s = "Definir ecuación"
Case 4: s = ""
Case Else: s = "Define equation"
End Select
RibDefEquation = s
End Property
Public Property Get RibDefConstants() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer fysiske konstanter"
Case 3: s = "Definir constantes físicas"
Case 4: s = ""
Case Else: s = "Define physical constants"
End Select
RibDefConstants = s
End Property

Public Property Get RibReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omskriv"
Case 3: s = "Reducir"
Case 4: s = ""
Case Else: s = "Reduce"
End Select
RibReduce = s
End Property
Public Property Get RibInfinitesimal() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Infinites."
Case 3: s = "Infinitos"
Case 4: s = ""
Case Else: s = "Infinites."
End Select
RibInfinitesimal = s
End Property
Public Property Get RibDif() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Differentier"
Case 3: s = "Diferenciar"
Case 4: s = ""
Case Else: s = "Differentiate"
End Select
RibDif = s
End Property
Public Property Get RibIntegrate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Stamfunktion"
Case 3: s = "Integrar"
Case 4: s = ""
Case Else: s = "Integrate"
End Select
RibIntegrate = s
End Property

Public Property Get RibPlotting() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graftegning"
Case 3: s = "Graficar"
Case 4: s = ""
Case Else: s = "Plotting"
End Select
RibPlotting = s
End Property
Public Property Get RibShowGraph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis Graf"
Case 3: s = "Mostrar gráfico"
Case 4: s = ""
Case Else: s = "Show Graph"
End Select
RibShowGraph = s
End Property
Public Property Get RibDirectionField() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Hældningsfelt"
Case 3: s = "Dirección del campo"
Case 4: s = ""
Case Else: s = "Direction field"
End Select
RibDirectionField = s
End Property
Public Property Get RibInsertGeoGebra() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt GeoGebra fil"
Case 3: s = "Insertar GeoGebra archivo"
Case 4: s = ""
Case Else: s = "Insert GeoGebra file"
End Select
RibInsertGeoGebra = s
End Property
Public Property Get Rib3DPlot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "3D Plot"
Case 3: s = "Gráfico 3D"
Case 4: s = ""
Case Else: s = "3D Plot"
End Select
Rib3DPlot = s
End Property
Public Property Get Rib3DRotate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omdrejningslegeme"
Case 3: s = "Volumen de revolución - Gráfico 3D rota en el eje x"
Case 4: s = ""
Case Else: s = "Volume of revolution - 3D Rotate Graph about x-axis"
End Select
Rib3DRotate = s
End Property
Public Property Get RibStatistics() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Statistik"
Case 3: s = "Estadística"
Case 4: s = ""
Case Else: s = "Statistics"
End Select
RibStatistics = s
End Property
Public Property Get RibUgrup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ugrupperede observationer"
Case 3: s = "Observaciones no agrupadas"
Case 4: s = ""
Case Else: s = "Ungrouped observations"
End Select
RibUgrup = s
End Property
Public Property Get RibGrup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grupperede observationer"
Case 3: s = "Observaciones agrupadas"
Case 4: s = ""
Case Else: s = "Grouped observations"
End Select
RibGrup = s
End Property
Public Property Get RibStickChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Pindediagram"
Case 3: s = "Gráfico de Barras"
Case 4: s = ""
Case Else: s = "Stick chart"
End Select
RibStickChart = s
End Property
Public Property Get RibHistogram() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Histogram"
Case 3: s = "Histograma"
Case 4: s = ""
Case Else: s = "Histogram"
End Select
RibHistogram = s
End Property
Public Property Get RibStepChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trappediagram"
Case 3: s = "Gráfico escalonado"
Case 4: s = ""
Case Else: s = "Step chart"
End Select
RibStepChart = s
End Property
Public Property Get RibCumChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sumkurve"
Case 3: s = "Gráfico acumulativo"
Case 4: s = ""
Case Else: s = "Cumulative chart"
End Select
RibCumChart = s
End Property
Public Property Get RibBoxPlot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Boksplot"
Case 3: s = "Gráfico de cajas y conectores"
Case 4: s = ""
Case Else: s = "Box and whiskers plot"
End Select
RibBoxPlot = s
End Property
Public Property Get RibStatProb() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Statistik/sands."
Case 3: s = "Probabilidades y Estadística"
Case 4: s = ""
Case Else: s = "Statistics/prob."
End Select
RibStatProb = s
End Property
Public Property Get RibRegression() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Regression"
Case 3: s = "Regresión"
Case 4: s = ""
Case Else: s = "Regression"
End Select
RibRegression = s
End Property
Public Property Get RibInsertTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt tabel"
Case 3: s = "Insertar tabla"
Case 4: s = ""
Case Else: s = "Insert table"
End Select
RibInsertTable = s
End Property
Public Property Get RibLinRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Lineær"
Case 3: s = "Lineal"
Case 4: s = ""
Case Else: s = "Linear"
End Select
RibLinRegr = s
End Property
Public Property Get RibExpRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eksponentiel"
Case 3: s = "Exponencial"
Case 4: s = ""
Case Else: s = "Exponential"
End Select
RibExpRegr = s
End Property
Public Property Get RibPowRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Potens"
Case 3: s = "Potencias"
Case 4: s = ""
Case Else: s = "Power"
End Select
RibPowRegr = s
End Property
Public Property Get RibPolRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Polynomisk"
Case 3: s = "Polinomios"
Case 4: s = ""
Case Else: s = "Polynomial"
End Select
RibPolRegr = s
End Property
Public Property Get RibDistributions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fordelinger"
Case 3: s = "Distribuciones"
Case 4: s = ""
Case Else: s = "Distributions"
End Select
RibDistributions = s
End Property
Public Property Get RibBinomDist() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Binomialfordeling"
Case 3: s = "Distribución binomial"
Case 4: s = ""
Case Else: s = "Binomial distribution"
End Select
RibBinomDist = s
End Property
Public Property Get RibNormalDist() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Normalfordeling"
Case 3: s = "Distribución Normal"
Case 4: s = ""
Case Else: s = "Normal distribution"
End Select
RibNormalDist = s
End Property
Public Property Get RibChi2Dist() As String
Dim s As String
Select Case SprogNr
Case 1: s = VBA.ChrW(967) & VBA.ChrW(178) & "-fordeling"
Case 3: s = "Distribución " & VBA.ChrW(967) & VBA.ChrW(178)
Case 4: s = ""
Case Else: s = VBA.ChrW(967) & VBA.ChrW(178) & " distribution"
End Select
RibChi2Dist = s
End Property
Public Property Get RibTest() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Test"
Case 3: s = "Prueba"
Case 4: s = ""
Case Else: s = "Test"
End Select
RibTest = s
End Property
Public Property Get RibBinomTest() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Binomialtest"
Case 3: s = "Prueba binomial"
Case 4: s = ""
Case Else: s = "Binomial test"
End Select
RibBinomTest = s
End Property
Public Property Get RibChi2Test() As String
Dim s As String
Select Case SprogNr
Case 1: s = VBA.ChrW(967) & VBA.ChrW(178) & "-test for sammenhæng"
Case 3: s = "Prueba " & VBA.ChrW(967) & VBA.ChrW(178) & ""
Case 4: s = ""
Case Else: s = VBA.ChrW(967) & VBA.ChrW(178) & " test for independence"
End Select
RibChi2Test = s
End Property
Public Property Get RibSimulation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Simulering"
Case 3: s = "Simulación"
Case 4: s = ""
Case Else: s = "Simulation"
End Select
RibSimulation = s
End Property
Public Property Get RibGroup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grupp" & VBA.ChrW(233) & "r"
Case 3: s = "Grupo" & VBA.ChrW(233) & "r"""
Case 4: s = ""
Case Else: s = "Group"
End Select
RibGroup = s
End Property
Public Property Get RibDiverse() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Diverse"
Case 3: s = "Otros"
Case 4: s = ""
Case Else: s = "Others"
End Select
RibDiverse = s
End Property
Public Property Get RibNewEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ny ligning"
Case 3: s = "Nueva ecuación"
Case 4: s = ""
Case Else: s = "New Equation"
End Select
RibNewEquation = s
End Property
Public Property Get RibSymbols() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Symboler"
Case 3: s = "Símbolos"
Case 4: s = ""
Case Else: s = "Symbols"
End Select
RibSymbols = s
End Property
Public Property Get RibMultsign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Prik gangetegn"
Case 3: s = "Signo de multiplicación"
Case 4: s = ""
Case Else: s = "Dot multiplication sign"
End Select
RibMultsign = s
End Property
Public Property Get RibFigurs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Figurer"
Case 3: s = "Figuras"
Case 4: s = ""
Case Else: s = "Shapes"
End Select
RibFigurs = s
End Property
Public Property Get RibDraw() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tegn"
Case 3: s = "Dibujar"
Case 4: s = ""
Case Else: s = "Draw"
End Select
RibDraw = s
End Property
Public Property Get RibTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tabel"
Case 3: s = "Tabla"
Case 4: s = ""
Case Else: s = "Table"
End Select
RibTable = s
End Property
Public Property Get RibInsertExcel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Excel"
Case 3: s = "Insertar Excel"
Case 4: s = ""
Case Else: s = "Insert Excel"
End Select
RibInsertExcel = s
End Property
Public Property Get RibTableToList() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tabel " & VBA.ChrW(8594) & " Liste"
Case 3: s = "Tabla " & VBA.ChrW(8594) & " Lista "
Case 4: s = ""
Case Else: s = "Table " & VBA.ChrW(8594) & " List"
End Select
RibTableToList = s
End Property
Public Property Get RibListToTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Liste " & VBA.ChrW(8594) & " Tabel"
Case 3: s = ""
Case 4: s = ""
Case Else: s = "List " & VBA.ChrW(8594) & " Table"
End Select
RibListToTable = s
End Property
Public Property Get RibTriangle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trekant"
Case 3: s = "Triángulo"
Case 4: s = ""
Case Else: s = "Triangle"
End Select
RibTriangle = s
End Property

''''''''''''''''''''''''''''''''''''
''''''''''''''Forms'''''''''''''''''
''''''''''''''''''''''''''''''''''''

''''''''''''''MaximaSettings'''''''''
Public Property Get CASSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "CAS indstillinger"
Case 3: s = "Configuración de CAS"
Case 4: s = ""
Case Else: s = "CAS settings"
End Select
CASSettings = s
End Property
Public Property Get Notation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Notation"
Case 3: s = "Notación"
Case 4: s = ""
Case Else: s = "Notation"
End Select
Notation = s
End Property
Public Property Get Graph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graf"
Case 3: s = "Gráfico"
Case 4: s = ""
Case Else: s = "Graph"
End Select
Graph = s
End Property
Public Property Get Advanced() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Avanceret"
Case 3: s = "Avanzado"
Case 4: s = ""
Case Else: s = "Advanced"
End Select
Advanced = s
End Property
Public Property Get ScientificNotation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Videnskabelig notation"
Case 3: s = "Notación científica"
Case 4: s = ""
Case Else: s = "Scientific notation"
End Select
ScientificNotation = s
End Property
Public Property Get Complex() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Komplekse tal"
Case 3: s = "Números complejos"
Case 4: s = ""
Case Else: s = "Complex numbers"
End Select
Complex = s
End Property
Public Property Get SignificantFigures() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Betydende cifre"
Case 3: s = "Cifras significativas"
Case 4: s = ""
Case Else: s = "Significant figures"
End Select
SignificantFigures = s
End Property
Public Property Get Units() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Enheder"
Case 3: s = "Unidades"
Case 4: s = ""
Case Else: s = "Units"
End Select
Units = s
End Property
Public Property Get PolarNotation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Polær notation output"
Case 3: s = "Salida en notación polar"
Case 4: s = ""
Case Else: s = "Polar notation output"
End Select
PolarNotation = s
End Property
Public Property Get Degrees() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grader"
Case 3: s = "Grados"
Case 4: s = ""
Case Else: s = "Degrees"
End Select
Degrees = s
End Property
Public Property Get Radians() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Radianer"
Case 3: s = "Radianes"
Case 4: s = ""
Case Else: s = "Radians"
End Select
Radians = s
End Property
Public Property Get AngleUnit() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vinkelenhed"
Case 3: s = "Unidad angular"
Case 4: s = ""
Case Else: s = "Angle unit"
End Select
AngleUnit = s
End Property
Public Property Get ShowSolutionConditions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis løsningsbetingelser (beta)"
Case 3: s = "Ver condiciones de solución"
Case 4: s = ""
Case Else: s = "Show solution conditions (beta)"
End Select
ShowSolutionConditions = s
End Property
Public Property Get Exact() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eksakt"
Case 3: s = "Exacto"
Case 4: s = ""
Case Else: s = "Exact"
End Select
Exact = s
End Property
Public Property Get Numeric() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Numerisk"
Case 3: s = "Numérico"
Case 4: s = ""
Case Else: s = "Numeric"
End Select
Numeric = s
End Property
Public Property Get Auto() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Automatisk"
Case 3: s = "Automático"
Case 4: s = ""
Case Else: s = "Automatic"
End Select
Auto = s
End Property
Public Property Get HighPrecision() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Meget høj præcision (beta)"
Case 3: s = "Muy alta precisión"
Case 4: s = ""
Case Else: s = "Very high precision (beta)"
End Select
HighPrecision = s
End Property
Public Property Get InsertExplanation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt forklaring"
Case 3: s = "Insertar explicación"
Case 4: s = ""
Case Else: s = "Insert explanation"
End Select
InsertExplanation = s
End Property
Public Property Get InsertMaximaCommand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Maxima kommando"
Case 3: s = "Insertar comando Maxima"
Case 4: s = ""
Case Else: s = "Insert Maxima command"
End Select
InsertMaximaCommand = s
End Property
Public Property Get Separators() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Separatorer"
Case 3: s = "Separadores"
Case 4: s = ""
Case Else: s = "Separators"
End Select
Separators = s
End Property
Public Property Get List() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Liste"
Case 3: s = "Lista"
Case 4: s = ""
Case Else: s = "List"
End Select
List = s
End Property
Public Property Get Multsign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gangetegn"
Case 3: s = "Signos de multiplicación"
Case 4: s = ""
Case Else: s = "Mult.sign"
End Select
Multsign = s
End Property
Public Property Get Dot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Prik"
Case 3: s = "Punto"
Case 4: s = ""
Case Else: s = "Dot"
End Select
Dot = s
End Property
Public Property Get SolutionType() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsningsangivelse"
Case 3: s = "Solución tipo"
Case 4: s = ""
Case Else: s = "Solution type"
End Select
SolutionType = s
End Property
Public Property Get SolutionBoolean() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Boolsk"
Case 3: s = "Booleano"
Case 4: s = ""
Case Else: s = "Boolean"
End Select
SolutionBoolean = s
End Property
Public Property Get SolutionSet() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsningsmængde"
Case 3: s = "Conjunto de soluciones"
Case 4: s = ""
Case Else: s = "Solution Set"
End Select
SolutionSet = s
End Property
Public Property Get IndexSubscript() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sænket skrift"
Case 3: s = "Subíndice"
Case 4: s = ""
Case Else: s = "Subscript"
End Select
IndexSubscript = s
End Property
Public Property Get JustText() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Er bare tekst som er del af variabelnavn"
Case 3: s = "Es simplemente el texto que forma parte del nombre de la variable"
Case 4: s = ""
Case Else: s = "Is just text part of variable name"
End Select
JustText = s
End Property
Public Property Get IsIndex() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Er liste/matrix indeks"
Case 3: s = "Indice de lista / matriz"
Case 4: s = ""
Case Else: s = "Is list/matrix index"
End Select
IsIndex = s
End Property
Public Property Get StandardGraph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Standard grafprogram"
Case 3: s = "Programa de gráficos"
Case 4: s = ""
Case Else: s = "Standard graph program"
End Select
StandardGraph = s
End Property
Public Property Get ExcelEmbed() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Excelark indlejret"
Case 3: s = "Insertar hoja Excel"
Case 4: s = ""
Case Else: s = "Embed Excel sheets"
End Select
ExcelEmbed = s
End Property
Public Property Get OutputUnits() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Outputenheder"
Case 3: s = "Unidades de salida"
Case 4: s = ""
Case Else: s = "Output units"
End Select
OutputUnits = s
End Property
Public Property Get UnitHelp() As String
Dim s As String
Select Case SprogNr
Case 1
s = "Angiv hvilke enheder som skal være i output adskilt ad komma" & VbCrLfMac & "eksempler:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ, Timer" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "Hvis intet angives returneres alle enheder som base SI-enheder eller afledte SI-enheder uden præfiks."
Case 3: s = "Establece unidades de salida " & VbCrLfMac & " separadas por coma ejemplos:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "Si no se indican las unidades se devuelve en unidades básicas del SI o unidades del SI derivadas sin prefijo."
Case 4: s = ""
Case Else:
s = "State output units separated by comma" & VbCrLfMac & "examples:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "If nothing is stated units are returned in basic SI-units or SI-derived units without prefix."
End Select
UnitHelp = s
End Property
Public Property Get UnitExamples() As String
Dim s As String
Select Case SprogNr
Case 1:
s = "Eksempler på understøttede enheder:" & VbCrLfMac
s = s & "(Se manual for flere)" & VbCrLfMac & VbCrLfMac
s = s & "Længde: cm , AU, ly, pc" & VbCrLfMac
s = s & "Masse: u , G, mg" & VbCrLfMac
s = s & "Tid: min , Timer, dage, år" & VbCrLfMac
s = s & "Frekvens: Hz , Bq" & VbCrLfMac
s = s & "Energi: kWh , eV, aJ" & VbCrLfMac
s = s & "Tryk: Bar , mmHg, atm, torr" & VbCrLfMac
Case 3: s = "Ejemplos de unidades compatibles:" & VbCrLfMac
s = s & "(Para mayor información consulte el manual)" & VbCrLfMac & VbCrLfMac
s = s & "Longitud: cm , AU, ly, pc" & VbCrLfMac
s = s & "Masa: u , G, mg" & VbCrLfMac
s = s & "Frecuencia: Hz , Bq" & VbCrLfMac
s = s & "Energía: kWh , eV, aJ" & VbCrLfMac
s = s & "Presión: Bar , mmHg, atm, torr" & VbCrLfMac
Case 4: s = ""
Case Else:
s = "Examples of supported units:" & VbCrLfMac
s = s & "(See manual for more)" & VbCrLfMac & VbCrLfMac
s = s & "Length: cm , AU, ly, pc" & VbCrLfMac
s = s & "Mass: u , G, mg" & VbCrLfMac
s = s & "Frequency: Hz , Bq" & VbCrLfMac
s = s & "Energy: kWh , eV, aJ" & VbCrLfMac
s = s & "Pressure: Bar , mmHg, atm, torr" & VbCrLfMac
End Select
UnitExamples = s
End Property
Public Property Get TrigEquations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trignonometriske ligninger"
Case 3: s = "Ecuaciones trigonométricas"
Case 4: s = ""
Case Else: s = "Trigonometric equations"
End Select
TrigEquations = s
End Property
Public Property Get AllSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle løsninger"
Case 3: s = "Todas las soluciones"
Case 4: s = ""
Case Else: s = "All solutions"
End Select
AllSolutions = s
End Property
Public Property Get OnlyOneSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Kun 1 løsning"
Case 3: s = "Sólo una solución"
Case 4: s = ""
Case Else: s = "Only 1 solution"
End Select
OnlyOneSolution = s
End Property
Public Property Get AutoUpdate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Check automatisk for opdateringer"
Case 3: s = "Comprobar actualizaciones automáticamente"
Case 4: s = ""
Case Else: s = "Automatically check for updates"
End Select
AutoUpdate = s
End Property
Public Property Get GenerateShortcuts() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Reparér tastaturgenveje"
Case 3: s = "Crear atajos de teclado"
Case 4: s = ""
Case Else: s = "Repair Keyboard shortcuts"
End Select
GenerateShortcuts = s
End Property
Public Property Get RestartMaxima() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genstart Maxima"
Case 3: s = "Reiniciar Maxima"
Case 4: s = ""
Case Else: s = "Restart Maxima"
End Select
RestartMaxima = s
End Property
Public Property Get RestartWordMat() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genstart WordMat"
Case 3: s = "Reiniciar WordMat"
Case 4: s = ""
Case Else: s = "Restart WordMat"
End Select
RestartWordMat = s
End Property
Public Property Get CalculationCount() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Antal beregninger"
Case 3: s = "Número de cálculos"
Case 4: s = ""
Case Else: s = "No. of calculations"
End Select
CalculationCount = s
End Property
Public Property Get ResetFigurDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil figurdokument"
Case 3: s = "Restablecer documento figura"
Case 4: s = ""
Case Else: s = "Reset Figur doc"
End Select
ResetFigurDoc = s
End Property
Public Property Get ResetMathDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil Matdokument"
Case 3: s = "Restablecer documento matemático"
Case 4: s = ""
Case Else: s = "Reset math doc"
End Select
ResetMathDoc = s
End Property
Public Property Get ResetPhysicsDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil fysikdokument"
Case 3: s = "Restablecer documento físico"
Case 4: s = ""
Case Else: s = "Reset physics doc"
End Select
ResetPhysicsDoc = s
End Property
Public Property Get ResetChemistryDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil kemidokument"
Case 3: s = "Restablecer documento químico"
Case 4: s = ""
Case Else: s = "Reset chemistry doc"
End Select
ResetChemistryDoc = s
End Property
Public Property Get SignDepends() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Resultatet af beregningen afhænger af fortegnet på følgende"
Case 3: s = "El resultado del cálculo depende del signo de"
Case 4: s = ""
Case Else: s = "The result of the calculation depends on the sign of"
End Select
SignDepends = s
End Property
Public Property Get StateSign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Angiv fortegn"
Case 3: s = "Signo de estado"
Case 4: s = ""
Case Else: s = "State sign"
End Select
StateSign = s
End Property
Public Property Get DifferentialEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Differentialligning"
Case 3: s = "Ecuación Diferencial"
Case 4: s = ""
Case Else: s = "Differential equation"
End Select
DifferentialEquation = s
End Property
Public Property Get SolveDE() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsning af differentialligning"
Case 3: s = "Resolviendo ecuación diferencial"
Case 4: s = ""
Case Else: s = "Solving Differential equation"
End Select
SolveDE = s
End Property
Public Property Get IndepVar() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Uafh. var"
Case 3: s = "Variable independiente"
Case 4: s = ""
Case Else: s = "Indep. var"
End Select
IndepVar = s
End Property
Public Property Get DepVar() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Afh. var/funktion"
Case 3: s = "Variable Dependiente"
Case 4: s = ""
Case Else: s = "Dep. var/function"
End Select
DepVar = s
End Property
Public Property Get StartCond() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Startbetingelse"
Case 3: s = "Condición inicial"
Case 4: s = ""
Case Else: s = "Initial condition"
End Select
StartCond = s
End Property
Public Property Get MaximaError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Maxima fejlmeddelse"
Case 3: s = "Error de Maxima"
Case 4: s = ""
Case Else: s = "Maxima error"
End Select
MaximaError = s
End Property
Public Property Get MathError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Matematikfejl"
Case 3: s = "Error de matemáticas"
Case 4: s = ""
Case Else: s = "Math error"
End Select
MathError = s
End Property
Public Property Get SyntaxError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Syntaksfejl"
Case 3: s = "Error de sintaxis"
Case 4: s = ""
Case Else: s = "Syntax error"
End Select
SyntaxError = s
End Property
Public Property Get SyntaxErrorLong() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en syntaksfejl. Dvs. WordMat har ikke forstået dit input."
Case 3: s = "Error de sintaxis. WordMat no entiende su entrada"
Case 4: s = ""
Case Else: s = "Syntax error. WordMat did not understand your input."
End Select
SyntaxErrorLong = s
End Property
Public Property Get Missing() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der mangler en"
Case 3: s = "Está perdiendo una"
Case 4: s = ""
Case Else: s = "You're missing a"
End Select
Missing = s
End Property
Public Property Get TooMany() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er for mange"
Case 3: s = "Hay demasiados"
Case 4: s = ""
Case Else: s = "There is too many"
End Select
TooMany = s
End Property
Public Property Get LispError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en fejl i programafviklingen."
Case 3: s = "Ocurrió un error en el programa"
Case 4: s = ""
Case Else: s = "An error occured in the code."
End Select
LispError = s
End Property
Public Property Get DivisionByZero() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en division med 0."
Case 3: s = "División por cero"
Case 4: s = ""
Case Else: s = "Division by zero"
End Select
DivisionByZero = s
End Property
Public Property Get DefError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der skete en fejl. Det skyldes formentlig en fejl i en af dine definitioner." & VbCrLfMac & VbCrLfMac & "Måske kan du se ud af nedenstående hvor den er."
Case 3: s = "Ocurrió un error. El error, probablemente, lo causó una de sus definiciones." & VbCrLfMac & VbCrLfMac & "Tal vez usted puede detectar el error de la información abajo."
Case 4: s = ""
Case Else: s = "An errror occured. The error was probably caused by one of your definitions." & VbCrLfMac & VbCrLfMac & "Maybe you can spot the error from the info below."
End Select
DefError = s
End Property
Public Property Get DefError3() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Du har skrevet noget som WordMat ikke kan forstå."
Case 3: s = "Hay un error en una de sus definiciones! " & VbCrLfMac & " Vea la lista de definiciones a continuación y compruebe si puede localizar el error. También puede tratar de insertar el comando 'delete def' para depurar."
Case 4: s = ""
Case Else: s = "You have written something which WordMat does not understand."
End Select
DefError3 = s
End Property
Public Property Get DefError2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er en fejl i en af dine definitioner." & VbCrLfMac & "Definitionen vil blive markeret i Word."
Case 3: s = "Un error de sintaxis puede haber sido causado por un error en una de sus definiciones. Vea la lista de definiciones abajo y verifique si puede localizar el error. También puede tratar de insertar el comando 'delete def' para depurar."
Case 4: s = ""
Case Else: s = "There is an error in one of your definitions." & VbCrLfMac & "The definition will now be selected in Word."
End Select
DefError2 = s
End Property
Public Property Get IllegalSymbol() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er et ulovligt tegn der er placeret ved"
Case 3: s = "Hay un carácter ilegal que se encuentra en"
Case 4: s = ""
Case Else: s = "An illegal character has been placed by"
End Select
IllegalSymbol = s
End Property
Public Property Get SolveEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs Ligning"
Case 3: s = "Resolver ecuación"
Case 4: s = ""
Case Else: s = "Solve equation"
End Select
SolveEquation = s
End Property
Public Property Get ChooseVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Variabel"
Case 3: s = "Variable"
Case 4: s = ""
Case Else: s = "Variable"
End Select
ChooseVariable = s
End Property
Public Property Get WriteVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eller skriv selv:"
Case 3: s = "O establezca variable:"
Case 4: s = ""
Case Else: s = "Or state variable:"
End Select
WriteVariable = s
End Property
Public Property Get WriteVariables() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eller skriv selv adskilt ad komma"
Case 3: s = "O establezca variables separadas por coma"
Case 4: s = ""
Case Else: s = "Or state variables separated by comma"
End Select
WriteVariables = s
End Property
Public Property Get PresentDefs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gældende definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
PresentDefs = s
End Property
Public Property Get TempDefs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Midl. definitioner"
Case 3: s = "Definiciones temporales"
Case 4: s = ""
Case Else: s = "Temp. definitions"
End Select
TempDefs = s
End Property
Public Property Get UnitWarning() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Advarsel: Enheder er slået til." & VbCrLfMac & "Ligningsløsning påvirkes negativt." & VbCrLfMac & "Pas også på variable, der er enheder fx A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
Case 3: s = "Advertencia: Las unidades están encendidas. " & VbCrLfMac & "Cuidado con las variables que son unidades como: A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
Case 4: s = ""
Case Else: s = "Warning: Units are turned on." & VbCrLfMac & "Equation solving will be poorer" & VbCrLfMac & "Watch out for variables which are units like: A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
End Select
UnitWarning = s
End Property
Public Property Get SolveSystem() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs ligningssystem"
Case 3: s = "Resolver sistemas de ecuaciones"
Case 4: s = ""
Case Else: s = "Solve system of equations"
End Select
SolveSystem = s
End Property
Public Property Get ChooseVariables() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vælg variable"
Case 3: s = "Escoger variables"
Case 4: s = ""
Case Else: s = "Choose variables"
End Select
ChooseVariables = s
End Property
Public Property Get WordMatStarting() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat starter op"
Case 3: s = "WordMat está iniciando"
Case 4: s = ""
Case Else: s = "WordMat is starting"
End Select
WordMatStarting = s
End Property
Public Property Get PleaseWait() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent venligst"
Case 3: s = "Por favor, espere"
Case 4: s = ""
Case Else: s = "Please wait"
End Select
PleaseWait = s
End Property
Public Property Get TriangleSolver() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trekantsløser"
Case 3: s = "Solucionario de triángulos"
Case 4: s = ""
Case Else: s = "Triangle solver"
End Select
TriangleSolver = s
End Property
Public Property Get RightAngled() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Retvinklet"
Case 3: s = "Triángulos rectángulos"
Case 4: s = ""
Case Else: s = "Right angled"
End Select
RightAngled = s
End Property
Public Property Get AnyTriangle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vilkårlig"
Case 3: s = "Cualquier triángulo"
Case 4: s = ""
Case Else: s = "Any"
End Select
AnyTriangle = s
End Property
Public Property Get right() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ret"
Case 3: s = "Recto"
Case 4: s = ""
Case Else: s = "Right"
End Select
right = s
End Property
Public Property Get Naming() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Navngivning"
Case 3: s = "Denominación"
Case 4: s = ""
Case Else: s = "Naming"
End Select
Naming = s
End Property
Public Property Get Manuel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manuel"
Case 3: s = "Manuel"
Case 4: s = ""
Case Else: s = "Manuel"
End Select
Manuel = s
End Property
Public Property Get AngleNaming1() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vinkler store, sider små bogstaver"
Case 3: s = "Angulos - mayúsculas, lados - minúsculas"
Case 4: s = ""
Case Else: s = "Angles capital, sides lowercase letters"
End Select
AngleNaming1 = s
End Property
Public Property Get AngleNaming2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sider navngives efter hjørner AB"
Case 3: s = "Los lados se nombran a partir del ángulo AB"
Case 4: s = ""
Case Else: s = "Sides named by corners AB"
End Select
AngleNaming2 = s
End Property
Public Property Get InsertNumbers() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt tal på figuren i Word"
Case 3: s = "Inserte números en el triángulo en Word"
Case 4: s = ""
Case Else: s = "Insert numbers on the triangle in Word"
End Select
InsertNumbers = s
End Property
Public Property Get ShowCalculations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis mellemregninger i Word"
Case 3: s = "Mostrar cálculos en Word"
Case 4: s = ""
Case Else: s = "Show calculations in Word"
End Select
ShowCalculations = s
End Property
Public Property Get TriangleSolverExplanation1() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt de kendte sider og vinkler så beregnes de resterende sider og vinkler, og indsættes på figur i Word." & VbCrLfMac & VbCrLfMac & "Det er også muligt at ændre betegnelserne på figuren. bare klik på bogstaverne."
Case 3: s = "Insertar los lados y ángulos conocidos. los restantes lados y ángulos serán calculados e insertados en Word." & VbCrLfMac & VbCrLfMac & "Es posible cambiar los nombres de los lados y ángulos en el triángulo. Simplemente haga clic en las letras."
Case 4: s = ""
Case Else: s = "Insert the known sides and angles. The remaining sides and angles will then be calculated, and inserted in Word." & VbCrLfMac & VbCrLfMac & "It is possible to change the names of sides and angles on the triangle. Just click the letters."
End Select
TriangleSolverExplanation1 = s
End Property
Public Property Get TriangleSolverExplanation2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Figuren viser ikke de korrekte størrelser, men det gør den når den indsættes i Word."
Case 3: s = "La figura no muestra el tamaño correcto, pero lo hace cuando se inserta en Word."
Case 4: s = ""
Case Else: s = "The triangle is not to size. But a correctly scaled triangle will be inserted into Word."
End Select
TriangleSolverExplanation2 = s
End Property
Public Property Get TriangleSolverExplanation3() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat's trekantsløser anvendes med input: "
Case 3: s = "El solucionario de triángulos de WordMat se usa con la entrada:"
Case 4: s = ""
Case Else: s = "WordMat's trianglesolver is used with input: "
End Select
TriangleSolverExplanation3 = s
End Property
Public Property Get TS2Solutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Bemærk der er to mulige trekanter ud fra dine oplysninger"
Case 3: s = "Con su información se pueden construir dos triángulos"
Case 4: s = ""
Case Else: s = "Two triangles can be constructed from your information"
End Select
TS2Solutions = s
End Property
Public Property Get TS2Solutions2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "To Trekanter"
Case 3: s = "Dos triángulos"
Case 4: s = ""
Case Else: s = "Two triangles"
End Select
TS2Solutions2 = s
End Property
Public Property Get TSNoSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der findes ikke en trekant der kan konstrueres ud fra dine oplysninger."
Case 4: s = ""
Case Else: s = "No triangle can be found from your information"
End Select
TSNoSolution = s
End Property
Public Property Get TSMissingInfo() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er ikke nok information til at beregne de resterende sider og vinkler."
Case 3: s = "Se requiere más información para calcular los restantes lados y ángulos"
Case 4: s = ""
Case Else: s = "More information is needed to calculate the remaining sides and angles"
End Select
TSMissingInfo = s
End Property
Public Property Get TSInfoOK() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle sider og vinkler kan nu findes."
Case 3: s = "Se han calculado todos los lados y ángulos."
Case 4: s = ""
Case Else: s = "All sides and angles can now be calculated."
End Select
TSInfoOK = s
End Property
Public Property Get NumericSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Numerisk løsning"
Case 3: s = "Solución numérica"
Case 4: s = ""
Case Else: s = "Numeric solution"
End Select
NumericSolution = s
End Property
Public Property Get InsertSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt løsning(er)"
Case 3: s = "Insertar solución"
Case 4: s = ""
Case Else: s = "Insert solution(s)"
End Select
InsertSolutions = s
End Property
Public Property Get SeekGraphically() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grafisk løsning"
Case 3: s = "Soluciones gráficamente"
Case 4: s = ""
Case Else: s = "Graphical solutions"
End Select
SeekGraphically = s
End Property
Public Property Get InsertCASresult() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt omskrevet resultat"
Case 3: s = "Inserte resultado reducido"
Case 4: s = ""
Case Else: s = "Insert reduced result"
End Select
InsertCASresult = s
End Property
Public Property Get NoAnalyticResult() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen kunne desværre ikke løses analytisk"
Case 3: s = "No se puede resolver la ecuación analíticamente"
Case 4: s = ""
Case Else: s = "The equation could not be solved analytically"
End Select
NoAnalyticResult = s
End Property
Public Property Get FoundNumericSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OBS: Der blev anvendt numeriske metoder til at løse ligningen. Der kan muligvis findes flere løsninger."
Case 3: s = "Se han encontrado soluciones utilizando métodos numéricos, pero posiblemente hay más."
Case 4: s = ""
Case Else: s = "Numerical methods were applied when solving the equation. More solutions may exist."
End Select
FoundNumericSolutions = s
End Property
Public Property Get NoNumericSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat kunne ikke løse ligningen analytisk, men har angivet et omskrevet resultat der måske kan arbejdes videre med."
Case 3: s = "WordMat no podía resolver la ecuación analíticamente, pero ha establecido un resultado reducido que podría funcionar."
Case 4: s = ""
Case Else: s = "WordMat could not solve the equation analytically. A reduced result has been returned though."
End Select
NoNumericSolution = s
End Property
Public Property Get SearchingSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent der søges efter løsninger ..."
Case 3: s = "Espere, buscando soluciones ..."
Case 4: s = ""
Case Else: s = "Wait, searching for solutions ..."
End Select
SearchingSolutions = s
End Property
Public Property Get SearchingMoreSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent der søges efter flere løsninger ..."
Case 3: s = "Espere, buscando más soluciones ..."
Case 4: s = ""
Case Else: s = "Wait, searching for more solutions ..."
End Select
SearchingMoreSolutions = s
End Property
Public Property Get NoNumericSolution2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der blev ikke fundet nogen løsninger vha. numeriske metoder, men der findes måske løsninger som du kan søge efter grafisk."
Case 3: s = "No se encontró ninguna solución utilizando métodos numéricos, pero, tal vez, haya soluciones que se pueden buscar gráficamente."
Case 4: s = ""
Case Else: s = "No solutions where found using numerical method, but you could try searching graphically for solutions."
End Select
NoNumericSolution2 = s
End Property
Public Property Get TempSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Midl. indstillinger"
Case 3: s = "Configuración temporal"
Case 4: s = ""
Case Else: s = "Temp. settings"
End Select
TempSettings = s
End Property
Public Property Get AutoReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Automatisk reducering"
Case 3: s = "Reducción automática"
Case 4: s = ""
Case Else: s = "Auto reduce"
End Select
AutoReduce = s
End Property
Public Property Get AutoReduceExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Forsøg at reducere udtrykket til kortest mulig form vha. flere forskellige metoder."
Case 3: s = "Intenta reducir la expresión a la forma más breve posible utilizando diferentes métodos"
Case 4: s = ""
Case Else: s = "Tries to reduce the expression to shortest possible form using different methods."
End Select
AutoReduceExplain = s
End Property
Public Property Get FactorExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sæt uden for parantes" & VbCrLfMac & "Forsøger at omdanne udtrykket til produkter"
Case 3: s = "Intenta factorizar la expresión"
Case 4: s = ""
Case Else: s = "Tries to factor the expression"
End Select
FactorExplain = s
End Property
Public Property Get ExpandExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gang ind i paranteser." & VbCrLfMac & "Alle paranteser ganges ud."
Case 3: s = "Expande la expresión"
Case 4: s = ""
Case Else: s = "Expands the expression."
End Select
ExpandExplain = s
End Property
Public Property Get Rationalize() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omskriv til brøker"
Case 3: s = "Racionaliza"
Case 4: s = ""
Case Else: s = "Rationalize"
End Select
Rationalize = s
End Property
Public Property Get RationalizeExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle decimaltal omskrives til brøker."
Case 3: s = "Convierte todos los números decimales a fracciones"
Case 4: s = ""
Case Else: s = "Converts all decimal numbers to fractions"
End Select
RationalizeExplain = s
End Property
Public Property Get TrigReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Reducer trigonometrisk"
Case 3: s = "Reduce trigonometricamente"
Case 4: s = ""
Case Else: s = "Reduce trigonometrically"
End Select
TrigReduce = s
End Property
Public Property Get TrigReduceExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Udtrykket forsøges omskrevet vha. trigonometriske identiteter."
Case 3: s = "Reduce la expresión usando identidades trigonométricas"
Case 4: s = ""
Case Else: s = "Reduce the expression using trigonometric identities."
End Select
TrigReduceExplain = s
End Property
Public Property Get OKSame() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK, samme linje"
Case 3: s = "OK, la misma línea"
Case 4: s = ""
Case Else: s = "OK, same line"
End Select
OKSame = s
End Property
Public Property Get OKNew() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK, ny linje"
Case 3: s = "OK, una nueva línea"
Case 4: s = ""
Case Else: s = "OK, new line"
End Select
OKNew = s
End Property
Public Property Get PlotCaption() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graftegning mm. i planen"
Case 3: s = "Graficar en el plano"
Case 4: s = ""
Case Else: s = "2D plotting"
End Select
PlotCaption = s
End Property
Public Property Get ResetAll() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil alt"
Case 3: s = "Reiniciar todo"
Case 4: s = ""
Case Else: s = "Reset all"
End Select
ResetAll = s
End Property
Public Property Get AxisTitle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "aksetitel"
Case 3: s = "Título del eje"
Case 4: s = ""
Case Else: s = "axis title"
End Select
AxisTitle = s
End Property
Public Property Get LineType() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Linjetype"
Case 3: s = "Tipo de línea"
Case 4: s = ""
Case Else: s = "Line type"
End Select
LineType = s
End Property
Public Property Get EquationMissingError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "For at løse en ligning skal du først indtaste en ligning vha. indsæt ligning i Word. Lad cursoren stå i ligningen og tryk igen på løs ligning. Marker flere ligninger for at løse flere ligninger med flere ubekendte."
Case 3: s = "Para resolver una ecuación debe, primero, insertar una ecuación en Word. Ponga el cursor en la ecuación y presione 'Resolver ecuación(es)' otra vez. Marque más ecuaciones para resolver un sistema de ecuaciones."
Case 4: s = ""
Case Else: s = "To solve an equation you must first insert an equation into Word. Place the cursor in the equation and press 'Solve equation(s)' again. Mark more equations to solve a system of equations"
End Select
EquationMissingError = s
End Property
Public Property Get EquationMissingError2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ingen ligning valgt"
Case 3: s = "No hay ecuación seleccionada"
Case 4: s = ""
Case Else: s = "No equation selected"
End Select
EquationMissingError2 = s
End Property


''''' Any ''''''''''''''

Public Property Get Simplify() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Reducer"
Case 3: s = "Simplificar"
Case 4: s = ""
Case Else: s = "Simplify"
End Select
Simplify = s
End Property
Public Property Get Factor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Faktoriser"
Case 3: s = "Factorizar"
Case 4: s = ""
Case Else: s = "Factor"
End Select
Factor = s
End Property
Public Property Get Expand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Udvid"
Case 3: s = "Expandir"
Case 4: s = ""
Case Else: s = "Expand"
End Select
Expand = s
End Property
Public Property Get Help() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Hjælp"
Case 3: s = "Ayuda"
Case 4: s = ""
Case Else: s = "Help"
End Select
Help = s
End Property
Public Property Get Manual() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manual"
Case 3: s = "Manual"
Case 4: s = ""
Case Else: s = "Manual"
End Select
Manual = s
End Property
Public Property Get ManualDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manual som Worddokument"
Case 3: s = "Manual como documento Word"
Case 4: s = ""
Case Else: s = "Manual as Word document"
End Select
ManualDoc = s
End Property
Public Property Get ManualOnline() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Online hjemmeside-Manual"
Case 3: s = "Manual como página web on-line"
Case 4: s = ""
Case Else: s = "Online webpage Manual"
End Select
ManualOnline = s
End Property
Public Property Get About() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Om"
Case 3: s = "Acerca de"
Case 4: s = ""
Case Else: s = "About"
End Select
About = s
End Property
Public Property Get Update() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Opdater"
Case 3: s = "Actualizar"
Case 4: s = ""
Case Else: s = "Update"
End Select
Update = s
End Property
Public Property Get Shortcuts() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genveje"
Case 3: s = "Atajos"
Case 4: s = ""
Case Else: s = "Shortcuts"
End Select
Shortcuts = s
End Property
Public Property Get Clear() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil"
Case 3: s = "Borrar"
Case 4: s = ""
Case Else: s = "Clear"
End Select
Clear = s
End Property
Public Property Get Logarithm() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Logaritme"
Case 3: s = "Logaritmo"
Case 4: s = ""
Case Else: s = "Logarithm"
End Select
Logarithm = s
End Property
Public Property Get Language() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sprog"
Case 3: s = "Lenguaje"
Case 4: s = ""
Case Else: s = "Language"
End Select
Language = s
End Property
Public Property Get Sign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fortegn"
Case 3: s = "Signo"
Case 4: s = ""
Case Else: s = "Sign"
End Select
Sign = s
End Property
Public Property Get Positive() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Positiv"
Case 3: s = "Positivo"
Case 4: s = ""
Case Else: s = "Positive"
End Select
Positive = s
End Property
Public Property Get Negative() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Negativ"
Case 3: s = "Negativo"
Case 4: s = ""
Case Else: s = "Negative"
End Select
Negative = s
End Property
Public Property Get Zero() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nul"
Case 3: s = "Cero"
Case 4: s = ""
Case Else: s = "Zero"
End Select
Zero = s
End Property
Public Property Get Error() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fejl"
Case 3: s = "Error"
Case 4: s = ""
Case Else: s = "Error"
End Select
Error = s
End Property
Public Property Get Definitions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
Definitions = s
End Property
Public Property Get OK() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK"
Case 3: s = "OK"
Case 4: s = ""
Case Else: s = "OK"
End Select
OK = s
End Property
Public Property Get Cancel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Annuller"
Case 3: s = ""
Case 4: s = ""
Case Else: s = "Cancel"
End Select
Cancel = s
End Property
Public Property Get Wait() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent"
Case 3: s = "Cancelar"
Case 4: s = ""
Case Else: s = "Wait"
End Select
Wait = s
End Property
Public Property Get Activity() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Aktivitet"
Case 3: s = "Esperar"
Case 4: s = ""
Case Else: s = "Activity"
End Select
Activity = s
End Property
Public Property Get StopLabel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Stop"
Case 3: s = "Parar"
Case 4: s = ""
Case Else: s = "Stop"
End Select
StopLabel = s
End Property


Public Property Get EquationSolvedFor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen løses for "
Case 3: s = "La ecuación se resuelve para "
Case 4: s = ""
Case Else: s = "The equation is solved for "
End Select
EquationSolvedFor = s
End Property
Public Property Get EquationSolvedNumFor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen løses numerisk for "
Case 3: s = "La ecuación se resuelve numéricamente para "
Case 4: s = ""
Case Else: s = "The equation is solved numerically for "
End Select
EquationSolvedNumFor = s
End Property
Public Property Get ByCAS() As String
Dim s As String
Select Case SprogNr
Case 1: s = " vha. " & AppNavn & ". "
Case 3: s = "por" & AppNavn & ". "
Case 4: s = ""
Case Else: s = " by " & AppNavn & ". "
End Select
ByCAS = s
End Property
Public Property Get Functions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Forskrifter"
Case 3: s = "Funciones"
Case 4: s = ""
Case Else: s = "Functions"
End Select
Functions = s
End Property
Public Property Get Equation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligning"
Case 3: s = "Ecuación"
Case 4: s = ""
Case Else: s = "Equation"
End Select
Equation = s
End Property
Public Property Get Equations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligninger"
Case 3: s = "Ecuaciones"
Case 4: s = ""
Case Else: s = "Equations"
End Select
Equations = s
End Property
Public Property Get Points() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Punkter"
Case 3: s = "Puntos"
Case 4: s = ""
Case Else: s = "Points"
End Select
Points = s
End Property
Public Property Get Vectors() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vektorer"
Case 3: s = "Vectores"
Case 4: s = ""
Case Else: s = "Vectors"
End Select
Vectors = s
End Property
Public Property Get Title() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Titel"
Case 3: s = "Título"
Case 4: s = ""
Case Else: s = "Title"
End Select
Title = s
End Property
Public Property Get Reset() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil"
Case 3: s = "Reiniciar"
Case 4: s = ""
Case Else: s = "Reset"
End Select
Reset = s
End Property
Public Property Get RetryNum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Komplicerede beregninger der udføres eksakt kan tage meget lang tid. Vil du istedet prøve at regne numerisk?"
Case 3: s = "Cálculos complicados se realizan exactamente y puede tomar un tiempo mulargo. øPrefiere intentar numéricamente?"
Case 4: s = ""
Case Else: s = "Complex calculation done exactly can take a very long time. Do you want to retry numerically?"
End Select
RetryNum = s
End Property

Public Property Get dAsDiffChr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Brug alm. d som differential d"
Case 3: s = "Use la letra d como el diferencial d"
Case 4: s = ""
Case Else: s = "Use regular d as differential d"
End Select
dAsDiffChr = s
End Property

Public Property Get dAsDiffChrHelp() As String
Dim s As String
Select Case SprogNr
Case 1: s = "dy/dx notation kan normalt kun bruges vha. et specielt differential d (\dd) eller symbolet for partielt afledte (\partial). Hvis det ikke forventes at d/dx kan forveksles med et andet udtryk er denne indstilling en fordel idet det er nemmere at taste normalt d"
Case 3: s = "La notación dy/dx se usa normalmente representada por una diferencial especial d (\dd) o por el símbolo para las derivadas parciales (\partial). Si usted cree que no se confundirá con otras expresiones puede utilizar esta opción, ya que es más rápido escribir dy / dx usando d"
Case 4: s = ""
Case Else: s = "dy/dx notation can normally only be used with a special differential d (\dd) or the symbol for partial derivatives (\partial). If you dont expect d/dx to be confused with other expression, you can use this setting since it is faster to type dy/dx using regular d."
End Select
dAsDiffChrHelp = s
End Property

Public Property Get RegrUser() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Brugerdefineret regression"
Case 3: s = "Regresión definida por el usuario"
Case 4: s = ""
Case Else: s = "User defined regression"
End Select
RegrUser = s
End Property

Public Property Get RibSolveDEnum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs koblede differentialligning(er) numerisk"
Case 3: s = "Resuelve sistemas de ecuaciones diferenciales"
Case 4: s = ""
Case Else: s = "Solve system of differential equation(s) numerically"
End Select
RibSolveDEnum = s
End Property

Public Property Get ErrorGeneral() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der skete en fejl"
Case 3: s = "Ha ocurrido un error"
Case 4: s = ""
Case Else: s = "An error occured"
End Select
ErrorGeneral = s
End Property

Public Property Get Variable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Variabel"
Case 3: s = "Variable"
Case 4: s = ""
Case Else: s = "Variable"
End Select
Variable = s
End Property

Public Property Get Warning() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Advarsel"
Case 3: s = "Advertencia"
Case 4: s = ""
Case Else: s = "Warning"
End Select
Warning = s
End Property
