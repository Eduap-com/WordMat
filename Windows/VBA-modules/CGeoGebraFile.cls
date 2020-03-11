VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeoGebraFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Show3D As Boolean
#If Mac Then
    Private GeoXMLdom As New CDOMDocument30
    Private GeoGebraelement As IXMLDOMElement
    Private Constructionelement As IXMLDOMElement
#Else
    Private GeoXMLdom As New CDOMDocument30
    Private GeoGebraelement As IXMLDOMElement
    Private Constructionelement As IXMLDOMElement
#End If

Private Sub Class_Initialize()
    CreateXML
End Sub
Private Sub CreateXML()
Dim xmlstarttag As IXMLDOMProcessingInstruction
GeoXMLdom.preserveWhiteSpace = False
GeoXMLdom.Async = False

'Set xmlstarttag = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""windows-1252""")
'Set xmlstarttag = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")
' kun denne virker. ved specielle tegn
'Set xmlstarttag = GeoXMLdom.createProcessingInstruction("xml", "version=""1.0""")
Set xmlstarttag = GeoXMLdom.createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
GeoXMLdom.appendChild xmlstarttag

' Creates root element
Set GeoGebraelement = GeoXMLdom.createElement("geogebra")
GeoGebraelement.setAttribute "format", "5.0"
GeoXMLdom.appendChild GeoGebraelement

CreateGuiElement
CreateEuclidianViewElement
CreateConstructionElement
'CreatePoint "A", 1.4, 2
'CreateFunction "f", "x"
'CreateFunction "g", "2*x"


End Sub


Public Sub Save(filnavn As String)
    If Show3D Then CreateEuclidianView3DElement
    CreateKernelElement
    GeoXMLdom.Save (filnavn)
End Sub

Sub CreateGuiElement()
    Dim guielement As IXMLDOMElement
    Dim showelement As IXMLDOMElement
    Dim splitdividerelement As IXMLDOMElement
    Dim fontelement As IXMLDOMElement
    
    Set guielement = GeoXMLdom.createElement("gui")
    GeoGebraelement.appendChild guielement
    
    Set showelement = GeoXMLdom.createElement("show")
    guielement.appendChild showelement
    showelement.setAttribute "algebraView", "true"
    showelement.setAttribute "spreadsheetView", "false"
    showelement.setAttribute "auxiliaryObjects", "false"
    showelement.setAttribute "algebraInput", "true"
    showelement.setAttribute "cmdList", "true"
    
    Set splitdividerelement = GeoXMLdom.createElement("splitDivider")
    guielement.appendChild splitdividerelement
    splitdividerelement.setAttribute "loc", "650"
    splitdividerelement.setAttribute "locVertical", "400"
    splitdividerelement.setAttribute "loc2", "250"
    splitdividerelement.setAttribute "locVertical2", "300"
    splitdividerelement.setAttribute "horizontal", "true"
    
    Set fontelement = GeoXMLdom.createElement("font")
    guielement.appendChild fontelement
    fontelement.setAttribute "size", "12"
    
    
    If Not Show3D Then Exit Sub
    
    Dim perspectiveselement As IXMLDOMElement
    Set perspectiveselement = GeoXMLdom.createElement("perspectives")
    guielement.appendChild perspectiveselement
    
    Dim perspectiveelement As IXMLDOMElement
    Set perspectiveelement = GeoXMLdom.createElement("perspective")
    perspectiveselement.appendChild perspectiveelement
    
    Dim paneselement As IXMLDOMElement
    Set paneselement = GeoXMLdom.createElement("panes")
    perspectiveelement.appendChild paneselement
    perspectiveelement.setAttribute "id", "tmp"
    
    Dim pane1element As IXMLDOMElement
    Set pane1element = GeoXMLdom.createElement("pane")
    paneselement.appendChild pane1element
    pane1element.setAttribute "location", ""
    pane1element.setAttribute "divider", "0.25"
    pane1element.setAttribute "orientation", "1"

    Dim pane2element As IXMLDOMElement
    Set pane2element = GeoXMLdom.createElement("pane")
    paneselement.appendChild pane2element
    pane2element.setAttribute "location", "1"
    pane2element.setAttribute "divider", "0.35375"
    pane2element.setAttribute "orientation", "1"

'der mangler views, toolbar,input, dockbar elementer
'      <views>
'        <view id="4097" visible="false" inframe="true" stylebar="true" location="1,1,1,1" size="400" window="100,100,700,550" />
'        <view id="1" visible="true" inframe="false" stylebar="false" location="1,3" size="283" window="100,100,600,400" />
'        <view id="4" toolbar="0 || 2020 , 2021 , 2022 , 66 || 2001 , 2003 , 2002 , 2004 , 2005 || 2040 , 2041 , 2042 , 2044 , 2043" visible="false" inframe="false" stylebar="false" location="1,1" size="300" window="100,100,600,400" />
'        <view id="8" toolbar="1001 | 1002 | 1003  || 1005 | 1004 || 1006 | 1007 | 1010 || 1008 1009 || 66 68 || 6" visible="false" inframe="false" stylebar="false" location="1,3" size="300" window="100,100,600,400" />
'        <view id="512" toolbar="0 | 1 501 5 19 , 67 | 2 15 45 18 , 7 37 | 514 3 9 , 13 44 , 47 | 16 | 551 550 11 , 22 , 23 , 55 56 57 , 12 | 69 | 510 511 , 512 513 | 533 531 , 534 532 , 522 523 , 537 536 , 535 | 521 520 | 36 , 38 49 560 | 571 30 29 570 31 33 | 17 | 540 40 41 42 , 27 28 35 , 6 , 502" visible="true" inframe="false" stylebar="false" location="1,1" size="278" window="100,100,600,400" />
'        <view id="2" visible="true" inframe="false" stylebar="false" location="3" size="200" window="100,100,250,400" />
'        <view id="16" visible="false" inframe="false" stylebar="false" location="1" size="150" window="50,50,500,500" />
'        <view id="32" visible="false" inframe="false" stylebar="true" location="1" size="150" window="50,50,500,500" />
'        <view id="64" visible="false" inframe="true" stylebar="true" location="1" size="150" window="50,50,500,500" />
'        <view id="70" visible="false" inframe="true" stylebar="true" location="1" size="150" window="50,50,500,500" />
'      </views>
'      <toolbar show="true" items="0 39 | 1 501 67 , 5 19 , 72 | 2 15 45 , 18 65 , 7 37 | 4 3 8 9 , 13 44 , 58 , 47 | 16 51 64 , 70 | 10 34 53 11 , 24  20 22 , 21 23 | 55 56 57 , 12 | 36 46 , 38 49  50 , 71 | 30 29 54 32 31 33 | 17 26 62 73 , 14 68 | 25 52 60 61 | 40 41 42 , 27 28 35 , 6" position="1" help="false" />
'      <input show="true" cmd="true" top="false" />
'      <dockBar show="true" east="true" />
    
End Sub

Sub CreateEuclidianViewElement()
    Dim euclidianviewelement As IXMLDOMElement
    Dim sizeelement As IXMLDOMElement
    
    
    Set euclidianviewelement = GeoXMLdom.createElement("euclidianView")
    GeoGebraelement.appendChild euclidianviewelement

    Set sizeelement = GeoXMLdom.createElement("size")
    euclidianviewelement.appendChild sizeelement
    sizeelement.setAttribute "width", "628"
    sizeelement.setAttribute "height", "508"

    Dim coordSystemelement As IXMLDOMElement
    Set coordSystemelement = GeoXMLdom.createElement("coordSystem")
    euclidianviewelement.appendChild coordSystemelement
    coordSystemelement.setAttribute "xZero", "215.0"
    coordSystemelement.setAttribute "yZero", "315.0"
    coordSystemelement.setAttribute "scale", "50.0"
    coordSystemelement.setAttribute "yscale", "50.0"

    Dim evSettingselement As IXMLDOMElement
    Set evSettingselement = GeoXMLdom.createElement("evSettings")
    euclidianviewelement.appendChild evSettingselement
    evSettingselement.setAttribute "axes", "true"
    evSettingselement.setAttribute "grid", "false"
    evSettingselement.setAttribute "gridIsBold", "false"
    evSettingselement.setAttribute "pointCapturing", "3"
    evSettingselement.setAttribute "pointStyle", "0"
    evSettingselement.setAttribute "rightAngleStyle", "1"
    evSettingselement.setAttribute "checkboxSize", "13"
    evSettingselement.setAttribute "gridType", "0"

    Dim bgColorelement As IXMLDOMElement
    Set bgColorelement = GeoXMLdom.createElement("bgColor")
    euclidianviewelement.appendChild bgColorelement
    bgColorelement.setAttribute "r", "255"
    bgColorelement.setAttribute "g", "255"
    bgColorelement.setAttribute "b", "255"

    Dim axesColorelement As IXMLDOMElement
    Set axesColorelement = GeoXMLdom.createElement("axesColor")
    euclidianviewelement.appendChild axesColorelement
    axesColorelement.setAttribute "r", "0"
    axesColorelement.setAttribute "g", "0"
    axesColorelement.setAttribute "b", "0"

    Dim gridColorelement As IXMLDOMElement
    Set gridColorelement = GeoXMLdom.createElement("gridColor")
    euclidianviewelement.appendChild gridColorelement
    gridColorelement.setAttribute "r", "192"
    gridColorelement.setAttribute "g", "192"
    gridColorelement.setAttribute "b", "192"

    Dim linestyleelement As IXMLDOMElement
    Set linestyleelement = GeoXMLdom.createElement("lineStyle")
    euclidianviewelement.appendChild linestyleelement
    linestyleelement.setAttribute "axes", "1"
    linestyleelement.setAttribute "grid", "10"

    Dim axiselement As IXMLDOMElement
    Set axiselement = GeoXMLdom.createElement("axis")
    euclidianviewelement.appendChild axiselement
    axiselement.setAttribute "id", "0"
    axiselement.setAttribute "show", "true"
    axiselement.setAttribute "label", ""
    axiselement.setAttribute "unitLabel", ""
    axiselement.setAttribute "tickStyle", "1"
    axiselement.setAttribute "showNumbers", "true"

    Dim axis2element As IXMLDOMElement
    Set axis2element = GeoXMLdom.createElement("axis")
    euclidianviewelement.appendChild axis2element
    axis2element.setAttribute "id", "1"
    axis2element.setAttribute "show", "true"
    axis2element.setAttribute "label", ""
    axis2element.setAttribute "unitLabel", ""
    axis2element.setAttribute "tickStyle", "1"
    axis2element.setAttribute "showNumbers", "true"


End Sub
Sub CreateEuclidianView3DElement()
    Dim euclidianview3delement As IXMLDOMElement
    
    
    Set euclidianview3delement = GeoXMLdom.createElement("euclidianView3D")
    GeoGebraelement.appendChild euclidianview3delement
    
    Dim coordSystemelement As IXMLDOMElement
    Set coordSystemelement = GeoXMLdom.createElement("coordSystem")
    euclidianview3delement.appendChild coordSystemelement
    coordSystemelement.setAttribute "xZero", "0.0"
    coordSystemelement.setAttribute "yZero", "0.0"
    coordSystemelement.setAttribute "zZero", "-1.5"
    coordSystemelement.setAttribute "scale", "50.0"
    coordSystemelement.setAttribute "xAngle", "20.0"
    coordSystemelement.setAttribute "zAngle", "-276.0"

    Dim evSettingselement As IXMLDOMElement
    Set evSettingselement = GeoXMLdom.createElement("evSettings")
    euclidianview3delement.appendChild evSettingselement
    evSettingselement.setAttribute "axes", "true"
    evSettingselement.setAttribute "grid", "false"
    evSettingselement.setAttribute "gridIsBold", "false"
    evSettingselement.setAttribute "pointCapturing", "3"
'    evSettingselement.setAttribute "pointStyle", "0"
'    evSettingselement.setAttribute "rightAngleStyle", "1"
'    evSettingselement.setAttribute "checkboxSize", "13"
    evSettingselement.setAttribute "gridType", "0"

    Dim bgColorelement As IXMLDOMElement
    Set bgColorelement = GeoXMLdom.createElement("bgColor")
    euclidianview3delement.appendChild bgColorelement
    bgColorelement.setAttribute "r", "255"
    bgColorelement.setAttribute "g", "255"
    bgColorelement.setAttribute "b", "255"

'    Dim axesColorelement As IXMLDOMElement
'    Set axesColorelement = GeoXMLdom.CreateElement("axesColor")
'    euclidianviewelement.appendChild axesColorelement
'    axesColorelement.setAttribute "r", "0"
'    axesColorelement.setAttribute "g", "0"
'    axesColorelement.setAttribute "b", "0"

'    Dim gridColorelement As IXMLDOMElement
'    Set gridColorelement = GeoXMLdom.CreateElement("gridColor")
'    euclidianviewelement.appendChild gridColorelement
'    gridColorelement.setAttribute "r", "192"
'    gridColorelement.setAttribute "g", "192"
'    gridColorelement.setAttribute "b", "192"

'    Dim linestyleelement As IXMLDOMElement
'    Set linestyleelement = GeoXMLdom.CreateElement("lineStyle")
'    euclidianviewelement.appendChild linestyleelement
'    linestyleelement.setAttribute "axes", "1"
'    linestyleelement.setAttribute "grid", "10"

    Dim plateelement As IXMLDOMElement
    Set plateelement = GeoXMLdom.createElement("plate")
    euclidianview3delement.appendChild plateelement
    plateelement.setAttribute "show", "false"

    Dim clippingelement As IXMLDOMElement
    Set clippingelement = GeoXMLdom.createElement("clipping")
    euclidianview3delement.appendChild clippingelement
    clippingelement.setAttribute "use", "true"
    clippingelement.setAttribute "show", "true"
    clippingelement.setAttribute "size", "1"

    Dim projectionelement As IXMLDOMElement
    Set projectionelement = GeoXMLdom.createElement("projection")
    euclidianview3delement.appendChild projectionelement
    projectionelement.setAttribute "type", "0"

    Dim axiselement As IXMLDOMElement
    Set axiselement = GeoXMLdom.createElement("axis")
    euclidianview3delement.appendChild axiselement
    axiselement.setAttribute "id", "0"
    axiselement.setAttribute "show", "true"
    axiselement.setAttribute "label", ""
    axiselement.setAttribute "unitLabel", ""
    axiselement.setAttribute "tickStyle", "1"
    axiselement.setAttribute "showNumbers", "true"

    Dim axis2element As IXMLDOMElement
    Set axis2element = GeoXMLdom.createElement("axis")
    euclidianview3delement.appendChild axis2element
    axis2element.setAttribute "id", "1"
    axis2element.setAttribute "show", "true"
    axis2element.setAttribute "label", ""
    axis2element.setAttribute "unitLabel", ""
    axis2element.setAttribute "tickStyle", "1"
    axis2element.setAttribute "showNumbers", "true"

    Dim axis3element As IXMLDOMElement
    Set axis3element = GeoXMLdom.createElement("axis")
    euclidianview3delement.appendChild axis3element
    axis3element.setAttribute "id", "2"
    axis3element.setAttribute "show", "true"
    axis3element.setAttribute "label", ""
    axis3element.setAttribute "unitLabel", ""
    axis3element.setAttribute "tickStyle", "1"
    axis3element.setAttribute "showNumbers", "true"


End Sub
Sub CreateKernelElement()
    Dim kernelelement As IXMLDOMElement
    Set kernelelement = GeoXMLdom.createElement("kernel")
    GeoGebraelement.appendChild kernelelement

    Dim uses3delement As IXMLDOMElement
    Set uses3delement = GeoXMLdom.createElement("uses3D")
    kernelelement.appendChild uses3delement
    uses3delement.setAttribute "val", VBA.LCase(CStr(Show3D))

    Dim continuouselement As IXMLDOMElement
    Set continuouselement = GeoXMLdom.createElement("continuous")
    kernelelement.appendChild continuouselement
    continuouselement.setAttribute "val", "false"

    Dim decimalselement As IXMLDOMElement
    Set decimalselement = GeoXMLdom.createElement("decimals")
    kernelelement.appendChild decimalselement
    decimalselement.setAttribute "val", "2"

    Dim angleUnitelement As IXMLDOMElement
    Set angleUnitelement = GeoXMLdom.createElement("angleUnit")
    kernelelement.appendChild angleUnitelement
    If Radians Then
        angleUnitelement.setAttribute "val", "radiant"
    Else
        angleUnitelement.setAttribute "val", "degree"
    End If
    
    Dim algebraStyleelement As IXMLDOMElement
    Set algebraStyleelement = GeoXMLdom.createElement("algebraStyle")
    kernelelement.appendChild algebraStyleelement
    algebraStyleelement.setAttribute "val", "0"

    Dim coordStyleelement As IXMLDOMElement
    Set coordStyleelement = GeoXMLdom.createElement("coordStyle")
    kernelelement.appendChild coordStyleelement
    coordStyleelement.setAttribute "val", "0"

End Sub
Sub CreateConstructionElement()
    Set Constructionelement = GeoXMLdom.createElement("construction")
    GeoGebraelement.appendChild Constructionelement
    Constructionelement.setAttribute "title", ""
    Constructionelement.setAttribute "author", ""
    Constructionelement.setAttribute "date", ""

End Sub
Public Sub CreatePoint(pointname As String, xkoord As Double, ykoord As Double)
    Dim pointelement As IXMLDOMElement
    Set pointelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild pointelement
    pointelement.setAttribute "type", "point"
    pointelement.setAttribute "label", pointname

    Dim showelement As IXMLDOMElement
    Set showelement = GeoXMLdom.createElement("show")
    pointelement.appendChild showelement
    showelement.setAttribute "object", "true"
    showelement.setAttribute "label", "true"

    Dim objColorelement As IXMLDOMElement
    Set objColorelement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorelement
    objColorelement.setAttribute "r", "0"
    objColorelement.setAttribute "g", "0"
    objColorelement.setAttribute "b", "0"
    objColorelement.setAttribute "alpha", "0.0"

    Dim layerelement As IXMLDOMElement
    Set layerelement = GeoXMLdom.createElement("layer")
    pointelement.appendChild layerelement
    layerelement.setAttribute "val", "0"

    Dim labelModeelement As IXMLDOMElement
    Set labelModeelement = GeoXMLdom.createElement("labelMode")
    pointelement.appendChild labelModeelement
    labelModeelement.setAttribute "val", "0"

    Dim animationelement As IXMLDOMElement
    Set animationelement = GeoXMLdom.createElement("animation")
    pointelement.appendChild animationelement
    animationelement.setAttribute "step", "0.1"
    animationelement.setAttribute "speed", "1"
    animationelement.setAttribute "type", "0"
    animationelement.setAttribute "playing", "false"

    Dim coordselement As IXMLDOMElement
    Set coordselement = GeoXMLdom.createElement("coords")
    pointelement.appendChild coordselement
    coordselement.setAttribute "x", xkoord
    coordselement.setAttribute "y", ykoord
    coordselement.setAttribute "z", "1.0"

    Dim pointSizeelement As IXMLDOMElement
    Set pointSizeelement = GeoXMLdom.createElement("pointSize")
    pointelement.appendChild pointSizeelement
    pointSizeelement.setAttribute "val", "3"


End Sub
Public Sub CreateList(listname As String, List As String)


    Dim expressionelement As IXMLDOMElement
    Set expressionelement = GeoXMLdom.createElement("expression")
    Constructionelement.appendChild expressionelement
    expressionelement.setAttribute "label", listname
    expressionelement.setAttribute "exp", List

    Dim listelement As IXMLDOMElement
    Set listelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild listelement
    listelement.setAttribute "type", "list"
    listelement.setAttribute "label", listname

    Dim linestyleelement As IXMLDOMElement
    Set linestyleelement = GeoXMLdom.createElement("lineStyle")
    listelement.appendChild linestyleelement
    linestyleelement.setAttribute "type", "0"
    linestyleelement.setAttribute "object", "true"
    
    Dim showelement As IXMLDOMElement
    Set showelement = GeoXMLdom.createElement("show")
    listelement.appendChild showelement
    showelement.setAttribute "object", "true"
    showelement.setAttribute "label", "true"

    Dim objColorelement As IXMLDOMElement
    Set objColorelement = GeoXMLdom.createElement("objColor")
    listelement.appendChild objColorelement
    objColorelement.setAttribute "r", "0"
    objColorelement.setAttribute "g", "110"
    objColorelement.setAttribute "b", "0"
    objColorelement.setAttribute "alpha", "0.0"

    Dim layerelement As IXMLDOMElement
    Set layerelement = GeoXMLdom.createElement("layer")
    listelement.appendChild layerelement
    layerelement.setAttribute "val", "0"

    Dim labelModeelement As IXMLDOMElement
    Set labelModeelement = GeoXMLdom.createElement("labelMode")
    listelement.appendChild labelModeelement
    labelModeelement.setAttribute "val", "0"

    Dim animationelement As IXMLDOMElement
    Set animationelement = GeoXMLdom.createElement("animation")
    listelement.appendChild animationelement
    animationelement.setAttribute "step", "0.1"
    animationelement.setAttribute "speed", "1"
    animationelement.setAttribute "type", "0"
    animationelement.setAttribute "playing", "false"

    Dim pointSizeelement As IXMLDOMElement
    Set pointSizeelement = GeoXMLdom.createElement("pointSize")
    listelement.appendChild pointSizeelement
    pointSizeelement.setAttribute "val", "3"

    Dim pointStyleelement As IXMLDOMElement
    Set pointStyleelement = GeoXMLdom.createElement("pointStyle")
    listelement.appendChild pointStyleelement
    pointStyleelement.setAttribute "val", "-1"


End Sub
Public Sub CreateFunction(funkname As String, Expr As String, ulighed As Boolean, Optional ConvertMaxima As Boolean = True)
    Dim pos As Integer
    
    If InStr(Expr, "matrix") > 0 Then GoTo Slut ' matrix/vektor ikke understøttet endnu
    
    pos = InStr(funkname, "(")
    If pos > 1 Then
        funkname = Left(funkname, pos - 1)
    End If
    
    Expr = ConvertToGeoGebraSyntax(Expr, ConvertMaxima)
    
    Dim expressionelement As IXMLDOMElement
    Set expressionelement = GeoXMLdom.createElement("expression")
    Constructionelement.appendChild expressionelement
'    expressionelement.setAttribute "type", "point"
    expressionelement.setAttribute "label", funkname
'    expressionelement.setAttribute "exp", "f(x)=2*x"
'    expressionelement.setAttribute "exp", "2*x"
    expressionelement.setAttribute "exp", Expr
    
    
    Dim pointelement As IXMLDOMElement
    Set pointelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild pointelement
    pointelement.setAttribute "type", "function"
    pointelement.setAttribute "label", funkname

    Dim showelement As IXMLDOMElement
    Set showelement = GeoXMLdom.createElement("show")
    pointelement.appendChild showelement
    showelement.setAttribute "object", "true"
    showelement.setAttribute "label", "true"

    Dim objColorelement As IXMLDOMElement
    Set objColorelement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorelement
    objColorelement.setAttribute "r", "0"
    objColorelement.setAttribute "g", "0"
    If ulighed Then
        objColorelement.setAttribute "b", "255"
        objColorelement.setAttribute "alpha", "0.25"
    Else
        objColorelement.setAttribute "b", "0"
        objColorelement.setAttribute "alpha", "0.0"
    End If

    Dim layerelement As IXMLDOMElement
    Set layerelement = GeoXMLdom.createElement("layer")
    pointelement.appendChild layerelement
    layerelement.setAttribute "val", "0"

    Dim labelModeelement As IXMLDOMElement
    Set labelModeelement = GeoXMLdom.createElement("labelMode")
    pointelement.appendChild labelModeelement
    labelModeelement.setAttribute "val", "0"

    Dim animationelement As IXMLDOMElement
    Set animationelement = GeoXMLdom.createElement("animation")
    pointelement.appendChild animationelement
    animationelement.setAttribute "step", "0.1"
    animationelement.setAttribute "speed", "1"
    animationelement.setAttribute "type", "0"
    animationelement.setAttribute "playing", "false"

    Dim pointSizeelement As IXMLDOMElement
    Set pointSizeelement = GeoXMLdom.createElement("lineStyle")
    pointelement.appendChild pointSizeelement
    pointSizeelement.setAttribute "thickness", "2"
    pointSizeelement.setAttribute "type", "0"
    
Slut:
End Sub

'Sub CreateElement()
'    Dim zelement As IXMLDOMElement
'    Set zelement = GeoXMLdom.CreateElement("z")
'    pointelement.appendChild zelement
'    zelement.setAttribute "", ""

'End Sub

Function ConvertToGeoGebraSyntax(text As String, Optional ConvertMaxima As Boolean = True)
   Dim p As Integer, p2 As Integer, arr() As String, p3 As Integer, sp As Integer, ep As Integer
   Dim ea As ExpressionAnalyser, s As String, t As String, gexpr As String, i As Integer, n As Integer

   text = Replace(text, "log", "lg")
   If ConvertMaxima Then
      text = omax.CodeForMaxima(text)
      'græske bogstaver
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
      text = Replace(text, "eta", VBA.ChrW(951)) ' skal være sidst da eta indgår i andre
      text = Replace(text, "increment", VBA.ChrW(8710))  ' specielt delta increment
      text = Replace(text, "Symhalf", VBA.ChrW(189)) ' ½
      text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
      text = Replace(text, "<=", VBA.ChrW(8804))  '
      text = Replace(text, ">=", VBA.ChrW(8805)) '
      text = Replace(text, "CVinkelO", VBA.ChrW(8736)) '
      text = Replace(text, "CVinkel", VBA.ChrW(8736)) '
      text = Replace(text, "Symangle", VBA.ChrW(8736))  '
    
      'Else
      '    text = Replace(text, "log", "ln")
   End If
   
   
   text = Replace(text, "%e", "e")
   text = Replace(text, "%pi", "pi")
   text = Replace(text, "%", "")
   text = Replace(text, "[", "{")
   text = Replace(text, "]", "}")
   text = Replace(text, "matrix", "")


      p = InStr(text, "if")
      p2 = InStr(text, "then")
      If p > 0 And p2 > 0 Then
         sp = p
         Set ea = New ExpressionAnalyser
         ea.SetNormalBrackets
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
         text = Replace(text, " and ", " âˆ§ ") '&&
         text = Replace(text, " or ", " âˆ¨ ") '||
      End If

   ConvertToGeoGebraSyntax = text
End Function


