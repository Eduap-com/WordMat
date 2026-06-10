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

Public Sub Initialize()
    CreateXML
End Sub
Private Sub CreateXML()
Dim xmlstarttag As IXMLDOMProcessingInstruction
GeoXMLdom.preserveWhiteSpace = False
GeoXMLdom.Async = False

'Set xmlstarttag = GeoXMLdom.createProcessingInstruction("xml", "version=""1.0"" encoding=""windows-1252""")
'Set xmlstarttag = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")
' kun denne virker. ved specielle tegn
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


Public Sub Save(FilNavn As String)
    If Show3D Then CreateEuclidianView3DElement
    CreateKernelElement
        
    GeoXMLdom.Save FilNavn
End Sub

Sub CreateGuiElement()
    
    Dim GuiElement As IXMLDOMElement
    Set GuiElement = GeoXMLdom.createElement("gui")
    GeoGebraelement.appendChild GuiElement
    
    Dim WindowElement As IXMLDOMElement
    Set WindowElement = GeoXMLdom.createElement("window")
    GuiElement.appendChild WindowElement
    WindowElement.setAttribute "width", "1600"
    WindowElement.setAttribute "height", "1000"
    
'    Dim showElement As IXMLDOMElement
'    Set showelement = GeoXMLdom.createElement("show") ' legacy
'    guielement.appendChild showelement
'    showelement.setAttribute "algebraView", "true"
'    showelement.setAttribute "spreadsheetView", "false"
'    showelement.setAttribute "auxiliaryObjects", "false"
'    showelement.setAttribute "algebraInput", "true"
'    showelement.setAttribute "cmdList", "true"
    
If Show3D Then
    Dim splitdividerelement As IXMLDOMElement
    Set splitdividerelement = GeoXMLdom.createElement("splitDivider")
    GuiElement.appendChild splitdividerelement
    splitdividerelement.setAttribute "loc", "650"
    splitdividerelement.setAttribute "locVertical", "400"
    splitdividerelement.setAttribute "loc2", "250"
    splitdividerelement.setAttribute "locVertical2", "300"
    splitdividerelement.setAttribute "horizontal", "true"

End If
    
    
    Dim FontElement As IXMLDOMElement
    Set FontElement = GeoXMLdom.createElement("font")
    GuiElement.appendChild FontElement
    FontElement.setAttribute "size", "14"

    Dim LabelingStyleElement As IXMLDOMElement
    Set LabelingStyleElement = GeoXMLdom.createElement("labelingStyle")
    GuiElement.appendChild LabelingStyleElement
    LabelingStyleElement.setAttribute "val", "0"

    
    Dim perspectiveselement As IXMLDOMElement
    Set perspectiveselement = GeoXMLdom.createElement("perspectives")
    GuiElement.appendChild perspectiveselement
    
    Dim perspectiveelement As IXMLDOMElement
    Set perspectiveelement = GeoXMLdom.createElement("perspective")
    perspectiveselement.appendChild perspectiveelement
    perspectiveelement.setAttribute "id", "tmp"
    
    Dim paneselement As IXMLDOMElement
    Set paneselement = GeoXMLdom.createElement("panes")
    perspectiveelement.appendChild paneselement
    
    Dim pane1element As IXMLDOMElement
    Set pane1element = GeoXMLdom.createElement("pane")
    paneselement.appendChild pane1element
    pane1element.setAttribute "location", ""
    pane1element.setAttribute "divider", "0.2"
    pane1element.setAttribute "orientation", "1"

'    Dim pane2element As IXMLDOMElement
'    Set pane2element = GeoXMLdom.createElement("pane")
'    paneselement.appendChild pane2element
'    pane2element.setAttribute "location", "1"
'    pane2element.setAttribute "divider", "0.35375"
'    pane2element.setAttribute "orientation", "1"

    Dim viewselement As IXMLDOMElement
    Set viewselement = GeoXMLdom.createElement("views")
    perspectiveelement.appendChild viewselement

    Dim viewelement1 As IXMLDOMElement
    Set viewelement1 = GeoXMLdom.createElement("view")
    viewselement.appendChild viewelement1
    viewelement1.setAttribute "id", "2"
    viewelement1.setAttribute "visible", "true"
    viewelement1.setAttribute "inframe", "false"
    viewelement1.setAttribute "stylebar", "false"
    viewelement1.setAttribute "location", "3"
    viewelement1.setAttribute "size", "200" '331
    viewelement1.setAttribute "tab", "ALGEBRA"
    viewelement1.setAttribute "window", "100,100,600,400" '0,0,200,400

    Dim viewelement2 As IXMLDOMElement
If Show3D Then
    Set viewelement2 = GeoXMLdom.createElement("view")
    viewselement.appendChild viewelement2
    viewelement2.setAttribute "id", "512"
    viewelement2.setAttribute "visible", "true"
    viewelement2.setAttribute "inframe", "false"
    viewelement2.setAttribute "stylebar", "false"
    viewelement2.setAttribute "location", "1" '1
    viewelement2.setAttribute "size", "1111" ' 546
    viewelement2.setAttribute "window", "100,100,600,400" ' '50,50,500,500
    viewelement2.setAttribute "toolbar", "0 | 1 501 5 19 , 67 | 2 15 45 18 , 7 37 | 514 3 9 , 13 44 , 47 | 16 51 | 551 550 11 , 20 22 21 23 , 55 56 57 , 12 | 69 | 510 511 , 512 513 | 533 531 , 534 532 , 522 523 , 537 536 , 535 , 538 | 521 520 | 36 , 38 49 560 | 571 30 29 570 31 33 | 17 | 540 40 41 42 , 27 28 35 , 6 , 502"
Else
    Set viewelement2 = GeoXMLdom.createElement("view")
    viewselement.appendChild viewelement2
    viewelement2.setAttribute "id", "1"
    viewelement2.setAttribute "visible", "true"
    viewelement2.setAttribute "inframe", "false"
    viewelement2.setAttribute "stylebar", "false"
    viewelement2.setAttribute "location", "1"
    viewelement2.setAttribute "size", "568"
    viewelement2.setAttribute "window", "100,100,600,400"
End If
    
    Dim toolbarelement As IXMLDOMElement
    Set toolbarelement = GeoXMLdom.createElement("toolbar")
    perspectiveelement.appendChild toolbarelement
    toolbarelement.setAttribute "show", "true"
    toolbarelement.setAttribute "items", "0 39 73 62 | 1 501 67 , 5 19 , 72 75 76 | 2 15 45 , 18 65 , 7 37 | 4 3 8 9 , 13 44 , 58 , 47 | 16 51 64 , 70 | 10 34 53 11 , 24 20 22 , 21 23 | 55 56 57 , 12 | 36 46 , 38 49 50 , 71 14 68 | 30 29 54 32 31 33 | 25 17 26 60 52 61 | 40 41 42 , 27 28 35 , 6"
    toolbarelement.setAttribute "position", "1"
    toolbarelement.setAttribute "help", "false"
    
    Dim inputelement As IXMLDOMElement
    Set inputelement = GeoXMLdom.createElement("input")
    perspectiveelement.appendChild inputelement
    inputelement.setAttribute "show", "true"
    inputelement.setAttribute "cmd", "true"
    inputelement.setAttribute "top", "algebra"
    
    Dim dockelement As IXMLDOMElement
    Set dockelement = GeoXMLdom.createElement("dock")
    perspectiveelement.appendChild dockelement
    dockelement.setAttribute "show", "false"
    dockelement.setAttribute "east", "false"
    
'der mangler views, toolbar,input, dockbar elementer
'<perspectives>
'<perspective id="tmp">
'<panes>
'<pane location="" divider="0.3685968819599109" orientation="1"/>
'</panes>

'<views>
'<view id="1" visible="false" inframe="false" stylebar="false" location="3,1" size="323" window="0,0,400,400"/>
'<view id="4" toolbar="0 || 2020 , 2021 , 2022 , 66 || 2001 , 2003 , 2002 , 2004 , 2005 || 2040 , 2041 , 2042 , 2044 , 2043" visible="false" inframe="false" stylebar="false" location="1,1" size="200" window="0,0,400,400"/>
'<view id="2" visible="true" inframe="false" stylebar="false" location="3" size="331" tab="ALGEBRA" window="0,0,200,400"/>
'<view id="8" visible="false" inframe="false" stylebar="false" location="1" size="150" window="50,50,500,500"/>
'<view id="16" visible="false" inframe="false" stylebar="false" location="1" size="150" window="50,50,500,500"/>
'<view id="32" visible="false" inframe="false" stylebar="true" location="1" size="150" window="50,50,500,500"/>
'<view id="64" visible="false" inframe="false" stylebar="false" location="1" size="150" window="50,50,500,500"/>
'<view id="4097" visible="false" inframe="false" stylebar="false" location="1" size="150" window="50,50,500,500"/>
'<view id="70" visible="false" inframe="false" stylebar="true" location="1" size="150" window="50,50,500,500"/>
'<view id="512" visible="true" inframe="false" stylebar="false" location="1" size="546" window="50,50,500,500"/>
'</views>
'<toolbar show="true" items="0 39 73 62 | 1 501 67 , 5 19 , 72 75 76 | 2 15 45 , 18 65 , 7 37 | 4 3 8 9 , 13 44 , 58 , 47 | 16 51 64 , 70 | 10 34 53 11 , 24 20 22 , 21 23 | 55 56 57 , 12 | 36 46 , 38 49 50 , 71 14 68 | 30 29 54 32 31 33 | 25 17 26 60 52 61 | 40 41 42 , 27 28 35 , 6" position="1" help="false"/>
'<input show="true" cmd="true" top="algebra"/>
'<dockBar show="false" east="false"/>
'</perspective>
'</perspectives>
'<labelingStyle val="3"/>
'<font size="12"/>
    
End Sub

Sub CreateEuclidianViewElement()
    Dim EuclidianviewElement As IXMLDOMElement
    Set EuclidianviewElement = GeoXMLdom.createElement("euclidianView")
    GeoGebraelement.appendChild EuclidianviewElement

    Dim viewNoElement As IXMLDOMElement
    Set viewNoElement = GeoXMLdom.createElement("viewNumber")
    EuclidianviewElement.appendChild viewNoElement
    viewNoElement.setAttribute "viewNo", "1"

    Dim sizeelement As IXMLDOMElement
    Set sizeelement = GeoXMLdom.createElement("size")
    EuclidianviewElement.appendChild sizeelement
    sizeelement.setAttribute "width", "568" '628
    sizeelement.setAttribute "height", "427" '427

    Dim coordSystemelement As IXMLDOMElement
    Set coordSystemelement = GeoXMLdom.createElement("coordSystem")
    EuclidianviewElement.appendChild coordSystemelement
    coordSystemelement.setAttribute "xZero", "640.0"
    coordSystemelement.setAttribute "yZero", "400.0"
    coordSystemelement.setAttribute "scale", "50.0"
    coordSystemelement.setAttribute "yscale", "50.0"

    Dim evSettingselement As IXMLDOMElement
    Set evSettingselement = GeoXMLdom.createElement("evSettings")
    EuclidianviewElement.appendChild evSettingselement
    evSettingselement.setAttribute "axes", "true"
    evSettingselement.setAttribute "grid", "true"
    evSettingselement.setAttribute "gridIsBold", "false"
    evSettingselement.setAttribute "pointCapturing", "3"
    evSettingselement.setAttribute "pointStyle", "0"
    evSettingselement.setAttribute "rightAngleStyle", "1"
    evSettingselement.setAttribute "checkboxSize", "26"
    evSettingselement.setAttribute "gridType", "3"

    Dim bgColorelement As IXMLDOMElement
    Set bgColorelement = GeoXMLdom.createElement("bgColor")
    EuclidianviewElement.appendChild bgColorelement
    bgColorelement.setAttribute "r", "255"
    bgColorelement.setAttribute "g", "255"
    bgColorelement.setAttribute "b", "255"

    Dim axesColorelement As IXMLDOMElement
    Set axesColorelement = GeoXMLdom.createElement("axesColor")
    EuclidianviewElement.appendChild axesColorelement
    axesColorelement.setAttribute "r", "28" '0 0 0
    axesColorelement.setAttribute "g", "28"
    axesColorelement.setAttribute "b", "31"

    Dim gridColorelement As IXMLDOMElement
    Set gridColorelement = GeoXMLdom.createElement("gridColor")
    EuclidianviewElement.appendChild gridColorelement
    gridColorelement.setAttribute "r", "180" ' 192 192 192
    gridColorelement.setAttribute "g", "179"
    gridColorelement.setAttribute "b", "186"

    Dim lineStyleElement As IXMLDOMElement
    Set lineStyleElement = GeoXMLdom.createElement("lineStyle")
    EuclidianviewElement.appendChild lineStyleElement
    lineStyleElement.setAttribute "axes", "1"
    lineStyleElement.setAttribute "grid", "0" '10

    Dim axiselement As IXMLDOMElement
    Set axiselement = GeoXMLdom.createElement("axis")
    EuclidianviewElement.appendChild axiselement
    axiselement.setAttribute "id", "0"
    axiselement.setAttribute "show", "true"
    axiselement.setAttribute "label", "x"
    axiselement.setAttribute "unitLabel", ""
    axiselement.setAttribute "tickStyle", "1"
    axiselement.setAttribute "showNumbers", "true"

    Dim axis2element As IXMLDOMElement
    Set axis2element = GeoXMLdom.createElement("axis")
    EuclidianviewElement.appendChild axis2element
    axis2element.setAttribute "id", "1"
    axis2element.setAttribute "show", "true"
    axis2element.setAttribute "label", "y"
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
    uses3delement.setAttribute "val", VBA.LCase$(CStr(Show3D))

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

    Dim showElement As IXMLDOMElement
    Set showElement = GeoXMLdom.createElement("show")
    pointelement.appendChild showElement
    showElement.setAttribute "object", "true"
    showElement.setAttribute "label", "true"

    Dim objColorElement As IXMLDOMElement
    Set objColorElement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorElement
    objColorElement.setAttribute "r", "0"
    objColorElement.setAttribute "g", "0"
    objColorElement.setAttribute "b", "0"
    objColorElement.setAttribute "alpha", "0.0"

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
    Dim expressionElement As IXMLDOMElement
    Set expressionElement = GeoXMLdom.createElement("expression")
    Constructionelement.appendChild expressionElement
    expressionElement.setAttribute "label", listname
    expressionElement.setAttribute "exp", List

    Dim listelement As IXMLDOMElement
    Set listelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild listelement
    listelement.setAttribute "type", "list"
    listelement.setAttribute "label", listname

    Dim lineStyleElement As IXMLDOMElement
    Set lineStyleElement = GeoXMLdom.createElement("lineStyle")
    listelement.appendChild lineStyleElement
    lineStyleElement.setAttribute "type", "0"
    lineStyleElement.setAttribute "object", "true"
    
    Dim showElement As IXMLDOMElement
    Set showElement = GeoXMLdom.createElement("show")
    listelement.appendChild showElement
    showElement.setAttribute "object", "true"
    showElement.setAttribute "label", "true"

    Dim objColorElement As IXMLDOMElement
    Set objColorElement = GeoXMLdom.createElement("objColor")
    listelement.appendChild objColorElement
    objColorElement.setAttribute "r", "0"
    objColorElement.setAttribute "g", "110"
    objColorElement.setAttribute "b", "0"
    objColorElement.setAttribute "alpha", "0.0"

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
    
    If InStr(Expr, "matrix") > 0 Then GoTo slut ' matrix/vektor not supported yet
    
    pos = InStr(funkname, "(")
    If pos > 1 Then
        funkname = Left$(funkname, pos - 1)
    End If
    
    Expr = ConvertToGeogebraSyntax(Expr, ConvertMaxima)
    ConvertToGeoGebra5Syntax Expr

    Dim expressionElement As IXMLDOMElement
    Set expressionElement = GeoXMLdom.createElement("expression")
    Constructionelement.appendChild expressionElement
    expressionElement.setAttribute "label", funkname
'    expressionelement.setAttribute "exp", "f(x)=2*x"
'    expressionelement.setAttribute "exp", "2*x"
    expressionElement.setAttribute "exp", Expr
    expressionElement.setAttribute "type", "function"
    
    
    Dim pointelement As IXMLDOMElement
    Set pointelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild pointelement
    pointelement.setAttribute "type", "function"
    pointelement.setAttribute "label", funkname

    Dim showElement As IXMLDOMElement
    Set showElement = GeoXMLdom.createElement("show")
    pointelement.appendChild showElement
    showElement.setAttribute "object", "true"
    showElement.setAttribute "label", "true"

    Dim objColorElement As IXMLDOMElement
    Set objColorElement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorElement
    objColorElement.setAttribute "r", "0"
    objColorElement.setAttribute "g", "0"
    If ulighed Then
        objColorElement.setAttribute "b", "255"
        objColorElement.setAttribute "alpha", "0.25"
    Else
        objColorElement.setAttribute "b", "0"
        objColorElement.setAttribute "alpha", "0.0"
    End If

    Dim layerelement As IXMLDOMElement
    Set layerelement = GeoXMLdom.createElement("layer")
    pointelement.appendChild layerelement
    layerelement.setAttribute "val", "0"

    Dim labelModeelement As IXMLDOMElement
    Set labelModeelement = GeoXMLdom.createElement("labelMode")
    pointelement.appendChild labelModeelement
    labelModeelement.setAttribute "val", "0"

    Dim fixedElement As IXMLDOMElement
    Set fixedElement = GeoXMLdom.createElement("fixed")
    pointelement.appendChild fixedElement
    fixedElement.setAttribute "val", "true"

    Dim lineStyleElement As IXMLDOMElement
    Set lineStyleElement = GeoXMLdom.createElement("lineStyle")
    pointelement.appendChild lineStyleElement
    lineStyleElement.setAttribute "thickness", "5"
    lineStyleElement.setAttribute "type", "0"
    lineStyleElement.setAttribute "typeHidden", "1"
    lineStyleElement.setAttribute "opacity", "178"

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
    
slut:
End Sub
Public Sub CreateEquation(EquationName As String, Equation As String, ulighed As Boolean, Optional ConvertMaxima As Boolean = True)
    
    If InStr(Equation, "matrix") > 0 Then GoTo slut ' matrix/vektor not supported yet
        
    Equation = ConvertToGeogebraSyntax(Equation, ConvertMaxima)
    ConvertToGeoGebra5Syntax Equation

    Dim expressionElement As IXMLDOMElement
    Set expressionElement = GeoXMLdom.createElement("expression")
    Constructionelement.appendChild expressionElement
    expressionElement.setAttribute "label", EquationName
    expressionElement.setAttribute "exp", Equation
    expressionElement.setAttribute "type", "line"
    
    
    Dim pointelement As IXMLDOMElement
    Set pointelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild pointelement
    pointelement.setAttribute "type", "function"
    pointelement.setAttribute "label", EquationName

    Dim showElement As IXMLDOMElement
    Set showElement = GeoXMLdom.createElement("show")
    pointelement.appendChild showElement
    showElement.setAttribute "object", "true"
    showElement.setAttribute "label", "true"

    Dim objColorElement As IXMLDOMElement
    Set objColorElement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorElement
    objColorElement.setAttribute "r", "20"
    objColorElement.setAttribute "g", "20"
    If ulighed Then
        objColorElement.setAttribute "b", "255"
        objColorElement.setAttribute "alpha", "0.25"
    Else
        objColorElement.setAttribute "b", "20"
        If Show3D Then
            objColorElement.setAttribute "alpha", "0.5"
        Else
            objColorElement.setAttribute "alpha", "0.0"
        End If
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
    
slut:
End Sub
Public Sub CreateVector(VectorName As String, VectorExpr As String, ulighed As Boolean, Optional ConvertMaxima As Boolean = True)
            
    VectorExpr = ConvertToGeogebraSyntax(VectorExpr, ConvertMaxima)
    ConvertToGeoGebra5Syntax VectorExpr


    Dim commandelement As IXMLDOMElement
    Set commandelement = GeoXMLdom.createElement("command")
    Constructionelement.appendChild commandelement
    commandelement.setAttribute "name", "Vector"
    
    Dim inputelement As IXMLDOMElement
    Set inputelement = GeoXMLdom.createElement("input")
    inputelement.setAttribute "a0", "(0,0)"
    inputelement.setAttribute "a1", VectorExpr
    commandelement.appendChild inputelement
    
    Dim outputelement As IXMLDOMElement
    Set outputelement = GeoXMLdom.createElement("output")
    outputelement.setAttribute "a0", VectorName
    commandelement.appendChild outputelement
    
    
    
    Dim pointelement As IXMLDOMElement
    Set pointelement = GeoXMLdom.createElement("element")
    Constructionelement.appendChild pointelement
    pointelement.setAttribute "type", "Vector"
    pointelement.setAttribute "label", VectorName

    Dim showElement As IXMLDOMElement
    Set showElement = GeoXMLdom.createElement("show")
    pointelement.appendChild showElement
    showElement.setAttribute "object", "true"
    showElement.setAttribute "label", "true"

    Dim objColorElement As IXMLDOMElement
    Set objColorElement = GeoXMLdom.createElement("objColor")
    pointelement.appendChild objColorElement
    objColorElement.setAttribute "r", "0"
    objColorElement.setAttribute "g", "0"
    If ulighed Then
        objColorElement.setAttribute "b", "255"
        objColorElement.setAttribute "alpha", "0.25"
    Else
        objColorElement.setAttribute "b", "0"
        objColorElement.setAttribute "alpha", "0.0"
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
End Sub


'Sub CreateElement()
'    Dim zelement As IXMLDOMElement
'    Set zelement = GeoXMLdom.CreateElement("z")
'    pointelement.appendChild zelement
'    zelement.setAttribute "", ""

'End Sub

Function ConvertToGeogebraSyntaxFile(ByVal text As String, Optional ConvertMaxima As Boolean = True) As String
   ConvertToGeogebraSyntaxFile = ConvertToGeogebraSyntax(text, ConvertMaxima, False)
End Function
Sub ConvertToGeoGebra5Syntax(ByRef Expr As String)
' takes a a string that has already been converted using ConvertToGeogebraSyntax, and makes changes needed for GeoGebra 5
    If InStrB(Expr, ChrW$(8804)) <> 0 Then Expr = Replace(Expr, ChrW$(8804), "<=")
    If InStrB(Expr, ChrW$(8805)) <> 0 Then Expr = Replace(Expr, ChrW$(8805), ">=")
    If InStrB(Expr, ChrW$(960)) <> 0 Then Expr = Replace(Expr, ChrW$(960), "pi")
'    If InStrB(Expr, ChrW$(956)) <> 0 Then Expr = Replace(Expr, ChrW$(956), ChrW$(&HCEBC)) ' mu
'    Expr = Replace(Expr, ChrW(&H3BC), ChrW(&HB5))
End Sub
