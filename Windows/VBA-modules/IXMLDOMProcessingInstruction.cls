VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IXMLDOMProcessingInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Dummy for IXMLDOMProcessingInstruction
Option Explicit
Implements IXMLDOMNode
Private m_target As String
Private m_data As String
'
Public Property Let target(RHS As String)
    m_target = RHS
End Property
Public Property Let Data(RHS As String)
    m_data = RHS
End Property
'
Private Sub Class_Initialize()
    m_target = ""
    m_data = ""
End Sub
'
Public Function IXMLDOMNode_nodeType() As Integer
    IXMLDOMNode_nodeType = 1
End Function
Public Property Get IXMLDOMNode_nodeValue() As String
'
End Property
Public Property Let IXMLDOMNode_nodeValue(RHS As String)
'
End Property
'
Public Function IXMLDOMNode_myChildren() As Collection
    Set IXMLDOMNode_myChildren = Nothing
End Function
'
Public Function IXMLDOMNode_myAttNames() As Collection
    Set IXMLDOMNode_myAttNames = Nothing
End Function
'
Public Function IXMLDOMNode_myAttValues() As Collection
    Set IXMLDOMNode_myAttValues = Nothing
End Function
'
Public Sub IXMLDOMNode_Save(ByRef dest As Variant)
    Put #dest, , StringToUTF8Bytes("<?" & m_target & " " & m_data & "?>")
End Sub

