VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IXMLDOMElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Dummy for IXMLDOMElement
Implements IXMLDOMNode
Private m_tagname As String
Private m_value As String
Private m_children As Collection
Private m_attnames As Collection
Private m_attvalues As Collection
'
Public Sub setAttribute(Name As String, Value)
    m_attnames.Add Name
    m_attvalues.Add Value
End Sub
'
Public Function appendChild(Obj As IXMLDOMNode) As IXMLDOMNode
    m_children.Add Obj
    Set appendChild = Obj
End Function
Public Property Get tagName() As String
    tagName = m_tagname
End Property
Public Property Let tagName(rhs As String)
    m_tagname = rhs
End Property
'
Private Sub Class_Initialize()
    m_tagname = ""
    m_value = ""
    Set m_children = New Collection
    Set m_attnames = New Collection
    Set m_attvalues = New Collection
End Sub
'
Public Function IXMLDOMNode_nodeType() As Integer
    IXMLDOMNode_nodeType = 2
End Function
Public Property Get IXMLDOMNode_nodeValue() As String
    IXMLDOMNode_nodeValue = m_value
End Property
Public Property Let IXMLDOMNode_nodeValue(rhs As String)
    m_value = rhs
End Property
'
Public Function IXMLDOMNode_myChildren() As Collection
    Set IXMLDOMNode_myChildren = m_children
End Function
'
Public Function IXMLDOMNode_myAttNames() As Collection
    Set IXMLDOMNode_myAttNames = m_attnames
End Function
'
Public Function IXMLDOMNode_myAttValues() As Collection
    Set IXMLDOMNode_myAttValues = m_attvalues
End Function
'
Public Sub IXMLDOMNode_Save(dest)
    Dim s As String
    
    s = "<" & m_tagname
    
    Dim n As Integer
    Dim i As Integer
    n = m_attnames.Count
    
    For i = 1 To n
        s = s & " " & m_attnames(i) & "=""" & m_attvalues(i) & """"
    Next
    
    n = m_children.Count
    If n < 1 Then
        Print #dest, s & " />"
    Else
        Print #dest, s & ">"
        For i = 1 To n
            Dim node As IXMLDOMNode
            Set node = m_children(i)
            node.Save (dest)
        Next
        Print #dest, "</" & m_tagname & ">"
    End If
    
End Sub

