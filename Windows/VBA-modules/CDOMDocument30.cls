VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOMDocument30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Dummy for DOMDocument30 needed on Mac for GeoGebra files
Private m_children As Collection
'
Public Function createProcessingInstruction(target As String, Data As String) As IXMLDOMProcessingInstruction
    Dim inst As IXMLDOMProcessingInstruction
    Set inst = New IXMLDOMProcessingInstruction
    inst.target = target
    inst.Data = Data
    Set createProcessingInstruction = inst
End Function
Public Function appendChild(Obj As IXMLDOMNode) As IXMLDOMNode
    m_children.Add Obj
    Set appendChild = Obj
End Function
Public Function createElement(tagName As String) As IXMLDOMElement
    Dim elm As IXMLDOMElement
    Set elm = New IXMLDOMElement
    elm.tagName = tagName
    Set createElement = elm
End Function
Public Sub Save(Destination As Variant)
' saves the xml in utf-8 encoding
    On Error GoTo Err
    Dim n As Integer
    Dim i As Integer
        
    Dim fh As Integer
    fh = FreeFile
    Open Destination For Binary As #fh
    
    n = m_children.Count
    For i = 1 To n
        Dim node As IXMLDOMNode
        Set node = m_children(i)
        node.Save (fh)
    Next
    
    Close fh

    GoTo slut
Err:
    MsgBox "Error saving DOM documet " & Destination & " " & Err.Description, vbOKOnly, TT.Error
    On Error Resume Next
#If Mac Then
    Close fh
#End If
slut:
End Sub
Public Property Let preserveWhiteSpace(ByVal RHS As Boolean)
'
End Property
Public Property Let Async(ByVal RHS As Boolean)
'
End Property
Private Sub Class_Initialize()
    Set m_children = New Collection
End Sub


