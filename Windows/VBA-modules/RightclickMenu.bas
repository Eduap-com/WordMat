Attribute VB_Name = "RightclickMenu"
Option Explicit
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'
'     'remove our custom menu before we leave
'    Run ("DeleteCustomMenu")
'
'End Sub
 
'Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'
'    Run ("DeleteCustomMenu") 'remove possible duplicates
'    Run ("BuildCustomMenu") 'build new menu
'
'End Sub
 '### code for the ThisWorkbook code sheet - END
 
Sub LavRCMenu()
'    CustomizationContext = ActiveDocument.AttachedTemplate
#If Mac Then ' giver ikke fejl, men der kommer ikke noget i menuen
#Else
    Dim cmdb As CommandBar
    Dim but As CommandBarControl
On Error GoTo slut
    SletRCMenu ' sikrer at der ikke oprettes dobbelt
    
    Set cmdb = Application.CommandBars("Equation Popup")
'    Set ctrl = Application.CommandBars("Equation Popup").Controls.Add _
'    (Type:=msoControlPopup, Before:=1)
    If cmdb Is Nothing Then GoTo slut
    
    Set but = cmdb.Controls.Add(Type:=msoControlButton)
    If but Is Nothing Then GoTo slut
    With but
        .Caption = Sprog.RibBeregn '"Beregn"
        .begingroup = True
        .Tag = "cust"
        .TooltipText = Sprog.A(396)
        .FaceId = 50 ' lommeregner
        .OnAction = "beregn"
    End With
        
    Set but = Application.CommandBars("Equation Popup").Controls.Add _
    (Type:=msoControlButton)
    With but
        .Caption = Sprog.RibSolve ' "Løs ligning(er)"
        .Tag = "cust"
        .TooltipText = Sprog.A(397)  '"Løser ligning"
        .FaceId = 26 ' kvadratrod a
        .OnAction = "MaximaSolve"
    End With
        
    Set but = Application.CommandBars("Equation Popup").Controls.Add _
    (Type:=msoControlButton)
    With but
        .Caption = Sprog.RibShowGraph ' "Vis graf"
        .Tag = "cust"
        .TooltipText = Sprog.RibShowGraph
        .FaceId = 42 ' kvadratrod a
        .OnAction = "Plot2DGraph"
    End With
        
'        .FaceId = 251 ' stort lig med
'        .FaceId = 212 ' Geometri
'        .FaceId = 17 ' Diagram
'        .FaceId = 477 ' integrale
'        .FaceId = 42 ' graf
slut:
#End If
    End Sub
Public Sub SletRCMenu()
    Exit Sub ' RCmenu bruges ikke
#If Mac Then
#Else
On Error Resume Next
Dim Ctrl As Object
    For Each Ctrl In Application.CommandBars("Equation Popup").Controls
        If Ctrl.Tag = "cust" Then Ctrl.Delete
    Next
#End If
End Sub
