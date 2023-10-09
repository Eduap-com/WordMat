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
    Dim i As Integer
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
#If Mac Then
#Else
On Error Resume Next
Dim ctrl As Object
    For Each ctrl In Application.CommandBars("Equation Popup").Controls
        If ctrl.Tag = "cust" Then ctrl.Delete
    Next
#End If
End Sub
Private Sub BuildCustomMenu()
     
    Dim ctrl As CommandBarControl
    Dim btn As CommandBarControl
    Dim i As Integer
     
     'add a 'popup' control to the cell commandbar (menu)
    Set ctrl = Application.CommandBars("Cell").Controls.Add _
    (Type:=msoControlPopup, Before:=1)
    ctrl.Caption = "Insert Shape..."
     'add the submenus
    For i = 50 To 250 Step 50 'add a few menu items
        Set btn = ctrl.Controls.Add
        btn.Caption = i & " x " & (i / 2) 'give them a name
        btn.Tag = i 'we'll use the tag property to hold a value
        btn.OnAction = "InsertShape" 'the routine called by the control
    Next
     
End Sub
 
Private Sub DeleteCustomMenu()
     
    Dim ctrl As CommandBarControl
     
     'go thru all the cell commandbar controls and delete our menu item
    For Each ctrl In Application.CommandBars("Cell").Controls
        If ctrl.Caption = "Insert Shape..." Then ctrl.Delete
    Next
     
End Sub
 
Private Sub InsertShape()
     
    Dim t As Long
    Dim shp As Shape
     
     'get the tag property of the clicked control
    t = CLng(Application.CommandBars.ActionControl.Tag)
     
     'use the value of t and the active cell as size and position parameters
     'for adding a rectangle to the worksheet
'    Set shp = ActiveSheet.Shapes.AddShape _
    (msoShapeRectangle, ActiveCell.Left, ActiveCell.Top, t, t / 2)
     'do something with our shape
    Randomize 'make it a random color from the workbook
    shp.Fill.ForeColor.SchemeColor = Int((56 - 1 + 1) * Rnd + 1)
     
End Sub

Sub ShowcmdBars()
Dim ct As CommandBar

For Each ct In Application.CommandBars
    MsgBox ct.Name
Next

End Sub
Sub ShowRC()
Dim ct As CommandBar
'Dim b As CommandBarButton
Dim b As control
For Each b In Application.CommandBars("Equation Popup").Controls
    MsgBox b.Caption
Next

End Sub

