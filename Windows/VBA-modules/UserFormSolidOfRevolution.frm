VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSolidOfRevolution 
   Caption         =   "Omdrejningslegeme"
   ClientHeight    =   3885
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   7020
   OleObjectBlob   =   "UserFormSolidOfRevolution.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSolidOfRevolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This forms provides input to GnuPlot for plotting a solid of revolution

Private Sub CommandButton_ok_Click()
    Dim funk As String
    Dim funk2 As String
    Dim dx As Double
    Dim xmin As Double
    Dim xmax As Double
    Dim Sxmin As String
    Dim Sxmax As String
    Dim grafobj As String
    On Error GoTo Fejl
    
    Label_vent.visible = True
    Sxmin = Replace(TextBox_xmin.Text, ",", ".")
    Sxmax = Replace(TextBox_xmax.Text, ",", ".")
    
    xmin = val(Sxmin)
    xmax = val(Sxmax)
    funk = TextBox_forskrift.Text
    funk = omax.CodeForMaxima(funk)
    funk2 = TextBox_forskrift2.Text
    funk2 = omax.CodeForMaxima(funk2)
    dx = (xmax - xmin) / 5
    omax.MaximaInputStreng = omax.MaximaInputStreng & "fomd(x):=" & funk & "$"
    If Len(funk2) > 0 Then
        omax.MaximaInputStreng = omax.MaximaInputStreng & "fomd2(x):=" & funk2 & "$"
    End If
    omax.MaximaInputStreng = omax.MaximaInputStreng & "fomdm:1.2*max(abs(fomd(" & Sxmin & ")),abs(fomd(" & Replace(ConvertNumberToString(xmin + dx), ",", ".") & ")),abs(fomd(" & Replace(ConvertNumberToString(xmin + 2 * dx), ",", ".") & ")),abs(fomd(" & Replace(ConvertNumberToString(xmin + 3 * dx), ",", ".") & ")),abs(fomd(" & Replace(ConvertNumberToString(xmin + 4 * dx), ",", ".") & ")),abs(fomd(" & Replace(ConvertNumberToString(xmax), ",", ".") & ")))$"
'    omax.MaximaInputStreng = omax.MaximaInputStreng & "fomdm:1.2*max(fomd(" & TextBox_xmin.text & "),fomd(" & dx & "),fomd(" & TextBox_xmax.text & "))$"
    
'    funk = "sqrt(fomd(x)^2 - y^2)"
'    funk = funk & ",-(" & funk & ")"
        
If Not CheckBox_linjer.Value Then
    grafobj = "palette=[""#55FF55"",""#5555FF"",""#222222""],enhanced3d=[x-z/10-y/10, x, y, z]," & grafobj
End If
If ComboBox_kvalitet.ListIndex = 0 Then 'super
    grafobj = "x_voxel=18,y_voxel=18,z_voxel=18," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 1 Then 'very high
    grafobj = "x_voxel=15,y_voxel=15,z_voxel=15," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 2 Then ' high
    grafobj = "x_voxel=12,y_voxel=12,z_voxel=12," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 4 Then 'low
    grafobj = "x_voxel=5,y_voxel=5,z_voxel=5," & grafobj
End If
If CheckBox_grid.Value Then
    grafobj = grafobj & "xaxis_width = 2,xaxis_color = blue,xaxis_type  = solid,xaxis=true,yaxis_width = 2,yaxis_color = orange,yaxis_type=solid,yaxis=true,zaxis_width = 2,zaxis_color = orange,zaxis_type=solid,zaxis=true,user_preamble = ""set xyplane at 0"","
End If

grafobj = grafobj & "user_preamble = ""set xyplane at 0"",colorbox=false,surface_hide = true,"
'grafobj = grafobj & "user_preamble = ""set pm3d at s depthorder;set xyplane at 0"",colorbox=false,surface_hide = false,"
    
    
'    Call omax.Plot3D(funk, "x", Replace(TextBox_xmin.text, ",", "."), Replace(TextBox_xmax.text, ",", "."), "y", "-fomdm", "fomdm", "", "", CheckBox_linjer.value, gridlines)
    grafobj = grafobj & "implicit(z^2+y^2=fomd(x)^2,x," & Sxmin & "," & Sxmax & ",y,-fomdm,fomdm,z,-fomdm,fomdm)"
    If Len(funk2) > 0 Then
        grafobj = grafobj & ",implicit(z^2+y^2=fomd2(x)^2,x," & Sxmin & "," & Sxmax & ",y,-fomdm,fomdm,z,-fomdm,fomdm)"
    End If
    
    omax.Draw3D grafobj, 3
#If Mac Then
    ShowPreviewMac
#Else
    Me.hide
#End If
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    Label_vent.visible = False
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = Sprog.A(194)
    ComboBox_kvalitet.AddItem Sprog.A(185)
    ComboBox_kvalitet.AddItem Sprog.A(184)
    ComboBox_kvalitet.AddItem Sprog.A(183)
    ComboBox_kvalitet.AddItem Sprog.A(182)
    ComboBox_kvalitet.AddItem Sprog.A(181)
    ComboBox_kvalitet.ListIndex = 3
    TextBox_forskrift.SetFocus
    
    Label1.Caption = Sprog.A(186)
    Label_vent.Caption = Sprog.A(187)
    Label5.Caption = Sprog.A(188)
    CheckBox_linjer.Caption = Sprog.A(189)
    CheckBox_grid.Caption = Sprog.A(190)
    Label_quality.Caption = Sprog.A(191)
#If Mac Then
    Label7.Caption = ""
    Label9.Caption = ""
#Else
    Label9.Caption = Sprog.A(192)
    Label7.Caption = Sprog.A(193)
#End If

End Sub
Sub ShowPreviewMac()
#If Mac Then
    RunScript "OpenPreview", GetTempDir() & "WordMatGraf.pdf"
#End If
End Sub

