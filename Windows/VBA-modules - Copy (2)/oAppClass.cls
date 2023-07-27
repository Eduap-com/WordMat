VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oAppClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oApp As Word.Application
Attribute oApp.VB_VarHelpID = -1
Private Sub oapp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
Dim d As Document
On Error Resume Next
If AntalB > Antalberegninger Then
    Antalberegninger = AntalB
End If
If Application.Documents.Count <= 2 Then
LukTempDoc
MaxProc.CloseProcess
cxl.CloseExcel
#If Mac Then
#Else
    Set WebV = Nothing
    Set MaxProc = Nothing
    Set MaxProcUnit = Nothing
#End If
'    For Each d In Application.Documents ' får WOrd til altid at spørge om der ikke skal gemmes
'       If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
'           d.Close (False)
'       End If
'    Next
    SletRCMenu
End If
End Sub

'Private Sub oapp_NewDocument(ByVal Doc As Document)
'If AutoStart Then
'    PrepareMaxima
'End If
'End Sub

Private Sub oapp_WindowBeforeDoubleClick(ByVal Sel As Selection, Cancel As Boolean)
Dim s As String
Dim arr As Variant
On Error GoTo slut
If Sel.InlineShapes.Count = 0 Then GoTo slut
If Sel.InlineShapes(1).Type = wdInlineShapePicture Then
    Set UF2Dgraph = New UserForm2DGraph
    s = Sel.InlineShapes(1).AlternativeText
    arr = Split(s, "|")
    PicOpen = False
    If arr(0) = "WordMat" Then
        ' arr(1) indeholder version. Hvis der kommer flere felter kræves nok check
        PrepareMaxima
        On Error Resume Next
        UF2Dgraph.TextBox_definitioner.Text = arr(2)
        UF2Dgraph.TextBox_titel.Text = arr(3)
        UF2Dgraph.TextBox_xaksetitel.Text = arr(4)
        UF2Dgraph.TextBox_yaksetitel.Text = arr(5)
        
        UF2Dgraph.TextBox_xmin.Text = arr(6)
        UF2Dgraph.TextBox_xmax.Text = arr(7)
        UF2Dgraph.TextBox_ymin.Text = arr(8)
        UF2Dgraph.TextBox_ymax.Text = arr(9)
        
        UF2Dgraph.TextBox_ligning1.Text = arr(10)
        UF2Dgraph.TextBox_var1.Text = arr(11)
        UF2Dgraph.TextBox_xmin1.Text = arr(12)
        UF2Dgraph.TextBox_xmax1.Text = arr(13)
        UF2Dgraph.ComboBox_ligning1.ListIndex = CInt(arr(14))
        
        UF2Dgraph.TextBox_ligning2.Text = arr(15)
        UF2Dgraph.TextBox_var2.Text = arr(16)
        UF2Dgraph.TextBox_xmin2.Text = arr(17)
        UF2Dgraph.TextBox_xmax2.Text = arr(18)
        UF2Dgraph.ComboBox_ligning2.ListIndex = CInt(arr(19))
        
        UF2Dgraph.TextBox_ligning3.Text = arr(20)
        UF2Dgraph.TextBox_var3.Text = arr(21)
        UF2Dgraph.TextBox_xmin3.Text = arr(22)
        UF2Dgraph.TextBox_xmax3.Text = arr(23)
        UF2Dgraph.ComboBox_ligning3.ListIndex = CInt(arr(24))
        
        UF2Dgraph.TextBox_ligning4.Text = arr(25)
        UF2Dgraph.TextBox_var4.Text = arr(26)
        UF2Dgraph.TextBox_xmin4.Text = arr(27)
        UF2Dgraph.TextBox_xmax4.Text = arr(28)
        UF2Dgraph.ComboBox_ligning4.ListIndex = CInt(arr(29))
        
        UF2Dgraph.TextBox_ligning5.Text = arr(30)
        UF2Dgraph.TextBox_var5.Text = arr(31)
        UF2Dgraph.TextBox_xmin5.Text = arr(32)
        UF2Dgraph.TextBox_xmax5.Text = arr(33)
        UF2Dgraph.ComboBox_ligning5.ListIndex = CInt(arr(34))
        
        UF2Dgraph.TextBox_ligning6.Text = arr(35)
        UF2Dgraph.TextBox_var6.Text = arr(36)
        UF2Dgraph.TextBox_xmin6.Text = arr(37)
        UF2Dgraph.TextBox_xmax6.Text = arr(38)
        UF2Dgraph.ComboBox_ligning6.ListIndex = CInt(arr(39))
        
        UF2Dgraph.TextBox_lig1.Text = arr(40)
        UF2Dgraph.TextBox_lig2.Text = arr(41)
        UF2Dgraph.TextBox_Lig3.Text = arr(42)
        
        UF2Dgraph.TextBox_parametric1x.Text = arr(43)
        UF2Dgraph.TextBox_parametric1y.Text = arr(44)
        UF2Dgraph.TextBox_tmin1.Text = arr(45)
        UF2Dgraph.TextBox_tmax1.Text = arr(46)
        UF2Dgraph.TextBox_parametric2x.Text = arr(47)
        UF2Dgraph.TextBox_parametric2y.Text = arr(48)
        UF2Dgraph.TextBox_tmin2.Text = arr(49)
        UF2Dgraph.TextBox_tmax2.Text = arr(50)
        UF2Dgraph.TextBox_parametric3x.Text = arr(51)
        UF2Dgraph.TextBox_parametric3y.Text = arr(51)
        UF2Dgraph.TextBox_tmin3.Text = arr(53)
        UF2Dgraph.TextBox_tmax3.Text = arr(54)
        
        UF2Dgraph.TextBox_punkter.Text = arr(55)
        UF2Dgraph.TextBox_punkter2.Text = arr(56)
        UF2Dgraph.TextBox_markerpunkter.Text = arr(57)
        UF2Dgraph.CheckBox_pointsjoined.Value = CBool(arr(58))
        UF2Dgraph.CheckBox_pointsjoined2.Value = CBool(arr(59))
        UF2Dgraph.TextBox_pointsize.Text = arr(60)
        UF2Dgraph.TextBox_pointsize2.Text = arr(61)
        
        UF2Dgraph.TextBox_vektorer.Text = arr(62)
        UF2Dgraph.TextBox_labels.Text = arr(63)
        
        UF2Dgraph.CheckBox_gitter.Value = CBool(arr(64))
        UF2Dgraph.CheckBox_logx.Value = CBool(arr(65))
        UF2Dgraph.CheckBox_logy.Value = CBool(arr(66))
        UF2Dgraph.CheckBox_visforklaring.Value = CBool(arr(67))
                
        
        Sel.InlineShapes(1).Select
        UF2Dgraph.Show vbModeless
        PicOpen = True
        Cancel = True
    End If
slut:
End If

End Sub


