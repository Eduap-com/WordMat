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
    On Error Resume Next
    If AntalB > Antalberegninger Then
        Antalberegninger = AntalB
    End If
End Sub

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
        ' arr(1) contains version. If more fields are required, enough check
        PrepareMaxima
        On Error Resume Next
        UF2Dgraph.TextBox_definitioner.text = arr(2)
        UF2Dgraph.TextBox_titel.text = arr(3)
        UF2Dgraph.TextBox_xaksetitel.text = arr(4)
        UF2Dgraph.TextBox_yaksetitel.text = arr(5)
        
        UF2Dgraph.TextBox_xmin.text = arr(6)
        UF2Dgraph.TextBox_xmax.text = arr(7)
        UF2Dgraph.TextBox_ymin.text = arr(8)
        UF2Dgraph.TextBox_ymax.text = arr(9)
        
        UF2Dgraph.TextBox_ligning1.text = arr(10)
        UF2Dgraph.TextBox_var1.text = arr(11)
        UF2Dgraph.TextBox_xmin1.text = arr(12)
        UF2Dgraph.TextBox_xmax1.text = arr(13)
        UF2Dgraph.ComboBox_ligning1.ListIndex = CInt(arr(14))
        
        UF2Dgraph.TextBox_ligning2.text = arr(15)
        UF2Dgraph.TextBox_var2.text = arr(16)
        UF2Dgraph.TextBox_xmin2.text = arr(17)
        UF2Dgraph.TextBox_xmax2.text = arr(18)
        UF2Dgraph.ComboBox_ligning2.ListIndex = CInt(arr(19))
        
        UF2Dgraph.TextBox_ligning3.text = arr(20)
        UF2Dgraph.TextBox_var3.text = arr(21)
        UF2Dgraph.TextBox_xmin3.text = arr(22)
        UF2Dgraph.TextBox_xmax3.text = arr(23)
        UF2Dgraph.ComboBox_ligning3.ListIndex = CInt(arr(24))
        
        UF2Dgraph.TextBox_ligning4.text = arr(25)
        UF2Dgraph.TextBox_var4.text = arr(26)
        UF2Dgraph.TextBox_xmin4.text = arr(27)
        UF2Dgraph.TextBox_xmax4.text = arr(28)
        UF2Dgraph.ComboBox_ligning4.ListIndex = CInt(arr(29))
        
        UF2Dgraph.TextBox_ligning5.text = arr(30)
        UF2Dgraph.TextBox_var5.text = arr(31)
        UF2Dgraph.TextBox_xmin5.text = arr(32)
        UF2Dgraph.TextBox_xmax5.text = arr(33)
        UF2Dgraph.ComboBox_ligning5.ListIndex = CInt(arr(34))
        
        UF2Dgraph.TextBox_ligning6.text = arr(35)
        UF2Dgraph.TextBox_var6.text = arr(36)
        UF2Dgraph.TextBox_xmin6.text = arr(37)
        UF2Dgraph.TextBox_xmax6.text = arr(38)
        UF2Dgraph.ComboBox_ligning6.ListIndex = CInt(arr(39))
        
        UF2Dgraph.TextBox_lig1.text = arr(40)
        UF2Dgraph.TextBox_lig2.text = arr(41)
        UF2Dgraph.TextBox_Lig3.text = arr(42)
        
        UF2Dgraph.TextBox_parametric1x.text = arr(43)
        UF2Dgraph.TextBox_parametric1y.text = arr(44)
        UF2Dgraph.TextBox_tmin1.text = arr(45)
        UF2Dgraph.TextBox_tmax1.text = arr(46)
        UF2Dgraph.TextBox_parametric2x.text = arr(47)
        UF2Dgraph.TextBox_parametric2y.text = arr(48)
        UF2Dgraph.TextBox_tmin2.text = arr(49)
        UF2Dgraph.TextBox_tmax2.text = arr(50)
        UF2Dgraph.TextBox_parametric3x.text = arr(51)
        UF2Dgraph.TextBox_parametric3y.text = arr(51)
        UF2Dgraph.TextBox_tmin3.text = arr(53)
        UF2Dgraph.TextBox_tmax3.text = arr(54)
        
        UF2Dgraph.TextBox_punkter.text = arr(55)
        UF2Dgraph.TextBox_punkter2.text = arr(56)
        UF2Dgraph.TextBox_markerpunkter.text = arr(57)
        UF2Dgraph.CheckBox_pointsjoined.Value = CBool(arr(58))
        UF2Dgraph.CheckBox_pointsjoined2.Value = CBool(arr(59))
        UF2Dgraph.TextBox_pointsize.text = arr(60)
        UF2Dgraph.TextBox_pointsize2.text = arr(61)
        
        UF2Dgraph.TextBox_vektorer.text = arr(62)
        UF2Dgraph.TextBox_labels.text = arr(63)
        
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


