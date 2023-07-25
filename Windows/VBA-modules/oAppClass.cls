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
Dim Arr As Variant
On Error GoTo slut
If Sel.InlineShapes.Count = 0 Then GoTo slut
If Sel.InlineShapes(1).Type = wdInlineShapePicture Then
    Set UF2Dgraph = New UserForm2DGraph
    s = Sel.InlineShapes(1).AlternativeText
    Arr = Split(s, "|")
    PicOpen = False
    If Arr(0) = "WordMat" Then
        ' arr(1) indeholder version. Hvis der kommer flere felter kræves nok check
        PrepareMaxima
        On Error Resume Next
        UF2Dgraph.TextBox_definitioner.Text = Arr(2)
        UF2Dgraph.TextBox_titel.Text = Arr(3)
        UF2Dgraph.TextBox_xaksetitel.Text = Arr(4)
        UF2Dgraph.TextBox_yaksetitel.Text = Arr(5)
        
        UF2Dgraph.TextBox_xmin.Text = Arr(6)
        UF2Dgraph.TextBox_xmax.Text = Arr(7)
        UF2Dgraph.TextBox_ymin.Text = Arr(8)
        UF2Dgraph.TextBox_ymax.Text = Arr(9)
        
        UF2Dgraph.TextBox_ligning1.Text = Arr(10)
        UF2Dgraph.TextBox_var1.Text = Arr(11)
        UF2Dgraph.TextBox_xmin1.Text = Arr(12)
        UF2Dgraph.TextBox_xmax1.Text = Arr(13)
        UF2Dgraph.ComboBox_ligning1.ListIndex = CInt(Arr(14))
        
        UF2Dgraph.TextBox_ligning2.Text = Arr(15)
        UF2Dgraph.TextBox_var2.Text = Arr(16)
        UF2Dgraph.TextBox_xmin2.Text = Arr(17)
        UF2Dgraph.TextBox_xmax2.Text = Arr(18)
        UF2Dgraph.ComboBox_ligning2.ListIndex = CInt(Arr(19))
        
        UF2Dgraph.TextBox_ligning3.Text = Arr(20)
        UF2Dgraph.TextBox_var3.Text = Arr(21)
        UF2Dgraph.TextBox_xmin3.Text = Arr(22)
        UF2Dgraph.TextBox_xmax3.Text = Arr(23)
        UF2Dgraph.ComboBox_ligning3.ListIndex = CInt(Arr(24))
        
        UF2Dgraph.TextBox_ligning4.Text = Arr(25)
        UF2Dgraph.TextBox_var4.Text = Arr(26)
        UF2Dgraph.TextBox_xmin4.Text = Arr(27)
        UF2Dgraph.TextBox_xmax4.Text = Arr(28)
        UF2Dgraph.ComboBox_ligning4.ListIndex = CInt(Arr(29))
        
        UF2Dgraph.TextBox_ligning5.Text = Arr(30)
        UF2Dgraph.TextBox_var5.Text = Arr(31)
        UF2Dgraph.TextBox_xmin5.Text = Arr(32)
        UF2Dgraph.TextBox_xmax5.Text = Arr(33)
        UF2Dgraph.ComboBox_ligning5.ListIndex = CInt(Arr(34))
        
        UF2Dgraph.TextBox_ligning6.Text = Arr(35)
        UF2Dgraph.TextBox_var6.Text = Arr(36)
        UF2Dgraph.TextBox_xmin6.Text = Arr(37)
        UF2Dgraph.TextBox_xmax6.Text = Arr(38)
        UF2Dgraph.ComboBox_ligning6.ListIndex = CInt(Arr(39))
        
        UF2Dgraph.TextBox_lig1.Text = Arr(40)
        UF2Dgraph.TextBox_lig2.Text = Arr(41)
        UF2Dgraph.TextBox_Lig3.Text = Arr(42)
        
        UF2Dgraph.TextBox_parametric1x.Text = Arr(43)
        UF2Dgraph.TextBox_parametric1y.Text = Arr(44)
        UF2Dgraph.TextBox_tmin1.Text = Arr(45)
        UF2Dgraph.TextBox_tmax1.Text = Arr(46)
        UF2Dgraph.TextBox_parametric2x.Text = Arr(47)
        UF2Dgraph.TextBox_parametric2y.Text = Arr(48)
        UF2Dgraph.TextBox_tmin2.Text = Arr(49)
        UF2Dgraph.TextBox_tmax2.Text = Arr(50)
        UF2Dgraph.TextBox_parametric3x.Text = Arr(51)
        UF2Dgraph.TextBox_parametric3y.Text = Arr(51)
        UF2Dgraph.TextBox_tmin3.Text = Arr(53)
        UF2Dgraph.TextBox_tmax3.Text = Arr(54)
        
        UF2Dgraph.TextBox_punkter.Text = Arr(55)
        UF2Dgraph.TextBox_punkter2.Text = Arr(56)
        UF2Dgraph.TextBox_markerpunkter.Text = Arr(57)
        UF2Dgraph.CheckBox_pointsjoined.Value = CBool(Arr(58))
        UF2Dgraph.CheckBox_pointsjoined2.Value = CBool(Arr(59))
        UF2Dgraph.TextBox_pointsize.Text = Arr(60)
        UF2Dgraph.TextBox_pointsize2.Text = Arr(61)
        
        UF2Dgraph.TextBox_vektorer.Text = Arr(62)
        UF2Dgraph.TextBox_labels.Text = Arr(63)
        
        UF2Dgraph.CheckBox_gitter.Value = CBool(Arr(64))
        UF2Dgraph.CheckBox_logx.Value = CBool(Arr(65))
        UF2Dgraph.CheckBox_logy.Value = CBool(Arr(66))
        UF2Dgraph.CheckBox_visforklaring.Value = CBool(Arr(67))
                
        
        Sel.InlineShapes(1).Select
        UF2Dgraph.Show vbModeless
        PicOpen = True
        Cancel = True
    End If
slut:
End If

End Sub


