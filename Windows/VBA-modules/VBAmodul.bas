Attribute VB_Name = "VBAmodul"
Option Explicit
' add reference 'Microsoft Visual Basic for Applications Extensibility 5.3'

Sub ReplaceToNonUnicode()
   Dim VBC As Object  'VBComponent
   Dim i As Long, s As String
   
   For Each VBC In ActiveDocument.VBProject.VBComponents
      If VBC.Name = "CSprog" Then
         For i = 2 To VBC.CodeModule.CountOfLines
            s = ReplaceLineToNonUnicode(VBC.CodeModule.Lines(i, 1))
            VBC.CodeModule.DeleteLines i, 1
            VBC.CodeModule.InsertLines i, s
         Next
      End If
   Next
'   ActiveDocument.VBProject.VBComponents(i).CodeModule.InsertLines(

End Sub
Sub ReplaceToUnicode()
   Dim VBC As Object  'VBComponent
   Dim i As Long, s As String
   
   For Each VBC In ActiveDocument.VBProject.VBComponents
      If VBC.Name = "CSprog" Then
         For i = 2 To VBC.CodeModule.CountOfLines
            s = ReplaceLineToUnicode(VBC.CodeModule.Lines(i, 1))
            VBC.CodeModule.DeleteLines i, 1
            VBC.CodeModule.InsertLines i, s
         Next
      End If
   Next
'   ActiveDocument.VBProject.VBComponents(i).CodeModule.InsertLines(

End Sub

Private Function ReplaceLineToNonUnicode(s As String) As String
   s = Replace(s, ChrW(230), "*ae*") '�
   s = Replace(s, ChrW(248), "*oe*") '�
   s = Replace(s, ChrW(229), "*aa*") '�
   s = Replace(s, ChrW(198), "*AE*") ' �
   s = Replace(s, ChrW(216), "*OE*") ' �
   s = Replace(s, ChrW(197), "*AA*") ' �
   s = Replace(s, ChrW(225), "*a-*") ' �
   s = Replace(s, ChrW(233), "*e-*") ' �
   s = Replace(s, ChrW(243), "*o-*") ' �
   s = Replace(s, ChrW(191), "*?-*") ' �
   s = Replace(s, ChrW(8230), "*._.*") ' ...
   '
   ReplaceLineToNonUnicode = s
End Function
Private Function ReplaceLineToUnicode(s As String) As String
   s = Replace(s, "*ae*", ChrW(230))
   s = Replace(s, "*oe*", ChrW(248))
   s = Replace(s, "*aa*", ChrW(229))
   s = Replace(s, "*AE*", ChrW(198))
   s = Replace(s, "*OE*", ChrW(216))
   s = Replace(s, "*AA*", ChrW(197))
   s = Replace(s, "*a-*", ChrW(225))
   s = Replace(s, "*e-*", ChrW(233))
   s = Replace(s, "*o-*", ChrW(243))
   s = Replace(s, "*?-*", ChrW(191))
   s = Replace(s, "*._.*", ChrW(8230))
   ReplaceLineToUnicode = s
End Function

