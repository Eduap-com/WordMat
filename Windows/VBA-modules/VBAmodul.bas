Attribute VB_Name = "VBAmodul"
Option Explicit
' To run these macros you need these two settings:
' * add reference 'Microsoft Visual Basic for Applications Extensibility 5.3'
' * Settings | Trust center | Settings for macros | always trust VBA project object model
' Don't reference other modules or userforms. This module must be self contained
Const VBAModulesFolder = "VBA-modules" ' the subfolder to import and export modules from/to

Sub ReplaceToASCIIseq()
   Dim VBC As Object 'VBComponent
   Dim i As Long, s As String
   
   If MsgBox("Do you want to replace all codetext to ASCII-sequences of 4 characters?" & vbCrLf & vbCrLf & "After use you can open on both Mac and Windows" & vbCrLf & vbCrLf & "The conversion can take 5-10s. You will be prompted upon completion", vbOKCancel, "Confirm") = vbCancel Then Exit Sub
   
   For Each VBC In ActiveDocument.VBProject.VBComponents
        If VBC.Name <> "VBAmodul" And VBC.Name <> "VBAmodul1" Then
'      If VBC.Name = "CSprog" Then
'        If MsgBox(VBC.Name, vbOKCancel, "Continue") = vbCancel Then Exit Sub
         For i = 1 To VBC.CodeModule.CountOfLines
            If i > VBC.CodeModule.CountOfLines Then Exit For
            s = ReplaceLineToASCIIseq(VBC.CodeModule.Lines(i, 1))
            If s <> "" Or i > 3 Then
                VBC.CodeModule.DeleteLines i, 1
                VBC.CodeModule.InsertLines i, s
            Else
                VBC.CodeModule.DeleteLines i, 1 ' import/export introduces a blank line at the top of the code for forms. This removes these blank lines
            End If
         Next
      End If
   Next
'   GenerateKeyboardShortcuts
'   ActiveDocument.VBProject.VBComponents(i).CodeModule.InsertLines(
    MsgBox "Conversion Done", vbOKOnly, "Done"
End Sub
Sub ReplaceToExtendedASCII()
   Dim VBC As Object  'VBComponent
   Dim i As Long, s As String
   
   If MsgBox("Do you want to replace all codetext to extended ASCII?" & vbCrLf & vbCrLf & "After use you can distribute the document" & vbCrLf & vbCrLf & "The conversion can take 5-10s. You will be prompted upon completion", vbOKCancel, "Confirm") = vbCancel Then Exit Sub
   
   For Each VBC In ActiveDocument.VBProject.VBComponents
        If VBC.Name <> "VBAmodul" And VBC.Name <> "VBAmodul1" Then
'      If VBC.Name = "CSprog" Then
        
         For i = 1 To VBC.CodeModule.CountOfLines
            s = ReplaceLineToExtendedASCII(VBC.CodeModule.Lines(i, 1))
            VBC.CodeModule.DeleteLines i, 1
            If s <> "" Or i > 3 Then VBC.CodeModule.InsertLines i, s
         Next
      End If
   Next
'   ActiveDocument.VBProject.VBComponents(i).CodeModule.InsertLines(
    MsgBox "Conversion Done", vbOKOnly, "Done"
End Sub

Private Function ReplaceLineToASCIIseq(s As String) As String
   s = Replace(s, ChrW(230), "*ae*") ' ae
   s = Replace(s, ChrW(248), "*oe*") ' oe
   s = Replace(s, ChrW(229), "*aa*") ' aa
   s = Replace(s, ChrW(198), "*AE*") ' AE
   s = Replace(s, ChrW(216), "*OE*") ' OE
   s = Replace(s, ChrW(197), "*AA*") ' AAA
   s = Replace(s, ChrW(225), "*a-*") ' a'
   s = Replace(s, ChrW(233), "*e-*") ' e'
   s = Replace(s, ChrW(243), "*o-*") ' o'
   s = Replace(s, ChrW(192), "*A~*") ' A~
   s = Replace(s, ChrW(191), "*?-*") ' (omvendt ?)
   s = Replace(s, ChrW(241), "*n-*") ' (n~)
   s = Replace(s, ChrW(237), "*i-*") ' (i')
   s = Replace(s, ChrW(250), "*u-*") ' (u')
   s = Replace(s, ChrW(176), "*gr*") ' (gradtegn)
   s = Replace(s, ChrW(167), "*pa*") ' (paragraf)
   s = Replace(s, ChrW(8364), "*eu*") ' (euro)
   s = Replace(s, ChrW(8230), "*._.*") ' ...
   '
   ReplaceLineToASCIIseq = s
End Function
Private Function ReplaceLineToExtendedASCII(s As String) As String
   s = Replace(s, "*ae*", ChrW(230)) ' ae
   s = Replace(s, "*oe*", ChrW(248)) ' oe
   s = Replace(s, "*aa*", ChrW(229)) ' aa
   s = Replace(s, "*AE*", ChrW(198)) ' AE
   s = Replace(s, "*OE*", ChrW(216)) ' OE
   s = Replace(s, "*AA*", ChrW(197)) ' AA
   s = Replace(s, "*a-*", ChrW(225)) ' a'
   s = Replace(s, "*e-*", ChrW(233)) ' e'
   s = Replace(s, "*o-*", ChrW(243)) ' o'
   s = Replace(s, "*A~*", ChrW(192)) ' A~
   s = Replace(s, "*?-*", ChrW(191)) ' (omvendt ?)
   s = Replace(s, "*n-*", ChrW(241)) ' (n~)
   s = Replace(s, "*i-*", ChrW(237)) ' (i')
   s = Replace(s, "*u-*", ChrW(250)) ' (u')
   s = Replace(s, "*gr*", ChrW(176)) ' (gradtegn)
   s = Replace(s, "*pa*", ChrW(167)) ' (paragraf)
   s = Replace(s, "*eu*", ChrW(8364)) ' (euro)
   s = Replace(s, "*._.*", ChrW(8230)) ' ...
   ReplaceLineToExtendedASCII = s
End Function
Sub ReplaceToANSI()
' if document has been edited on Mac this function must be run on Windows to convert
' special characters in the code. (ANSI-Unicode problem)
' There are few problems with characters from Windows to Mac. Only _ (shift 4) seems to be a problem. Just dont use it.
   Dim VBC As Object  'VBComponent
   Dim i As Long, s As String
   If MsgBox("Do you want to replace all codetext to ANSI?", vbOKCancel, "Comfirm") = vbCancel Then Exit Sub
   
   For Each VBC In ActiveDocument.VBProject.VBComponents
        If VBC.Name <> "VBAmodul" And VBC.Name <> "VBAmodul1" Then
'      If VBC.Name = "CSprog" Then
         For i = 2 To VBC.CodeModule.CountOfLines
            s = ReplaceLineToANSI(VBC.CodeModule.Lines(i, 1))
            VBC.CodeModule.DeleteLines i, 1
            VBC.CodeModule.InsertLines i, s
         Next
      End If
   Next
'   ActiveDocument.VBProject.VBComponents(i).CodeModule.InsertLines(
    MsgBox "Conversion Done", vbOKOnly, "Done"
End Sub
Private Function ReplaceLineToANSI(s As String) As String
    s = Replace(s, ChrW(190), ChrW(230)) ' ae
    s = Replace(s, ChrW(191), ChrW(248)) ' oe
    s = Replace(s, ChrW(338), ChrW(229)) ' aa
    s = Replace(s, ChrW(174), ChrW(198)) ' AE
    s = Replace(s, ChrW(175), ChrW(216)) ' OE
    s = Replace(s, ChrW(129), ChrW(197)) ' AA
    s = Replace(s, ChrW(8225), ChrW(225)) '(a')
    s = Replace(s, ChrW(381), ChrW(233)) '(e')
    s = Replace(s, ChrW(8212), ChrW(243)) '(o')
    s = Replace(s, ChrW(191), ChrW(192)) '(A') (Fra omvendt ?)
    s = Replace(s, ChrW(203), ChrW(192)) '(A')
    '   s = Replace(s, ChrW(192), ChrW(191)) '(omvendt ?) karambolerer med A' ovenfor
    s = Replace(s, ChrW(8211), ChrW(241)) '(n~)
    s = Replace(s, ChrW(8217), ChrW(237)) '(i')
    s = Replace(s, ChrW(339), ChrW(250)) '(u')
    s = Replace(s, ChrW(161), ChrW(176)) '(gradtegn)
    s = Replace(s, ChrW(164), ChrW(167)) ' paragraf (fra sol)
    s = Replace(s, ChrW(219), ChrW(8364)) ' Euro
    '   s = Replace(s, "*._.*", ChrW(8230)) ' tre prikker
    ReplaceLineToANSI = s
End Function
Function FolderExists(FolderPath As String) As Boolean

    If right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If
    
    If Dir(FolderPath, vbDirectory) <> vbNullString Then
        FolderExists = True
    Else
        FolderExists = False
    End If
End Function

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
'    Dim FSO As Object
    Dim SpecialPath As String

'    Set FSO = CreateObject("Scripting.FileSystemObject")
    SpecialPath = ActiveDocument.path
    
    If right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
'    SpecialPath = """" & SpecialPath & "VBAProjectFiles"""
    SpecialPath = SpecialPath & VBAModulesFolder ' "VBAProjectFiles"
    
'    If FSO.FolderExists(SpecialPath) = False Then
    If FolderExists(SpecialPath) = False Then
        On Error Resume Next
        MkDir SpecialPath
        On Error GoTo 0
    End If
    
    If FolderExists(SpecialPath) = True Then
        FolderWithVBAProjectFiles = SpecialPath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function
Public Sub ExportAllModules()
    Dim bExport As Boolean
    Dim wkbSource As Document
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String, FileList As String
    Dim cmpComponent As VBIDE.VBComponent

#If Mac Then
    MsgBox "This function is not meant to be run on Mac, as the export will not be Windows compatible", vbOKOnly, "No Mac"
    Exit Sub
#End If


   If MsgBox("Do you really want to export all VBA modules to folder '" & VBAModulesFolder & "'?" & vbCrLf & "(all current files in folder are deleted)", vbOKCancel) = vbCancel Then Exit Sub
    
    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveDocument.Name
    Set wkbSource = Application.ActiveDocument
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles '& "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

    ' n‘r der importeres oveni VBAmodul omdèbes til VBAmodul1. Det _ndres tilbage
        If cmpComponent.Name = "VBAmodul1" Then cmpComponent.Name = "VBAmodul"
        If cmpComponent.Name = "VBAmodul11" Then cmpComponent.Name = "VBAmodul"

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            FileList = FileList & szFileName & vbCrLf
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        End If
   
    Next cmpComponent
    
    ' save datefile
    Dim fs As Object, A As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.CreateTextFile(szExportPath & "A-ExportCreated " & Replace(Now(), ":", "") & ".txt", True)
    A.WriteLine ("VBA-exported of Project " & wkbSource.VBProject.Name & " created " & Now())
    A.Close

    MsgBox "Files exported to folder '" & VBAModulesFolder & "':" & vbCrLf & vbCrLf & FileList, vbOKOnly, "Export complete"
End Sub
Sub ImportAllModules()
    Dim bExport As Boolean, d As String, q As String
    Dim wkbSource As Document
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim StrFile As String, i As Integer
    Dim arr() As String, FileList As String, MBP As Integer
    Dim cmpComponent As VBIDE.VBComponent

#If Mac Then
    MsgBox "This function is not meant to be run on Mac, as the import is not Mac to Windows compatible", vbOKOnly, "No Mac"
    Exit Sub
#End If

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder does not exist"
        Exit Sub
    End If
    
    szSourceWorkbook = ActiveDocument.Name
    Set wkbSource = Application.ActiveDocument
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to import"
        Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles
    
    If right(szExportPath, 1) <> "\" Then szExportPath = szExportPath & "\"
    
    StrFile = Dir(szExportPath & "A-ExportCreated*")
    If StrFile <> "" Then d = Mid(StrFile, 17, Len(StrFile) - 20)
    d = Left(d, 13) & ":" & Mid(d, 14, 2) & ":" & right(d, 2)
    
    
    StrFile = Dir(szExportPath & "*")
    Do While Len(StrFile) > 0
        If Left(StrFile, 15) <> "A-ExportCreated" Then 'StrFile <> "VBAmodul.bas"
            FileList = FileList & StrFile & vbCrLf
        End If
        StrFile = Dir
    Loop
    arr = Split(FileList, vbCrLf)
'    FileList = Replace(FileList, vbCrLf, " | ")
    q = ""
    If Not ActiveDocument.Saved Then
        q = "You have unsaved changes!" & vbCrLf & vbCrLf
        MBP = vbExclamation
    End If
    
    If GetTimeString(ActiveDocument.BuiltInDocumentProperties("Last Save Time")) - 100 > GetTimeString(d) Then
        q = q & "Your document is newer than the export in '" & VBAModulesFolder & "'" & vbCrLf & "Export date: " & d & vbCrLf & "Save date: " & ActiveDocument.BuiltInDocumentProperties("Last Save Time") & vbCrLf
        If MBP = vbExclamation Then
            MBP = vbCritical
        Else
            MBP = vbExclamation
        End If
    End If
    
    q = q & vbCrLf & "Do you really want to import all VBA modules from folder " & VBAModulesFolder & "? (All existing VBA-modules will be removed before importing)" & vbCrLf
    If d <> "" Then q = q & "Export date: " & d & vbCrLf & vbCrLf
    q = q & "File list to import:" & vbCrLf & FileList
    
   If MsgBox(q, vbOKCancel + MBP, "Continue?") = vbCancel Then Exit Sub
    
   DeleteAllModules False
    
    For i = 0 To UBound(arr)
        If arr(i) <> "" And InStr(arr(i), ".frx") <= 0 Then  'Arr(i) <> "VBAmodul.bas" And
            wkbSource.VBProject.VBComponents.Import szExportPath & arr(i)
        End If
    Next
    
    MsgBox "Files imported from folder '" & VBAModulesFolder & "':" & vbCrLf & vbCrLf & FileList, vbOKOnly, "Import complete"
End Sub
Public Sub RemoveAllModules()
    DeleteAllModules True
End Sub
Public Sub DeleteAllModules(Optional PromptOk As Boolean = True)
    Dim bExport As Boolean
    Dim wkbSource As Document
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    If PromptOk Then
        If MsgBox("Do you really want to remove all VBA modules in this document?" & vbCrLf & "(Except VBAmodul.bas)", vbOKCancel) = vbCancel Then Exit Sub
    End If
    
    szSourceWorkbook = ActiveDocument.Name
    Set wkbSource = Application.ActiveDocument
    
    If wkbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected, not possible to delete the code"
        Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles
    If right(szExportPath, 1) <> "\" Then szExportPath = szExportPath & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                bExport = False
        End Select
        If PromptOk Then
            If szFileName = "VBAmodul.bas" Then bExport = False
        End If
        
        If bExport Then
            wkbSource.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent
    
    If PromptOk Then MsgBox "All modules has been removed (Except VBAmodul)"
End Sub

Function GetTimeString(ByVal d As Date) As String

GetTimeString = Year(d) & Month(d) & Day(d) & AddZero(Hour(d)) & AddZero(Minute(d)) & AddZero(Second(d))

End Function

Function AddZero(n As Integer) As String
If n < 10 Then
    AddZero = "0" & n
Else
    AddZero = n
End If
End Function


