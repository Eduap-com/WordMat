' Dette script renser WordMat.dotm for compileret kode mm.
' Det kræver at Ribbon Commander er installeret med licens
' Det er en hjælp hvis WordMat.dotm pludelig crasher ved åbning. 
' Der er andre løsninger, men denne er nemmest.
' Scriptet skal ligge i samme mappe som WordMat.dotm
' scriptet kopierer også de rensede filer til Mac-mappen

Option Explicit

Dim WordApp, Document, strPath, FilNavn, Arr, i, FL, strFile, OldSize
dim objFSO, objFile
SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

' Kan indeholde flere filnanve adskilt ad komma
FilNavn="WordMat.dotm,WordMatP.dotm,WordMatP2.dotm"

Set WordApp = CreateObject("Word.Application")
Wordapp.visible = False
strPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
'Msgbox strpath
Set Document = WordApp.Documents.Open(strPath & "VBACleaner.docm", 0, True)
'WordApp.Run "CleanVBA", strPath & WScript.Arguments.Item(0), FALSE, TRUE

Arr = split(FilNavn,",")

FL="Filename" & vbTab & vbTab & "Oldsize" & vbTab & "Newsize" & VbCrLf
For i=0 to Ubound(Arr)
    strFile = strPath & Arr(i)
    Set objFile=objFSO.GetFile(strFile)
    OldSize=objFile.size/1000
    WordApp.Run "CleanVBA", strFile, FALSE, TRUE
    FL = FL & Arr(i) & vbtab & oldsize & vbtab & objFile.size/1000 & VbCrLf
Next
'WordApp.Run "CleanVBA", strPath & FilNavn, FALSE, TRUE

Document.Close
WordApp.Quit

' Copy the cleaned files to Mac folder one level up. Dette giver problemer på Mac da git pull stoppes når der er ændringer i filerne
'objFSO.CopyFile strPath & "WordMat.dotm", strPath & "..\Mac\WordMat.dotm", TRUE
'objFSO.CopyFile strPath & "WordMatP.dotm", strPath & "..\Mac\WordMatP.dotm", TRUE
'objFSO.CopyFile strPath & "WordMatP2.dotm", strPath & "..\Mac\WordMatP2.dotm", TRUE

set objFSO = Nothing
Set Document = Nothing
Set WordApp = Nothing

msgbox "Following files has been cleaned:" & vbcrlf & vbcrlf  & FL & VbCrLf & VbCrLf & ""

WScript.Quit