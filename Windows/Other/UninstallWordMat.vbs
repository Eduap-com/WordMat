'' WordMat uninstall script
'' Mikael Sørensen, EDUAP
'' 28/1-2024
''

Option Explicit
Dim objFSO, objFolder, objSubFolder, objShell, FilePath, WMall, AdminPriv, AllOK

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject( "WScript.Shell" )

AllOK=True

' Prompt user to continue with uninstall
Dim intAnswer
intAnswer = MsgBox("This will remove WordMat from the Ribbon in Word, and delete all associated files." & vbcrlf & "Do you want to continue?", vbYesNo + vbQuestion, "Uninstall WordMat")
If intAnswer = vbNo Then
    Wscript.Quit
End If

' Check if Word is running. If so, prompt user to close Word
Dim objWMIService, colProcessList, objProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'WINWORD.EXE'")
For Each objProcess in colProcessList
    intAnswer = MsgBox("Word is running. Please close any open Word documents. If you continue by clicking OK. All Word documents will be forceclosed", vbOKCancel + vbExclamation, "Word is running")
Next
if intAnswer = vbCancel Then
    Wscript.Quit
End If

' Force close word if it is running
Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'WINWORD.EXE'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next

' Delete WordMat from the Ribbon for userinstall
Dim sourceFile, topLevelFolder
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word"
on error resume next
' objFSO.DeleteFile(topLevelFolder & "\STARTUP\WordMat.dotm")
' objFSO.DeleteFile(topLevelFolder & "\STARTUP\WordMatP.dotm")
' objFSO.DeleteFile(topLevelFolder & "\STARTUP\WordMatP2.dotm")
' objFSO.DeleteFile(topLevelFolder & "\START\WordMat.dotm")
' objFSO.DeleteFile(topLevelFolder & "\START\WordMatP.dotm")
' objFSO.DeleteFile(topLevelFolder & "\START\WordMatP2.dotm")
 Set objFolder = objFSO.GetFolder(topLevelFolder)
For Each objSubFolder in objFolder.Subfolders
    objFSO.DeleteFile(objSubFolder.Path & "\WordMat.dotm")
    objFSO.DeleteFile(objSubFolder.Path & "\WordMatP.dotm")
    objFSO.DeleteFile(objSubFolder.Path & "\WordMatP2.dotm")
Next
on error goto 0

' Delete the WordMat folder and all files and subfolders for userinstall
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\WordMat"
on error resume next ' Script cant delete it self, so ignore errors
objFSO.DeleteFolder(topLevelFolder), True
on error goto 0

' Delete WordMat from Start-menu
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\WordMat"  
on error resume next ' Script cant delete it self, so ignore errors
objFSO.DeleteFolder(topLevelFolder), True
on error goto 0

' Check if WordMat is installed for all users.
Dim WMallInstallFolder, WMallSTARTUPfolder
WMallInstallFolder = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") & "\WordMat"
WMallSTARTUPfolder = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Microsoft Office\root\Office16\STARTUP"
If objFSO.FolderExists(WMallSTARTUPfolder) Then
    if objFSO.fileexists(WMallSTARTUPfolder & "\WordMat.dotm") Then
        WMall=true
    End If
End If

' Check if user has admin privileges.
Dim colItems, objItem
Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount Where Name = '" & objShell.ExpandEnvironmentStrings("%USERNAME%") & "'")
For Each objItem in colItems
    If objItem.AccountType = 512 Then
        AdminPriv=true
    End If
Next

if WMall = true then
    if AdminPriv = false then
        MsgBox "WordMat is installed for all users." & VbCrLf & "WordMat can only be uninstalled by an administrator." & VbCrLf & "Please run this script again as administrator to remove WordMat for all users", vbOKOnly + vbExclamation, "Admin required"
        AllOK=false
    else
        on error resume next
        err.clear
        objFSO.DeleteFile(WMallSTARTUPfolder & "\WordMat.dotm")
        if err.number <> 0 then
            MsgBox "WordMat is installed for all users." & VbCrLf & "WordMat can only be uninstalled by an administrator." & VbCrLf & "Please run this script again as administrator to remove WordMat for all users", vbOKOnly + vbExclamation, "Admin required"
            AllOK=false
        else
            objFSO.DeleteFile(WMallSTARTUPfolder & "\WordMatP.dotm")
            objFSO.DeleteFile(WMallSTARTUPfolder & "\WordMatP2.dotm")
            objFSO.DeleteFolder(WMallInstallFolder), True
        end if
        on error goto 0
    end if
End If

if AllOK then MsgBox "Uninstall complete." & VbCrLf & "WordMat is now removed from the Ribbon in Word.", vbOKOnly + vbInformation, "Uninstall complete"
set objFSO = Nothing
set objShell = Nothing
set objWMIService = Nothing
set colProcessList = Nothing
set objProcess = Nothing
set colItems = Nothing
set objItem = Nothing
Wscript.Quit
