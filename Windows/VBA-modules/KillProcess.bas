Attribute VB_Name = "KillProcess"
Option Explicit
#If Mac Then
Sub KillMaxima()
End Sub
#Else
Sub KillMaxima()
        Shell "cmd.exe /c taskkill /IM sbcl.exe /F" ' Denne slår alt i hjel
        ' closeapp bruger WMI som kan bruges til exploit. Dette er måske sikrere.

'CloseAPP_B "maxima.exe"
'CloseAPP_B "xmaxima.exe"
End Sub

''**************************************
'Sub KillTest()
'MsgBox IIf(CloseAPP("notepad.exe", _
'True, False), _
'"Killed", "Failed")
'End Sub
''**************************************
'
'Sub KillTest_B()
'CloseAPP_B "notepad.exe"
'End Sub
''**************************************
'
''Close Application
''CloseApp KillAll=False -Only first occurrence
'' KillAll=True -All occurrences
'' NeedYesNo=True -Prompt to kill
'' NeedYesNo=False -Silent kill
'Private Function CloseAPP _
'( _
'AppNameOfExe _
'As String, _
'Optional _
'KillAll _
'As Boolean = False, _
'Optional _
'NeedYesNo _
'As Boolean = True _
') _
'As Boolean
'
'Dim oProcList As Object
'Dim oWMI As Object
'Dim oProc As Object
'
'CloseAPP = False
'' step 1: create WMI object instance:
'Set oWMI = GetObject("winmgmts:")
'If IsNull(oWMI) = False Then
'' step 2: create object collection of Win32 processes:
'Set oProcList = oWMI.InstancesOf("win32_process")
'' step 3: iterate through the enumerated collection:
'For Each oProc In oProcList
''MsgBox oProc.Name
'' option to close a process:
'If VBA.UCase(oProc.Name) = VBA.UCase(AppNameOfExe) Then
'If NeedYesNo Then
'If MsgBox("Kill " & _
'oProc.Name & vbNewLine & _
'"Are you sure?", _
'vbYesNo + vbCritical) _
'= vbYes Then
'oProc.Terminate (0)
''no test to see if this is really true
'CloseAPP = True
'End If 'MsgBox("Kill "
'Else 'NeedYesNo
'oProc.Terminate (0)
''no test to see if this is really true
'CloseAPP = True
'End If 'NeedYesNo
'
''continue search for more???
'If Not KillAll And CloseAPP Then
'Exit For 'oProc In oProcList
'End If 'Not KillAll And CloseAPP
'
'End If 'IsNull(oWMI) = False
'Next 'oProc In oProcList
'Else 'IsNull(oWMI) = False
''report error
'End If 'IsNull(oWMI) = False
'' step 4: close log file; clear out the objects:
'Set oProcList = Nothing
'Set oWMI = Nothing
'End Function
''**************************************
'
''No frills killer
'Private Function CloseAPP_B(AppNameOfExe As String)
'Dim oProcList As Object
'Dim oWMI As Object
'Dim oProc As Object
'
'' step 1: create WMI object instance:
'Set oWMI = GetObject("winmgmts:")
'If IsNull(oWMI) = False Then
'' step 2: create object collection of Win32 processes:
'Set oProcList = oWMI.InstancesOf("win32_process")
'' step 3: iterate through the enumerated collection:
'For Each oProc In oProcList
'' option to close a process:
'If VBA.UCase(oProc.Name) = VBA.UCase(AppNameOfExe) Then
'oProc.Terminate (0)
'End If 'IsNull(oWMI) = False
'Next
'End If
'End Function
'
#End If

