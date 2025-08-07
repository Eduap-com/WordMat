Attribute VB_Name = "ModuleRegistry"
Option Explicit
Option Private Module

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As LongPtr, phkResult As LongPtr, lpdwDisposition As Long) As Long
    
Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As String, lpcchValueName As Long, ByVal lpReserved As LongPtr, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    
Private Const HKEY_CURRENT_USER As LongPtr = &H80000001
Private Const HKEY_LOCAL_MACHINE As LongPtr = &H80000002
Private Const KEY_READ As Long = &H20019
Private Const KEY_WRITE As Long = &H20006
Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_NO_MORE_ITEMS As Long = 259

Public Const REG_SZ As Long = 1
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4

Private myWS As Object

Public Function GetRegistryValue(hive As String, path As String, valueName As String) As Variant
'Example: GetRegistryValue("HKCU", "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState")
'Example: GetRegistryValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
'?GetRegistryValue("HKCU", "Software\WordMat\Settings", "AntalBeregninger")

    Dim hRoot As LongPtr
    Dim hKey As LongPtr
    Dim result As Long
    Dim valueType As Long
    Dim dataSize As Long
    Dim dataBuffer() As Byte '0 To 1023
    Dim strData As String
    Dim dwordData As Long, i As Long
    Dim hexStr As String
    
    ' Map hive string to handle
    Select Case UCase(hive)
        Case "HKLM", "HKEY_LOCAL_MACHINE"
            hRoot = HKEY_LOCAL_MACHINE
        Case "HKCU", "HKEY_CURRENT_USER"
            hRoot = HKEY_CURRENT_USER
        Case Else
            GetRegistryValue = vbNullString 'CVErr(5) ' Invalid procedure call
            Exit Function
    End Select

    ' Open registry key
    result = RegOpenKeyEx(hRoot, path, 0, KEY_READ, hKey)
    If result <> ERROR_SUCCESS Then
        RegCloseKey hKey
        GetRegistryValue = vbNullString 'CVErr(91) ' Object variable or With block variable not set
        Exit Function
    End If
        
    ' Query value size first
    result = RegQueryValueEx(hKey, valueName, 0, valueType, ByVal 0&, dataSize)
    If result <> ERROR_SUCCESS Then
        RegCloseKey hKey
        GetRegistryValue = vbNullString ' CVErr(94) ' Invalid use of Null
        Exit Function
    End If

    ReDim dataBuffer(0 To dataSize - 1) As Byte

    ' Query actual data
    result = RegQueryValueEx(hKey, valueName, 0, valueType, dataBuffer(0), dataSize)
    If result = ERROR_SUCCESS Then
        Select Case valueType
            Case REG_SZ
                strData = Left$(StrConv(dataBuffer, vbUnicode), InStr(1, StrConv(dataBuffer, vbUnicode), vbNullChar) - 1)
                GetRegistryValue = strData
            Case REG_DWORD
                CopyMemory dwordData, dataBuffer(0), 4
                GetRegistryValue = dwordData
            Case REG_BINARY
                For i = 0 To dataSize - 1
                    hexStr = hexStr & right$("0" & Hex(dataBuffer(i)), 2) & " "
                Next i
                GetRegistryValue = Trim(hexStr)
            Case Else
                GetRegistryValue = vbNullString 'CVErr(13) ' Type mismatch
        End Select
    Else
        GetRegistryValue = vbNullString 'CVErr(11) ' Division by zero (used here to mean "read failed")
    End If

    RegCloseKey hKey
End Function

Public Function SetRegistryValue(hive As String, path As String, valueName As String, valueType As Long, valueData As Variant) As Boolean
' SetRegistryValue("HKCU", "Software\WordMat\Settings", "AntalBeregninger",REG_SZ,"1001")
    Dim hRoot As LongPtr
    Dim hKey As LongPtr
    Dim result As Long
    Dim disposition As Long
    Dim binData() As Byte
    Dim dwordVal As Long
    Dim strVal As String
    Dim dataLen As Long

    ' Map hive
    Select Case UCase(hive)
        Case "HKLM", "HKEY_LOCAL_MACHINE"
            hRoot = HKEY_LOCAL_MACHINE
        Case "HKCU", "HKEY_CURRENT_USER"
            hRoot = HKEY_CURRENT_USER
        Case Else
            SetRegistryValue = False
            Exit Function
    End Select

    ' Create or open the key
    result = RegCreateKeyEx(hRoot, path, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, hKey, disposition)
    If result <> 0 Then
        SetRegistryValue = False
        Exit Function
    End If

    Select Case valueType
        Case REG_SZ
            strVal = CStr(valueData) & vbNullChar
            dataLen = LenB(strVal)
            result = RegSetValueEx(hKey, valueName, 0, REG_SZ, ByVal strVal, dataLen)

        Case REG_DWORD
            dwordVal = CLng(valueData)
            result = RegSetValueEx(hKey, valueName, 0, REG_DWORD, dwordVal, 4)

        Case REG_BINARY
            If IsArray(valueData) Then
                binData = valueData
                dataLen = UBound(binData) - LBound(binData) + 1
                result = RegSetValueEx(hKey, valueName, 0, REG_BINARY, binData(LBound(binData)), dataLen)
            Else
                result = 1 ' Invalid data
            End If

        Case Else
            result = 1 ' Unsupported type
    End Select

    RegCloseKey hKey
    SetRegistryValue = (result = 0)
End Function

' **** These are mainly for mac use ******
Public Function RegKeyRead(i_RegKey As String) As String
'"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir"
#If Mac Then
    RegKeyRead = GetSetting("com.wordmat", "defaults", i_RegKey)
#Else
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
    If myWS Is Nothing Then
        Set myWS = CreateObject("WScript.Shell")
    End If
  'read key from registry
  If RegKeyExists(i_RegKey) Then
      RegKeyRead = myWS.regread(i_RegKey)
  End If
  Err.Clear
  Set myWS = Nothing
#End If
End Function

Public Function RegKeyExists(i_RegKey As String) As Boolean
#If Mac Then
    RegKeyExists = True
    If GetSetting("com.wordmat", "defaults", i_RegKey) = "" Then RegKeyExists = False
#Else

  On Error GoTo ErrorHandler
  'access Windows scripting
    If myWS Is Nothing Then
        Set myWS = CreateObject("WScript.Shell")
    End If
  'try to read the registry key
  myWS.regread i_RegKey
  'key was found
  RegKeyExists = True
    Set myWS = Nothing
    Exit Function

ErrorHandler:
  'key was not found
  RegKeyExists = False
'    Set myWS = Nothing
#End If
End Function

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
Public Sub RegKeySave(ByVal i_RegKey As String, ByVal i_Value As String, Optional ByVal i_Type As String = "REG_SZ")
#If Mac Then
    SaveSetting "com.wordmat", "defaults", i_RegKey, i_Value
#Else
    On Error Resume Next
  'access Windows scripting
    If myWS Is Nothing Then
        Set myWS = CreateObject("WScript.Shell")
    End If
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type
'  Set myWS = Nothing
#End If
End Sub

'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Private Function RegKeyDelete(i_RegKey As String) As Boolean
#If Mac Then
    DeleteSetting "com.wordmat", "defaults", i_RegKey
#Else

    On Error GoTo ErrorHandler
    'access Windows scripting
    If myWS Is Nothing Then
        Set myWS = CreateObject("WScript.Shell")
    End If
    'delete registry key
    On Error Resume Next
    myWS.RegDelete i_RegKey
    'deletion was successful
    RegKeyDelete = True
    Set myWS = Nothing
    Exit Function

ErrorHandler:
    'deletion wasn't successful
    RegKeyDelete = False
'    Set myWS = Nothing
#End If
End Function


