Attribute VB_Name = "MaximaDirectdll"
Option Explicit

Option Private Module
' Funktion til at kalde .net dll uden at de er registreret.
' Add reference mscorelib.dll
' add reference to mscoree.dll (Common language runtime execution engine) ligger i c:\Windows\Microsoft.NET\Framework64\v4.0.30319
' QlmCLRHost_x64.dll and QlmCLRHost_x86.dll  must be placed where it can found. Any subfolder to %appdata% (AppData\Roaming) is ok?
#If Mac Then
#Else

#If VBA7 Then
Private Declare PtrSafe Function GetShortPathName Lib "Kernel32.dll" Alias "GetShortPathNameW" (ByVal LongPath As LongPtr, ByVal ShortPath As LongPtr, ByVal Size As Long) As Long
Private Declare PtrSafe Function SetDllDirectory Lib "Kernel32.dll" Alias "SetDllDirectoryW" (ByVal Path As LongPtr) As Long
Private Declare PtrSafe Sub LoadClr_x64 Lib "QlmCLRHost_x64.dll" (ByVal clrVersion As String, ByVal verbose As Boolean, ByRef CorRuntimeHost As IUnknown)
Private Declare PtrSafe Sub LoadClr_x86 Lib "QlmCLRHost_x86.dll" (ByVal clrVersion As String, ByVal verbose As Boolean, ByRef CorRuntimeHost As IUnknown)

'Private Declare PtrSafe Function CorBindToRuntimeEx Lib "mscoree" ( _
'    ByVal pwszVersion As LongPtr, _
'    ByVal pwszBuildFlavor As LongPtr, _
'    ByVal startupFlags As Long, _
'    ByRef rclsid As Long, _
'    ByRef riid As Long, _
'    ByRef ppvObject As mscoree.CorRuntimeHost) As Long
    
#Else
Private Declare Function GetShortPathName Lib "Kernel32.dll" Alias "GetShortPathNameW" (ByVal LongPath As Long, ByVal ShortPath As Long, ByVal Size As Long) As Long
Private Declare Function SetDllDirectory Lib "Kernel32.dll" Alias "SetDllDirectoryW" (ByVal Path As Long) As Long

Private Declare Sub LoadClr_x64 Lib "QlmCLRHost_x64.dll" (ByVal clrVersion As String, ByVal verbose As Boolean, ByRef CorRuntimeHost As IUnknown)
Private Declare Sub LoadClr_x86 Lib "QlmCLRHost_x86.dll" (ByVal clrVersion As String, ByVal verbose As Boolean, ByRef CorRuntimeHost As IUnknown)
#End If ' WinAPI Declarations

' Class variables

Private Sub TestDllDotNet()
    Dim m_myobject As Object
    Set m_myobject = PGetMaxProc()
'    Set m_myobject = GetObjectFromDll("C:\Users\mikae\AppData\Roaming\WordMat\", "MathMenu.dll", "MaximaProcessClass") ' "namespace.class"  MathMenu.MaximaProcessClass   men det er en global klasse så intet namespace
'    m_myobject.SetMaximaPath "C:\Program Files (x86)\WordMat\Maxima-5.47.0"
'    m_myobject.SetMaximaPath Environ("AppData") & "\WordMat\Maxima-5.47.0"
    m_myobject.SetMaximaPath GetMaximaPath
    m_myobject.StartMaximaProcess
    MsgBox "errcode: " & m_myobject.ErrCode
    m_myobject.ExecuteMaximaCommand "2+3;", 5
    MsgBox "running 2+3;"
    MsgBox m_myobject.LastMaximaOutput
    Set m_myobject = Nothing
End Sub
Private Sub TestBrowserDll()
    Dim m_myobject As Object
    Set m_myobject = GetObjectFromDll("C:\Users\mikae\AppData\Roaming\WordMat\WebViewWrap\", "WebViewWrap.dll", "WebViewWrap.Browser", "C:\Users\mikae\AppData\Roaming\WordMat\") ' "namespace.class"  WebViewWrap.Browser   burde være global klasse så forstår ikke hvorfor der skal WebViewWrap med
    m_myobject.start
    m_myobject.Show
    m_myobject.navigate "https://www.eduap.com"
    MsgBox "ok for at lukke", vbOKOnly
    m_myobject.Close
    Set m_myobject = Nothing
End Sub
Public Function GetObjectFromDll(dllFolder As String, dllFileName As String, dllClass As String, Optional CLRdllFolder As String) As Object
' Henter object fra en .Net dll direkte uden at den er registreret
' dllFolder er den mappe hvor dll-filen ligger
' dllFileName er navnet på dll-filen
' CLRdllFolder er mappen hvor de to QlmCLRHost_x64.dll and QlmCLRHost_x86.dll er placeret. Hvis intet angives, så bruges dllFolder
On Error GoTo Slut
    Dim LongPath As String, PathLength As Integer, dllPath As String
    Dim ShortPath As String
    
    If CLRdllFolder = vbNullString Then CLRdllFolder = dllFolder
    
    dllFolder = Trim(dllFolder)
    If right(dllFolder, 1) <> "\" Then dllFolder = dllFolder & "\"
    dllPath = dllFolder & dllFileName
    
    LongPath = "\\?\" & CLRdllFolder
    ShortPath = String$(260, vbNull)

    PathLength = GetShortPathName(StrPtr(LongPath), StrPtr(ShortPath), 260)
    ShortPath = Mid$(ShortPath, 5, CLng(PathLength - 4))

    Call SetDllDirectory(StrPtr(ShortPath))
    Dim clr As mscoree.CorRuntimeHost
    
    ' If QlmCLRHost_xxx.dll 's cant be found try to chdir just before calling
    ' ChDrive Left$(Me.Path, 1)
    ' ChDir Me.Path
    
    If Is64BitApp() Then
        Call LoadClr_x64("v4.0", False, clr)
    Else
        Call LoadClr_x86("v4.0", False, clr)
    End If

    Call clr.start

    Dim domain As mscorlib.AppDomain
    Call clr.GetDefaultDomain(domain)

    Dim myInstanceOfDotNetClass As Object
    Dim handle As mscorlib.ObjectHandle

    Set handle = domain.CreateInstanceFrom(dllPath, dllClass)

    Dim clrObject As Object
    Set GetObjectFromDll = handle.Unwrap

Slut:
    On Error Resume Next
    Call clr.Stop
End Function

Private Function Is64BitApp() As Boolean
#If Win64 Then
    Is64BitApp = True
#End If
End Function

#End If
