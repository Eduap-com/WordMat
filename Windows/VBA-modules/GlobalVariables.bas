Attribute VB_Name = "GlobalVariables"
Option Explicit

#Const AVsafe = False ' Compiler constant. AntiVirus safe. If true will not compile functions that could potentially be blocked by AntiVirus

Public AppNavn As String

Public Const AppVersion = "1.30"
#If Mac Then ' WordMat shows patchVersion attached to AppVersion. if ".1" then shows as 1.24.1. or " beta" -> "1.24 beta" A new PatchVersion does not trigger autoupdate
    Public Const PatchVersion = "" ' Mac  ".1"
#Else
    Public Const PatchVersion = "" ' Windows ".1"
#End If
'Public DebugWM As Boolean = True
Public DebugWM As Boolean ' brug toggledebug funktion
Public Const PI As Double = 3.14159265358979

#If Mac Then ' VbCrLf which is mac friendly
    Public Const VbCrLfMac = vbCr
#Else
    Public Const VbCrLfMac = vbCrLf
#End If

'#If Mac Then
'#Else
    Public tempDoc As Document ' dokument hvor ligning kopieres over og l�ses. P� mac er der lavet en function . v.1.29 Mac kan nu h�ndteres som Windows
'#End If

'Public XLapp As Excel.Application
Public XLapp As Object

'Public MaxProc As MaximaProcessClass
#If Mac Then
Public MaxProc As MaximaProcess
#Else
Public MaxProc As Object
Public MaxProcUnit As Object
#End If

Public mSkoleNavn As String
Public mPartnerSkab As Integer ' 0=not checked, 1= true, 2=false

' Farvekoder til LabelButtons
Public Const LBColorFont = &H242424 'ForeColor
Public Const LBColorBorder = &H616161 'BorderColor
Public Const LBColorInactive = &HFFFFFF 'BackColor
Public Const LBColorHover = &HF5F5F5 'BackColor
Public Const LBColorPress = &HEBEBEB 'ForeColor
Public Const LBColorTABPress = &HE0E0E0 'ForeColor


Public ResFeltIndex As Integer
Public ResIndex As Integer
Public ResPos1 As Integer
Public ResPos2 As Integer
Public AntalB As LongPtr

Public TriangleAS As String
Public TriangleBS As String
Public TriangleCS As String
Public TriangleAV As String
Public TriangleBV As String
Public TriangleCV As String
Public TriangleNAS As String
Public TriangleNBS As String
Public TriangleNCS As String
Public TriangleNAV As String
Public TriangleNBV As String
Public TriangleNCV As String
Public TriangleSett1 As Integer
Public TriangleSett2 As Integer
Public TriangleSett3 As Boolean
Public TriangleSett4 As Boolean

Public ExcelVBAWarning As Integer

Public colindex As Integer

Public PicOpen As Boolean ' markerer at grafbillede er �bent s� det skal overskrives
Public Sprog As New CSprog

Public latexfil As New CLaTex

