Attribute VB_Name = "GlobalVariables"
Option Explicit

Public AppNavn As String

Public Const AppVersion = "1.34"
#If Mac Then ' WordMat shows patchVersion attached to AppVersion. if ".1" then shows as 1.24.1. or " beta" -> "1.24 beta" A new PatchVersion does not trigger autoupdate
    Public Const PatchVersion = ".3" ' Mac  ".1"
#Else
    Public Const PatchVersion = ".3" ' Windows ".1"
#End If
'Public DebugWM As Boolean = True
Public DebugWM As Boolean ' brug toggledebug funktion
Public Const PI As Double = 3.14159265358979

#If Mac Then ' VbCrLf which is mac friendly
    Public Const VbCrLfMac = vbCr
#Else
    Public Const VbCrLfMac = vbCrLf
#End If

Public TT As New CLang ' Text Translation

Public XLapp As Object ' XLapp As Excel.Application

#If Mac Then
#Else
Public MaxProc As Object
Public MaxProcUnit As Object
#End If

Public mSkoleNavn As String
Public mPartnerSkab As Integer ' 0=not checked, 1= true, 2=false

' colorcodes for LabelButtons
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

Public PicOpen As Boolean ' marks a graphimage is open, It must then be overwritten

Public latexfil As New CLaTex

