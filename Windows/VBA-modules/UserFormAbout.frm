VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAbout 
   Caption         =   "About"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "UserFormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_license_Click()
    OpenLink "https://www.gnu.org/licenses/gpl-3.0.html"
End Sub

Private Sub Label_Licensetitle_Click()
    OpenLink "https://www.gnu.org/licenses/gpl-3.0.html"
End Sub

Private Sub Label_link_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_logo_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_ok_Click()
    Me.Hide
End Sub

Private Sub Label_partnerskab_Click()
    OpenLink "https://www.eduap.com/da/partnerskab/"
End Sub

Private Sub Label_partnerskab2_Click()
    OpenLink "https://www.eduap.com/da/partnerskab/"
End Sub

Private Sub Label1_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub UserForm_Activate()
    SetCaptions
    
    If QActivePartnership Then
        If Sprog.SprogNr = 1 Then
            Label_partnerskab.Caption = "Partnerskab registreret med " & SkoleNavn
        Else
            Label_partnerskab.Caption = "Partnership registered with " & SkoleNavn
        End If
    Else
        Label_partnerskab.Caption = "Ingen Partnerskab registreret"
    End If
        If Sprog.SprogNr = 1 Then
'            Label_partnerskab.Caption = "Partnerskab registreret med " & SkoleNavn
        Else
            Label_partnerskab.Caption = "Schools can sign up for a Partnership with Eduap. Partnership gets access to support, WordMat+ and much more. Click to read more"
        End If

End Sub

Sub SetCaptions()
    Dim V As String

    V = AppVersion
    If PatchVersion <> "" Then
        V = V & PatchVersion
    End If
    
    Me.Caption = Sprog.About & " WordMat"
    Label_version.Caption = "Version: " & V
    Label_text.Caption = Sprog.A(20)

End Sub
'SA(20, 0) = AppNavn & " er lavet af Mikael Samsøe Sørensen, Nyborg Gymnasium, EDUAP 2009-2024." & VbCrLfMac & VbCrLfMac & "Programmet er gratis og stilles til rådighed under Gnu General public license." & VbCrLfMac & VbCrLfMac & "Med WordMat er følgende andre gratis programmer installeret: Maxima, GnuPlot, GeoGebra, Graph " & VbCrLfMac & VbCrLfMac & "www.eduap.com"
'SA(20, 1) = AppNavn & " was made by Mikael Samsøe Sørensen, Nyborg Gymnasium, EDUAP 2009-2024." & VbCrLfMac & VbCrLfMac & "The program is free and open source under Gnu General public license." & VbCrLfMac & VbCrLfMac & "WordMat uses the following programs: Maxima, GnuPlot, GeoGebra, Graph " & VbCrLfMac & VbCrLfMac & "www.eduap.com"

Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
End Sub
