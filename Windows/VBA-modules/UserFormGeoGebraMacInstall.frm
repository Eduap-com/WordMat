VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebraMacInstall 
   Caption         =   "GeoGebra 5 Installation"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "UserFormGeoGebraMacInstall.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGeoGebraMacInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_cancel_Click()
    Me.hide
End Sub

Private Sub CommandButton_downloads_Click()

End Sub

Private Sub CommandButton_moveApps_Click()
    Dim s As String
    s = RunScript("MoveGeoGebraToApplications", "")
    If s = "ok" Then
        RunScript "OpenApps", ""
        Label_step2.Caption = "1. Hold 'Control' nede mens du klikker på GeoGebra" & vbCrLf & "4. Klik Åben" & vbCrLf & "5. Klik Åben igen" & vbCrLf & "Så skulle GeoGebra gerne åbne, og WordMat vil fremover også kunne åbne GeoGebra."
        Label_step2.visible = True
    End If
End Sub

Private Sub UserForm_Activate()
    Dim FilePath As String, i As Integer
    StopNow = False
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    Do While Dir(FilePath, vbNormal) = vbNullString And i < 10
        DoEvents
        If StopNow Then Exit Do
        Wait 1
        Label_progress.Caption = Label_progress.Caption & "."
        i = i + 1
    Loop
End Sub

