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
Private StopNow As Boolean
Private TestDone As Boolean

Private Sub CommandButton_cancel_Click()
    Dim FilePath As String
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        If MsgBox("GeoGebra 5 er ikke blevet færdiginstalleret. Vil du afslutte?", vbYesNo, "Ikke færdig") = vbYes Then
            Me.hide
        End If
    End If
    If Not TestDone Then
        If MsgBox("GeoGebra 5 er ikke blevet testet. Vil du afslutte?", vbYesNo, "Ikke testet") = vbYes Then
            Me.hide
        End If
    End If

End Sub

Private Sub CommandButton_stop_Click()
    StopNow = True
End Sub

Private Sub CommandButton_test_Click()
    Dim FilePath As String
    
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        MsgBox "GeoGebra er ikke blevet installeret endnu. Vent til download er færdig.", vbOKOnly, "Vent"
        GoTo slut
    End If
    
    GeoGebra
'    RunScript "OpenGeoGebra", geogebrafilersti
    TestDone = True
slut:
End Sub

Private Sub UserForm_Activate()
    Dim FilePath As String, i As Integer
    Dim s As String
    StopNow = False
    TestDone = False
    CommandButton_test.visible = False
    Label_progress.Caption = "*"
    CommandButton_stop.visible = True
    
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    
    Do While Dir(FilePath, vbDirectory) = vbNullString And i < 30
        DoEvents
        If StopNow Then GoTo slut
        Wait 1
        Label_progress.Caption = Label_progress.Caption & "*"
        i = i + 1
    Loop
    If i = 30 Then
        Label1.Caption = "Fejl ved download af GeoGebra"
        Label1.Caption = "Filen kunne ikke findes. (Det kan skyldes, at det tager meget lang tid at hente filen)"
        GoTo slut
    Else
        CommandButton_stop.visible = False
        s = RunScript("MoveGeoGebraToApplications", "")
        If s = "ok" Then
            RunScript "OpenApps", ""
            Label1.Caption = "Åben GeoGebra på følgende måde"
            Label2.Caption = "Apps skulle nu gerne være blevet åbnet med Finder." & vbCrLf & vbCrLf & "1. Hold 'Control' nede mens du klikker på 'GeoGebra'" & vbCrLf & "2. Klik Åben" & vbCrLf & "3. Klik Åben igen" & vbCrLf & "4. Gentag punkt 1+2 Så skulle GeoGebra gerne åbne, og WordMat vil fremover også kunne åbne GeoGebra." & vbCrLf & "4. Slut af med at klikke på knappen 'Test' herunder for at se om WordMat kan starte GeoGebra 5"
            Label2.visible = True
            CommandButton_test.visible = True
        End If
    End If

slut:
End Sub

