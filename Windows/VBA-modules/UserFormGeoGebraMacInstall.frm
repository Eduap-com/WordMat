VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebraMacInstall 
   Caption         =   "GeoGebra 5 Installation"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840.001
   OleObjectBlob   =   "UserFormGeoGebraMacInstall.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserFormGeoGebraMacInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If Mac Then ' Hele formen bruges kun på Mac
Private StopNow As Boolean

Private Sub CommandButton_cancel_Click()

End Sub

Private Sub CommandButton_cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_cancel_Click()
    Dim FilePath As String
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        If MsgBox("GeoGebra 5 er ikke blevet færdiginstalleret. Vil du afslutte?", vbYesNo, "Ikke færdig") = vbYes Then
            Me.hide
        End If
    End If
'    If Not TestDone Then
'        If MsgBox("GeoGebra 5 er ikke blevet testet. Vil du afslutte?", vbYesNo, "Ikke testet") = vbYes Then
'            Me.Hide
'        End If
'    End If

End Sub

Private Sub CommandButton_stop_Click()
    StopNow = True
End Sub

Private Sub CommandButton_test_Click()
    Dim FilePath As String
    
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        MsgBox "GeoGebra er ikke blevet installeret endnu. Vent til download er færdig.", vbOKOnly, "Vent"
        GoTo slut
    End If
    
    Me.hide
    GeoGebra
'    RunScript "OpenGeoGebra", geogebrafilersti
slut:
End Sub
Private Sub UserForm_Activate()
    Dim FilePath As String, i As Integer
    Dim s As String
    StopNow = False
    CommandButton_retry.visible = False
    CommandButton_test.visible = False
    CommandButton_stop.visible = True
    Label_progress.Caption = "*"
    Label1.Caption = "GeoGebra 5 er ved at blive hentet"
    Label2.Caption = "Vent først til GeoGebra er blevet downloadet..." & vbCrLf & "Så vil du blive bedt om at tildele adgang til GeoGebra."
    i = 0
    
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    If Dir(FilePath, vbDirectory) <> vbNullString Then GoTo slut ' Efter test kan klik på formen aktivere dette event igen, så skal der ikke ventes på download igen.
    
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    
    Do While Dir(FilePath, vbDirectory) = vbNullString And i < 30
        DoEvents
        If StopNow Then
            Label1.Caption = "Stoppet"
            Label2.Caption = "Processen blev stoppet." & vbCrLf & vbCrLf & "Klik Luk og prøv evt. igen en anden gang."
            Label_progress.Caption = ""
            CommandButton_stop.visible = False
            GoTo slut
        End If
        Wait 1
        Label_progress.Caption = Label_progress.Caption & "*"
        i = i + 1
    Loop
    If i = 30 Then
        Label1.Caption = "Fejl ved download af GeoGebra"
        Label2.Caption = "Filen kunne ikke findes i overførsler." & vbCrLf & "Du kan klikke 'Retry' forneden, hvis WordMat skal lede igen."
        CommandButton_test.visible = False
        CommandButton_stop.visible = False
        GoTo slut
    Else
        CommandButton_stop.visible = False
        Label1.Caption = "Giv tilladelse til GeoGebra"
        Label2.Caption = "Der skulle nu gerne åbne et vindue der hedder 'Giv adgang til fil'." & vbCrLf & vbCrLf & "1. Klik 'Vælg...' " & vbCrLf & "2. Klik 'Tildel adgang' "
        Label_progress.Caption = ""
        If GrantAccessToMultipleFiles(Array(GetDownloadsFolder & "GeoGebra.app")) = "true" Then ' selvom der kun er en fil, er det en fordel at spørge om lov først, da man så får svar tilbage
            s = RunScript("MoveGeoGebraToApplications", "") ' flytter GeoGebra til Apps og fjerner quarantine
            If s = "ok" Then
                CommandButton_test.visible = True
                Label1.Caption = "Færdig"
                Label2.Caption = "GeoGebra 5 er nu installeret." & vbCrLf & "Klik 'Fortsæt' for at åbne GeoGebra."
'                If RunScript("OpenApps", "") = "OK" Then
'                    CommandButton_test.visible = True
'                Else
'                    Label1.Caption = "Fejl"
'                    Label2.Caption = "Der er sket en fejl. Måske er GeoGebra ikke blevet hentet korrekt, eller du har ikke tildelt adgang"
'                    Label2.visible = True
'                    CommandButton_test.visible = False
'                    TestDone = True
'                End If
            Else
            End If
        Else
            Label1.Caption = "Fejl"
            Label2.Caption = "Du fik ikke givet tilladelse til GeoGebra. Klik på knappen 'prøv igen nedenfor'."
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
        End If
    End If

slut:
End Sub
#End If


