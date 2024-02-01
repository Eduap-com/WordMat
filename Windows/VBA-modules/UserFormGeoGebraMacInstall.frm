VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebraMacInstall 
   Caption         =   "GeoGebra 5 Installation"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
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
    Dim FilePath As String
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        If MsgBox("GeoGebra 5 er ikke blevet færdiginstalleret. Vil du afslutte?", vbYesNo, "Ikke færdig") = vbYes Then
            Me.Hide
        End If
    End If
'    If Not TestDone Then
'        If MsgBox("GeoGebra 5 er ikke blevet testet. Vil du afslutte?", vbYesNo, "Ikke testet") = vbYes Then
'            Me.Hide
'        End If
'    End If
    Unload Me

End Sub

Private Sub CommandButton_retry_Click()
    Me.DownloadNow
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
    
    Me.Hide
    GeoGebra
'    RunScript "OpenGeoGebra", geogebrafilersti
slut:
End Sub
Public Sub DownloadNow()
    Dim FilePath As String, i As Integer
    Dim s As String, FN As String
    StopNow = False
    CommandButton_retry.visible = False
    CommandButton_test.visible = False
    CommandButton_stop.visible = True
    Label_progress.Caption = "*"
    If Sprog.SprogNr = 1 Then
        Label1.Caption = "Slet gamle version"
        Label2.Caption = "Før den nye version kan hentes, skal den gamle version i Overførsler fjernes." & vbCrLf & "1. Der skulle gerne åbne et vindue, der hedder 'Giv adgang til fil'." & vbCrLf & "2. Klik 'Vælg...'  " & vbCrLf & "3. Klik 'Tildel adgang' "
    Else
        Label1.Caption = "Remove old version"
        Label2.Caption = "Before the new version can be downloaded, the old version in Downloads must be removed." & vbCrLf & "You will be asked to grant access to GeoGebra."
    End If
    i = 0
    
    FN = GetDownloadsFolder & "GeoGebra.app"
    If GrantAccessToMultipleFiles(Array(FN)) = "true" Then ' true hvis filen ikke finde
        If Dir(FN, vbDirectory) <> "" Then
            '                RmDir FN
            RunScript "RemoveApp", FN
        End If
    Else
        If Sprog.SprogNr = 1 Then
            Label1.Caption = "Download stoppet"
            Label2.Caption = "Du skal give adgang til GeoGebra.app før end installationen kan fortsætte"
        Else
            Label1.Caption = "Download stopped"
            Label2.Caption = "You must grant access to GeoGebra.app for the installation to proceed"
        End If
        GoTo slut
    End If
    If Sprog.SprogNr = 1 Then
        Label1.Caption = "GeoGebra 5 er ved at blive hentet"
        Label2.Caption = "1. Vent først til GeoGebra 5 er blevet downloadet." & vbCrLf & "2. Når download er gennemført skulle der gerne åbne et vindue, der hedder 'Giv adgang til fil'." & vbCrLf & "3. Klik 'Vælg...'  " & vbCrLf & "4. Klik 'Tildel adgang' "
    Else
        Label1.Caption = "Downloading GeoGebra 5"
        Label2.Caption = "Wait until download is complete." & vbCrLf & "When the download is complete you will be asked to grant access to GeoGebra."
    End If
    OpenLink "https://download.geogebra.org/package/mac", True
    AppActivate "Microsoft Word"
    
    '    FilePath = GetProgramFilesDir & "GeoGebra.app"
    '    If Dir(FilePath, vbDirectory) <> vbNullString Then GoTo slut ' Efter test kan klik på formen aktivere dette event igen, så skal der ikke ventes på download igen.
    
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    
    Do While Dir(FilePath, vbDirectory) = vbNullString And i < 60 * 10
        DoEvents
        If StopNow Then
            If Sprog.SprogNr = 1 Then
                Label1.Caption = "Stoppet"
                Label2.Caption = "Processen blev stoppet." & vbCrLf & vbCrLf & "Klik Luk og prøv evt. igen en anden gang."
            Else
                Label1.Caption = "Stopped"
                Label2.Caption = "The installation was stopped" & vbCrLf & vbCrLf & "Klik Close and try again."
            End If
            Label_progress.Caption = ""
            CommandButton_stop.visible = False
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
            GoTo slut
        End If
        If GrantAccessToMultipleFiles(Array(FilePath)) = "false" Then
            If Sprog.SprogNr = 1 Then
                Label1.Caption = "Fejl"
                Label2.Caption = "Du fik ikke givet tilladelse til GeoGebra. Klik på 'Retry' for at prøve igen."
            Else
                Label1.Caption = "Error"
                Label2.Caption = "You did not grant access to GeGebra. Click 'Retry' to try again."
            End If
            CommandButton_stop.visible = False
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
            GoTo slut
        End If
        Wait 1
        Label_progress.Caption = Label_progress.Caption & "*"
        If i = 60 Then
            If Sprog.SprogNr = 1 Then
                Label2.Caption = Label2.Caption & vbCrLf & "Download tager lang tid. Check i Safari om det ser ud til at der er en download igang. Du kan klikke stop hvis der ikke er fremgang."
            Else
                Label2.Caption = Label2.Caption & vbCrLf & "Download is slow. Check Safari if the download is progressing. You can click stop if there is no progress."
            End If
        End If
        i = i + 1
    Loop
    If i >= 600 Then
        If Sprog.SprogNr = 1 Then
            Label1.Caption = "Fejl ved download af GeoGebra"
            Label2.Caption = "Filen kunne ikke findes i overførsler." & vbCrLf & "Du kan klikke 'Retry' forneden, hvis WordMat skal lede igen."
        Else
            Label1.Caption = "Error downloading GeoGebra"
            Label2.Caption = "The file could not be located in Downloads" & vbCrLf & "You can click 'Retry' to try again"
        End If
        CommandButton_test.visible = False
        CommandButton_stop.visible = False
        GoTo slut
    Else
        CommandButton_stop.visible = False
        If Sprog.SprogNr = 1 Then
            Label1.Caption = "Giv tilladelse til GeoGebra"
            Label2.Caption = "Der skulle nu gerne åbne et vindue der hedder 'Giv adgang til fil'." & vbCrLf & vbCrLf & "1. Klik 'Vælg...' " & vbCrLf & "2. Klik 'Tildel adgang' "
        Else
            Label1.Caption = "Grant Access to GeoGebra"
            Label2.Caption = ""
        End If
        Label_progress.Caption = ""
        DoEvents
        If GrantAccessToMultipleFiles(Array(GetDownloadsFolder & "GeoGebra.app")) = "true" Then ' selvom der kun er en fil, er det en fordel at spørge om lov først, da man så får svar tilbage
            If Sprog.SprogNr = 1 Then
                Label1.Caption = "Flytter GeoGebra til Apps"
                Label2.Caption = "Du skal muligvis give tilladelse til 'System events' bare klik OK."
            Else
                Label1.Caption = "Moving GeoGebra to Apps"
                Label2.Caption = "You may be required to allow permission for 'System events'. Just click OK."
            End If
            s = RunScript("MoveGeoGebraToApplications", "") ' flytter GeoGebra til Apps og fjerner quarantine
            If s = "ok" Then
                CommandButton_test.visible = True
                If Sprog.SprogNr = 1 Then
                    Label1.Caption = "Færdig"
                    Label2.Caption = "GeoGebra 5 er nu installeret." & vbCrLf & "Klik 'Fortsæt' for at åbne GeoGebra."
                Else
                    Label1.Caption = "Installation complete!"
                    Label2.Caption = "GeoGebra 5 is now installed." & vbCrLf & "Click 'Continue' to open GeoGebra."
                End If
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
            If Sprog.SprogNr = 1 Then
                Label1.Caption = "Fejl"
                Label2.Caption = "Du fik ikke givet tilladelse til GeoGebra. Klik på 'Retry' for at prøve igen."
            Else
                Label1.Caption = "Error"
                Label2.Caption = "You did not grant access to GeGebra. Click 'Retry' to try again."
            End If
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
        End If
    End If

slut:
End Sub

Private Sub UserForm_Activate()
    If Sprog.SprogNr = 1 Then
        CommandButton_cancel.Caption = "Luk"
        CommandButton_test.Caption = "Fortsæt"
    Else
        CommandButton_cancel.Caption = "Close"
        CommandButton_test.Caption = "Continue"
    End If
End Sub

#End If
