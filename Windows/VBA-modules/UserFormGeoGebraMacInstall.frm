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
#If Mac Then ' This form is only for Mac
Private StopNow As Boolean

Private Sub CommandButton_cancel_Click()
    Dim FilePath As String
    FilePath = GetProgramFilesDir & "GeoGebra.app"
    
    If Dir(FilePath, vbDirectory) = vbNullString Then
        MsgBox TT.A(849), vbYesNo, ""
        Me.hide
    End If
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
        MsgBox TT.A(850), vbOKOnly, TT.A(769)
        GoTo slut
    End If
    
    Me.hide
    GeoGebra
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
    Label1.Caption = TT.A(852)
    Label2.Caption = TT.A(853)
    i = 0
    
    FN = GetDownloadsFolder & "GeoGebra.app"
    If GrantAccessToMultipleFiles(Array(FN)) = "true" Then ' true if file does not exist
        If Dir(FN, vbDirectory) <> "" Then
            RunScript "RemoveApp", FN
        End If
    Else
        Label1.Caption = TT.A(854)
        Label2.Caption = TT.A(855)
        GoTo slut
    End If
    Label1.Caption = TT.A(856)
    Label2.Caption = TT.A(857)
    OpenLink "https://download.geogebra.org/package/mac", True
    AppActivate "Microsoft Word"
        
    FilePath = GetDownloadsFolder & "GeoGebra.app"
    
    Do While Dir(FilePath, vbDirectory) = vbNullString And i < 60 * 10
        DoEvents
        If StopNow Then
            Label1.Caption = TT.A(858)
            Label2.Caption = TT.A(859)
            Label_progress.Caption = ""
            CommandButton_stop.visible = False
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
            GoTo slut
        End If
        If GrantAccessToMultipleFiles(Array(FilePath)) = "false" Then
            Label1.Caption = TT.Error
            Label2.Caption = TT.A(860)
            CommandButton_stop.visible = False
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
            GoTo slut
        End If
        Wait 1
        Label_progress.Caption = Label_progress.Caption & "*"
        If i = 60 Then
            Label2.Caption = Label2.Caption & vbCrLf & TT.A(861)
        End If
        i = i + 1
    Loop
    If i >= 600 Then
        Label1.Caption = TT.A(862)
        Label2.Caption = TT.A(863)
        CommandButton_test.visible = False
        CommandButton_stop.visible = False
        GoTo slut
    Else
        CommandButton_stop.visible = False
        Label1.Caption = TT.A(864)
        Label2.Caption = TT.A(865)
        Label_progress.Caption = ""
        DoEvents
        If GrantAccessToMultipleFiles(Array(GetDownloadsFolder & "GeoGebra.app")) = "true" Then ' Even if there is only one file, it is an advantage to ask permission first, as you will then get a response back.
            Label1.Caption = TT.A(866)
            Label2.Caption = TT.A(867)
            s = RunScript("MoveGeoGebraToApplications", "") ' moves GeoGebra til Apps and removes quarantine
            If s = "ok" Then
                CommandButton_test.visible = True
                Label1.Caption = TT.A(868)
                Label2.Caption = TT.A(869)
            Else
            End If
        Else
            Label1.Caption = TT.Error
            Label2.Caption = TT.A(860)
            CommandButton_test.visible = False
            CommandButton_retry.visible = True
        End If
    End If

slut:
End Sub

Private Sub UserForm_Activate()
        CommandButton_cancel.Caption = TT.A(661)
        CommandButton_test.Caption = TT.A(851)
End Sub

#End If
