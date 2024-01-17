VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebra 
   Caption         =   "GeoGebra"
   ClientHeight    =   3510
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   8775.001
   OleObjectBlob   =   "UserFormGeoGebra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGeoGebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public ReturnVal As Integer ' 1=Install, 2=browser

Private Sub Label_Installer_Click()
    ReturnVal = 1
    Me.Hide
End Sub

Private Sub CommandButton_webstart_Click()
    ReturnVal = 2
    Me.Hide
End Sub


Private Sub Label_webstart_Click()
    CommandButton_webstart_Click
End Sub

Private Sub UserForm_Activate()
#If Mac Then
#Else
#End If
    
    If Sprog.SprogNr = 1 Then
        Label_title.Caption = "GeoGebra 5 er ikke installeret"
#If Mac Then
        Label2.Caption = "Knappen downloader GeoGebra 5 til 'overførsler' på din Mac. Du bliver efterfølgende guidet igennem opsætningen."
#Else
        Label2.Caption = "Knappen sender dig til hjemmesiden, hvor du kan installere GeoGebra. WordMat på Windows understøtter 'GeoGebra classic 5', 'GeoGebra Calculator Suite', 'GeoGebra Classic 6' samt de fleste andre App-udgaver af GeoGebra."
#End If
    Else
        Label_title.Caption = "GeoGebra 5 is not installed"
    End If
    
'    Label1.Caption = Sprog.A(292)
'    Label2.Caption = Sprog.A(293)
    Label3.Caption = Sprog.A(294)
    
    Label_webstart.Caption = Sprog.A(296)
    
#If Mac Then
    Label_Installer.Caption = Sprog.A(295) & " 5"
#Else
    Label_Installer.Caption = Sprog.A(295)
#End If
    
    ReturnVal = 0
End Sub

Private Sub Label_Installer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorPress
End Sub

Private Sub Label_Installer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorHover
End Sub

Private Sub Label_webstart_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_webstart.BackColor = LBColorPress
End Sub

Private Sub Label_webstart_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_webstart.BackColor = LBColorHover
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorInactive
    Label_webstart.BackColor = LBColorInactive
End Sub
