VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebra 
   Caption         =   "GeoGebra"
   ClientHeight    =   4035
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   6585
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

Private Sub CommandButton_Installer_Click()

'    explorersti = """" & Environ("ProgramFiles") & "\Internet Explorer\iexplore.exe"" http://www.geogebra.org/cms/da/installers"

'    appnr = Shell(explorersti, vbMaximizedFocus) 'vbNormalFocus vbMinimizedFocus
    ReturnVal = 1
    Me.hide
End Sub

Private Sub CommandButton_webstart_Click()
'Dim explorersti As String
'Dim appnr As Integer
'    OpenLink "https://www.geogebra.org/classic/"
'    explorersti = """" & Environ("ProgramFiles") & "\Internet Explorer\iexplore.exe"" http://www.geogebra.org/webstart/geogebra.html"
'    appnr = Shell(explorersti, vbMaximizedFocus) 'vbNormalFocus vbMinimizedFocus
    ReturnVal = 2
    Me.hide
End Sub


Private Sub UserForm_Activate()
#If Mac Then
#Else
#End If
    
    If Sprog.SprogNr = 1 Then
        Label_title.Caption = "GeoGebra 5 er ikke installeret"
#If Mac Then
        Label2.Caption = "Knappen downloader GeoGebra 5 til 'overførsler' på din Mac. Du skal efterfølgende åbne Finder og åbne GeoGebra med Control + Klik, for at tillade brugen."
#Else
        Label2.Caption = "Knappen sender dig til hjemmesiden, hvor du kan installere GeoGebra. WordMat på Windows understøtter 'GeoGebra classic 5', 'GeoGebra Calculator Suite', 'GeoGebra Classic 6' samt de fleste andre App-udgaver af GeoGebra."
#End If
    Else
        Label_title.Caption = "GeoGebra 5 is not installed"
    End If
    
    Label1.Caption = Sprog.A(292)
'    Label2.Caption = Sprog.A(293)
    Label3.Caption = Sprog.A(294)
    
    CommandButton_webstart.Caption = Sprog.A(296)
    
#If Mac Then
    CommandButton_Installer.Caption = Sprog.A(295) & " 5"
#Else
    CommandButton_Installer.Caption = Sprog.A(295)
#End If
    
    ReturnVal = 0
End Sub


