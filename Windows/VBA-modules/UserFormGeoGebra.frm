VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebra 
   Caption         =   "GeoGebra"
   ClientHeight    =   2640
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   5325
   OleObjectBlob   =   "UserFormGeoGebra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGeoGebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit

Private Sub CommandButton_Installer_Click()
'Dim explorersti As String
'Dim appnr As Integer
'    explorersti = """" & Environ("ProgramFiles") & "\Internet Explorer\iexplore.exe"" http://www.geogebra.org/webstart/geogebra.jnlp"
    MsgBox "The download page will now open. Install GeoGebra classic - not the graphing calculator"
#If Mac Then
    OpenLink "https://www.geogebra.org/download"
#Else
    OpenLink "https://www.geogebra.org/download"
#End If
'    explorersti = """" & Environ("ProgramFiles") & "\Internet Explorer\iexplore.exe"" http://www.geogebra.org/cms/da/installers"

'    appnr = Shell(explorersti, vbMaximizedFocus) 'vbNormalFocus vbMinimizedFocus
    Me.hide
End Sub

Private Sub CommandButton_webstart_Click()
'Dim explorersti As String
'Dim appnr As Integer
    OpenLink "https://www.geogebra.org/classic/"
'    explorersti = """" & Environ("ProgramFiles") & "\Internet Explorer\iexplore.exe"" http://www.geogebra.org/webstart/geogebra.html"
'    appnr = Shell(explorersti, vbMaximizedFocus) 'vbNormalFocus vbMinimizedFocus
    Me.hide
End Sub

Private Sub UserForm_Activate()
    Label1.Caption = Sprog.A(292)
    Label2.Caption = Sprog.A(293)
    Label3.Caption = Sprog.A(294)
    
    CommandButton_Installer.Caption = Sprog.A(295)
    CommandButton_webstart.Caption = Sprog.A(296)
    
End Sub


