VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormNumericQuestion 
   Caption         =   "Numerisk løsning"
   ClientHeight    =   4155
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10485
   OleObjectBlob   =   "UserFormNumericQuestion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormNumericQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsRunning As Boolean
Public result As String
Public variabel As String
Public FejlMeld As String
Public Ligning As String
Public Kommentar As String
Public Finished As Boolean
Public SorryNoAnalyticSol As Boolean ' Changes title

Private Sub Label_cancel_Click()
    result = "afbryd"
    FejlMeld = ""
    Kommentar = ""
    omax.StopNow = True
    Finished = True
#If Mac Then
#Else
    If CASengine = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
#End If

'    If MaxProc.Finished = 0 Then
    Me.hide
'    End If
End Sub

Private Sub Label_nsolveB_Click()
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
    result = "nsolve"
    FejlMeld = ""
    Kommentar = Sprog.FoundNumericSolutions
    Me.hide
End Sub

Private Sub Label_numeric_Click()
On Error GoTo Slut
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
'    omax.StopNow = True
Slut:
    result = "num"
    FejlMeld = ""
    Kommentar = ""
    Me.hide
End Sub

Private Sub CommandButton_Omskrevet_Click()
    omax.StopNow = True
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
    result = Sprog.A(180) '"omskriv"
    FejlMeld = ""
    Finished = True
    Kommentar = Sprog.NoNumericSolution
    omax.MaximaOutput = Label_omskrevet.Caption
    Me.hide
End Sub

Private Sub UserForm_Activate()
'    PrepareMaximaNoSplash
    Dim tid As Single
    If IsRunning Then Exit Sub
    IsRunning = True
    tid = Timer
#If Mac Then
    Label_numeric.visible = False ' grafisk løsning er gnuplot, som ikke længere understøttes på Mac.
#End If

    If CASengine > 0 Then
        Label2.Caption = ""
        Label_nsolve.Caption = Replace(Label_nsolve.Caption, "=", " = ")
        Label_overskrift.Caption = Sprog.NumericSolution
        Label_omskrevet.Caption = ""
        Finished = True
'        CommandButton_Omskrevet.visible = False
'        CommandButton_nsolve.SetFocus
        Exit Sub
    End If
    Dim TempDefs As String
    TempDefs = omax.TempDefs
    SetCaptions
    Finished = False
    Label_nsolveB.visible = False
    Label_nsolve.Caption = ""
'    If Label_omskrevet.Caption = "" Then
'        Label_omskrevet.visible = False
'        CommandButton_Omskrevet.visible = False
'    Else
'        Label_omskrevet.visible = True
'        CommandButton_Omskrevet.visible = True
'    End If
    DoEvents
    Label_fejlmeld.Caption = FejlMeld
    
    Label2.Caption = Sprog.SearchingSolutions
    omax.Kommando = Ligning
    omax.Nsolve variabel, 0, 2, 15, 1, 0, 0
    If omax.DefFejl Then
        Me.hide
        GoTo Slut
    End If
    If Finished Then GoTo Slut
    If omax.StopNow Or (Timer - tid) > 15 Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = Replace(omax.MaximaOutput, ChrW(8776), " " & ChrW(8776) & " ")
    
    If Len(omax.MaximaOutput) > 1 Then
        Label_nsolveB.visible = True
'        Label_nsolveB.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "..."
    Else
        Label2.Caption = Sprog.SearchingSolutions & "..."
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -3, 3, 15, 5, 0, 0
    If Finished Then GoTo Slut
    If omax.StopNow Or (Timer - tid) > 15 Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = Replace(omax.MaximaOutput, ChrW(8776), " " & ChrW(8776) & " ")
    If Len(omax.MaximaOutput) > 1 Then
        Label_nsolveB.visible = True
'        Label_nsolveB.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "......"
    Else
        Label2.Caption = Sprog.SearchingSolutions & "......"
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -6, 6, 15, 20, 0, 0
    If Finished Then GoTo Slut
    If omax.StopNow Or (Timer - tid) > 15 Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = Replace(omax.MaximaOutput, ChrW(8776), " " & ChrW(8776) & " ")
    If Len(omax.MaximaOutput) > 1 Then
        Label_nsolveB.visible = True
'        CommandButton_nsolve.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "........."
    Else
        Label2.Caption = Sprog.SearchingSolutions & "........."
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -15, 15, 15, 20, 30, 30, True
    If Finished Then GoTo Slut
    If omax.StopNow Or (Timer - tid) > 15 Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = Replace(omax.MaximaOutput, ChrW(8776), " " & ChrW(8776) & " ")
    If Len(omax.MaximaOutput) > 1 Then
        Label_nsolveB.visible = True
'        CommandButton_nsolve.SetFocus
    End If
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "..............."
    Else
        Label2.Caption = Sprog.SearchingSolutions & "..............."
    End If
afslut:
    omax.MaximaOutput = Label_nsolve.Caption
    
    If Len(Label_nsolve.Caption) > 0 Then
'        Label2.Caption = "Der blev fundet løsninger vha. automatiske numeriske metoder, men der findes måske flere. Du kan enten nøjes med de fundne løsninger eller forsøge at finde flere grafisk."
        Label2.Caption = Sprog.FoundNumericSolutions
        Label_nsolveB.visible = True
'        CommandButton_nsolve.SetFocus
    Else
        Label2.Caption = Sprog.NoNumericSolution2
    End If
'    CommandButton_nsolve.SetFocus
Slut:
    Finished = True
    IsRunning = False
End Sub

Private Sub UserForm_Initialize()
    result = ""
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.NumericSolution
    Label_nsolveB.Caption = Sprog.InsertSolutions
    Label_numeric.Caption = Sprog.SeekGraphically
'    CommandButton_Omskrevet.Caption = Sprog.InsertCASresult
    Label_cancel.Caption = Sprog.Cancel
    If SorryNoAnalyticSol Then
        Label_overskrift.Caption = Sprog.NoAnalyticResult
    Else
        Label_overskrift.Caption = Sprog.NumericSolution
    End If
    
    If Sprog.SprogNr = 1 Then
        Label_graphexplain.Caption = "Numerisk løsning giver ingen garanti for at alle løsninger er fundet. Overvej at forsøge grafisk løsning. Ved grafisk løsning viser man grafen for venstresiden og højresiden af ligningen i samme koordinatsystem. Løsningerne er så x-værdierne af grafernes skæringspunkter."
    Else
        Label_graphexplain.Caption = "Numerical solutions does not guarantee that alle solutions have been found. Consider trying graphical solution."
    End If
    
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_nsolveB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nsolveB.BackColor = LBColorPress
End Sub
Private Sub Label_nsolveB_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nsolveB.BackColor = LBColorHover
End Sub
Private Sub Label_numeric_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_numeric.BackColor = LBColorPress
End Sub
Private Sub Label_numeric_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_numeric.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nsolveB.BackColor = LBColorInactive
    Label_numeric.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub
