VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormNumericQuestion 
   Caption         =   "Numerisk løsning"
   ClientHeight    =   4700
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   8955.001
   OleObjectBlob   =   "UserFormNumericQuestion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormNumericQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public Result As String
Public variabel As String
Public FejlMeld As String
Public Ligning As String
Public Kommentar As String
Public Finished As Boolean

Private Sub CommandButton_cancel_Click()
    Result = "afbryd"
    FejlMeld = ""
    Kommentar = ""
    omax.StopNow = True
    Finished = True
    MaxProc.CloseProcess
    MaxProc.StartMaximaProcess
    
'    If MaxProc.Finished = 0 Then
    Me.hide
'    End If
End Sub

Private Sub CommandButton_nsolve_Click()
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
    Result = "nsolve"
    FejlMeld = ""
    Kommentar = Sprog.FoundNumericSolutions
    Me.hide
End Sub

Private Sub CommandButton_numeric_Click()
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
'    omax.StopNow = True
    Result = "num"
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
    Result = Sprog.A(180) '"omskriv"
    FejlMeld = ""
    Finished = True
    Kommentar = Sprog.NoNumericSolution
    omax.MaximaOutput = Label_omskrevet.Caption
    Me.hide
End Sub

Private Sub UserForm_Activate()
'    PrepareMaximaNoSplash
    Dim tempDefs As String
    tempDefs = omax.tempDefs
    SetCaptions
    Finished = False
    CommandButton_nsolve.visible = False
    Label_nsolve.Caption = ""
    If Label_omskrevet.Caption = "" Then
        Label_omskrevet.visible = False
        CommandButton_Omskrevet.visible = False
    Else
        Label_omskrevet.visible = True
        CommandButton_Omskrevet.visible = True
    End If
    DoEvents
    Label_fejlmeld.Caption = FejlMeld
    
    Label2.Caption = Sprog.SearchingSolutions
    omax.Kommando = Ligning
    omax.Nsolve variabel, 0, 2, 15, 1, 0, 0
    If Finished Then GoTo Slut
    If omax.StopNow Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = omax.MaximaOutput
    If Len(omax.MaximaOutput) > 1 Then
        CommandButton_nsolve.visible = True
        CommandButton_nsolve.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "..."
    Else
        Label2.Caption = Sprog.SearchingSolutions & "..."
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.tempDefs = tempDefs
    omax.Nsolve variabel, -3, 3, 15, 5, 0, 0
    If Finished Then GoTo Slut
    If omax.StopNow Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = omax.MaximaOutput
    If Len(omax.MaximaOutput) > 1 Then
        CommandButton_nsolve.visible = True
        CommandButton_nsolve.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "......"
    Else
        Label2.Caption = Sprog.SearchingSolutions & "......"
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.tempDefs = tempDefs
    omax.Nsolve variabel, -6, 6, 15, 20, 0, 0
    If Finished Then GoTo Slut
    If omax.StopNow Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = omax.MaximaOutput
    If Len(omax.MaximaOutput) > 1 Then
        CommandButton_nsolve.visible = True
        CommandButton_nsolve.SetFocus
    End If
    
    If Len(Label_nsolve.Caption) > 1 Then
        Label2.Caption = Sprog.SearchingMoreSolutions & "........."
    Else
        Label2.Caption = Sprog.SearchingSolutions & "........."
    End If
    If Finished Then GoTo Slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.tempDefs = tempDefs
    omax.Nsolve variabel, -15, 15, 15, 20, 30, 30, True
    If Finished Then GoTo Slut
    If omax.StopNow Then GoTo afslut
    If Len(Label_nsolve.Caption) < Len(omax.MaximaOutput) Then Label_nsolve.Caption = omax.MaximaOutput
    If Len(omax.MaximaOutput) > 1 Then
        CommandButton_nsolve.visible = True
        CommandButton_nsolve.SetFocus
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
        CommandButton_nsolve.visible = True
        CommandButton_nsolve.SetFocus
    Else
        Label2.Caption = Sprog.NoNumericSolution2
    End If
Slut:
    Finished = True
    
End Sub

Private Sub UserForm_Initialize()
    Result = ""
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.NumericSolution
    CommandButton_nsolve.Caption = Sprog.InsertSolutions
    CommandButton_numeric.Caption = Sprog.SeekGraphically
    CommandButton_Omskrevet.Caption = Sprog.InsertCASresult
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_overskrift.Caption = Sprog.NoAnalyticResult
    
End Sub
