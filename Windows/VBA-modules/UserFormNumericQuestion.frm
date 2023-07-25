VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormNumericQuestion 
   Caption         =   "Numerisk løsning"
   ClientHeight    =   4710
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
Public result As String
Public variabel As String
Public FejlMeld As String
Public Ligning As String
Public Kommentar As String
Public Finished As Boolean

Private Sub CommandButton_cancel_Click()
    result = "afbryd"
    FejlMeld = ""
    Kommentar = ""
    omax.StopNow = True
    Finished = True
    If CASengine = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
    
'    If MaxProc.Finished = 0 Then
    Me.Hide
'    End If
End Sub

Private Sub CommandButton_nsolve_Click()
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
    result = "nsolve"
    FejlMeld = ""
    Kommentar = Sprog.FoundNumericSolutions
    Me.Hide
End Sub

Private Sub CommandButton_numeric_Click()
    If Not Finished Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
        Finished = True
        omax.MaximaOutput = Label_nsolve.Caption
    End If
'    omax.StopNow = True
    result = "num"
    FejlMeld = ""
    Kommentar = ""
    Me.Hide
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
    Me.Hide
End Sub

Private Sub UserForm_Activate()
'    PrepareMaximaNoSplash

    If CASengine > 0 Then
        Label2.Caption = ""
        Label_overskrift.Caption = "Hvordan vil du løse ligningen numerisk?"
        Label_omskrevet.Caption = ""
        Finished = True
        CommandButton_Omskrevet.visible = False
        CommandButton_nsolve.SetFocus
        Exit Sub
    End If
    Dim TempDefs As String
    TempDefs = omax.TempDefs
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
    If Finished Then GoTo slut
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
    If Finished Then GoTo slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -3, 3, 15, 5, 0, 0
    If Finished Then GoTo slut
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
    If Finished Then GoTo slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -6, 6, 15, 20, 0, 0
    If Finished Then GoTo slut
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
    If Finished Then GoTo slut
    omax.PrepareNewCommand
    omax.Kommando = Ligning
    omax.TempDefs = TempDefs
    omax.Nsolve variabel, -15, 15, 15, 20, 30, 30, True
    If Finished Then GoTo slut
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
    CommandButton_nsolve.SetFocus
slut:
    Finished = True
    
End Sub

Private Sub UserForm_Initialize()
    result = ""
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.NumericSolution
    CommandButton_nsolve.Caption = Sprog.InsertSolutions
    CommandButton_numeric.Caption = Sprog.SeekGraphically
    CommandButton_Omskrevet.Caption = Sprog.InsertCASresult
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_overskrift.Caption = Sprog.NoAnalyticResult
    
End Sub
