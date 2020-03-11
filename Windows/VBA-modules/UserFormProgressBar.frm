VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ProgressBar 
   Caption         =   "Please wait... (Press ESC to abort.)"
   ClientHeight    =   2580
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   7320
   OleObjectBlob   =   "UserFormProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================================================
' == XL TOOLBOX   (c) 2008-2010 Daniel Kraus   Licensed under GPLv3
' ========================================================================
' == Form_ProgressBar
' ==
' == A modeless progress bar.
' == Introduced in version 2.14.

' Usage:
' 1. StartProgress [Number of steps]
' 2. Info = "Informative text about the current action"
'    Optionally: Caption = "My caption"
' 3. MakeProgress for each step



Implements i_Form_ProgressBar
Option Explicit

    Private CurrentProgress As Long
    Private MaxProgress As Long
    Private StartTickCount As Long
    Private PreviousCancelMode As Long
    
    Private mAbort As Boolean
    Private mBusy As Boolean


Private Sub i_Form_ProgressBar_Reset()
    
    BlueBar.Width = 0
    CurrentProgress = 0
    MaxProgress = 0
    Label_RemainingTime.Caption = ""
    Label_Info.Caption = ""

End Sub

Private Sub i_Form_ProgressBar_Show()
    Me.Show
End Sub

Sub i_Form_progressBar_StartProgress(aMaxValue As Long)
    
    With Application
        PreviousCancelMode = .EnableCancelKey
        .EnableCancelKey = xlDisabled
    End With
    
    i_Form_ProgressBar_Reset
    StartTickCount = GetTickCount
    MaxProgress = aMaxValue
    If Not Me.Visible Then
        Me.Show
        DoEvents
    End If
End Sub



Sub i_Form_progressBar_MakeProgress(Optional Amount As Long = 1)
    
    Dim NewWidth As Long
    Dim RemainSec As Long
    Dim Elapsed As Long
        
    CurrentProgress = CurrentProgress + Amount
    If CurrentProgress > MaxProgress Then CurrentProgress = MaxProgress
    
    NewWidth = Round(CurrentProgress / MaxProgress * GrayBar.Width)
    
    If NewWidth <> BlueBar.Width Then ' Try to avoid unnecessary repainting of the progress bar to speed up the process
        
        With BlueBar
            .Width = NewWidth
            .Caption = " " & Round(CurrentProgress / MaxProgress * 100) & " %"
        End With
        
        With GrayBar
            .Caption = BlueBar.Caption
            If (CurrentProgress / MaxProgress) > 0.2 Then  ' Remaining time estimate is very inaccurate at the beginning, so let's
                    ' wait until 20 % of the task is completed before we show the time
                Elapsed = GetTickCount - StartTickCount
                RemainSec = (Elapsed / CurrentProgress) * (MaxProgress - CurrentProgress) / 1000
                Label_RemainingTime.Caption = Round(RemainSec / 60) & ":" & Format(RemainSec, "00")
            End If
        End With
        
        ' The doevents command is necessary to let the progress
        ' bar form repaint itself.
        DoEvents
    End If
End Sub



Property Get i_Form_progressBar_UserAborts() As Boolean
    Dim SuspendTicks As Long
    
    DoEvents
    
    SuspendTicks = GetTickCount
    
    If (GetAsyncKeyState(&H1B) < 0) Or mAbort Then
        If MsgBox("Are you sure you want to abort?", _
            vbInformation + vbYesNo, "User Abort") = vbYes Then
                i_Form_progressBar_UserAborts = True
                ' Version 2.40+: Removed the Unload me statement
                'Unload Me
        End If
        mBusy = True ' Suppress the Toggle_Abort_Click execution
        Toggle_Abort.Value = False
        mBusy = False
        mAbort = False
    End If
    StartTickCount = StartTickCount + (GetTickCount - SuspendTicks)
End Property



Property Get i_Form_progressBar_Progress() As Single
    If MaxProgress <> 0 Then i_Form_progressBar_Progress = CurrentProgress / MaxProgress
End Property



Property Let i_Form_progressBar_Info(ByRef aInfo As String)
    Label_Info = aInfo
End Property


Private Property Let i_Form_ProgressBar_Caption(RHS As String)
    Me.Caption = RHS
End Property




Private Sub Toggle_Abort_Click()
    ' The user should be allowed to "press" the button
    ' only once, but not "unpress" it.
    ' mBusy will be true if this method is called as a result
    ' of the Ttoggle_Abort value being changed within UserAborts.
    If Not mBusy Then
        Toggle_Abort.Value = True
        mAbort = True
    End If
End Sub


Private Sub UserForm_Initialize()
    Label_RemainingTime.Caption = ""
    With BlueBar
        .Caption = ""
        .Width = 0
    End With ' BlueBar
    With GrayBar
        .Caption = ""
    End With ' GrayBar
End Sub

Private Sub UserForm_Terminate()
    ' Restore the previous cancel mode
    ' when this object is destroyed
    Application.EnableCancelKey = PreviousCancelMode
End Sub
