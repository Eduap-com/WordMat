VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGraphFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If Mac Then
#Else
Private GraphText As String
Public CustomFunctions As String
Public funkno As Integer
Public pointno As Integer
Public relationno As Integer

Private Sub Class_Initialize()
    Dim DefAxes As String

    'MsgBox RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\Ivan\Graph\Property\DefaultLabelFont")
    DefAxes = RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\Ivan\Graph\DefaultAxes")
    If Len(DefAxes) > 1 Then
        GraphText = DefAxes
        GraphText = GraphText & "" & vbCrLf
    Else
        GraphText = ";This file was created by Graph (padowan.dk)" & vbCrLf
        'GraphText = GraphText & ";Do not change this file from other programs." & vbCrLf
        GraphText = GraphText & "[Graph]" & vbCrLf
        GraphText = GraphText & "Version = 4.3.0.384" & vbCrLf
        GraphText = GraphText & "MinVersion = 2.5" & vbCrLf
        GraphText = GraphText & "OS = Windows XP 5.1 Service Pack 2" & vbCrLf
        GraphText = GraphText & "" & vbCrLf

        GraphText = GraphText & "[Axes]" & vbCrLf
        GraphText = GraphText & "xmin = -10" & vbCrLf
        GraphText = GraphText & "xmax = 10" & vbCrLf
        GraphText = GraphText & "xTickUnit = 1" & vbCrLf
        GraphText = GraphText & "xGridUnit = 1" & vbCrLf
        GraphText = GraphText & "ymin = -10" & vbCrLf
        GraphText = GraphText & "ymax = 10" & vbCrLf
        GraphText = GraphText & "yTickUnit = 2" & vbCrLf
        GraphText = GraphText & "yGridUnit = 2" & vbCrLf
        'GraphText = GraphText & "AxesColor = clBlack" & vbCrLf
        'GraphText = GraphText & "GridColor = clBlack" & vbCrLf
        GraphText = GraphText & "AxesColor = clBlue" & vbCrLf
        GraphText = GraphText & "GridColor = 0x00FF9999" & vbCrLf
        GraphText = GraphText & "ShowLegend = 1" & vbCrLf
        GraphText = GraphText & "Radian = 1" & vbCrLf

        GraphText = GraphText & "" & vbCrLf
    End If
    funkno = 0
    pointno = 0
    relationno = 0
    'GraphText = GraphText & "" & vbCrLf
End Sub

Sub InsertDataSection()

    GraphText = GraphText & "[Data]" & vbCrLf
    GraphText = GraphText & "TextLabelCount = 0" & vbCrLf
    GraphText = GraphText & "FuncCount = " & funkno - 1 & vbCrLf
    GraphText = GraphText & "PointSeriesCount = " & pointno & vbCrLf
    GraphText = GraphText & "RelationCount = " & relationno & vbCrLf
    GraphText = GraphText & "ShadeCount = 0" & vbCrLf
    GraphText = GraphText & "RelationCount = 0" & vbCrLf
    GraphText = GraphText & "OleObjectCount = 0" & vbCrLf

    GraphText = GraphText & "" & vbCrLf

End Sub

Public Sub Save(FilNavn As String)
    Dim filnr As Integer
    On Error GoTo Fejl

    InsertDataSection
    InsertCustomFunctions

    filnr = FreeFile()
    Open FilNavn For Output As filnr    ' Open file for output.

    Print #filnr, GraphText  ' print skriver uden " "
    'Print #filnr, "Hello World"  ' print skriver uden " "

    'Write #filnr, "Hello World", 234    ' Write comma-delimited data.
    'Write #filnr,    ' Write blank line.
    'Write #filnr, MyError; " is an error value"

    GoTo slut
Fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
    Close #filnr    ' Close file.

End Sub

Sub InsertFunction(fkt As String, Optional visible As Integer = 1)
    fkt = ConvertToGraphSyntax(fkt, True)
    funkno = funkno + 1

    GraphText = GraphText & "[Func" & funkno & "]" & vbCrLf
    GraphText = GraphText & "FuncType = 0" & vbCrLf
    GraphText = GraphText & "y = " & fkt & vbCrLf
    GraphText = GraphText & "Visible = " & visible & vbCrLf
    GraphText = GraphText & "Color = clRed" & vbCrLf
    GraphText = GraphText & "" & vbCrLf

End Sub
Sub InsertPointSeries(serie As String)

    pointno = pointno + 1

    GraphText = GraphText & "[PointSeries" & pointno & "]" & vbCrLf
    GraphText = GraphText & "FillColor = clBlack" & vbCrLf
    GraphText = GraphText & "LineColor = clBlue" & vbCrLf
    GraphText = GraphText & "Size = 3" & vbCrLf
    GraphText = GraphText & "Style = 0" & vbCrLf
    GraphText = GraphText & "LabelPosition = 2" & vbCrLf
    'GraphText = GraphText & "Points = 32,1.93;40,1.67;50,1.4;60,1.22;70,1.06;80,0.93;90,0.825;100,0.74;150,0.477;200,0.341;250,0.269;300,0.22;350,0.189;400,0.17;450,0.155;500,0.145;550,0.139;600,0.137;"
    GraphText = GraphText & "Points = " & serie & vbCrLf
    GraphText = GraphText & "LegendText = Serie 1" & vbCrLf
    GraphText = GraphText & "" & vbCrLf

End Sub
Sub InsertCustomFunctions()
If Len(CustomFunctions) > 0 Then
    GraphText = GraphText & "[CustomFunctions]" & vbCrLf
    GraphText = GraphText & CustomFunctions & vbCrLf
    GraphText = GraphText & "" & vbCrLf
End If

End Sub

Sub AddCustomFunction(fkt As String)
    fkt = ConvertToGraphSyntax(fkt, False)
    CustomFunctions = CustomFunctions & fkt & vbCrLf
End Sub

Sub InsertRelation(relation As String)
    relation = ConvertToGraphSyntax(relation, True)
    relationno = relationno + 1
    
    GraphText = GraphText & "[Relation" & relationno & "]" & vbCrLf
    GraphText = GraphText & "Relation = " & relation & vbCrLf
    GraphText = GraphText & "Style = 5" & vbCrLf
    GraphText = GraphText & "Color = clGreen" & vbCrLf
    GraphText = GraphText & "Size = 1" & vbCrLf
    GraphText = GraphText & "" & vbCrLf

End Sub

Function ConvertToGraphSyntax(Text As String, Optional ConvertMaxima As Boolean = True)
    If ConvertMaxima Then
        On Error Resume Next
        omax.ConvertLnLog = False
        Text = omax.CodeForMaxima(Text)
        omax.ConvertLnLog = True
    End If
    Text = Replace(Text, "%e", "e")
    Text = Replace(Text, "%pi", "pi")
    Text = Replace(Text, "%", "")

    ConvertToGraphSyntax = Text
End Function
#End If

