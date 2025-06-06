VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3DGraph 
   Caption         =   "3D grafer"
   ClientHeight    =   5235
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm3DGraph.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3DGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form is used for 3d plots using gnuplot for GeoGebra

Private palindex As Integer
Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
    For Each c In ControlColl ' Me.Controls
        TN = TypeName(c)
        If TN = "CheckBox" Then
            Set CE = New CEvents: Set CE.CheckBoxControl = c: EventsCol.Add CE
        ElseIf TN = "OptionButton" Then
            Set CE = New CEvents: Set CE.OptionButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ComboBox" Then
            Set CE = New CEvents: Set CE.ComboBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Label" Then
            Set CE = New CEvents: Set CE.LabelControl = c: EventsCol.Add CE
        ElseIf TN = "TextBox" Then
            Set CE = New CEvents: Set CE.TextBoxControl = c: EventsCol.Add CE
        ElseIf TN = "CommandButton" Then
            Set CE = New CEvents: Set CE.CommandButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ListBox" Then
            Set CE = New CEvents: Set CE.ListBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub CommandButton_geogebra_Click()
    Dim s As String, vekt As String, Arr() As String, i As Integer, j As Integer
    Dim ea As New ExpressionAnalyser, punkttekst As String, parx As String, pary As String, parz As String, cmd As String
    Dim sl As New CSortList, DefList As String, Var As String, k As Integer, fktudtryk As String, UrlLink As String, p As Integer
    
    ea.SetNormalBrackets

    'definitions
    For i = 0 To omax.defindex - 1
        DefList = DefList & "," & omax.DefName(i)
        ea.Text = omax.DefValue(i)
        Var = ea.GetNextVar
        If Var = "" Then
            sl.Add omax.DefName(i), omax.DefValue(i), 0
        Else
            k = 0
            For j = 0 To sl.Length - 1
                ea.Text = sl.GetVal(j)
                If ea.ContainsVar(omax.DefName(i)) Then
                    Exit For
                End If
                k = k + 1
            Next
            sl.Add omax.DefName(i), omax.DefValue(i), k
        End If
    Next

    ' define variables not defined
    omax.FindVariable
    ea.Text = DefList
    For i = 0 To sl.Length - 1
        fktudtryk = ReplaceIndepvarX(sl.GetVal(i))
        If sl.GetVal(i) <> ReplacedVar Then
            DefinerKonstanter sl.GetVal(i), DefList, Nothing, UrlLink
            p = InStr(sl.GetName(i), "(")
            If p > 0 Then
                cmd = Left(sl.GetName(i), p) & Replace(sl.GetName(i), ReplacedVar, "x", p + 1) & "=" & fktudtryk
            Else
                cmd = sl.GetName(i) & "=" & fktudtryk
            End If
        Else
            cmd = sl.GetName(i) & "=" & fktudtryk
        End If
        cmd = Replace(ConvertToGeogebraSyntax(cmd), "+", "%2B") & ";"
        UrlLink = UrlLink & cmd
    Next


    s = UrlLink

    ' functions
    If TextBox_forskrift1.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift1.Text) & ";"
    End If
    If TextBox_forskrift2.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift2.Text) & ";"
    End If
    If TextBox_forskrift3.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift3.Text) & ";"
    End If

    'equations
    If TextBox_ligning1.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning1.Text) & ";"
    End If
    If TextBox_ligning2.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning2.Text) & ";"
    End If
    If TextBox_ligning3.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning3.Text) & ";"
    End If
    
    'vectors
    If TextBox_vektorer.Text <> "" Then
        vekt = TextBox_vektorer.Text
        Arr = Split(vekt, VbCrLfMac)
        For i = 0 To UBound(Arr)
            If Arr(i) <> "" Then
                If InStr(Arr(i), ";") > 0 Then
                    Arr(i) = Replace(Arr(i), ",", ".")
                    Arr(i) = Replace(Arr(i), ";", ",")
                Else
                    ea.Text = Arr(i)
                    j = ea.CountText(",")
                    If Not (j = 2 Or j = 4) Then
                        ea.ConvertDecSeparator
                        Arr(i) = ea.Text
                    End If
                End If
                If InStr(Arr(i), ")(") > 0 Then
                    Arr(i) = Replace(Arr(i), ")(", "),(")
                Else
                    Arr(i) = "(0,0,0)," & Arr(i)
                End If
                Arr(i) = Replace(Arr(i), "(", "(")
                Arr(i) = Replace(Arr(i), ")", ")")
            
                s = s & "vector(" & Arr(i) & ");"
            End If
        Next
    End If

    'parametric plots
    If TextBox_parametric1x.Text <> "" Then
        parx = ConvertToGeogebraSyntax(TextBox_parametric1x.Text)
        pary = ConvertToGeogebraSyntax(TextBox_parametric1y.Text)
        parz = ConvertToGeogebraSyntax(TextBox_parametric1z.Text)
        s = s & "(" & parx & " , " & pary & " , " & parz & ");"
    End If
    If TextBox_parametric2x.Text <> "" Then
        parx = ConvertToGeogebraSyntax(TextBox_parametric2x.Text)
        pary = ConvertToGeogebraSyntax(TextBox_parametric2y.Text)
        parz = ConvertToGeogebraSyntax(TextBox_parametric2z.Text)
        s = s & "(" & parx & " , " & pary & " , " & parz & ");"
    End If
    If TextBox_parametric3x.Text <> "" Then
        parx = ConvertToGeogebraSyntax(TextBox_parametric3x.Text)
        pary = ConvertToGeogebraSyntax(TextBox_parametric3y.Text)
        parz = ConvertToGeogebraSyntax(TextBox_parametric3z.Text)
        s = s & "(" & parx & " , " & pary & " , " & parz & ");"
    End If

    'points
    If TextBox_punkter.Text <> "" Then
        punkttekst = TextBox_punkter.Text
        If InStr(punkttekst, ";") > 0 Then
            punkttekst = Replace(punkttekst, ",", ".")
            punkttekst = Replace(punkttekst, ";", ",")
        End If
        punkttekst = Replace(punkttekst, ")(", ");(")
        punkttekst = Replace(punkttekst, vbCrLf, ";")
        punkttekst = Replace(punkttekst, vbCr, ";")
        punkttekst = Replace(punkttekst, " ", "")
        If Right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
        s = s & punkttekst & ";"
    End If
    s = Left(s, Len(s) - 1)
    
    OpenGeoGebraWeb s, "3d", True, False
    
End Sub

Private Sub CommandButton_insertplan_Click()
Dim plan As String
    plan = "1*(x-0)+1*(y-0)+1*(z-0)=0"
    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = plan
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = plan
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = plan
    End If
End Sub

Private Sub CommandButton_kugle_Click()
Dim kugle As String
    kugle = "(x-0)^2+(y-0)^2+(z-0)^2=1^2"
    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = kugle
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = kugle
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = kugle
    End If
End Sub

Private Sub CommandButton_nulstilalt_Click()
    TextBox_forskrift1.Text = ""
    TextBox_forskrift2.Text = ""
    TextBox_forskrift3.Text = ""
    TextBox_ligning1.Text = ""
    TextBox_ligning2.Text = ""
    TextBox_ligning3.Text = ""
    TextBox_vektorer.Text = ""
    TextBox_parametric1x.Text = ""
    TextBox_parametric1y.Text = ""
    TextBox_parametric1z.Text = ""
    TextBox_tmin1.Text = ""
    TextBox_tmax1.Text = ""
    TextBox_parametric2x.Text = ""
    TextBox_parametric2y.Text = ""
    TextBox_parametric2z.Text = ""
    TextBox_tmin2.Text = ""
    TextBox_tmax2.Text = ""
    TextBox_parametric3x.Text = ""
    TextBox_parametric3y.Text = ""
    TextBox_parametric3z.Text = ""
    TextBox_tmin3.Text = ""
    TextBox_tmax3.Text = ""
    TextBox_punkter.Text = ""
End Sub

Private Sub CommandButton_nulstilforsk1_Click()
    TextBox_forskrift1.Text = ""
End Sub
Private Sub CommandButton_nulstilforsk2_Click()
    TextBox_forskrift2.Text = ""
End Sub

Private Sub CommandButton_nulstilforsk3_Click()
    TextBox_forskrift3.Text = ""
End Sub

Private Sub CommandButton_nulstillign1_Click()
    TextBox_ligning1.Text = ""
End Sub

Private Sub CommandButton_nulstillign2_Click()
    TextBox_ligning2.Text = ""
End Sub

Private Sub CommandButton_nulstillign3_Click()
    TextBox_ligning3.Text = ""
End Sub

Private Sub CommandButton_nulstilpar1_Click()
TextBox_parametric1x.Text = ""
TextBox_parametric1y.Text = ""
TextBox_parametric1z.Text = ""
TextBox_tmin1.Text = ""
TextBox_tmax1.Text = ""
TextBox_smin1.Text = ""
TextBox_smax1.Text = ""
End Sub
Private Sub CommandButton_nulstilpar2_Click()
TextBox_parametric2x.Text = ""
TextBox_parametric2y.Text = ""
TextBox_parametric2z.Text = ""
TextBox_tmin2.Text = ""
TextBox_tmax2.Text = ""
TextBox_smin2.Text = ""
TextBox_smax2.Text = ""
End Sub
Private Sub CommandButton_nulstilpar3_Click()
TextBox_parametric3x.Text = ""
TextBox_parametric3y.Text = ""
TextBox_parametric3z.Text = ""
TextBox_tmin3.Text = ""
TextBox_tmax3.Text = ""
TextBox_smin3.Text = ""
TextBox_smax3.Text = ""
End Sub

Private Sub CommandButton_nulstilpunkter_Click()
    TextBox_punkter.Text = ""
End Sub

Private Sub CommandButton_nulstilvektorer_Click()
    TextBox_vektorer.Text = ""
End Sub

Private Sub CommandButton_nyparplan_Click()
Dim px As String
Dim py As String
Dim pz As String
px = "0+1*s+1*t"
py = "0+1*s-1*t"
pz = "0+1*s+1*t"

If TextBox_parametric1x.Text = "" Then
    TextBox_parametric1x.Text = px
    TextBox_parametric1y.Text = py
    TextBox_parametric1z.Text = pz
    TextBox_tmin1.Text = "0"
    TextBox_tmax1.Text = "1"
    TextBox_smin1.Text = "0"
    TextBox_smax1.Text = "1"
ElseIf TextBox_parametric2x.Text = "" Then
    TextBox_parametric2x.Text = px
    TextBox_parametric2y.Text = py
    TextBox_parametric2z.Text = pz
    TextBox_tmin2.Text = "0"
    TextBox_tmax2.Text = "1"
    TextBox_smin2.Text = "0"
    TextBox_smax2.Text = "1"
ElseIf TextBox_parametric3x.Text = "" Then
    TextBox_parametric3x.Text = px
    TextBox_parametric3y.Text = py
    TextBox_parametric3z.Text = pz
    TextBox_tmin3.Text = "0"
    TextBox_tmax3.Text = "1"
    TextBox_smin3.Text = "0"
    TextBox_smax3.Text = "1"
End If

End Sub

Private Sub CommandButton_nyvektor_Click()
    If TextBox_vektorer.Text <> "" Then
        If Right(TextBox_vektorer.Text, 1) = ")" Then
            TextBox_vektorer.Text = TextBox_vektorer.Text & vbCr
        End If
    End If
    TextBox_vektorer.Text = TextBox_vektorer.Text & "(0 , 0 , 0)(1 , 1 , 1)"
End Sub

Private Sub CommandButton_ok_Click()
Dim grafobj As String, s As String
Dim vekt As String
Dim lign As String
Dim parx As String
Dim pary As String
Dim parz As String
Dim xmin As String
Dim xmax As String
Dim Ymin As String
Dim Ymax As String
Dim zmin As String
Dim zmax As String
Dim tmin As String
Dim tmax As String
Dim smin As String
Dim smax As String
Dim punkttekst As String
Dim antalobj As Integer
Dim Arr As Variant
Dim i As Integer, j As Integer
Dim ea As New ExpressionAnalyser

On Error GoTo fejl
'Me.Hide
Label_vent.visible = True
colindex = 0
palindex = 0

xmin = TextBox_xmin.Text
xmax = TextBox_xmax.Text
Ymin = TextBox_ymin.Text
Ymax = TextBox_ymax.Text
zmin = TextBox_zmin.Text
zmax = TextBox_zmax.Text

'functions
If TextBox_forskrift1.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_forskrift1.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & "),"
    antalobj = antalobj + 1
End If
If TextBox_forskrift2.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_forskrift2.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & "),"
    antalobj = antalobj + 1
End If
If TextBox_forskrift3.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_forskrift3.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & "),"
    antalobj = antalobj + 1
End If

'Equations
If TextBox_ligning1.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_ligning1.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    If OptionButton_kunflader.Value Then
        grafobj = grafobj & "enhanced3d=[" & GetNextEnhFunc() & ", x, y, z],"
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & ",z," & zmin & "," & zmax & "),"
    antalobj = antalobj + 1
End If
If TextBox_ligning2.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_ligning2.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    If OptionButton_kunflader.Value Then
        grafobj = grafobj & "enhanced3d=[" & GetNextEnhFunc() & ", x, y, z],"
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & ",z," & zmin & "," & zmax & "),"
    antalobj = antalobj + 1
End If
If TextBox_ligning3.Text <> "" Then
    lign = omax.CodeForMaxima(TextBox_ligning3.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""" & lign & ""","
    End If
    If OptionButton_kunflader.Value Then
        grafobj = grafobj & "enhanced3d=[" & GetNextEnhFunc() & ", x, y, z],"
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xmin & "," & xmax & ",y," & Ymin & "," & Ymax & ",z," & zmin & "," & zmax & "),"
    antalobj = antalobj + 1
End If


'parametric plots
If TextBox_parametric1x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric1x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric1y.Text)
    parz = omax.CodeForMaxima(TextBox_parametric1z.Text)
    tmin = omax.CodeForMaxima(TextBox_tmin1.Text)
    tmax = omax.CodeForMaxima(TextBox_tmax1.Text)
    smin = omax.CodeForMaxima(TextBox_smin1.Text)
    smax = omax.CodeForMaxima(TextBox_smax1.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""paramterfremstilling: (" & parx & "," & pary & "," & parz & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ","
    If smin <> "" Then
        grafobj = grafobj & "parametric_surface(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & ",s," & smin & "," & smax & "),"
    Else
        grafobj = grafobj & "parametric(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & "),"
    End If
    antalobj = antalobj + 1
End If
If TextBox_parametric2x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric2x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric2y.Text)
    parz = omax.CodeForMaxima(TextBox_parametric2z.Text)
    tmin = TextBox_tmin2.Text
    tmax = TextBox_tmax2.Text
    smin = TextBox_smin2.Text
    smax = TextBox_smax2.Text
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""paramterfremstilling: (" & parx & "," & pary & "," & parz & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ","
    If smin <> "" Then
        grafobj = grafobj & "parametric_surface(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & ",s," & smin & "," & smax & "),"
    Else
        grafobj = grafobj & "parametric(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & "),"
    End If
    antalobj = antalobj + 1
End If
If TextBox_parametric3x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric3x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric3y.Text)
    parz = omax.CodeForMaxima(TextBox_parametric3z.Text)
    tmin = omax.CodeForMaxima(TextBox_tmin3.Text)
    tmax = omax.CodeForMaxima(TextBox_tmax3.Text)
    smin = omax.CodeForMaxima(TextBox_smin3.Text)
    smax = omax.CodeForMaxima(TextBox_smax3.Text)
    If CheckBox_udtryk.Value Then
        grafobj = grafobj & "key=""paramterfremstilling: (" & parx & "," & pary & "," & parz & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "color=" & GetNextColor & ","
    If smin <> "" Then
        grafobj = grafobj & "parametric_surface(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & ",s," & smin & "," & smax & "),"
    Else
        grafobj = grafobj & "parametric(" & parx & "," & pary & "," & parz & ",t," & tmin & "," & tmax & "),"
    End If
    antalobj = antalobj + 1
End If

'vectors
If TextBox_vektorer.Text <> "" Then
    If antalobj = 0 Then
        grafobj = grafobj & "surface_hide = false,"
    End If
    vekt = TextBox_vektorer.Text
    Arr = Split(vekt, VbCrLfMac)
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
            If InStr(Arr(i), ";") > 0 Then
                Arr(i) = Replace(Arr(i), ",", ".")
                Arr(i) = Replace(Arr(i), ";", ",")
            Else
                ea.Text = Arr(i)
                j = ea.CountText(",")
                If Not (j = 2 Or j = 4) Then
                    ea.ConvertDecSeparator
                    Arr(i) = ea.Text
                End If
            End If
            If InStr(Arr(i), ")(") > 0 Then
                Arr(i) = Replace(Arr(i), ")(", "],[")
            Else
                Arr(i) = "[0,0,0]," & Arr(i)
            End If
            Arr(i) = Replace(Arr(i), "(", "[")
            Arr(i) = Replace(Arr(i), ")", "]")
            
            If CheckBox_udtryk.Value Then
                grafobj = grafobj & "key=""Vektor: " & Arr(i) & ""","
            Else
                grafobj = grafobj & "key="""","
            End If
            grafobj = grafobj & "color=" & GetNextColor & ","
            grafobj = grafobj & "vector(" & Arr(i) & "),"
        End If
    Next
    antalobj = antalobj + 1
End If

'points
If TextBox_punkter.Text <> "" Then
    punkttekst = TextBox_punkter.Text
    If InStr(punkttekst, ";") > 0 Then
        punkttekst = Replace(punkttekst, ",", ".")
        punkttekst = Replace(punkttekst, ";", ",")
    End If
    punkttekst = Replace(punkttekst, ")(", "],[")
    punkttekst = Replace(punkttekst, "),(", "],[")
    punkttekst = Replace(punkttekst, "(", "[")
    punkttekst = Replace(punkttekst, ")", "]")
    punkttekst = Replace(punkttekst, vbCrLf, ",")
    punkttekst = Replace(punkttekst, vbCr, ",")
    punkttekst = Replace(punkttekst, " ", "")
    If Right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    grafobj = grafobj & "key="""","
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & TextBox_pointsize.Text & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ",points([" & punkttekst & "]),"
    antalobj = antalobj + 1
End If



If grafobj = "" Then GoTo slut

If OptionButton_kunlinjer Then
    If Not CheckBox_transp.Value Then
        grafobj = "surface_hide = true," & grafobj
    End If
    grafobj = "user_preamble = ""set xyplane at 0""," & grafobj 'palette=gray,
ElseIf OptionButton_kunflader Then
    grafobj = "user_preamble = ""set xyplane at 0"",palette=color,enhanced3d=[z, x, y, z]," & grafobj 'palette=gray,
ElseIf OptionButton_linjerogflader Then
    grafobj = "user_preamble = ""set pm3d at s depthorder;set xyplane at 0""," & grafobj 'palette=gray,
End If

If CheckBox_grid.Value Then
    grafobj = grafobj & "xaxis_width = 2,xaxis_color = orange,xaxis_type  = solid,xaxis=true,yaxis_width = 2,yaxis_color = orange,yaxis_type=solid,yaxis=true,zaxis_width = 2,zaxis_color = blue,zaxis_type=solid,zaxis=true,grid=true,user_preamble = ""set xyplane at 0"","
End If

If ComboBox_kvalitet.ListIndex = 0 Then 'super
    grafobj = "xu_grid=200,yv_grid=200,x_voxel=18,y_voxel=18,z_voxel=18," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 1 Then 'very high
    grafobj = "xu_grid=100,yv_grid=100,x_voxel=15,y_voxel=15,z_voxel=15," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 2 Then ' high
    grafobj = "xu_grid=50,yv_grid=50,x_voxel=12,y_voxel=12,z_voxel=12," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 4 Then 'low
    grafobj = "xu_grid=15,yv_grid=15,x_voxel=5,y_voxel=5,z_voxel=5," & grafobj
End If


If TextBox_titel.Text <> "" Then
    grafobj = grafobj & "title=""" & TextBox_titel.Text & ""","
End If
'grafobj = grafobj & "user_preamble = ""set pm3d at s depthorder;set xyplane at 0"",colorbox=false,surface_hide = true"

If ComboBox_farver.ListIndex = 0 Then ' standard
    grafobj = grafobj & "palette=color,"
ElseIf ComboBox_farver.ListIndex = 1 Then ' bluw
    grafobj = grafobj & "palette=[4,5,7],"
ElseIf ComboBox_farver.ListIndex = 2 Then ' brown
    grafobj = grafobj & "palette=[4,5,6],"
ElseIf ComboBox_farver.ListIndex = 3 Then ' Grey
    grafobj = grafobj & "palette=gray,"
End If

grafobj = "font=""Arial"",font_size=8," & grafobj
grafobj = grafobj & "colorbox=false"

If CheckBox_maximakommando Then
    omax.MaximaOutput = "draw3d(" & grafobj & ")"
    omax.InsertMaximaOutput
End If
    s = "set_draw_defaults("
    If TextBox_xmin.Text <> vbNullString And TextBox_xmax.Text <> vbNullString Then
        s = s & "xrange=[" & xmin & "," & xmax & "],"
    End If
    If TextBox_ymin.Text <> vbNullString And TextBox_ymax.Text <> vbNullString Then
        s = s & "yrange=[" & Ymin & "," & Ymax & "],"
    End If
    If TextBox_zmin.Text <> vbNullString And TextBox_zmax.Text <> vbNullString Then
        s = s & "zrange=[" & zmin & "," & zmax & "],"
    End If
    If Right(s, 1) = "," Then s = Left(s, Len(s) - 1)
    grafobj = s & ")," & grafobj
    omax.Draw3D grafobj, antalobj
    omax.PrepareNewCommand
    Label_vent.visible = False
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Sub
Function GetNextEnhFunc() As String
palindex = palindex + 1
If palindex = 1 Then
    GetNextEnhFunc = "z"
ElseIf palindex = 2 Then
    GetNextEnhFunc = "-z"
ElseIf palindex = 3 Then
    GetNextEnhFunc = "1.5*z"
Else
    GetNextEnhFunc = "z*x*y"
End If
End Function

Private Sub CommandButton_parlinje_Click()
Dim px As String
Dim py As String
Dim pz As String
px = "0+1*t"
py = "0+1*t"
pz = "0+1*t"

If TextBox_parametric1x.Text = "" Then
    TextBox_parametric1x.Text = px
    TextBox_parametric1y.Text = py
    TextBox_parametric1z.Text = pz
    TextBox_tmin1.Text = "0"
    TextBox_tmax1.Text = "1"
    TextBox_smin1.Text = ""
    TextBox_smax1.Text = ""
ElseIf TextBox_parametric2x.Text = "" Then
    TextBox_parametric2x.Text = px
    TextBox_parametric2y.Text = py
    TextBox_parametric2z.Text = pz
    TextBox_tmin2.Text = "0"
    TextBox_tmax2.Text = "1"
    TextBox_smin2.Text = ""
    TextBox_smax2.Text = ""
ElseIf TextBox_parametric3x.Text = "" Then
    TextBox_parametric3x.Text = px
    TextBox_parametric3y.Text = py
    TextBox_parametric3z.Text = pz
    TextBox_tmin3.Text = "0"
    TextBox_tmax3.Text = "1"
    TextBox_smin3.Text = ""
    TextBox_smax3.Text = ""
End If

End Sub

Private Sub OptionButton_kunflader_Click()
    ComboBox_farver.Enabled = True
    CheckBox_transp.Enabled = False
End Sub

Private Sub OptionButton_kunlinjer_Click()
    ComboBox_farver.Enabled = False
    CheckBox_transp.Enabled = True
End Sub

Private Sub OptionButton_linjerogflader_Click()
    ComboBox_farver.Enabled = True
    CheckBox_transp.Enabled = False
End Sub

Private Sub TextBox_forskrift1_Change()
    UpdateBoldTabs
End Sub

Sub UpdateBoldTabs()
    If TextBox_forskrift1.Text <> vbNullString Or TextBox_forskrift2.Text <> vbNullString Or TextBox_forskrift3.Text <> vbNullString Then
        MultiPage1.Pages(0).Caption = TT.A(319) & "*"
    Else
        MultiPage1.Pages(0).Caption = TT.A(319)
    End If
    If TextBox_ligning1.Text <> vbNullString Or TextBox_ligning2.Text <> vbNullString Or TextBox_ligning3.Text <> vbNullString Then
        MultiPage1.Pages(1).Caption = TT.A(834) & "*"
    Else
        MultiPage1.Pages(1).Caption = TT.A(834)
    End If
    If TextBox_vektorer.Text <> vbNullString Then ' vectors
        MultiPage1.Pages(2).Caption = TT.A(836) & "*"
    Else
        MultiPage1.Pages(2).Caption = TT.A(836)
    End If
    If (TextBox_parametric1x.Text <> vbNullString And TextBox_parametric1y.Text <> vbNullString And TextBox_parametric1z.Text <> vbNullString) Or TextBox_parametric2x.Text <> vbNullString Or TextBox_parametric3x.Text <> vbNullString Then
        MultiPage1.Pages(3).Caption = TT.A(320) & "*"
    Else
        MultiPage1.Pages(3).Caption = TT.A(320)
    End If
    If TextBox_punkter.Text <> vbNullString Then
        MultiPage1.Pages(4).Caption = TT.A(835) & "*"
    Else
        MultiPage1.Pages(4).Caption = TT.A(835)
    End If
    
End Sub

Private Sub TextBox_forskrift2_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_forskrift3_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_ligning1_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_ligning2_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_ligning3_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_parametric1x_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_parametric1y_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_parametric1z_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_parametric2x_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_parametric3x_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_punkter_Change()
    UpdateBoldTabs
End Sub

Private Sub TextBox_vektorer_Change()
    UpdateBoldTabs
End Sub

Private Sub UserForm_Activate()
    SetCaptions
    colindex = 0
    palindex = 0
#If Mac Then
    CommandButton_ok.visible = False
    CheckBox_maximakommando.visible = False
    Frame1.visible = False
    CheckBox_transp.visible = False
    CheckBox_udtryk.visible = False
    CheckBox_grid.visible = False
    ComboBox_kvalitet.visible = False
    Label45.visible = False
    TextBox_titel.visible = False
    Label42.visible = False
    TextBox_xmin.visible = False
    TextBox_ymin.visible = False
    TextBox_zmin.visible = False
    TextBox_xmax.visible = False
    TextBox_ymax.visible = False
    TextBox_zmax.visible = False
    Label10.visible = False
    Label11.visible = False
    Label13.visible = False
    Label12.visible = False
    Label15.visible = False
    Label14.visible = False
    Label46.visible = False
    ComboBox_farver.visible = False
    TextBox_tmin1.visible = False
    TextBox_tmax1.visible = False
    TextBox_smin1.visible = False
    TextBox_smax1.visible = False
    TextBox_tmin2.visible = False
    TextBox_tmax2.visible = False
    TextBox_smin2.visible = False
    TextBox_smax2.visible = False
    TextBox_tmin3.visible = False
    TextBox_tmax3.visible = False
    TextBox_smin3.visible = False
    TextBox_smax3.visible = False
    Label23.visible = False
    Label24.visible = False
    Label36.visible = False
    Label37.visible = False
    Label28.visible = False
    Label29.visible = False
    Label38.visible = False
    Label39.visible = False
    Label33.visible = False
    Label34.visible = False
    Label40.visible = False
    Label41.visible = False
    CheckBox_pointsjoined.visible = False
    TextBox_pointsize.visible = False
    Label44.visible = False
#Else
#End If
End Sub

Private Sub UserForm_Initialize()
    colindex = 0
    palindex = 0
    ComboBox_kvalitet.AddItem TT.A(185)
    ComboBox_kvalitet.AddItem TT.A(184)
    ComboBox_kvalitet.AddItem TT.A(183)
    ComboBox_kvalitet.AddItem TT.A(182)
    ComboBox_kvalitet.AddItem TT.A(181)
    ComboBox_kvalitet.ListIndex = 3
    
    ComboBox_farver.AddItem TT.A(321)
    ComboBox_farver.AddItem TT.A(322)
    ComboBox_farver.AddItem TT.A(323)
    ComboBox_farver.AddItem TT.A(324)
    ComboBox_farver.ListIndex = 0

    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = 1
    Me.hide
    CommandButton_nulstilalt_Click
  End If
End Sub

Sub SetCaptions()
    Me.Caption = TT.A(302)
'    CommandButton_ok.Caption = TT.OK ' gnuplot not ok
    Label6.Caption = TT.A(186) & " 1"
    Label7.Caption = TT.A(186) & " 2"
    Label8.Caption = TT.A(186) & " 3"
    Label9.Caption = TT.A(198)
    CommandButton_nulstilforsk1.Caption = TT.Reset
    CommandButton_nulstilforsk2.Caption = TT.Reset
    CommandButton_nulstilforsk3.Caption = TT.Reset
    Label42.Caption = TT.A(837) & ":"
    Label45.Caption = TT.A(191)
    CommandButton_nulstilalt.Caption = TT.A(800)
    Label_vent.Caption = TT.A(826) & "..."
    CheckBox_grid.Caption = TT.A(190)
    CheckBox_udtryk.Caption = TT.A(303)
    CheckBox_transp.Caption = TT.A(304)
    Label46.Caption = TT.A(305)
    CheckBox_maximakommando.Caption = TT.A(306)
    Frame1.Caption = TT.A(307)
    OptionButton_kunlinjer.Caption = TT.A(308)
    OptionButton_kunflader.Caption = TT.A(309)
    OptionButton_linjerogflader.Caption = TT.A(310)
    Label2.Caption = TT.A(833) & " 1"
    Label3.Caption = TT.A(833) & " 2"
    Label4.Caption = TT.A(833) & " 3"
    CommandButton_nulstillign1.Caption = TT.Reset
    CommandButton_nulstillign2.Caption = TT.Reset
    CommandButton_nulstillign3.Caption = TT.Reset
    Label1.Caption = TT.A(311)
    CommandButton_kugle.Caption = TT.A(312)
    CommandButton_insertplan.Caption = TT.A(313)
    CommandButton_nulstilvektorer.Caption = TT.Reset
    CommandButton_nyvektor.Caption = TT.A(314)
    Label5.Caption = TT.A(315)
    CommandButton_nulstilpar1.Caption = TT.Reset
    CommandButton_nulstilpar2.Caption = TT.Reset
    CommandButton_nulstilpar3.Caption = TT.Reset
    CommandButton_parlinje.Caption = TT.A(200)
    CommandButton_nyparplan.Caption = TT.A(313)
    CommandButton_nulstilpunkter.Caption = TT.Reset
    Label44.Caption = TT.A(316) & ":"
    CheckBox_pointsjoined.Caption = TT.A(89)
    Label43.Caption = TT.A(317)
    Label16.Caption = TT.A(318)
'    MultiPage1.Pages(0).Caption = TT.A(319)
'    MultiPage1.Pages(1).Caption = TT.A(834)
'    MultiPage1.Pages(2).Caption = TT.A(836)
'    MultiPage1.Pages(3).Caption = TT.A(320)
'    MultiPage1.Pages(4).Caption = TT.A(835)
    MultiPage1.Pages(5).Caption = TT.A(808)
End Sub
