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
Private palindex As Integer
Private Sub CommandButton_geogebra_Click()
Dim s As String, vekt As String, Arr() As String, i As Integer, j As Integer
Dim ea As New ExpressionAnalyser, punkttekst As String, parx As String, pary As String, parz As String, cmd As String
    Dim sl As New CSortList, DefList As String, Var As String, k As Integer, fktudtryk As String, UrlLink As String, p As Integer
    
    ea.SetNormalBrackets

'PrepareMaxima
'    omax.ConvertLnLog = False

'definitioner
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

'    ' definer variable der ikke er defineret
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

' forskrifter
    If TextBox_forskrift1.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift1.Text) & ";"
    End If
    If TextBox_forskrift2.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift2.Text) & ";"
    End If
    If TextBox_forskrift3.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_forskrift3.Text) & ";"
    End If

'ligninger
    If TextBox_ligning1.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning1.Text) & ";"
    End If
    If TextBox_ligning2.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning1.Text) & ";"
    End If
    If TextBox_ligning3.Text <> "" Then
        s = s & ConvertToGeogebraSyntax(TextBox_ligning1.Text) & ";"
    End If
    
'vektorer
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

'parameterfremstillinger
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

'punkter
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
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    s = s & punkttekst & ";"
End If
    s = Left(s, Len(s) - 1)
    
    OpenGeoGebraWeb s, "3d", True, False
    
End Sub

Private Sub CommandButton_insertplan_Click()
Dim plan As String
'    plan = "a*(x-x0)+b*(y-y0)+c*(z-z0)=0"
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
Private Sub CommandButton_nulstilligning2_Click()
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
        If right(TextBox_vektorer.Text, 1) = ")" Then
            TextBox_vektorer.Text = TextBox_vektorer.Text & vbCr
        End If
    End If
    TextBox_vektorer.Text = TextBox_vektorer.Text & "(0 , 0 , 0)(1 , 1 , 1)"
End Sub

Private Sub CommandButton_ok_Click()
Dim grafobj As String
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

On Error GoTo Fejl
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

'forskrifter
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

'ligninger
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


'parameterfremstillinger
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

'vektorer
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

'punkter
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
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    grafobj = grafobj & "key="""","
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & TextBox_pointsize.Text & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ",points([" & punkttekst & "]),"
    antalobj = antalobj + 1
End If



If grafobj = "" Then GoTo Slut

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

'gridno = TextBox_gridlines.text
'grafobj = "xu_grid=" & gridno & ",yv_grid=" & gridno & ",x_voxel=" & gridno & ",y_voxel=" & gridno & ",z_voxel=" & gridno & "," & grafobj
If ComboBox_kvalitet.ListIndex = 0 Then 'super
    grafobj = "xu_grid=200,yv_grid=200,x_voxel=18,y_voxel=18,z_voxel=18," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 1 Then 'meget høj
    grafobj = "xu_grid=100,yv_grid=100,x_voxel=15,y_voxel=15,z_voxel=15," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 2 Then ' høj
    grafobj = "xu_grid=50,yv_grid=50,x_voxel=12,y_voxel=12,z_voxel=12," & grafobj
ElseIf ComboBox_kvalitet.ListIndex = 4 Then 'lav
    grafobj = "xu_grid=15,yv_grid=15,x_voxel=5,y_voxel=5,z_voxel=5," & grafobj
End If


If TextBox_titel.Text <> "" Then
    grafobj = grafobj & "title=""" & TextBox_titel.Text & ""","
End If
'grafobj = grafobj & "user_preamble = ""set pm3d at s depthorder;set xyplane at 0"",colorbox=false,surface_hide = true"

If ComboBox_farver.ListIndex = 0 Then ' standard
    grafobj = grafobj & "palette=color,"
ElseIf ComboBox_farver.ListIndex = 1 Then ' blå
    grafobj = grafobj & "palette=[4,5,7],"
ElseIf ComboBox_farver.ListIndex = 2 Then ' brun
    grafobj = grafobj & "palette=[4,5,6],"
ElseIf ComboBox_farver.ListIndex = 3 Then ' Grå
    grafobj = grafobj & "palette=gray,"
End If

grafobj = "font=""Arial"",font_size=8," & grafobj
grafobj = grafobj & "colorbox=false"

' afslut
'If Right(grafobj, 1) = "," Then grafobj = Left(grafobj, Len(grafobj) - 1)

If CheckBox_maximakommando Then
    omax.MaximaOutput = "draw3d(" & grafobj & ")"
    omax.InsertMaximaOutput
End If


    omax.Draw3D grafobj, antalobj
'    If Len(omax.MaximaOutput) < 3 Then
'        MsgBox "Der skete en fejl. Der er nok en fejl i din syntaks et sted. Check alle faner.", vbOKOnly, "Fejl"
'    End If
    omax.PrepareNewCommand
    Label_vent.visible = False
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:

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
'    Label16.visible = False
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
    ComboBox_kvalitet.AddItem Sprog.A(185)
    ComboBox_kvalitet.AddItem Sprog.A(184) '("Meget høj")
    ComboBox_kvalitet.AddItem Sprog.A(183) '("Høj")
    ComboBox_kvalitet.AddItem Sprog.A(182) '("Normal")
    ComboBox_kvalitet.AddItem Sprog.A(181) '("Lav")
    ComboBox_kvalitet.ListIndex = 3
    
    ComboBox_farver.AddItem Sprog.A(321) '("Gul/rød/lilla")
    ComboBox_farver.AddItem Sprog.A(322) '("Blå")
    ComboBox_farver.AddItem Sprog.A(323) '("Brun")
    ComboBox_farver.AddItem Sprog.A(324) '("Grå")
    ComboBox_farver.ListIndex = 0

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = 1
    Me.Hide
  End If
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(302)
'    CommandButton_ok.Caption = Sprog.OK ' gnuplot not ok
    Label6.Caption = Sprog.Equation & " 1"
    Label7.Caption = Sprog.Equation & " 2"
    Label8.Caption = Sprog.Equation & " 3"
    Label9.Caption = Sprog.A(198)
    CommandButton_nulstilforsk1.Caption = Sprog.Reset
    CommandButton_nulstilforsk2.Caption = Sprog.Reset
    CommandButton_nulstilforsk3.Caption = Sprog.Reset
    Label42.Caption = Sprog.Title & ":"
    Label45.Caption = Sprog.A(191)
    CommandButton_nulstilalt.Caption = Sprog.ResetAll
    Label_vent.Caption = Sprog.Wait & "..."
    CheckBox_grid.Caption = Sprog.A(190)
    CheckBox_udtryk.Caption = Sprog.A(303)
    CheckBox_transp.Caption = Sprog.A(304)
    Label46.Caption = Sprog.A(305)
    CheckBox_maximakommando.Caption = Sprog.A(306)
    Frame1.Caption = Sprog.A(307)
    OptionButton_kunlinjer.Caption = Sprog.A(308)
    OptionButton_kunflader.Caption = Sprog.A(309)
    OptionButton_linjerogflader.Caption = Sprog.A(310)
    Label2.Caption = Sprog.Equation & " 1"
    Label3.Caption = Sprog.Equation & " 2"
    Label4.Caption = Sprog.Equation & " 3"
    CommandButton_nulstillign1.Caption = Sprog.Reset
    CommandButton_nulstillign2.Caption = Sprog.Reset
    CommandButton_nulstillign3.Caption = Sprog.Reset
    Label1.Caption = Sprog.A(311)
    CommandButton_kugle.Caption = Sprog.A(312)
    CommandButton_insertplan.Caption = Sprog.A(313)
    CommandButton_nulstilvektorer.Caption = Sprog.Reset
    CommandButton_nyvektor.Caption = Sprog.A(314)
    Label5.Caption = Sprog.A(315)
    CommandButton_nulstilpar1.Caption = Sprog.Reset
    CommandButton_nulstilpar2.Caption = Sprog.Reset
    CommandButton_nulstilpar3.Caption = Sprog.Reset
    CommandButton_parlinje.Caption = Sprog.A(200)
    CommandButton_nyparplan.Caption = Sprog.A(313)
    CommandButton_nulstilpunkter.Caption = Sprog.Reset
    Label44.Caption = Sprog.A(316) & ":"
    CheckBox_pointsjoined.Caption = Sprog.A(89)
    Label43.Caption = Sprog.A(317)
    Label16.Caption = Sprog.A(318)
    MultiPage1.Pages(0).Caption = Sprog.A(319)
    MultiPage1.Pages(1).Caption = Sprog.Equations
    MultiPage1.Pages(2).Caption = Sprog.Vectors
    MultiPage1.Pages(3).Caption = Sprog.A(320)
    MultiPage1.Pages(4).Caption = Sprog.Points
    MultiPage1.Pages(5).Caption = Sprog.Help
    
    
End Sub
