Attribute VB_Name = "RibbonSubs"
Option Explicit
Public WoMatRibbon As IRibbonUI
#If Mac Then
' not testet. Taken from a forum on macrumors.com. dest and source set to longptr but try any without byval. libc.dylib is a guess.
'    Private Declare PtrSafe Function memcpy Lib "libsystem_c.dylib" Alias "CopyMemory" (Dest As Any, source As Any, ByVal length As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

'Callback for customUI.onLoad
Sub LoadRibbon(ribbon As IRibbonUI)
#If Mac Then
    Set WoMatRibbon = ribbon
#Else
    Set WoMatRibbon = ribbon
    Dim lngRibPtr As LongPtr
    lngRibPtr = ObjPtr(ribbon)
    
    SetRegSettingLong "RibbonPointer", lngRibPtr
#End If
End Sub

Public Sub ribbonLoaded(ribbon As IRibbonUI)

End Sub
Function GetRibbon(lngRibPtr As LongPtr) As Object
#If Mac Then
#Else
   Dim objRibbon As Object
   CopyMemory objRibbon, lngRibPtr, 4
   Set GetRibbon = objRibbon
   Set objRibbon = Nothing
#End If
End Function
Sub RefreshRibbon()
#If Mac Then
    If Not (WoMatRibbon Is Nothing) Then
        WoMatRibbon.Invalidate
    End If
#Else
    On Error GoTo fejl
   Dim lngRibPtr As LongPtr
   Dim lngRibPtrBackup As LongPtr
   Dim objRibbon As Object
   
  If Not (WoMatRibbon Is Nothing) Then
        WoMatRibbon.Invalidate
    Else
        lngRibPtrBackup = ObjPtr(WoMatRibbon)
        lngRibPtr = CLng(GetRegSettingLong("RibbonPointer"))
        If lngRibPtr > 0 Then
          Set WoMatRibbon = GetRibbon(lngRibPtr)
          WoMatRibbon.Invalidate
        End If
        ' The static guiRibbon-variable was meanwhile lost
'        MsgBox "Due to a design flaw in the architecture of the MS ribbon UI you have to close " & _
'            "and reopen this workbook." & vbNewLine & vbNewLine & _
'            "Very sorry about that.", vbExclamation + vbOKOnly
        ' Note: In the help we can find
        ' guiRibbon.Refresh
        ' but unfortunately this is not implemented.
        ' It is exactly what we should have instead of that brute force reload mechanism.
    End If

GoTo slut
fejl:
    MsgBox Sprog.A(394), vbOKOnly, Sprog.Error
    Set WoMatRibbon = GetRibbon(lngRibPtrBackup)
    lngRibPtr = 0
slut:
#End If
End Sub
' events der fyres når der trykkes på ribbon
Sub insertribformel(Kommentar As String, ByVal formel As String)
    On Error GoTo fejl
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If

    Application.ScreenUpdating = False
    If Kommentar <> "" Then
    Selection.InsertAfter (Kommentar)
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    End If
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText formel
    Selection.OMaths.BuildUp
'    Selection.OMaths(1).BuildUp
    
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If
    
    GoTo slut
fejl:
    MsgBox Sprog.A(395), vbOKOnly, Sprog.Error
slut:
End Sub

Public Sub Rib_Settings(control As IRibbonControl)
    Call MaximaSettings
End Sub

'Callback for proc1 onAction
Sub Rib_FSfremskriv(control As IRibbonControl)
    insertribformel "", "S=B" & VBA.ChrW(183) & "(1+r)"
End Sub

'Callback for Button4 onAction
Sub Rib_FSkapital(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        insertribformel "", "K_n=K_0" & VBA.ChrW(183) & "(1+r)^n"
    Else
        insertribformel "", "A=P" & VBA.ChrW(183) & "(1+r/n)^(n" & VBA.ChrW(183) & "t)"
    End If
End Sub

Sub Rib_FSannuitet1(control As IRibbonControl)
    insertribformel "", "A=b" & VBA.ChrW(183) & "((1+r)^n-1)/r"
End Sub
Sub Rib_FSannuitet2(control As IRibbonControl)
    insertribformel "", "y=G" & VBA.ChrW(183) & "r/(1-(1+r)^(-n))"
End Sub

'Callback for lin1 onAction
Sub Rib_FSlinligning(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        insertribformel "", "y=a" & VBA.ChrW(183) & "x+b"
    Else
        insertribformel "", "y=m" & VBA.ChrW(183) & "x+c"
    End If
End Sub

'Callback for lin2 onAction
Sub Rib_FSberegna(control As IRibbonControl)
    insertribformel "", "a=(y_2-y_1)/(x_2-x_1)"
End Sub

'Callback for lin3 onAction
Sub Rib_FSlinligning2(control As IRibbonControl)
    insertribformel "", "y=a" & VBA.ChrW(183) & "(x-x_0)+y_0"
End Sub

'Callback for difftangent onAction
Sub Rib_FSdiff(control As IRibbonControl)
    insertribformel "", "y=f'(x_0)" & VBA.ChrW(183) & "(x-x_0)+f(x_0)"
End Sub

'Callback for eksp1 onAction
Sub Rib_FSekspligning(control As IRibbonControl)
    insertribformel "", "y=b" & VBA.ChrW(183) & "a^x"
End Sub
'Callback for eksp5 onAction
Sub Rib_FSekspligning2(control As IRibbonControl)
    insertribformel "", "y=b" & VBA.ChrW(183) & "e^(k" & VBA.ChrW(183) & "x)"
End Sub

'Callback for eksp6 onAction
Sub Rib_FSekspligning3(control As IRibbonControl)
    insertribformel "", "y=b" & VBA.ChrW(183) & "2^(x/T_2)"
End Sub
'Callback for eksp6 onAction
Sub Rib_FSekspligning4(control As IRibbonControl)
    insertribformel "", "y=b" & VBA.ChrW(183) & "(1/2)^(x/T_" & ChrW(189) & ")"
End Sub

'Callback for eksp2 onAction
Sub Rib_FSekspa(control As IRibbonControl)
    insertribformel "", "a=" & VBA.ChrW(&H221A) & "(x_2-x_1&y_2/y_1)"
End Sub
'Callback for eksp3 onAction
Sub Rib_FSford(control As IRibbonControl)
    insertribformel "", "T_2=ln" & VBA.ChrW(8289) & "(2)/ln" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(2)/k"
End Sub

'Callback for eksp4 onAction
Sub Rib_FShalv(control As IRibbonControl)
#If Mac Then
    insertribformel "", "T_(1/2) =ln" & VBA.ChrW(8289) & "(1/2)/(ln" & VBA.ChrW(8289) & "(a))=ln" & VBA.ChrW(8289) & "(1/2)/k"
#Else
    insertribformel "", "T_" & ChrW(189) & "=ln" & VBA.ChrW(8289) & "(1/2)/ln" & VBA.ChrW(8289) & "(a)=ln" & VBA.ChrW(8289) & "(1/2)/k"
#End If
End Sub

'Callback for pot1 onAction
Sub Rib_FSpotligning(control As IRibbonControl)
    insertribformel "", "y=b" & VBA.ChrW(183) & "x^a"
End Sub

'Callback for pot2 onAction
Sub Rib_FSpota(control As IRibbonControl)
    insertribformel "", "a=log" & VBA.ChrW(8289) & "(y_2/y_1)/log" & VBA.ChrW(8289) & "(x_2/x_1)"
End Sub

'Callback for pot3 onAction
Sub Rib_FSpotprocvaekst(control As IRibbonControl)
    insertribformel "", "1+r_y=(1+r_x)^a"
End Sub
'Callback for pol1 onAction
Sub Rib_FSpol(control As IRibbonControl)
    insertribformel "", "x_t=-b/2a"
    Selection.TypeParagraph
    insertribformel "", "y_t=-(b^2-4" & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "c)/4a"
End Sub

'Callback for geo1 onAction
Sub Rib_FSsinrel(control As IRibbonControl)
    insertribformel "", "a/sin" & VBA.ChrW(8289) & "(A)=b/sin" & VBA.ChrW(8289) & "(B)"
End Sub
'Callback for geo2 onAction
Sub Rib_FScosrel(control As IRibbonControl)
    insertribformel "", "c^2=a^2+b^2-2" & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "cos(C)"
End Sub
'Callback for geo3 onAction
Sub Rib_FSarealtrekant(control As IRibbonControl)
    insertribformel "", "T=1/2" & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "sin(C)"
End Sub
'Callback for sandbin1 onAction
Sub Rib_FSbinfrekvens(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    If Not InStr(omax.DefString, "K(n,r):=") > 0 Then
        insertribformel "", "K(n,r)" & VBA.ChrW(8797) & "n!/(r!" & VBA.ChrW(183) & "(n-r)!)" ' chrw8801 er 3 streget =
        Selection.TypeParagraph
    End If
    
    If Not InStr(omax.DefString, "p:") > 0 Then
        s = "p=0,5 ; "
    End If
    If Not InStr(omax.DefString, "n:") > 0 Then
        s = s & "n=20"
    End If
    s = Trim(s)
    If right(s, 1) = ";" Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        s = InputBox("Indtast nødvendige definitioner", "Definitioner", s)
        s = Replace(s, ";", " , ")
        s = Replace(s, "  ", " ")
    
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If
    
    insertribformel "", "P(r)" & VBA.ChrW(8797) & "K(n,r)" & VBA.ChrW(183) & "p^r" & VBA.ChrW(183) & "(1-p)^(n-r)"
    Selection.TypeParagraph
End Sub
'Callback for sandbin5 onAction
Sub Rib_FSbinkum(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    If Not InStr(omax.DefString, "K(n,r):=") > 0 Then
        insertribformel "", "K(n,r)" & VBA.ChrW(8797) & "n!/(r!" & VBA.ChrW(183) & "(n-r)!)" ' chrw8801 er 3 streget =
        Selection.TypeParagraph
    End If
    
    If Not InStr(omax.DefString, "p:") > 0 Then
        s = "p=0,5 ; "
    End If
    If Not InStr(omax.DefString, "n:") > 0 Then
        s = s & "n=20"
    End If
    s = Trim(s)
    If right(s, 1) = ";" Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        s = InputBox("Indtast nødvendige definitioner", "Definitioner", s)
        s = Replace(s, ";", " , ")
        s = Replace(s, "  ", " ")
    
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If
    insertribformel "", "P_kum (m)" & VBA.ChrW(8797) & "" & VBA.ChrW(8721) & "_(r=0)^m" & VBA.ChrW(9618) & VBA.ChrW(12310) & "K(n,r)" & VBA.ChrW(183) & "p^r" & VBA.ChrW(183) & "(1-p)^(n-r)" & VBA.ChrW(12311)
    Selection.TypeParagraph
End Sub

'Callback for sandbin2 onAction
Sub Rib_FSbinkoeff(control As IRibbonControl)
    insertribformel "", "K(n,r)" & VBA.ChrW(8797) & "n!/(r!" & VBA.ChrW(183) & "(n-r)!)" ' chrw8801 er 3 streget =
End Sub

'Callback for sandbin3 onAction
Sub Rib_FSbinmid(control As IRibbonControl)
    insertribformel "", VBA.ChrW(956) & "=n" & VBA.ChrW(183) & "p"
End Sub

'Callback for sandbin4 onAction
Sub Rib_FSbinspred(control As IRibbonControl)
    insertribformel "", VBA.ChrW(963) & "=" & VBA.ChrW(&H221A) & "(n" & VBA.ChrW(183) & "p" & VBA.ChrW(183) & "(1-p))"
End Sub

Sub Rib_FSbinusik(control As IRibbonControl)
    insertribformel "", "p" & VBA.ChrW(770) & "±2" & VBA.ChrW(183) & VBA.ChrW(8730) & "((p" & VBA.ChrW(770) & "" & VBA.ChrW(183) & "(1-p" & VBA.ChrW(770) & "))/n)"
End Sub

'Callback for sandnorm1 onAction
Sub Rib_FSnormfrekvens(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    
    If Not InStr(omax.DefString, "mu:") > 0 Then ' sigma
        s = VBA.ChrW(181) & "=0 ; "
    End If
    If Not InStr(omax.DefString, "sigma:") > 0 Then ' mu
        s = s & "s=1"
    End If
    s = Trim(s)
    If right(s, 1) = ";" Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        s = InputBox("Indtast nødvendige definitioner", "Definitioner", s)
        #If Mac Then
        #Else
            s = Replace(s, VBA.ChrW(181) & "=", VBA.ChrW(956) & "=")
        #End If
        s = Replace(s, "s=", VBA.ChrW(963) & "=")
        s = Replace(s, ";", " , ")
        s = Replace(s, "  ", " ")
    
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If

    insertribformel "", "f(x)" & VBA.ChrW(8797) & "1/(" & VBA.ChrW(&H221A) & "(2" & VBA.ChrW(960) & ")" & VBA.ChrW(183) & VBA.ChrW(963) & ")" & VBA.ChrW(183) & "e^(-1/2" & VBA.ChrW(183) & "((x-" & VBA.ChrW(956) & ")/" & VBA.ChrW(963) & ")^2)"
    Selection.TypeParagraph
End Sub

'Callback for sandnorm2 onAction
Sub Rib_FSnormkum(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    
    If Not InStr(omax.DefString, "mu:") > 0 Then ' sigma
        s = VBA.ChrW(181) & "=0 ; "
    End If
    If Not InStr(omax.DefString, "sigma:") > 0 Then ' mu
        s = s & "s=1"
    End If
    s = Trim(s)
    If right(s, 1) = ";" Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        s = InputBox("Indtast nødvendige definitioner", "Definitioner", s)
        s = Replace(s, "s=", VBA.ChrW(963) & "=")
        s = Replace(s, VBA.ChrW(181) & "=", VBA.ChrW(956) & "=")
        s = Replace(s, ";", " , ")
        s = Replace(s, "  ", " ")
    
        insertribformel "", "Definer: " & VBA.ChrW(963) & ">0"
        Selection.TypeParagraph
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If
    insertribformel "", "F(x)" & VBA.ChrW(8797) & VBA.ChrW(8747) & "_(-" & VBA.ChrW(8734) & ")^x" & VBA.ChrW(9618) & "1/(" & VBA.ChrW(&H221A) & "(2" & VBA.ChrW(960) & ")" & VBA.ChrW(183) & VBA.ChrW(963) & ")" & VBA.ChrW(183) & "e^(-1/2" & VBA.ChrW(183) & "((y-" & VBA.ChrW(956) & ")/" & VBA.ChrW(963) & ")^2) dy"
    Selection.TypeParagraph
End Sub
'Callback for sandchiford onAction
Sub Rib_FSchi2ford(control As IRibbonControl)
    Chi2Fordeling
End Sub
'Callback for omdrejlegeme onAction
Sub Rib_FSomdrejlegeme(control As IRibbonControl)
    insertribformel "", "V=" & VBA.ChrW(&H3C0) & VBA.ChrW(183) & VBA.ChrW(8747) & "_a^b" & VBA.ChrW(9618) & "(f(x))" & VBA.ChrW(&HB2) & " dx"
End Sub
'Callback for kurvel onAction
Sub Rib_FSkurvelaengde(control As IRibbonControl)
    insertribformel "", "s=" & VBA.ChrW(8747) & "_a^b" & VBA.ChrW(9618) & VBA.ChrW(&H221A) & "(1+(f'(x))" & VBA.ChrW(&HB2) & ") dx"
End Sub
'Callback for middelv onAction
Sub Rib_FSmiddelv(control As IRibbonControl)
    insertribformel "", "<f(x)>=1/(b-a) " & VBA.ChrW(8747) & "_a^b" & VBA.ChrW(9618) & "f(x) dx"
End Sub
'Callback for planparamlinje onAction
Sub Rib_FSplanlinjelign(control As IRibbonControl) '
    insertribformel "", "a" & VBA.ChrW(183) & "(x-x_0)+b" & VBA.ChrW(183) & "(y-y_0)=0"
End Sub
'Callback for planparamlinje onAction
Sub Rib_FSplanparamlinje(control As IRibbonControl) '
    insertribformel "", "(" & VBA.ChrW(9632) & "(x@y))=(" & VBA.ChrW(9632) & "(x_0@y_0 ))+t" & VBA.ChrW(183) & "(" & VBA.ChrW(9632) & "(r_1@r_2 ))"
End Sub

Sub Rib_FSvektorvinkel(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    
    If Not InStr(omax.DefString, "aSymVecta:") > 0 Then ' vector a
        s = "a" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9608) & "(1@2)), "
    End If
    If Not InStr(omax.DefString, "bSymVecta:") > 0 Then ' vector b
        s = s & "b" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9608) & "(1@2))"
    End If
    s = Trim(s)
    If right(s, 1) = "," Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If
   
   If CASengine = 0 Then
    insertribformel "", "cos(v)=(a" & VBA.ChrW(8407) & ChrW(183) & "b" & VBA.ChrW(8407) & ")/(|a" & VBA.ChrW(8407) & "|" & ChrW(183) & "|b" & VBA.ChrW(8407) & "|)"
   Else
    insertribformel "", "cos(v)=(dot(a" & VBA.ChrW(8407) & ";b" & VBA.ChrW(8407) & "))/(|a" & VBA.ChrW(8407) & "|" & ChrW(183) & "|b" & VBA.ChrW(8407) & "|)"
   End If
End Sub
Sub Rib_FSvektorproj(control As IRibbonControl)
    Dim s As String, arr() As String
    PrepareMaxima
    omax.FindDefinitions
    
    If Not InStr(omax.DefString, "aSymVecta:") > 0 Then ' vector a
        s = "a" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9608) & "(1@2)), "
    End If
    If Not InStr(omax.DefString, "bSymVecta:") > 0 Then ' vector b
        s = s & "b" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9608) & "(1@2))"
    End If
    s = Trim(s)
    If right(s, 1) = "," Then s = Left(s, Len(s) - 1)
    If s <> vbNullString Then
        insertribformel "", "Definer: " & s
        Selection.TypeParagraph
    End If
    
    insertribformel "", "b" & VBA.ChrW(8407) & "_a=(a" & VBA.ChrW(8407) & ChrW(183) & "b" & VBA.ChrW(8407) & ")/(|a" & VBA.ChrW(8407) & "|^2) a" & VBA.ChrW(8407)
End Sub
Sub Rib_FSdistpunkt(control As IRibbonControl)
    insertribformel "", "dist(P,l)=|a" & ChrW(183) & "x_1+b" & ChrW(183) & "y_1+c|/" & VBA.ChrW(&H221A) & "(a^2+b^2)"
End Sub

'Callback for cirkelligning onAction
Sub Rib_FScirklensligning(control As IRibbonControl)
    insertribformel "", "(x-x_0)^2+(y-y_0)^2=r^2"
End Sub

'Callback for rumparamlinje onAction
Sub Rib_FSrumlinjelign(control As IRibbonControl) '
    insertribformel "", "a" & VBA.ChrW(183) & "(x-x_0)+b" & VBA.ChrW(183) & "(y-y_0)+c" & VBA.ChrW(183) & "(z-z_0)=0"
End Sub
'Callback for rumparamlinje onAction
Sub Rib_FSrumparamlinje(control As IRibbonControl)
    insertribformel "", "(" & VBA.ChrW(9632) & "(x@y@z))=(" & VBA.ChrW(9632) & "(x_0@y_0@z_0 ))+t" & VBA.ChrW(183) & "(" & VBA.ChrW(9632) & "(r_1@r_2@r_3 ))"
End Sub

'Callback for rumafstandpunktlinje onAction
Sub Rib_FSrumpunktlinje(control As IRibbonControl)
    insertribformel "", "definer: r" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9632) & "(r_1@r_2@r_3)) ,  (P0P)" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9632) & "(x_1-x_0@y_1-y_0@z_1-z_0))"
    Selection.TypeParagraph
'    insertribformel "", "dist(P,l)=(|r" & VBA.ChrW(8407) & "_(P_0 P)" & VBA.ChrW(8407) & "|)/(|r" & VBA.ChrW(8407) & "|)"
    insertribformel "", "dist(P,l)=(|r" & VBA.ChrW(8407) & VBA.ChrW(215) & "(P0P)" & VBA.ChrW(8407) & "|)/(|r" & VBA.ChrW(8407) & "|)"
End Sub

'Callback for rumligningplan onAction
Sub Rib_FSrumligningplan(control As IRibbonControl)
    insertribformel "", "definer: n" & VBA.ChrW(8407) & "=(" & VBA.ChrW(9632) & "(a@b@c))"
    Selection.TypeParagraph
    insertribformel "", "n" & VBA.ChrW(8407) & VBA.ChrW(183) & "(" & VBA.ChrW(9632) & "(x-x_0@y-y_0@z-z_0))=0"
End Sub

'Callback for rumligningplan2 onAction
Sub Rib_FSrumligningplan2(control As IRibbonControl)
    insertribformel "", "a" & ChrW(183) & "(x-x_0)+b" & ChrW(183) & "(y-y_0)+c" & ChrW(183) & "(z-z_0)=0"
End Sub

'Callback for rumafstandpunktplan onAction
Sub Rib_FSrumafstandpunktplan(control As IRibbonControl)
    insertribformel "", "dist(P," & VBA.ChrW(945) & ")=|n" & VBA.ChrW(8407) & ChrW(183) & "(" & VBA.ChrW(9632) & "(x_1-x_0@y_1-y_0@z_1-z_0 ))|/(|n" & VBA.ChrW(8407) & "|)"
End Sub

'Callback for rumafstandpunktplan2 onAction
Sub Rib_FSrumafstandpunktplan2(control As IRibbonControl)
    insertribformel "", "dist(P," & VBA.ChrW(945) & ")=(|a" & ChrW(183) & "x_1+b" & ChrW(183) & "y_1+c" & ChrW(183) & "z_1+d|)/" & VBA.ChrW(&H221A) & "(a^2+b^2+c^2)"
End Sub

'Callback for kugleligning onAction
Sub Rib_FSkuglensligning(control As IRibbonControl)
    insertribformel "", "(x-x_0)^2+(y-y_0)^2+(z-z_0)^2=r^2"
End Sub

'Callback for matformler onAction
Sub Rib_matformler(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        OpenFormulae ("MatFormler.docx")
    ElseIf LanguageSetting = 3 Then
        OpenFormulae ("MatFormler_spansk.docx")
    Else
        OpenFormulae ("MatFormler_english.docx")
    End If
End Sub
'Callback for fysikformler onAction
Sub Rib_fysikformler(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        OpenFormulae ("FysikFormler.docx")
    ElseIf LanguageSetting = 3 Then
        OpenFormulae ("FysikFormler_spansk.docx")
    Else
        OpenFormulae ("FysikFormler.docx")
    End If
End Sub
'Callback for kemiformler onAction
Sub Rib_kemiformler(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        OpenFormulae ("KemiFormler.docx")
    ElseIf Sprog.SprogNr = 3 Then
        OpenFormulae ("KemiFormler_spansk.docx")
    Else
        OpenFormulae ("KemiFormler.docx")
    End If
End Sub
'Callback for togglebuttonAuto getPressed
Sub Rib_GetPressedAuto(control As IRibbonControl, ByRef returnedVal)
    If MaximaExact = 0 Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for togglebuttonAuto onAction
Sub Rib_Auto(control As IRibbonControl, pressed As Boolean)
On Error Resume Next
    MaximaExact = 0
    RefreshRibbon
End Sub

'Callback for togglebuttonExact getPressed
Sub Rib_GetPressedExact(control As IRibbonControl, ByRef returnedVal)
    If MaximaExact = 1 Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for togglebuttonExact onAction
Sub Rib_Exact(control As IRibbonControl, pressed As Boolean)
On Error Resume Next
    MaximaExact = 1
    RefreshRibbon
End Sub

'Callback for togglebuttonNum getPressed
Sub Rib_GetPressedNum(control As IRibbonControl, ByRef returnedVal)
    If MaximaExact = 2 Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for togglebuttonNum onAction
Sub Rib_num(control As IRibbonControl, pressed As Boolean)
On Error Resume Next
    MaximaExact = 2
    RefreshRibbon
End Sub
'Callback for dcifre onAction
Sub Rib_cifre(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
    MaximaCifre = Index + 2
End Sub
'Callback for dcifre getSelectedItemIndex
Sub Rib_GetSelectedItemIndexCifre(control As IRibbonControl, ByRef returnedVal)
On Error Resume Next
    returnedVal = CInt(MaximaCifre) - 2
    RefreshRibbon
End Sub
'Callback for togglebuttonRad getPressed
Sub Rib_GetPressedRad(control As IRibbonControl, ByRef returnedVal)
On Error Resume Next
    returnedVal = Radians
    RefreshRibbon
End Sub

'Callback for togglebuttonRad onAction
Sub Rib_rad(control As IRibbonControl, pressed As Boolean)
    Radians = pressed
End Sub
'Callback for togglebuttonVid getPressed
Sub Rib_GetPressedVid(control As IRibbonControl, ByRef returnedVal)
On Error Resume Next
    returnedVal = MaximaVidNotation
    RefreshRibbon
End Sub

'Callback for togglebuttonVid onAction
Sub Rib_vid(control As IRibbonControl, pressed As Boolean)
    MaximaVidNotation = pressed
End Sub
Public Sub Rib_Beregn(control As IRibbonControl)
    beregn
End Sub
'Callback for ButtonMaxima onAction
Sub Rib_MaximaKommando(control As IRibbonControl)
    MaximaCommand
End Sub
'Callback for ButtonSolve onAction
Sub Rib_Solve(control As IRibbonControl)
    MaximaSolve
End Sub
'Callback for solvenum onAction
Sub Rib_solvenum(control As IRibbonControl)
    MaximaNsolve
    'MaximaSolveNumeric
End Sub
'Callback for eliminate onAction
Sub Rib_eliminate(control As IRibbonControl)
    MaximaEliminate
End Sub
'Callback for test sand/falsk onAction
Sub Rib_test(control As IRibbonControl)
    CompareTest
End Sub
'Callback for solvede onAction
Sub Rib_solvede(control As IRibbonControl)
    SolveDE
End Sub
'Callback for solvedenum onAction
Sub Rib_solvedenum(control As IRibbonControl)
    SolveDENumeric
End Sub
'Callback for ButtonOmskriv onAction
Sub Rib_Omskriv(control As IRibbonControl)
    Omskriv
End Sub
'Callback for simp onAction
Sub Rib_reducer(control As IRibbonControl)
    reducer
End Sub

'Callback for factor onAction
Sub Rib_faktoriser(control As IRibbonControl)
    faktoriser
End Sub

'Callback for udvid onAction
Sub Rib_udvid(control As IRibbonControl)
    udvid
End Sub
'Callback for ButtonDefiner onAction
Sub Rib_Definitioner(control As IRibbonControl)
    MsgBox VisDef, vbOKOnly, Sprog.Definitions
End Sub
'Callback for defsletdef onAction
Sub Rib_sletdef(control As IRibbonControl)
    InsertSletDef
'insertribformel "", "slet definitioner:"
End Sub

'Callback for defvar onAction
Sub Rib_defvar(control As IRibbonControl)
    DefinerVar
End Sub

'Callback for deffunkt onAction
Sub Rib_deffunk(control As IRibbonControl)
    DefinerFunktion
End Sub
Sub Rib_deflign(control As IRibbonControl)
    DefinerLigning
End Sub
Sub Rib_defkonstanter(control As IRibbonControl)
Dim UFkonstanter As New UserFormKonstanter
    UFkonstanter.Show vbModeless
'    UserFormKonstanter.Show
End Sub
'Callback for Differentier onAction
Sub Rib_diff(control As IRibbonControl)
    Differentier
End Sub
'Callback for Stamfunktion onAction
Sub Rib_stam(control As IRibbonControl)
    Integrer
End Sub
'Callback for gnuplot onAction
Sub Rib_gnuplot(control As IRibbonControl)
    Plot2DGraph
End Sub
'Callback for graphobj onAction
Sub Rib_graphobj(control As IRibbonControl)
    Call InsertGraphOleObject
End Sub
'Callback for ButtonGeoGebra onAction
Sub Rib_GeoGebraB(control As IRibbonControl)
    GeoGebraWeb
End Sub
'Callback for excelobj onAction
Sub Rib_excelobj(control As IRibbonControl)
#If Mac Then
    InsertOpenExcel "Graphs.xltm"
#Else
    Call InsertChart
#End If
End Sub
'Callback for ButtonGraf onAction
Sub Rib_graf(control As IRibbonControl)
    StandardPlot
End Sub
'Callback for uGrupperedeobservationer onAction
Sub Rib_ugrupobs(control As IRibbonControl)
    InsertUGrupObs
End Sub
'Callback for Grupperedeobservationer onAction
Sub Rib_grupobs(control As IRibbonControl)
    InsertGrupObs
End Sub
'Callback for pindediagram onAction
Sub Rib_pindediagram(control As IRibbonControl)
    InsertPindeDiagram
End Sub
'Callback for boksplot onAction
Sub Rib_boksplot(control As IRibbonControl)
    InsertBoksplot
End Sub
'Callback for histogram onAction
Sub Rib_histogram(control As IRibbonControl)
    InsertHistogram
End Sub
'Callback for trappediagram onAction
Sub Rib_trappediagram(control As IRibbonControl)
    InsertTrappediagram
End Sub
'Callback for sumkurve onAction
Sub Rib_sumkurve(control As IRibbonControl)
    InsertSumkurve
End Sub

'Callback for ButtonGeoGebra onAction
Sub Rib_GeoGebra(control As IRibbonControl)
    GeoGebra
End Sub
'Callback for geogebrainsert onAction
Sub Rib_insertgeogebra(control As IRibbonControl)
    InsertGeoGeobraObject
End Sub
'Callback for ButtonStatistik onAction
Sub Rib_Statistik(control As IRibbonControl)
    InsertOpenExcel FilNavn:="statistik.xltm", WorkBookName:=Sprog.A(563)
End Sub
Sub Rib_plot3D(control As IRibbonControl)
'#If Mac Then
'    MsgBox "This function is not avaiable on Mac. Use the 3D functions in GeoGebra", vbOKOnly, "Mac issue"
'#Else
    Plot3DGraph
'#End If
End Sub
'Callback for omdrejlegm onAction
Sub Rib_omdrejningslegeme(control As IRibbonControl)
'#If Mac Then
'    MsgBox "This function is not avaiable on Mac.", vbOKOnly, "Mac issue"
'#Else
    OmdrejningsLegeme
'#End If
End Sub
'Callback for retningsfeltm onAction
Sub Rib_retningsfelt(control As IRibbonControl)
'#If Mac Then
'    MsgBox "This function is not avaiable on Mac. Use GeoGebra", vbOKOnly, "Mac issue"
'#Else
    PlotDF
'#End If
End Sub
'Callback for regrtabel onAction
Sub Rib_regrtabel(control As IRibbonControl)
    InsertTabel
End Sub

'Callback for regrlin onAction
Sub Rib_regrlin(control As IRibbonControl)
    linregression
End Sub

'Callback for regreksp onAction
Sub Rib_regreksp(control As IRibbonControl)
    ekspregression
End Sub

'Callback for regrpot onAction
Sub Rib_regrpot(control As IRibbonControl)
    potregression
End Sub

'Callback for regrpol onAction
Sub Rib_regrpol(control As IRibbonControl)
    polregression
End Sub

Sub Rib_regrexcel(control As IRibbonControl)
#If Mac Then
    InsertOpenExcel "Graphs.xltm"
#Else
    Call InsertChart
#End If
End Sub

'Callback for regruser onAction
Sub Rib_regruser(control As IRibbonControl)
    UserRegression
End Sub

'Callback for ButtonBinomtest onAction
Sub Rib_binomialtest(control As IRibbonControl)
    BinomialTest
End Sub
'Callback for ButtonChi2 onAction
Sub Rib_chi2test(control As IRibbonControl)
    Chi2Test
End Sub
'Callback for ButtonGoodnessoffit onAction
Sub Rib_goodnessoffit(control As IRibbonControl)
    GoodnessofFit
End Sub
'Callback for ButtonSimulering onAction
Sub Rib_simulering(control As IRibbonControl)
    InsertOpenExcel FilNavn:="Simulering.xltm", WorkBookName:=Sprog.A(599)
End Sub
'Callback for Buttonbinomialfordeling onAction
Sub Rib_binomialfordeling(control As IRibbonControl)
    BinomialFordeling
End Sub
'Callback for Buttonnormalfordeling onAction
Sub Rib_normalfordelinggraf(control As IRibbonControl)
    NormalFordelingGraf
End Sub
'Callback for ButtonChi2fordeling onAction
Sub Rib_chi2fordelinggraf(control As IRibbonControl)
    Chi2Graf
End Sub
Sub Rib_tfordelinggraf(control As IRibbonControl)
    InsertOpenExcel FilNavn:="studenttFordeling.xltm", WorkBookName:="t" 'Sprog.A(483)
End Sub
'Callback for ButtonGrupper onAction
Sub Rib_Grupper(control As IRibbonControl)
    GrupperIntervaller
End Sub

'Callback for ButtonNyLig onAction
Sub Rib_nylign(control As IRibbonControl)
    On Error GoTo fejl

    Application.ScreenUpdating = False
    Selection.OMaths.Add Range:=Selection.Range
'    Selection.OMaths(1).BuildUp
    
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
'Callback for ButtonNumEq onAction
Sub Rib_nynumlign(control As IRibbonControl)
    InsertNumberedEquation
End Sub
'Callback for ButtonNumEqRef onAction
Sub Rib_nynumlignref(control As IRibbonControl)
    InsertNumberedEquation True
End Sub
'Callback for ButtonRefEq onAction
Sub Rib_reflign(control As IRibbonControl)
    InsertEquationRef
End Sub
'Callback for ButtonSetEqNo onAction
Sub Rib_seteqno(control As IRibbonControl)
    SetEquationNumber
End Sub
'Callback for ButtonInsertEquationSection onAction
Sub Rib_inserteqsec(control As IRibbonControl)
    InsertEquationHeadingNo
End Sub
'Callback for ButtonUpdateEqNo onAction
Sub Rib_updateeqno(control As IRibbonControl)
    UpdateEquationNumbers
End Sub
Sub Rib_LatexTemplate(control As IRibbonControl)
    OpenLatexTemplate
End Sub
'Callback for prik onAction
Sub Rib_gangetegn(control As IRibbonControl)
    Gange
End Sub
'Callback for tilprik onAction
Sub Rib_Tilgangetegn(control As IRibbonControl)
    ReplaceStarMult
End Sub
'Callback for Tilstjerne onAction
Sub Rib_Tilstjerne(control As IRibbonControl)
    ReplaceStarMultBack
End Sub

'Callback for Tilkomma onAction
Sub Rib_Tilkomma(control As IRibbonControl)
    ErstatKomma
End Sub
'Callback for Tilpunktum onAction
Sub Rib_Tilpunktum(control As IRibbonControl)
    ErstatPunktum
End Sub
'Callback for Tilkomma onAction
Sub Rib_TilLaTex(control As IRibbonControl)
    KonverterTilLaTex
End Sub
Sub Rib_ConvertLatex(control As IRibbonControl)
    ToggleLatex
End Sub
Sub Rib_ConvTexLatex(control As IRibbonControl)
#If Mac Then
    MsgBox "This function is not avaiable on Mac.", vbOKOnly, "Mac issue"
#Else
    SaveDocToLatexTex
#End If
End Sub
Sub Rib_ConvPDFLatex(control As IRibbonControl)
#If Mac Then
    MsgBox "This function is not avaiable on Mac.", vbOKOnly, "Mac issue"
#Else
    SaveDocToLatexPdf
#End If
End Sub

'Callback for OpenFigurer onAction
Sub Rib_figurer(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        OpenWordFile ("Figurer.docx")
    ElseIf Sprog.SprogNr = 3 Then
        OpenWordFile ("Figurer_spansk.docx")
    Else
        OpenWordFile ("Figurer_english.docx")
    End If
End Sub
'Callback for InsertExcel onAction
Sub Rib_insertexcel(control As IRibbonControl)
    Call InsertIndlejretExcel
End Sub
'Callback for TabelToList onAction
Sub Rib_TabelToList(control As IRibbonControl)
    TabelToList
End Sub
'Callback for TabelToList onAction
Sub Rib_ListToTabel(control As IRibbonControl)
    ListToTabel
End Sub
'Callback for InsertExcel onAction
Sub Rib_trianglesolver(control As IRibbonControl)
    Dim UFtriangle As New UserFormTrekant
    UFtriangle.Show vbModeless
End Sub
'Callback for Wolfram onAction
Sub Rib_WolframAlpha(control As IRibbonControl)
    OpenWolframAlpha
End Sub
Sub Rib_wasolve(control As IRibbonControl)
    OpenWolframAlpha "solve"
End Sub
'Callback for ButtonOm onAction
Sub Rib_om(control As IRibbonControl)
    OmMathMenu
End Sub
'Callback for ButtonHelp onAction
Sub Rib_Help(control As IRibbonControl)
    If Sprog.SprogNr = 1 Then
        OpenWordFile ("WordMatManual.docx")
    Else
        OpenWordFile ("WordMatManual_english.docx")
    End If
'    hjælpeMenu
End Sub
Sub Rib_HelpOnline(control As IRibbonControl)
    OpenLink "https://sites.google.com/site/wordmat/"
End Sub
Sub Rib_HelpMaxima(control As IRibbonControl)
    OpenLink "https://maxima.sourceforge.io/docs/manual/maxima_toc.html#SEC_Contents"
' Der er ikke længere lokal hjælp med i Maxima distributionen
'#If Mac Then
'    OpenLink "https://maxima.sourceforge.io/docs/manual/maxima_toc.html#SEC_Contents"
'#Else
'Dim ReturnValue
'Dim sti As String
'    On Error GoTo fejl
'    OpenLink GetProgramFilesDir & "\WordMat\Maxima-5.47.0\share\maxima\5.47.0\doc\html\index.html", True
''    OpenLink GetProgramFilesDir & "\WordMat\Maxima-5.47.0\share\maxima\5.47.0\doc\html\index.html", True
''    sti = "cmd /C """ & GetProgramFilesDir & "\WordMat\Maxima-5.45.1\share\maxima\5.45.1\doc\chm\maxima.chm"""
''    ReturnValue = Shell(sti, vbHide)
'GoTo slut
'fejl:
'    OpenLink "https://maxima.sourceforge.io/docs/manual/maxima_toc.html#SEC_Contents"
'slut:
'#End If
End Sub
'Callback for ButtonCheckUpdate onAction
Sub Rib_CheckForUpdate(control As IRibbonControl)
    CheckForUpdate
End Sub
Sub Rib_CheckForUpdateGeoGebra(control As IRibbonControl)
    InstallGeoGebra False
End Sub

'Callback for ButtonGenveje onAction
Sub Rib_Genveje(control As IRibbonControl)
    UserFormGenveje.Show
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''Language labels''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Rib_GetLabelFormulae(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibFormulae
End Sub
Sub Rib_GetLabelPercentage(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibPercentage
End Sub
'Callback for proc1 getLabel
Sub Rib_FSpercentage1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(435)
End Sub
Sub Rib_FSpercentage2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(436)
End Sub
Sub Rib_FSpercentage3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "A=b" & ChrW(183) & "((1+r)" & ChrW(&H207F) & "- 1) / r" & "     Annuitetsopsparing"
End Sub
Sub Rib_FSpercentage4(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "y=G" & ChrW(183) & "r/(1-(1+r)" & ChrW(&H207B) & ChrW(&H207F) & ")     Annuitetslån"
End Sub

Sub Rib_GetLabelFunctions(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibFunctions
End Sub
Sub Rib_FSlinear1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(437)
End Sub
Sub Rib_FSlinear2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(438)
End Sub
Sub Rib_FSlinear3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(439)
End Sub
Sub Rib_FSlinear4(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(440)
End Sub
Sub Rib_FSexp1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(441)
End Sub
Sub Rib_FSexp2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(442)
End Sub
Sub Rib_FSexp3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(443)
End Sub
Sub Rib_FSexp4(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(444)
End Sub
Sub Rib_FSexp5(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(445)
End Sub
Sub Rib_FSexp6(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(446)
End Sub
Sub Rib_FSexp7(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(447)
End Sub
Sub Rib_FSpow1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(448)
End Sub
Sub Rib_FSpow2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(449)
End Sub
Sub Rib_FSpow3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(450)
End Sub
Sub Rib_FSpol1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(451)
End Sub

Sub Rib_GetLabelGeometry(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibGeometry
End Sub
Sub Rib_FSgeo1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(452)
End Sub
Sub Rib_FSgeo2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(453)
End Sub
Sub Rib_FSgeo3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(454)
End Sub

Sub Rib_GetLabelProbabilityShort(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibProbabilityShort
End Sub
Sub Rib_FSBinomial(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(481)
End Sub
Sub Rib_FSNormaldist(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(482)
End Sub
Sub Rib_FSChi2dist(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(483)
End Sub
Sub Rib_FSprob1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(455)
End Sub
Sub Rib_FSprob2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(456)
End Sub
Sub Rib_FSprob3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(457)
End Sub
Sub Rib_FSprob4(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(458)
End Sub
Sub Rib_FSprob5(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(459)
End Sub
Sub Rib_FSprob5a(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(687)
End Sub
Sub Rib_FSprob6(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(460)
End Sub
Sub Rib_FSprob7(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(461)
End Sub
Sub Rib_FSprob8(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(462)
End Sub

Sub Rib_GetLabelInfinitesimalShort(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibInfinitesimalShort
End Sub
Sub Rib_FSinf1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(463)
End Sub
Sub Rib_FSinf2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(464)
End Sub
Sub Rib_FSinf3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(465)
End Sub

Sub Rib_GetLabelVector(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibVector
End Sub
Sub Rib_FS2D(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(484)
End Sub
Sub Rib_FS3D(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(485)
End Sub
Sub Rib_FSvec1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(466)
End Sub
Sub Rib_FSvec2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(467)
End Sub
Sub Rib_FSvec3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(468)
End Sub
Sub Rib_FSvec4(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(469)
End Sub
Sub Rib_FSvec5(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(470)
End Sub
Sub Rib_FSvec6(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(471)
End Sub
Sub Rib_FSvec7(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(472)
End Sub
Sub Rib_FSvec8(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(473)
End Sub
Sub Rib_FSvec9(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(474)
End Sub
Sub Rib_FSvec10(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(475)
End Sub
Sub Rib_FSvec11(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(476)
End Sub
Sub Rib_FSvec12(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(477)
End Sub
Sub Rib_FSvec13(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(478)
End Sub
Sub Rib_FSvec14(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(479)
End Sub
Sub Rib_FSvec15(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(480)
End Sub

Sub Rib_GetLabelMath(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibMath
End Sub
Sub Rib_GetLabelPhysics(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibPhysics
End Sub
Sub Rib_GetLabelChemistry(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibChemistry
End Sub

Sub Rib_GetLabelSettings(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibSettings
End Sub
Sub Rib_GetLabelSettingsShort(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibSettingsShort
End Sub
Sub Rib_GetLabelSciNot(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibSciNot
End Sub

Sub Rib_GetLabelBeregn(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibBeregn
End Sub
Sub Rib_GetLabelMaximaCommand(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibMaximaCommand
End Sub
Sub Rib_GetLabelSolve(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibSolve
End Sub
Sub Rib_GetLabelSolveNum(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibSolveNum
End Sub
Sub Rib_GetLabelEliminate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibEliminate
End Sub
Sub Rib_GetLabelTestTF(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.TestTF
End Sub
Sub Rib_GetLabelSolveDE(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibSolveDE
End Sub
'Callback for solvedenum getLabel
Sub Rib_GetLabelSolveDEnum(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RibSolveDEnum
End Sub
Sub Rib_GetLabelDefinitions(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefinitions
End Sub
Sub Rib_GetLabelDeleteDefs(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefDelete
End Sub
Sub Rib_GetLabelDefineVariable(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefVariable
End Sub
Sub Rib_GetLabelDefineFunction(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefFunction
End Sub
Sub Rib_GetLabelDefineEquation(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefEquation
End Sub
Sub Rib_GetLabelDefineConstants(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDefConstants
End Sub
Sub Rib_GetLabelReduce(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibReduce
End Sub
Sub Rib_GetLabelSimplify(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Simplify
End Sub
Sub Rib_GetLabelFactor(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Factor
End Sub
Sub Rib_GetLabelExpand(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Expand
End Sub
Sub Rib_GetLabelInfinitesimal(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibInfinitesimal
End Sub
Sub Rib_GetLabelIntegrate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibIntegrate
End Sub
Sub Rib_GetLabelDifferentiate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDif
End Sub

Sub Rib_GetLabelPlotting(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibPlotting
End Sub
Sub Rib_GetLabelShowGraph(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibShowGraph
End Sub

Sub Rib_getVisibleGnuPlot(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub
Sub Rib_getVisibleGraph(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub
Sub Rib_GetLabelDirectionField(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDirectionField
End Sub
Sub Rib_GetLabelInsertGeoGebra(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibInsertGeoGebra
End Sub
Sub Rib_getVisibleInsertGeoGebra(control As IRibbonControl, ByRef returnedVal)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub
Sub Rib_GetLabel3DPlot(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Rib3DPlot
End Sub
Sub Rib_GetLabel3dRotate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Rib3DRotate
End Sub
Sub Rib_GetLabelStatistics(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibStatistics
End Sub

Sub Rib_GetLabelUgrup(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibUgrup
End Sub
Sub Rib_GetLabelGrup(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibGrup
End Sub
Sub Rib_GetLabelStickChart(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibStickChart
End Sub
Sub Rib_GetLabelHistogram(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibHistogram
End Sub
Sub Rib_GetLabelStepChart(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibStepChart
End Sub
Sub Rib_GetLabelCumChart(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibCumChart
End Sub
Sub Rib_GetLabelBoxPlot(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibBoxPlot
End Sub

Sub Rib_GetLabelStatProb(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibStatProb
End Sub
Sub Rib_GetLabelRegression(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibRegression
End Sub
Sub Rib_GetLabelInsertTable(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibInsertTable
End Sub
Sub Rib_GetLabelLinRegr(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibLinRegr
End Sub
Sub Rib_GetLabelExpRegr(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibExpRegr
End Sub
Sub Rib_GetLabelPowRegr(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibPowRegr
End Sub
Sub Rib_GetLabelPolRegr(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibPolRegr
End Sub

Sub Rib_GetLabelDistributions(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDistributions
End Sub
Sub Rib_GetLabelBinomDist(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibBinomDist
End Sub
Sub Rib_GetLabelNormalDist(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibNormalDist
End Sub
Sub Rib_GetLabelChi2Dist(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibChi2Dist
End Sub
Sub Rib_GetLabeltDist(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = "t-fordeling"
End Sub

Sub Rib_GetLabelTest(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibTest
End Sub
Sub Rib_GetLabelBinomTest(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibBinomTest
End Sub
Sub Rib_GetLabelChi2Test(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibChi2Test
End Sub
Sub Rib_GetLabelSimulation(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibSimulation
End Sub
Sub Rib_GetLabelGroup(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibGroup
End Sub

Sub Rib_GetLabelDiverse(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibDiverse
End Sub
Sub Rib_GetLabelNewEquation(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibNewEquation
End Sub
Sub Rib_GetLabelNumEquation(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(1)
End Sub
Sub Rib_GetLabelNumEquationRef(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(2)
End Sub
Sub Rib_GetLabelRefEquation(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(3)
End Sub
Sub Rib_GetLabelSetEquationNo(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(6)
End Sub
Sub Rib_GetLabelInsertEquationSection(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(8)
End Sub
Sub Rib_GetLabelUpdateEqNo(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(9)
End Sub
Sub Rib_GetLabelSymbols(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibSymbols
End Sub
Sub Rib_GetLabelLatexTemplate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.A(655)
End Sub
Sub Rib_GetLabelMultSign(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibMultsign
End Sub
Sub Rib_GetLabelFigurs(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibFigurs
End Sub
Sub Rib_GetLabelTable(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibTable
End Sub
Sub Rib_GetLabelInsertExcel(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibInsertExcel
End Sub
Sub Rib_GetLabelTableToList(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibTableToList
End Sub
Sub Rib_GetLabelListToTable(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibListToTable
End Sub
Sub Rib_GetLabelTriangle(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.RibTriangle
End Sub
Sub Rib_GetLabelHelp(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Help
End Sub
Sub Rib_GetLabelManual(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Manual
End Sub
Sub Rib_GetLabelManualDoc(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.ManualDoc
End Sub
Sub Rib_GetLabelManualOnline(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.ManualOnline
End Sub
Sub Rib_GetLabelMaximaHelp(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(0)
End Sub
Sub Rib_GetLabelAbout(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.About
End Sub
Sub Rib_GetLabelUpdate(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Update
End Sub
Sub Rib_GetLabelShortcuts(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Sprog.Shortcuts
End Sub
Sub Rib_GetLabelUserRegr(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.RegrUser
End Sub
Sub Rib_GetLabelRegrExcel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Excel regression"
End Sub

' screentips
Sub Rib_STmathformula(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(486)
End Sub
Sub Rib_STphysicsformula(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(487)
End Sub
Sub Rib_STchemistryformula(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(488)
End Sub
Sub Rib_STauto1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(489)
End Sub
Sub Rib_STauto2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(490)
End Sub
Sub Rib_STexact1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(491)
End Sub
Sub Rib_STexact2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(492)
End Sub
Sub Rib_STnum1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(493)
End Sub
Sub Rib_STnum2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(494)
End Sub
Sub Rib_STbetcif1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(495)
End Sub
Sub Rib_STbetcif2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(496)
End Sub
Sub Rib_STrad1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(497)
End Sub
Sub Rib_STrad2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(498)
End Sub
Sub Rib_STsci1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(499)
End Sub
Sub Rib_STsci2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(500)
End Sub
Sub Rib_STset1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(501)
End Sub
Sub Rib_STset2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(502)
End Sub
Sub Rib_STcalc1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(503)
End Sub
Sub Rib_STcalc2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(504)
End Sub
Sub Rib_STmaxima1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(505)
End Sub
Sub Rib_STmaxima2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(506)
End Sub
Sub Rib_STsolve1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(507)
End Sub
Sub Rib_STsolve2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(508)
End Sub
Sub Rib_STsolvenum1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(509)
End Sub
Sub Rib_STsolvenum2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(510)
End Sub
Sub Rib_STeliminate1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(511)
End Sub
Sub Rib_STeliminate2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(512)
End Sub
Sub Rib_STtest1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(513)
End Sub
Sub Rib_STtest2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(514)
End Sub
Sub Rib_STsolvede1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(515)
End Sub
Sub Rib_STsolvede2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(516)
End Sub
Sub Rib_STsolvedenum1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(517)
End Sub
Sub Rib_STsolvedenum2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(518)
End Sub
Sub Rib_STdef1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(519)
End Sub
Sub Rib_STdef2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(520)
End Sub
Sub Rib_STsletdef1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(521)
End Sub
Sub Rib_STsletdef2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(522)
End Sub
Sub Rib_STdefine1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(523)
End Sub
Sub Rib_STdefine2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(524)
End Sub
Sub Rib_STdefconst1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(525)
End Sub
Sub Rib_STdefconst2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(526)
End Sub
Sub Rib_STreduce1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(527)
End Sub
Sub Rib_STreduce2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(528)
End Sub
Sub Rib_STsimplify1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(529)
End Sub
Sub Rib_STsimplify2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(530)
End Sub
Sub Rib_STfactor1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(531)
End Sub
Sub Rib_STfactor2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(532)
End Sub
Sub Rib_STexpand1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(533)
End Sub
Sub Rib_STexpand2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(534)
End Sub
Sub Rib_STinf1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(535)
End Sub
Sub Rib_STinf2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(536)
End Sub
Sub Rib_STdiff1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(537)
End Sub
Sub Rib_STdiff2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(538)
End Sub
Sub Rib_STint1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(539)
End Sub
Sub Rib_STint2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(540)
End Sub
Sub Rib_STwolfram1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(541)
End Sub
Sub Rib_STwolfram2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(542)
End Sub
Sub Rib_STwasolve1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(543)
End Sub
Sub Rib_STwasolve2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(544)
End Sub
Sub Rib_STplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(545)
End Sub
Sub Rib_STplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(546)
End Sub
Sub Rib_STgnuplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(547)
End Sub
Sub Rib_STgnuplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(548)
End Sub
Sub Rib_STgraphplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(549)
End Sub
Sub Rib_STgraphplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(550)
End Sub
Sub Rib_STgeogebraplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(551)
End Sub
Sub Rib_STgeogebraplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(552)
End Sub
Sub Rib_STexcelplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(553)
End Sub
Sub Rib_STexcelplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(554)
End Sub
Sub Rib_STretnfelt1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(555)
End Sub
Sub Rib_STretnfelt2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(556)
End Sub
Sub Rib_STinsertgeo1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(557)
End Sub
Sub Rib_STinsertgeo2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(558)
End Sub
Sub Rib_ST3dplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(559)
End Sub
Sub Rib_ST3dplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(560)
End Sub
Sub Rib_STomdleg1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(561)
End Sub
Sub Rib_STomdleg2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(562)
End Sub
Sub Rib_STstat1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(563)
End Sub
Sub Rib_STstat2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(564)
End Sub
Sub Rib_STugrup1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(565)
End Sub
Sub Rib_STugrup2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(566)
End Sub
Sub Rib_STgrup1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(567)
End Sub
Sub Rib_STgrup2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(568)
End Sub
Sub Rib_STpinde1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(569)
End Sub
Sub Rib_STpinde2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(570)
End Sub
Sub Rib_SThist1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(571)
End Sub
Sub Rib_SThist2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(572)
End Sub
Sub Rib_STtrap1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(573)
End Sub
Sub Rib_STtrap2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(574)
End Sub
Sub Rib_STsumkurve1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(575)
End Sub
Sub Rib_STsumkurve2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(576)
End Sub
Sub Rib_STboksplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(577)
End Sub
Sub Rib_STboksplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(578)
End Sub
Sub Rib_STregr1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(579)
End Sub
Sub Rib_STregr2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(580)
End Sub
Sub Rib_STtable1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(581)
End Sub
Sub Rib_STtable2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(582)
End Sub
Sub Rib_STdistrib1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(583)
End Sub
Sub Rib_STdistrib2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(584)
End Sub
Sub Rib_STbinom1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(585)
End Sub
Sub Rib_STbinom2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(586)
End Sub
Sub Rib_STnorm1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(587)
End Sub
Sub Rib_STnorm2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(588)
End Sub
Sub Rib_STchi21(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(589)
End Sub
Sub Rib_STchi22(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(590)
End Sub
Sub Rib_STt1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Student's t-distribution"
End Sub
Sub Rib_STt2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Student's t-distribution"
End Sub

Sub Rib_STtestmenu1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(591)
End Sub
Sub Rib_STtestmenu2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(592)
End Sub
Sub Rib_STbinomtest1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(593)
End Sub
Sub Rib_STbinomtest2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(594)
End Sub
Sub Rib_STchitest1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(595)
End Sub
Sub Rib_STchitest2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(596)
End Sub
Sub Rib_STgof1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(597)
End Sub
Sub Rib_STgof2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(598)
End Sub
Sub Rib_STsim1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(599)
End Sub
Sub Rib_STsim2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(600)
End Sub
Sub Rib_STneweq1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(601)
End Sub
Sub Rib_STneweq2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(602)
End Sub
Sub Rib_STnumeq1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(603)
End Sub
Sub Rib_STnumeq2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(604)
End Sub
Sub Rib_STrefeq1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(605)
End Sub
Sub Rib_STrefeq2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(606)
End Sub
Sub Rib_STinsrefeq1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(607)
End Sub
Sub Rib_STinsrefeq2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(608)
End Sub
Sub Rib_STseteqno1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(609)
End Sub
Sub Rib_STseteqno2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(610)
End Sub
Sub Rib_STeqsection1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(611)
End Sub
Sub Rib_STeqsection2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(612)
End Sub
Sub Rib_STequpdate1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(613)
End Sub
Sub Rib_STequpdate2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(614)
End Sub
Sub Rib_STsymbols1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(615)
End Sub
Sub Rib_STsymbols2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(616)
End Sub
Sub Rib_STtilprik1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(617)
End Sub
Sub Rib_STtilprik2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(618)
End Sub
Sub Rib_STlatex1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(619)
End Sub
Sub Rib_STlatex2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(620)
End Sub
Sub Rib_STconvlatex1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(621)
End Sub
Sub Rib_STconvlatex2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(622)
End Sub
Sub Rib_STtostar1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(623)
End Sub
Sub Rib_STtostar2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(624)
End Sub
Sub Rib_STtilkomma1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(625)
End Sub
Sub Rib_STtilkomma2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(626)
End Sub
Sub Rib_STtilpunktum1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(627)
End Sub
Sub Rib_STtilpunktum2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(628)
End Sub
Sub Rib_STfigur1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(629)
End Sub
Sub Rib_STfigur2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(630)
End Sub
Sub Rib_STtables1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(633)
End Sub
Sub Rib_STtables2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(634)
End Sub
Sub Rib_STembedexcel1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(635)
End Sub
Sub Rib_STembedexcel2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(636)
End Sub
Sub Rib_STtolist1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(637)
End Sub
Sub Rib_STtolist2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(638)
End Sub
Sub Rib_STtotable1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(639)
End Sub
Sub Rib_STtotable2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(640)
End Sub
Sub Rib_STtriangle1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(641)
End Sub
Sub Rib_STtriangle2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(642)
End Sub
Sub Rib_STmanual1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(643)
End Sub
Sub Rib_STmanual2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(644)
End Sub
Sub Rib_STwebmanual1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(645)
End Sub
Sub Rib_STwebmanual2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(646)
End Sub
Sub Rib_STmaxmanual1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(647)
End Sub
Sub Rib_STmaxmanual2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(648)
End Sub
Sub Rib_STabout1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(649)
End Sub
Sub Rib_STabout2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(650)
End Sub
Sub Rib_STupdate1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(651)
End Sub
Sub Rib_STupdate2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(652)
End Sub
Sub Rib_STgenveje1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(653)
End Sub
Sub Rib_STgenveje2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(654)
End Sub
Sub Rib_STlatextemplate1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(656)
End Sub
Sub Rib_STlatextemplate2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Sprog.A(657)
End Sub
'Callback for ButtonGeoGebra getScreentip
Sub Rib_STgeogebraBplot1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "GeoGebra webapplet in a browser"
End Sub

'Callback for ButtonGeoGebra getSupertip
Sub Rib_STgeogebraBplot2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Plot graphs and points using GeoGebra webapplet in a browser. Does not require internet access. Quite fast. Many functions."
End Sub

