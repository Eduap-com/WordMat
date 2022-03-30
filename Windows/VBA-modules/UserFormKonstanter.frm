VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormKonstanter 
   Caption         =   "Konstanter"
   ClientHeight    =   7020
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   7890
   OleObjectBlob   =   "UserFormKonstanter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormKonstanter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_annuller_Click()
Unload Me
'    Me.Hide
End Sub

Private Sub CommandButton_ok_Click()
Dim text As String, mult As String

mult = MaximaGangeTegn

If Sprog.SprogNr = 1 Then
    text = "Definer:  "
Else
    text = "Define:  "
End If

If CheckBox_planck.Value = True Then text = text & "h=6" & DecSeparator & "62606896" & MaximaGangeTegn & "10^(-34) J" & MaximaGangeTegn & "s" & ListSeparator
If CheckBox_planckev.Value = True Then text = text & "h=4" & DecSeparator & "13566733" & MaximaGangeTegn & "10^(-15) eV" & MaximaGangeTegn & "s" & ListSeparator
If CheckBox_c.Value = True Then text = text & "c=299792458 m" & MaximaGangeTegn & "s^(-1)" & ListSeparator
If CheckBox_G.Value = True Then text = text & "G=6" & DecSeparator & "67428" & MaximaGangeTegn & "10^(-11) N" & MaximaGangeTegn & "m^2" & MaximaGangeTegn & "kg^-2" & ListSeparator
If CheckBox_ge.Value = True Then text = text & "g_" & Sprog.A(255) & "=9" & DecSeparator & "82m" & MaximaGangeTegn & "s^-2" & ListSeparator
If CheckBox_e.Value = True Then text = text & "e_l=1" & DecSeparator & "602176487" & MaximaGangeTegn & "10^-19 C" & ListSeparator
If CheckBox_NA.Value = True Then text = text & "N_A=6" & DecSeparator & "02214179" & MaximaGangeTegn & "10^23 mol^(-1) " & ListSeparator
If CheckBox_R.Value = True Then text = text & "R=8" & DecSeparator & "314472 J" & MaximaGangeTegn & "mol^-1" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_R2.Value = True Then text = text & "R=0" & DecSeparator & "0821 L" & MaximaGangeTegn & "atm" & MaximaGangeTegn & "mol^-1" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_k.Value = True Then text = text & "k=1" & DecSeparator & "3806504" & MaximaGangeTegn & "10^-23 J" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_mu0.Value = True Then text = text & VBA.ChrW(956) & "_0=4" & VBA.ChrW(960) & MaximaGangeTegn & "10^-7 N" & MaximaGangeTegn & "A^-2" & ListSeparator
If CheckBox_e0.Value = True Then text = text & VBA.ChrW(1013) & "_0=8" & DecSeparator & "854187817" & MaximaGangeTegn & "10^-12 F" & MaximaGangeTegn & "m^-1" & ListSeparator
If CheckBox_sigma.Value = True Then text = text & VBA.ChrW(963) & "=5" & DecSeparator & "670400" & MaximaGangeTegn & "10^-8 W" & MaximaGangeTegn & "m^-2" & MaximaGangeTegn & "K^-4" & ListSeparator
If CheckBox_coulomb.Value = True Then text = text & "k=8" & DecSeparator & "99" & MaximaGangeTegn & "10^9 N" & MaximaGangeTegn & "m^2" & MaximaGangeTegn & "C^-2" & ListSeparator

If CheckBox_u.Value = True Then text = text & "u=1" & DecSeparator & "660538782" & MaximaGangeTegn & "10^-27 kg" & ListSeparator
If CheckBox_uev.Value = True Then text = text & "u=931" & DecSeparator & "494028 MeV" & MaximaGangeTegn & "c^-2" & ListSeparator

If CheckBox_cvand.Value = True Then text = text & "c_" & Sprog.A(256) & "=4181 J" & MaximaGangeTegn & "kg^-1" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_calu.Value = True Then text = text & "c_alu=897 J" & MaximaGangeTegn & "kg^-1" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_ccop.Value = True Then text = text & "c_" & Sprog.A(257) & "=385 J" & MaximaGangeTegn & "kg^-1" & MaximaGangeTegn & "K^-1" & ListSeparator
If CheckBox_cjern.Value = True Then text = text & "c_" & Sprog.A(258) & "=450 J" & MaximaGangeTegn & "kg^-1" & MaximaGangeTegn & "K^-1" & ListSeparator

If CheckBox_me.Value = True Then text = text & "m_e=5" & DecSeparator & "4857990943" & MaximaGangeTegn & "10^-4 u" & ListSeparator
If CheckBox_mekg.Value = True Then text = text & "m_e=9" & DecSeparator & "10938215" & MaximaGangeTegn & "10^-31 kg" & ListSeparator
If CheckBox_mp.Value = True Then text = text & "m_p=1" & DecSeparator & "00727646677 u" & ListSeparator
If CheckBox_mn.Value = True Then text = text & "m_n=1" & DecSeparator & "00866491597 u" & ListSeparator

If CheckBox_mj.Value = True Then text = text & "m_" & Sprog.A(255) & "=5" & DecSeparator & "9737" & MaximaGangeTegn & "10^24 kg" & ListSeparator
If CheckBox_rjord.Value = True Then text = text & "r_" & Sprog.A(255) & "=6371km" & ListSeparator
If CheckBox_AU.Value = True Then text = text & "AU=1" & DecSeparator & "50" & MaximaGangeTegn & "10^11 m" & ListSeparator
If CheckBox_mmoon.Value = True Then text = text & "m_" & Sprog.A(259) & "=7" & DecSeparator & "3477" & MaximaGangeTegn & "10^22 kg" & ListSeparator
If CheckBox_rmoon.Value = True Then text = text & "r_" & Sprog.A(259) & "=1737km" & ListSeparator
If CheckBox_msol.Value = True Then text = text & "m_" & Sprog.A(260) & "=1" & DecSeparator & "98892" & MaximaGangeTegn & "10^30 kg" & ListSeparator


'If Len(text) > 12 Then
'    text = Left(text, Len(text) - 1)
'
'    If DecSeparator = "." Then
'        text = Replace(text, ",", ".")
'    End If

    insertribformel "", text
'End If

slut:
    Me.Hide
End Sub

Private Sub Label1_Click()
    OpenLink "http://physics.nist.gov/cuu/Constants/index.html"
End Sub

Private Sub UserForm_Activate()
    SetCaptions
End Sub

Sub SetCaptions()
' ChrW(&H2070)  opløftet 0
' 185  opløftet i 1.
' 178  opløftet i 2.
' chrw(179)  opløftet i 3.
' ChrW(&H207B)  opløftet -
' ChrW(&H2074)  opløftet 4
' ChrW(&H2075)  opløftet 5
' ChrW(&H2076)  opløftet 6
' ChrW(&H2077)  opløftet 7
' ChrW(&H2078)  opløftet 8
' ChrW(&H2079)  opløftet 9
' ChrW(&H2092)  sænket 0
    Dim g As String
    g = MaximaGangeTegn
    
    Me.Caption = Sprog.A(70)
    MultiPage1.Pages(0).Caption = Sprog.A(261)
    MultiPage1.Pages(1).Caption = Sprog.A(262)
    MultiPage1.Pages(2).Caption = Sprog.A(263)
    MultiPage1.Pages(3).Caption = Sprog.A(264)
    MultiPage1.Pages(4).Caption = Sprog.A(265)
    CommandButton_annuller.Caption = Sprog.Cancel
    CommandButton_ok.Caption = Sprog.OK
    Label1.Caption = Sprog.A(71)
    
    CheckBox_c.Caption = "c = 299792458 m s" & ChrW(&H207B) & ChrW(185) & "   -   " & Sprog.A(266)
    CheckBox_planck.Caption = "h = 6" & DecSeparator & "63 " & g & " 10" & ChrW(&H207B) & ChrW(179) & ChrW(&H2074) & " Js   -   " & Sprog.A(267)
    CheckBox_planckev.Caption = "h = 4" & DecSeparator & "13566733 " & g & " 10" & ChrW(&H207B) & ChrW(185) & ChrW(&H2075) & " eVs   -   " & Sprog.A(268)
    CheckBox_G.Caption = "G = 6" & DecSeparator & "67428 " & g & " 10" & ChrW(&H207B) & ChrW(185) & ChrW(185) & " Nm" & ChrW(178) & " kg" & ChrW(&H207B) & ChrW(178) & "   -   " & Sprog.A(269)
    CheckBox_ge.Caption = "g_" & Sprog.A(255) & " = 9" & DecSeparator & "82 m s" & ChrW(&H207B) & ChrW(178) & "   -   " & Sprog.A(270)
    CheckBox_e.Caption = "e" & ChrW(&H2097) & " = 1" & DecSeparator & "602176487 " & g & " 10" & ChrW(&H207B) & ChrW(185) & ChrW(&H2079) & " C    -   " & Sprog.A(271)
    CheckBox_NA.Caption = "N_A = 6" & DecSeparator & "02214179 " & g & " 10" & ChrW(178) & ChrW(179) & " mol" & ChrW(&H207B) & ChrW(185) & "   -   " & Sprog.A(272)
    CheckBox_R.Caption = "R = 8" & DecSeparator & "314472 J/(mol K)    -    " & Sprog.A(273)
    CheckBox_R2.Caption = "R = 0" & DecSeparator & "0821 l atm/(mol K)    -    " & Sprog.A(273)
    CheckBox_k.Caption = "k = 1" & DecSeparator & "3806504 " & g & " 10" & ChrW(&H207B) & ChrW(178) & ChrW(179) & " J K" & ChrW(&H207B) & ChrW(185) & "   -   " & Sprog.A(274)
    CheckBox_mu0.Caption = ChrW(956) & ChrW(&H2092) & " = 4" & ChrW(960) & g & " 10" & ChrW(&H207B) & ChrW(&H2077) & " N A" & ChrW(&H207B) & ChrW(178) & "   -   " & Sprog.A(275)
    CheckBox_e0.Caption = "e" & ChrW(&H2092) & " = 8" & DecSeparator & "854187817 " & g & " 10" & ChrW(&H207B) & ChrW(185) & ChrW(178) & " F m" & ChrW(&H207B) & ChrW(185) & "   -   " & Sprog.A(276) & " = 1/(" & ChrW(956) & ChrW(&H2092) & "c" & ChrW(178) & ")"
    CheckBox_coulomb.Caption = "k = 8" & DecSeparator & "99 " & g & " 10" & ChrW(&H2079) & " N m" & ChrW(178) & " C" & ChrW(&H207B) & ChrW(178) & "   -   " & Sprog.A(277)
    CheckBox_sigma.Caption = ChrW(963) & " = 5" & DecSeparator & "670400 " & g & " 10" & ChrW(&H207B) & ChrW(&H2078) & " W m" & ChrW(&H207B) & ChrW(178) & " K" & ChrW(&H207B) & ChrW(&H2074) & "   -   " & Sprog.A(278)
    
    CheckBox_u.Caption = "u = 1" & DecSeparator & "660 538 782 " & g & " 10" & ChrW(&H207B) & ChrW(178) & ChrW(&H2077) & " kg    -   " & Sprog.A(279)
    CheckBox_uev.Caption = "u = 931" & DecSeparator & "494 028 MeV c" & ChrW(&H207B) & ChrW(178) & "    -   " & Sprog.A(280)
    
    CheckBox_cvand.Caption = "c_" & Sprog.A(256) & " = 4181 J kg" & ChrW(&H207B) & ChrW(185) & " K" & ChrW(&H207B) & ChrW(185) & "    -   " & Sprog.A(281) & " " & Sprog.A(256)
    CheckBox_calu.Caption = "c_alu = 897 J kg" & ChrW(&H207B) & ChrW(185) & " K" & ChrW(&H207B) & ChrW(185) & "    -   " & Sprog.A(281) & " aluminium"
    CheckBox_ccop.Caption = "c_" & Sprog.A(257) & " = 385 J kg" & ChrW(&H207B) & ChrW(185) & " K" & ChrW(&H207B) & ChrW(185) & "    -   " & Sprog.A(281) & " " & Sprog.A(257)
    CheckBox_cjern.Caption = "c_" & Sprog.A(258) & " = 450 J kg" & ChrW(&H207B) & ChrW(185) & " K" & ChrW(&H207B) & ChrW(185) & "    -   " & Sprog.A(281) & " " & Sprog.A(258)
    
    CheckBox_me.Caption = "m" & ChrW(&H2091) & " = 5" & DecSeparator & "485 799 0943 " & g & " 10" & ChrW(&H207B) & ChrW(&H2074) & " u   -   " & Sprog.A(282)
    CheckBox_mekg.Caption = "m" & ChrW(&H2091) & " = 9" & DecSeparator & "109 382 15 " & g & " 10" & ChrW(&H207B) & ChrW(179) & ChrW(185) & " kg   -   " & Sprog.A(282) & " " & Sprog.A(283)
    CheckBox_mp.Caption = "m" & ChrW(&H209A) & " = 1" & DecSeparator & "00727646677 u   -   " & Sprog.A(284)
    CheckBox_mn.Caption = "m" & ChrW(&H2099) & " = 1" & DecSeparator & "00866491597 u   -   " & Sprog.A(285)
    
    CheckBox_mj.Caption = "m_" & Sprog.A(255) & " = 5" & DecSeparator & "9737 " & g & " 10" & ChrW(178) & ChrW(&H2074) & " kg   -   " & Sprog.A(286)
    CheckBox_rjord.Caption = "r_" & Sprog.A(255) & " = 6371 km    -   " & Sprog.A(287)
    CheckBox_AU.Caption = "AU = 1" & DecSeparator & "50 " & g & " 10" & ChrW(185) & ChrW(185) & " m   -   " & Sprog.A(288)
    CheckBox_msol.Caption = "m_" & Sprog.A(260) & " = 1" & DecSeparator & "98892 " & g & " 10" & ChrW(179) & ChrW(&H2070) & " kg   -   " & Sprog.A(289)
    CheckBox_mmoon.Caption = "m_" & Sprog.A(259) & " = 7" & DecSeparator & "3477 " & g & " 10" & ChrW(178) & ChrW(178) & " kg   -   " & Sprog.A(290)
    CheckBox_rmoon.Caption = "r_" & Sprog.A(259) & " = 1737 km    -   " & Sprog.A(291)
    
    
End Sub

