VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SprogNr As Integer
Private SA(800, 5) As String
'0=WordLanguage, 1=Dansk, 2=English, 3=Spanish

Private Sub Class_Initialize()
    CheckSetting
    LoadSprogArray
End Sub
Public Sub CheckSetting()
Dim b As Boolean
    If SprogNr < 1 Then b = True
    If LanguageSetting = 0 Then
        SprogNr = GetOfficeLanguage
        If SprogNr <> 1 Then GenerateKeyboardShortcuts
    Else
        SprogNr = LanguageSetting
    End If
    If b = True Then
    End If
End Sub
Function GetOfficeLanguage() As Integer
'Reads the UI-language for Word
Select Case Application.International(wdProductLanguageID)
Case 1030 ' Dansk
    GetOfficeLanguage = 1
Case 1034 ' Spanish
    GetOfficeLanguage = 3
Case Else ' 1033 is english
    GetOfficeLanguage = 2
End Select

End Function

Public Function A(i As Long) As String
On Error Resume Next
    If SA(i, SprogNr - 1) <> vbNullString Then
        A = SA(i, SprogNr - 1)
    ElseIf SA(i, 1) <> vbNullString Then
        A = SA(i, 1)
    Else
        If SA(0, 0) = vbNullString Then LoadSprogArray
        A = SA(i, 0)
    End If
End Function
Public Sub LoadSprogArray()
' 0 - dansk , 1-engelsk, 2-spansk
' There is a limit to how long a sub can be, hence divided in two

' for label in menu
SA(0, 0) = "Hjælp til Maxima-kommandoer"
SA(0, 1) = "Help for Maxima commands"
SA(0, 2) = "Ayuda para comandos de Máxima"
SA(0, 3) = ""
SA(0, 4) = ""
SA(0, 5) = ""

SA(1, 0) = "Indsæt nummereret ligning"
SA(1, 1) = "Insert numbered equation"
SA(1, 2) = "Inserte ecuaciones numeradas"
SA(1, 3) = ""
SA(1, 4) = ""
SA(1, 5) = ""

SA(2, 0) = "Indsæt nummeret ligning til reference"
SA(2, 1) = "Insert numbered equation for reference"
SA(2, 2) = "Inserte ecuaciones numeradas para referencia"
SA(2, 3) = ""
SA(2, 4) = ""
SA(2, 5) = ""

SA(3, 0) = "Indsæt reference til ligning"
SA(3, 1) = "Insert reference to equation"
SA(3, 4) = ""
SA(3, 3) = ""
SA(3, 2) = "Inserte referencia a la ecuación"
SA(3, 5) = ""

SA(4, 0) = "Bogmærke"
SA(4, 1) = "Bookmark"
SA(4, 4) = ""
SA(4, 3) = ""
SA(4, 2) = "Marcador"
SA(4, 5) = ""

SA(5, 0) = "Indtast navn på ligningen"
SA(5, 1) = "Type name of the equation"
SA(5, 4) = ""
SA(5, 3) = ""
SA(5, 2) = "Escriba el nombre de la ecuación"
SA(5, 5) = ""

SA(6, 0) = "Sæt ligningsnummer"
SA(6, 1) = "Set equation number"
SA(6, 4) = ""
SA(6, 3) = ""
SA(6, 2) = "Numere la ecuación"
SA(6, 5) = ""

SA(7, 0) = "Nummerede ligninger"
SA(7, 1) = "Numbered equations"
SA(7, 4) = ""
SA(7, 3) = ""
SA(7, 2) = "Ecuaciones numeradas"
SA(7, 5) = ""

SA(8, 0) = "Indsæt Ligningssektion nr."
SA(8, 1) = "Insert equation section no."
SA(8, 4) = ""
SA(8, 3) = ""
SA(8, 2) = "Inserte ecuación de la sección N°"
SA(8, 5) = ""

SA(9, 0) = "Opdater ligningsnumre"
SA(9, 1) = "Update equation numbers"
SA(9, 4) = ""
SA(9, 3) = ""
SA(9, 2) = "Actualice los números de las ecuaciones"
SA(9, 5) = ""

SA(10, 0) = "Eksisterende navne på ligninger"
SA(10, 1) = "Existing names of equations"
SA(10, 4) = ""
SA(10, 3) = ""
SA(10, 2) = "Nombres existentes de ecuaciones"

SA(11, 0) = "For at indsætte reference til ligningen senere vælges 'Indsæt reference til ligning' fra WordMat-menuen under 'Indsæt ligning'."
SA(11, 1) = "To insert a reference to the equation later, choose 'Insert reference to equation' from the WordMat-menu under 'Insert equation'"
SA(11, 4) = ""
SA(11, 3) = ""
SA(11, 2) = "Para insertar una referencia a la ecuación, más tarde, escoja ""Insertar referencia a la ecuación"" desde el menú de WordMat en la pestaña ""Inserte Ecuación"""
SA(11, 5) = ""

SA(12, 0) = "Navnet eksisterer allerede"
SA(12, 1) = "The name already exists"
SA(12, 4) = ""
SA(12, 3) = ""
SA(12, 2) = "El nombre ya existe"
SA(12, 5) = ""

SA(13, 0) = "Navnet må ikke indeholde mellemrum"
SA(13, 1) = "The name may not contain spaces"
SA(13, 4) = ""
SA(13, 3) = ""
SA(13, 2) = "El nombre no debe contener espacios"
SA(13, 5) = ""

SA(14, 0) = "Ligningsreference"
SA(14, 1) = "Equation reference"
SA(14, 4) = ""
SA(14, 3) = ""
SA(14, 2) = "Referencia de ecuación"
SA(14, 5) = ""

SA(15, 0) = "Placering"
SA(15, 1) = "Placement"
SA(15, 4) = ""
SA(15, 3) = ""
SA(15, 2) = "Ubicación"
SA(15, 5) = ""

SA(16, 0) = "Venstre side"
SA(16, 1) = "Left side"
SA(16, 4) = ""
SA(16, 3) = ""
SA(16, 2) = "Lado izquierdo"
SA(16, 5) = ""

SA(17, 0) = "Højre side"
SA(17, 1) = "Right side"
SA(17, 4) = ""
SA(17, 3) = ""
SA(17, 2) = "Lado derecho"
SA(17, 5) = ""

SA(18, 0) = "Nummerering"
SA(18, 1) = "Numbering"
SA(18, 4) = ""
SA(18, 3) = ""
SA(18, 2) = "Numeración"
SA(18, 5) = ""

SA(19, 0) = "Indsæt reference med 2 x Alt+m"
SA(19, 1) = "Insert reference with 2 x Alt+m"
SA(19, 4) = ""
SA(19, 3) = ""
SA(19, 2) = "Inserte referencia con 2 x Alt+m"
SA(19, 5) = ""

SA(20, 0) = AppNavn & " er lavet af Mikael Samsøe Sørensen, Nyborg Gymnasium, EDUAP 2009-2022." & VbCrLfMac & VbCrLfMac & "Programmet er gratis og stilles til rådighed under Gnu General public license." & VbCrLfMac & VbCrLfMac & "Med WordMat er følgende andre gratis programmer installeret: Maxima, GnuPlot, GeoGebra, Graph " & VbCrLfMac & VbCrLfMac & "www.eduap.com"
SA(20, 1) = AppNavn & " was made by Mikael Samsøe Sørensen, Nyborg Gymnasium, EDUAP 2009-2022." & VbCrLfMac & VbCrLfMac & "The program is free and open source under Gnu General public license." & VbCrLfMac & VbCrLfMac & "WordMat uses the following programs: Maxima, GnuPlot, GeoGebra, Graph " & VbCrLfMac & VbCrLfMac & "www.eduap.com"
SA(20, 4) = ""
SA(20, 3) = ""
SA(20, 2) = ""
SA(20, 5) = ""

SA(21, 0) = "Der er en ny version af " & AppNavn & " klar til download." & VbCrLfMac & VbCrLfMac & "Nyheder: " & VbCrLfMac
SA(21, 1) = "There is a new version of " & AppNavn & " ready for download." & VbCrLfMac & VbCrLfMac & "News: " & VbCrLfMac
SA(21, 4) = ""
SA(21, 3) = ""
SA(21, 2) = "Hay una nueva versión de " & AppNavn & " lista para descargar"
SA(21, 5) = ""

SA(22, 0) = "Vil du hente den nye version nu?"
SA(22, 1) = "Do you want to download the new version now?"
SA(22, 4) = ""
SA(22, 3) = ""
SA(22, 2) = "¿Desea descargar ahora la nueva versión?"
SA(22, 5) = ""

SA(23, 0) = "Ny version"
SA(23, 1) = "New version"
SA(23, 4) = ""
SA(23, 3) = ""
SA(23, 2) = "Nueva versión"
SA(23, 5) = ""

SA(24, 0) = "Du skal indtaste en punktmængde enten i en tabel eller i et matematikfelt som f.eks. {(1,2;2),(2;5,3)}" & VbCrLfMac & VbCrLfMac & "Indtast et antal punkter, så oprettes automatisk en tabel som kan udfyldes med punterne. Når du har udfyldt den aktiverer du den ønskede regression."
SA(24, 1) = "Type the number of points you want to do regression on. Then a table will be created, where you can put your points. When you want to do the regression just select the table and choose your desired regression from the menu."
SA(24, 4) = ""
SA(24, 3) = ""
SA(24, 2) = "Indique el número de puntos con el que desea ejecutar la regresión. Se creará una tabla donde puede anotar sus puntos. Cuando quiera ejecutar la regresión sólo seleccione la tabla y escoja el tipo de regresión que desee ejecutar desde el menú."
SA(24, 5) = ""

SA(25, 0) = "Det anbefales ikke at lave regression med mere end 200 punkter i en Word-tabel. Brug Excel hvis du vil lave regression på så mange punkter."
SA(25, 1) = "It is not recommened to do regression on more than 200 points in a Word-table. Use Excel if you want to do regression on that many points"
SA(25, 4) = ""
SA(25, 3) = ""
SA(25, 2) = "No es recomendable ejecutar una regresión con más de 200 puntos en una tabla de Word. Use Excel si desea ejecutar una regresión con muchos puntos."
SA(25, 5) = ""

SA(26, 0) = "Der skete en fejl. Der er formentlig en fejl i inputtallene."
SA(26, 1) = "An error occured. There is probably an error in the input numbers"
SA(26, 4) = ""
SA(26, 3) = ""
SA(26, 2) = "Ha ocurrido un error. Hay, probablemente, un error en los números de entrada"
SA(26, 5) = ""

SA(27, 0) = "Du skal lave en tabel med to kolonner. x-værdier i venstre kolonne og y-værdier i højrekolonne. Bogstaver ignoreres. Der udføres regression på de markerede tal. Hvis ingen er markerede udføres regression på alle tal i tabellen."
SA(27, 1) = "You must create a table with two columns. x-values in the left column and y-values in the right column. Letters are ignored. Regression is performed on the numbers selected. If none are selected, regression is performed on all numbers in the table"
SA(27, 4) = ""
SA(27, 3) = ""
SA(27, 2) = "Usted debe crear una tabla con dos columnas. Los valores de x en la columna izquierda y los valores de y en la columna derecha. No se incluyen letras. La regresión se ejecutará sobre los números seleccionados. Si no hay selección, la regresión se ejecutará sobre todos los números de la tabla"
SA(27, 5) = ""

SA(28, 0) = "Den ønskede regression kan ikke udføres på de punkter."
SA(28, 1) = "Regression cannot be performed on those points"
SA(28, 4) = ""
SA(28, 3) = ""
SA(28, 2) = "La regresión no puede ejecutarse en esos puntos"
SA(28, 5) = ""

SA(29, 0) = "Lineær regression udført vha. " & AppNavn & ": " & vbTab
SA(29, 1) = "Linear regression performed by " & AppNavn & ": " & vbTab
SA(29, 4) = ""
SA(29, 3) = ""
SA(29, 2) = "Regresión lineal ejecutada por" & AppNavn & ": " & vbTab
SA(29, 5) = ""

SA(30, 0) = "Eksponentiel regression udført vha. " & AppNavn & ": " & vbTab
SA(30, 1) = "Exponential regression performed by " & AppNavn & ": " & vbTab
SA(30, 4) = ""
SA(30, 3) = ""
SA(30, 2) = "Regresión exponencial ejecutada por" & AppNavn & ": " & vbTab
SA(30, 5) = ""

SA(31, 0) = "Potens regression udført vha. " & AppNavn & ": " & vbTab
SA(31, 1) = "Power regression performed by " & AppNavn & ": " & vbTab
SA(31, 4) = ""
SA(31, 3) = ""
SA(31, 2) = "Regresión de potencias ejecutada por" & AppNavn & ": " & vbTab
SA(31, 5) = ""

SA(32, 0) = "Der kan ikke udføres potensregression på de punkter. Alle skal være positive og må ikke være 0."
SA(32, 1) = "Power regression cannot be performed on those points. All numbers must be positive (and not zero)."
SA(32, 4) = ""
SA(32, 3) = ""
SA(32, 2) = "La Regresión de potencias no puede ejecutarse sobre esos números. Todos los números deben ser positivos(y distintos de cero)."
SA(32, 5) = ""

SA(33, 0) = "Brugerdefineret regression med funktionen "
SA(33, 1) = "User defined regression with the function "
SA(33, 4) = ""
SA(33, 3) = ""
SA(33, 2) = "Regresión y función definida por el usuario"
SA(33, 5) = ""

SA(34, 0) = " udført vha. " & AppNavn
SA(34, 1) = " performed by " & AppNavn
SA(34, 4) = ""
SA(34, 3) = ""
SA(34, 2) = " ejecutada por" & AppNavn
SA(34, 5) = ""

SA(35, 0) = "Angiv orden af den polynomiske regression (1-"
SA(35, 1) = "State the order of the polynomial regression (1-"
SA(35, 4) = ""
SA(35, 3) = ""
SA(35, 2) = "Estado de orden de la regresión polinomial(1-"
SA(35, 5) = ""

SA(36, 0) = "(Den maksimale orden bestemmes af antal punkter)"
SA(36, 1) = "(The max order is determined by the number of points)"
SA(36, 4) = ""
SA(36, 3) = ""
SA(36, 2) = "(El máximo orden se determina por el número de puntos)"
SA(36, 5) = ""

' polynomial order
SA(37, 0) = "orden"
SA(37, 1) = "order"
SA(37, 4) = ""
SA(37, 3) = ""
SA(37, 2) = "orden"
SA(37, 5) = ""

SA(38, 0) = "Polynomisk regression af grad "
SA(38, 1) = "Polynomial regression of degree "
SA(38, 4) = ""
SA(38, 3) = ""
SA(38, 2) = "Regresión polinomial de grado"
SA(38, 5) = ""

SA(39, 0) = "Brugerdefineret regression"
SA(39, 1) = "User defined regression"
SA(39, 4) = ""
SA(39, 3) = ""
SA(39, 2) = "Regresión definida por el usuario"
SA(39, 5) = ""

SA(40, 0) = "Funktion"
SA(40, 1) = "Function"
SA(40, 4) = ""
SA(40, 3) = ""
SA(40, 2) = "Función"
SA(40, 5) = ""

SA(41, 0) = "Antal decimaler"
SA(41, 1) = "Decimals"
SA(41, 4) = ""
SA(41, 3) = ""
SA(41, 2) = "Decimales"
SA(41, 5) = ""

SA(42, 0) = "Startgæt (Skrives på formen a=2 ; c=10)"
SA(42, 1) = "Start guess (ex.: a=2 , c=10)"
SA(42, 4) = ""
SA(42, 3) = ""
SA(42, 2) = "Comience inferencia (ej.: a=2 , c=10)"
SA(42, 5) = ""

SA(43, 0) = "Præcisionen er sat meget højt. Beregningstiden kan blive meget lang eller fejle. Prøv evt. igen med en lavere indstilling."
SA(43, 1) = "The precision is set very high. The calculation time can be very long or cause an error. If it doesn't work out, try a lower setting."
SA(43, 4) = ""
SA(43, 3) = ""
SA(43, 2) = "La precisión está ajustada muy alta. El tiempo de cálculo puede ser muy largo o cusar un error. Si no funciona intente un ajuste más bajo."
SA(43, 5) = ""

SA(44, 0) = "Præcisionen er sat ret højt. Beregningstiden kan blive lang eller fejle. Prøv evt. igen med en lavere indstilling."
SA(44, 1) = "The precision is set quite high. The calculation time can be quite long or cause an error. If it doesn't work out, try a lower setting."
SA(44, 4) = ""
SA(44, 3) = ""
SA(44, 2) = "La precisión está ajustada bastante alta. El tiempo de cálculo puede ser muy largo o cusar un error. Si no funciona intente un ajuste más bajo."
SA(44, 5) = ""

SA(45, 0) = "Præcisionen kan ikke være negativ. Vælg et positivt heltal"
SA(45, 1) = "The precision cant be negative. Choose a positive integer"
SA(45, 4) = ""
SA(45, 3) = ""
SA(45, 2) = "La precisión puede ser negativa. Escoja un entero positivo"
SA(45, 5) = ""

SA(46, 0) = "Præcisionen er ugyldig. Vælg et positivt heltal."
SA(46, 1) = "The precision is not valid. Choose a positive integer."
SA(46, 4) = ""
SA(46, 3) = ""
SA(46, 2) = "La precisión no es válida. Escoja un entero positivo"
SA(46, 5) = ""

SA(47, 0) = "Placer cursoren i et matematik-felt"
SA(47, 1) = "Place the cursor in an equation"
SA(47, 4) = ""
SA(47, 3) = ""
SA(47, 2) = "Sitúe el cursor en una ecuación"
SA(47, 5) = ""

SA(48, 0) = "Der kan ikke udføres beregninger på definitionsfelter. Du må udnytte din defintion nedenunder til en beregning."
SA(48, 1) = "It is not possible to do calculation on definitions"
SA(48, 4) = ""
SA(48, 3) = ""
SA(48, 2) = "No se pueden hacer cálculos en los campos de definición"
SA(48, 5) = ""

SA(49, 0) = "Udtrykket differentieres vha. " & AppNavn & ". "
SA(49, 1) = "The expression is differentiated by " & AppNavn & ". "
SA(49, 4) = ""
SA(49, 3) = ""
SA(49, 2) = "La expresión se ha diferenciado mediante " & AppNavn & ". "
SA(49, 5) = ""

SA(50, 0) = "Udtrykket kunne ikke differentieres"
SA(50, 1) = "The expression could not be differetiated"
SA(50, 4) = ""
SA(50, 3) = ""
SA(50, 2) = "La expresión no se puede diferenciar"
SA(50, 5) = ""

SA(51, 0) = "Indtast variabel"
SA(51, 1) = "State variable"
SA(51, 4) = ""
SA(51, 3) = ""
SA(51, 2) = "Variable de estado"
SA(51, 5) = ""

SA(52, 0) = "Der findes en stamfunktion vha. " & AppNavn & "'s 'Stamfunktion' funktion"
SA(52, 1) = "An antiderivative is found by " & AppNavn & ""
SA(52, 4) = ""
SA(52, 3) = ""
SA(52, 2) = "Se ha encontrado una derivada mediante " & AppNavn & ""
SA(52, 5) = ""

SA(53, 0) = "Udtrykket kunne ikke integreres"
SA(53, 1) = "The expression could not be integrated"
SA(53, 4) = ""
SA(53, 3) = ""
SA(53, 2) = "La expresión no se puede integrar"
SA(53, 5) = ""

SA(54, 0) = "Der skete en fejl. WordMat er formentlig ikke installeret korrekt. Prøv at geninstallere, evt. med Antivirus deaktiveret"
SA(54, 1) = "An error occured. WordMat is probably not installed correctly. Try reinstalling, possibly with antivirus deactivated."
SA(54, 4) = ""
SA(54, 3) = ""
SA(54, 2) = "Ha ocurrido un error. Es probable que WordMat no se haya instalado correctamente. Intente re-instalarlo, preferiblemente con el antivirus desactivado."
SA(54, 5) = ""

SA(55, 0) = "Der skete en fejl i forsøget på at starte Maxima"
SA(55, 1) = "An error occured trying to start Maxima"
SA(55, 4) = ""
SA(55, 3) = ""
SA(55, 2) = "Ha ocurrido un error intentando arrancar Maxima"
SA(55, 5) = ""

SA(56, 0) = "Uligheden løses for "
SA(56, 1) = "The inequality is solved for "
SA(56, 4) = ""
SA(56, 3) = ""
SA(56, 2) = "La inegualdad se resuelve para "
SA(56, 5) = ""

SA(57, 0) = " vha. " & AppNavn & ". "
SA(57, 1) = " by " & AppNavn & ". "
SA(57, 4) = ""
SA(57, 3) = ""
SA(57, 2) = " mediante " & AppNavn & ". "
SA(57, 5) = ""

SA(58, 0) = "Uligheden har alle tal som løsning."
SA(58, 1) = "All numbers solve this inequality"
SA(58, 4) = ""
SA(58, 3) = ""
SA(58, 2) = "Todos los números satisfacen esta inegualdad"
SA(58, 5) = ""

SA(59, 0) = "Uligheden havde ingen løsninger indenfor " & VBA.ChrW(8450) & " for variabel "
SA(59, 1) = "The inequality had no solutions in the domain " & VBA.ChrW(8450) & " for variable "
SA(59, 4) = ""
SA(59, 3) = ""
SA(59, 2) = "La inegualdad no tiene soluciones en el dominio " & VBA.ChrW(8450) & " para la variable "
SA(59, 5) = ""

SA(60, 0) = "Uligheden havde ingen løsninger indenfor " & VBA.ChrW(8477) & " for variabel "
SA(60, 1) = "The inequality had no solutions in the domain " & VBA.ChrW(8477) & " for variable "
SA(60, 4) = ""
SA(60, 3) = ""
SA(60, 2) = "La inegualdad no tiene soluciones en el dominio " & VBA.ChrW(8450) & " para la variable "
SA(60, 5) = ""

SA(61, 0) = "med følgende antagelser/definitioner:  "
SA(61, 1) = "With the following assumptions/definitions:  "
SA(61, 4) = ""
SA(61, 3) = ""
SA(61, 2) = "Con los siguientes supuestos/definiciones:"
SA(61, 5) = ""

' Dont change these
SA(62, 0) = "Definer"
SA(62, 1) = "Define"
SA(62, 4) = "Define"
SA(62, 3) = "Define"
SA(62, 2) = "Define"
SA(62, 5) = "Define"

SA(63, 0) = "Vil du oprette tastaturgenveje?" & VbCrLfMac & "Som standard er de oprettet, men kan blive overskrevet. Du kan genoprette dem nu."
SA(63, 1) = "Do you want to create keyboard shortcuts?" & VbCrLfMac & "They are created by default, but they can be overwritten by other programs. You can restore them now."
SA(63, 4) = ""
SA(63, 3) = ""
SA(63, 2) = "øDesea crear un atajo de teclas?" & VbCrLfMac & "Este se crea por defecto, pero puede sobreescribirse por otros programas. Usted debe restablecerlo ahora"
SA(63, 5) = ""

SA(64, 0) = "Tasturgenveje"
SA(64, 1) = "Keyboard shortcuts"
SA(64, 4) = ""
SA(64, 3) = ""
SA(64, 2) = "Atajos de Teclado"
SA(64, 5) = ""

SA(65, 0) = "WordMat tastaturgenveje"
SA(65, 1) = "WordMat keyboard shortcuts"
SA(65, 4) = ""
SA(65, 3) = ""
SA(65, 2) = "Atajos de Teclado de WordMat"
SA(65, 5) = ""

SA(66, 0) = "Word matematikfelt-genveje"
SA(66, 1) = "Word equation shortcuts"
SA(66, 4) = ""
SA(66, 3) = ""
SA(66, 2) = "Atajos de ecuaciones de Word"
SA(66, 5) = ""

SA(67, 0) = "Man kan lave sine egne genveje til matematikfelter i Word eller ændre de eksisterende. Når man står i et matematikfelt finder man Ligningsfunktion/Design fanen. Under funktioner trykker man på den lille pil i nederste højre hjørne. Dernæst vælges 'matematisk autokorrektur'"
SA(67, 1) = "You can create your own typing short cuts in Word or change the existing shortcuts. Place the cursor in an equation. Open Design menu. Locate group called functions. Click the little arrow in lower right corner. Choose 'Math autoentries'"
SA(67, 4) = ""
SA(67, 3) = ""
SA(67, 2) = "Puede crear su propio atajo en Word o cambiar los existentes. Sitúe el cursor en una ecuación. Abra el menú Diseño. Ubique el grupo llamado Funciones. Pulse la flecha pequeña en la esquina derecha. Escoja Autoentradas Matemáticas"
SA(67, 5) = ""

SA(68, 0) = "Nyt Matematikfelt" & vbTab & "Alt+M" & VbCrLfMac & "Forrige resultat(er)" & vbTab & "Alt+R" & vbTab & VbCrLfMac & "Prik-gangetegn" & vbTab & "Alt+G" & vbTab & VbCrLfMac & "Beregn" & vbTab & vbTab & "Alt+b   eller   AltGr+Enter" & vbTab & VbCrLfMac & "Løs ligning" & vbTab & "Alt+L" & VbCrLfMac & "Indstillinger" & vbTab & "Alt + j" & VbCrLfMac & "Enheder til/fra" & vbTab & "Alt+E" & VbCrLfMac & "Konverter LaTex" & vbTab & "Alt+T" & VbCrLfMac & "Definer:" & vbTab & vbTab & "Alt+D" & VbCrLfMac & "Slet definitioner:" & vbTab & "Alt+S" & VbCrLfMac & "Plot Graf" & vbTab & vbTab & "Alt+P" & VbCrLfMac & "Omskriv" & vbTab & vbTab & "Alt+O" & VbCrLfMac & "Num/eksakt/auto" & vbTab & "Alt+N" & VbCrLfMac & "Latex PDF" & vbTab & "Alt+Q"
SA(68, 1) = "New Math equation" & vbTab & "Alt+M" & VbCrLfMac & "Previous result(s)" & vbTab & "Alt+R" & VbCrLfMac & "Multiplication sign" & vbTab & "Alt+G" & VbCrLfMac & "Calculate" & vbTab & vbTab & "Alt+B   or   AltGr+Enter" & vbTab & VbCrLfMac & "Solve equation" & vbTab & "Alt+E" & vbTab & VbCrLfMac & "Settings (Options)" & vbTab & "Alt+j" & vbTab & VbCrLfMac & "Units on/off" & vbTab & "Alt+U" & VbCrLfMac & "Convert LaTex" & vbTab & "Alt+T" & VbCrLfMac & "Define:" & vbTab & vbTab & "Alt+D" & VbCrLfMac & "Delete definitions:" & vbTab & "Alt+F       (clear)" & VbCrLfMac & "Plot Graph" & vbTab & "Alt+P" & VbCrLfMac & "Simplify" & vbTab & vbTab & "Alt+S" & VbCrLfMac & "Num/exact/auto" & vbTab & "Alt+N" & VbCrLfMac & "Latex PDF" & vbTab & "Alt+Q"
SA(68, 2) = "Nueva ecuación" & vbTab & "Alt+M" & VbCrLfMac & "Resultado(s) anterior(es)" & vbTab & "Alt+R" & VbCrLfMac & "Signo de multiplicación " & vbTab & "Alt+G" & VbCrLfMac & "Calcular" & vbTab & vbTab & "Alt+B   or   AltGr+Enter" & vbTab & VbCrLfMac & "Resolvee ecuación" & vbTab & "Alt+E" & vbTab & VbCrLfMac & "Ajustes" & vbTab & vbTab & "Alt + j" & vbTab & VbCrLfMac & "Unidades on/off" & vbTab & "Alt+U" & VbCrLfMac & "Convertidor a LaTex" & vbTab & "Alt+T" & VbCrLfMac & "Definir:" & vbTab & vbTab & "Alt+d" & VbCrLfMac & "Eliminar definiciones:" & vbTab & "Alt+F" & VbCrLfMac & "Plot Graph" & vbTab & "Alt+P" & VbCrLfMac & "Simplificar" & vbTab & vbTab & "Alt+S" & VbCrLfMac & "Num/exacto/auto" & vbTab & "Alt+N" & VbCrLfMac & "Latex PDF" & vbTab & "Alt+Q"

' Dont change this
SA(69, 0) = "Slet definitioner"
SA(69, 1) = "Delete definitions"
SA(69, 4) = "Delete definitions"
SA(69, 3) = "Delete definitions"
SA(69, 2) = "Delete definitions"
SA(69, 5) = "Delete definitions"

SA(70, 0) = "Konstanter"
SA(70, 1) = "Constants"
SA(70, 4) = ""
SA(70, 3) = ""
SA(70, 2) = "Constantes"
SA(70, 5) = ""

SA(71, 0) = "Find værdier på nettet"
SA(71, 1) = "Find values online"
SA(71, 4) = ""
SA(71, 3) = ""
SA(71, 2) = "Encontrar valores en línea"
SA(71, 5) = ""

SA(72, 0) = "Konverter ligninger til LaTex. (Genvejen alt+T konverterer en enkelt ligning)"
SA(72, 1) = "Convert equations to LaTex. (Shortcut alt+T converts a single equation)"
SA(72, 4) = ""
SA(72, 3) = ""
SA(72, 2) = "Convertir ecuaciones a LaTex. (El atajo alt+T convierte una ecuación simple)"
SA(72, 5) = ""

SA(73, 0) = "Kopier til clipboard"
SA(73, 1) = "Copy to clipboard"
SA(73, 4) = ""
SA(73, 3) = ""
SA(73, 2) = "Copiar al portapapeles"
SA(73, 5) = ""

SA(74, 0) = "Visningsform"
SA(74, 1) = "Display style"
SA(74, 4) = ""
SA(74, 3) = ""
SA(74, 2) = "Estilo pantalla"
SA(74, 5) = ""

SA(75, 0) = "Genkend enheder og skriv dem ikke i kursiv"
SA(75, 1) = "Dont write units in italic"
SA(75, 4) = ""
SA(75, 3) = ""
SA(75, 2) = "No escriba las unidades en itálica"
SA(75, 5) = ""

SA(76, 0) = "Brug Maxima til LaTex konvertering"
SA(76, 1) = "Use Maxima to convert LaTex"
SA(76, 4) = ""
SA(76, 3) = ""
SA(76, 2) = "Usar Maxima para convertir LaTex"
SA(76, 5) = ""

SA(77, 0) = "Konverter alle"
SA(77, 1) = "Convert all"
SA(77, 4) = ""
SA(77, 3) = ""
SA(77, 2) = "Convertir todo"
SA(77, 5) = ""

SA(78, 0) = "Konverter og gå til næste"
SA(78, 1) = "Convert and go to next"
SA(78, 4) = ""
SA(78, 3) = ""
SA(78, 2) = "Convertir y continuar en la siguiente"
SA(78, 5) = ""

SA(79, 0) = "Før"
SA(79, 1) = "Before"
SA(79, 4) = ""
SA(79, 3) = ""
SA(79, 2) = "Antes"
SA(79, 5) = ""

SA(80, 0) = "Efter"
SA(80, 1) = "After"
SA(80, 4) = ""
SA(80, 3) = ""
SA(80, 2) = "Después"
SA(80, 5) = ""

SA(81, 0) = "Ingen"
SA(81, 1) = "None"
SA(81, 4) = ""
SA(81, 3) = ""
SA(81, 2) = "Ninguno"
SA(81, 5) = ""

SA(82, 0) = "Skriv selv"
SA(82, 1) = "Custom"
SA(82, 4) = ""
SA(82, 3) = ""
SA(82, 2) = "Escriba usted mismo"
SA(82, 5) = ""

SA(83, 0) = "Omslutningstegn"
SA(83, 1) = "Embed code"
SA(83, 4) = ""
SA(83, 3) = ""
SA(83, 2) = "Código incrustado"
SA(83, 5) = ""

SA(84, 0) = "Der er ikke flere matematikfelter i resten af dette dokument."
SA(84, 1) = "There are no equations in the rest of this document to convert."
SA(84, 4) = ""
SA(84, 3) = ""
SA(84, 2) = "No hay más ecuaciones para convertir en el resto del documento"
SA(84, 5) = ""

SA(85, 0) = "Løs koblede differentialligning(er) numerisk"
SA(85, 1) = "Solve system of differential equations numerically"
SA(85, 4) = ""
SA(85, 3) = ""
SA(85, 2) = "Resolver el sistema de ecuaciones diferenciales numéricamente"
SA(85, 5) = ""

SA(86, 0) = "Differentialligninger"
SA(86, 1) = "Differential equations"
SA(86, 4) = ""
SA(86, 3) = ""
SA(86, 2) = "Ecuaciones diferenciales"
SA(86, 5) = ""

SA(87, 0) = "Startværdi"
SA(87, 1) = "Start value"
SA(87, 4) = ""
SA(87, 3) = ""
SA(87, 2) = "Valor de inicio"
SA(87, 5) = ""

SA(88, 0) = "Skridtlængde"
SA(88, 1) = "Step size"
SA(88, 4) = ""
SA(88, 3) = ""
SA(88, 2) = "Tamaño del paso"
SA(88, 5) = ""

SA(89, 0) = "Forbind punkter"
SA(89, 1) = "Connect dots"
SA(89, 4) = ""
SA(89, 3) = ""
SA(89, 2) = "Conectar Puntos"
SA(89, 5) = ""

SA(90, 0) = "Vis forskrifter"
SA(90, 1) = "Show equations"
SA(90, 4) = ""
SA(90, 3) = ""
SA(90, 2) = "Mostrar ecuaciones"
SA(90, 5) = ""

SA(91, 0) = "Indsæt liste"
SA(91, 1) = "Insert list"
SA(91, 4) = ""
SA(91, 3) = ""
SA(91, 2) = "Insertar lista"
SA(91, 5) = ""

SA(92, 0) = "Indsæt tabel"
SA(92, 1) = "Insert table"
SA(92, 4) = ""
SA(92, 3) = ""
SA(92, 2) = "Insertar tabla"
SA(92, 5) = ""

SA(93, 0) = "Indsæt graf"
SA(93, 1) = "Insert graph"
SA(93, 4) = ""
SA(93, 3) = ""
SA(93, 2) = "Insertar gráfico"
SA(93, 5) = ""

SA(94, 0) = "Formentlig Udenfor definitionsområde! " & VbCrLfMac & "Juster xmin/xmax"
SA(94, 1) = "Probably outside domain! " & VbCrLfMac & "adjust xmin/xmax"
SA(94, 4) = ""
SA(94, 3) = ""
SA(94, 2) = "Probablemente fuera del dominio!" & VbCrLfMac & "ajuste xmin/xmax"
SA(94, 5) = ""

SA(95, 0) = "Der skete en fejl. Se på tallene ved xmin og xmax"
SA(95, 1) = "An error occured. Check the numbers by xmin and xmax"
SA(95, 4) = ""
SA(95, 3) = ""
SA(95, 2) = "Ha ocurrido un error. Revise las números para xmin y xmax"
SA(95, 5) = ""

SA(96, 0) = "Fejl, regnearket med statistikdiagrammer kan ikke findes på stien: "
SA(96, 1) = "Error, cannot find the Excel file with statistic on the path: "
SA(96, 4) = ""
SA(96, 3) = ""
SA(96, 2) = "Error, no se encuentra el archivo Excel con la estadística en la ruta: "
SA(96, 5) = ""

SA(97, 0) = "Der skete en fejl i forsøget på at indsætte Graph. Der er formentlig en syntaksfejl et sted i en definition eller lign"
SA(97, 1) = "An error occured trying to insert Graph. You might have a syntax error in a definition"
SA(97, 4) = ""
SA(97, 3) = ""
SA(97, 2) = "Ha ocurrido un error intentando insertar Graph. Podría deberse a un error de sintaxis en una definición"
SA(97, 5) = ""

SA(98, 0) = "Der skete en fejl i forsøget på at indsætte Excelark. Det kan skyldes at Excel er gået ned tidligere.Luk evt. åbne Excelark og åben Excel inden du forsøger via WordMat igen."
SA(98, 1) = "An error occured trying to insert an Excel sheet. This can be caused by a previous crash of Excel. Close all open Excel sheets and then reopen Excel before trying with WordMat again."
SA(98, 4) = ""
SA(98, 3) = ""
SA(98, 2) = "Ha ocurrido un error intentando insertar una hoja Excel. Puede deberse a una falla previa de Excel. Cierre todas las hojas Excel y luego abra Excel antes de intentar con Wordmat nuevamente."
SA(98, 5) = ""

SA(99, 0) = "Der skete en fejl i forsøget på at oversætte din kommando til Maxima syntaks."
SA(99, 1) = "An error occured trying to convert your command to Maxima syntax."
SA(99, 4) = ""
SA(99, 3) = ""
SA(99, 2) = "Ha ocurrido un error intentando convertir  "
SA(99, 5) = ""

SA(100, 0) = "Der skete en fejl i forsøget på at oversætte Maxima output til Word synktaks."
SA(100, 1) = "An error occured trying to convert Maxima output to Word syntax."
SA(100, 4) = ""
SA(100, 3) = ""
SA(100, 2) = "Ha ocurrido un error al tratar de traducir la salida de Maxima para sintaxis de comandos  de Word."
SA(100, 5) = ""

SA(101, 0) = " Løsningen er usikker. Prøv evt. Numerisk løsning"
SA(101, 1) = " The solution is unsure. You might want to try using numerical methods as well"
SA(101, 4) = ""
SA(101, 3) = ""
SA(101, 2) = "a solución es incierta. Pruebe cualquier solución numérica."
SA(101, 5) = ""

SA(102, 0) = "Da det er en trigonometrisk ligning er der uendeligt mange løsninger. Kun en af disse løsninger er angivet."
SA(102, 1) = "Since this is a trigonometric equation there are infinitely many solutions. Only one of these are stated"
SA(102, 4) = ""
SA(102, 3) = ""
SA(102, 2) = "Dado que es una ecuación trigonométrica existen infinitas soluciones. Sólo se muestra una de ellas."
SA(102, 5) = ""

SA(103, 0) = " Alle løsninger er ikke fundet."
SA(103, 1) = " All solutions have not been found."
SA(103, 4) = ""
SA(103, 3) = ""
SA(103, 2) = "No se encontraron todas las soluciones"
SA(103, 5) = ""

SA(104, 0) = "Fejl: Division med nul."
SA(104, 1) = "Error: Division by zero."
SA(104, 4) = ""
SA(104, 3) = ""
SA(104, 2) = "Error: División por cero."
SA(104, 5) = ""

SA(105, 0) = "Syntaksfejl. Der er en fejl i notationen."
SA(105, 1) = "Syntax error."
SA(105, 4) = ""
SA(105, 3) = ""
SA(105, 2) = "Error de sintaxis"
SA(105, 5) = ""

SA(106, 0) = "Syntaksfejl. Der er brugt et forkert tegn. Måske manglende gangetegn eller komma istedet for ; som listeseparator."
SA(106, 1) = "Syntax error. An incorrect symbol was used. May be a missing multiplication sign or a , . issue"
SA(106, 4) = ""
SA(106, 3) = ""
SA(106, 2) = "Error de sintaxis. Se usó un símbolo incorrecto. Puede ser la falta de un signo de multiplicación o un punto o una coma en vez del punto y conma como separador de listas"
SA(106, 5) = ""

SA(107, 0) = "Advarsel: Du har bedt om at løse et ligningssystem, men angivet færre variable end ligninger. Kun de første ligninger er løst. Kontroller selv om de(n) sidste ligning(er) passer."
SA(107, 1) = "Warning: You have asked to solve a system of equation, but stated fewer variables than equations. Only the first equations are solved. Check manually the last equation(s)."
SA(107, 4) = ""
SA(107, 3) = ""
SA(107, 2) = "Advertencia: Se le consultará para resolver el sistema de ecuaciones, pero aparecen menos variables que ecuaciones. Sólo se resolverá la primera ecuación. Revise manualmente la(s) última(s) ecuación(es)."
SA(107, 5) = ""

SA(108, 0) = "Der kom følgende fejlmelding fra Maxima:"
SA(108, 1) = "Maxima reported the following error:"
SA(108, 4) = ""
SA(108, 3) = ""
SA(108, 2) = "Maxima indica el siguiente error"
SA(108, 5) = ""

SA(109, 0) = "Maxima fejlmelding"
SA(109, 1) = "Maxima error"
SA(109, 4) = ""
SA(109, 3) = ""
SA(109, 2) = "Error de maxima"
SA(109, 5) = ""

SA(110, 0) = "WordMat kunne ikke finde 'Program files' mappen"
SA(110, 1) = "WordMat was unable to locate the 'Program files' folder"
SA(110, 4) = ""
SA(110, 3) = ""
SA(110, 2) = "WordMat no pudo localizar la carpeta 'Archivos de Programa'"
SA(110, 5) = ""

SA(111, 0) = "Der skete en fejl i forsøget på at åbne: "
SA(111, 1) = "An error occured trying to open: "
SA(111, 4) = ""
SA(111, 3) = ""
SA(111, 2) = "Ocurrió un error intentando abrir: "
SA(111, 5) = ""

SA(112, 0) = "Der kunne ikke oprettes forbindelse til serveren"
SA(112, 1) = "Unable to connect to server"
SA(112, 4) = ""
SA(112, 3) = ""
SA(112, 2) = "No se puede conectar con el servidor"
SA(112, 5) = ""

SA(113, 0) = "På dette sted i dokumentet er der defineret følgende:"
SA(113, 1) = "At this place in the document the following is defined"
SA(113, 4) = ""
SA(113, 3) = ""
SA(113, 2) = "Aquí, en ésta parte del documento, está definido lo siguiente:"
SA(113, 5) = ""

SA(114, 0) = "Der er ingen definitioner på dette sted i dokumentet."
SA(114, 1) = "There aren't any definitions at this place in the document."
SA(114, 4) = ""
SA(114, 3) = ""
SA(114, 2) = "No hay ninguna definición en este lugar del documento"
SA(114, 5) = ""

SA(115, 0) = "Lignings definitioner indledes med en 'definer ligning:' kommando Indtast definitionen på den nye ligning" & VbCrLfMac & VbCrLfMac & "Definitionen kan benyttes i resten af dokumentet, men ikke før." & VbCrLfMac & " Hvis der indsættes en 'slet def:' kommando længere nede i dokumentet kan den ikke benyttes derefter." & VbCrLfMac & VbCrLfMac & "Der kan defineres flere ligninger ved at adskille definitionerne med " & ListSeparator & VbCrLfMac & "f.eks. lig1: c=3*d " & ListSeparator & " lig2: a=2*b"
SA(115, 1) = "Equation definitions are defined by 'definer ligning:' kommando Indtast definitionen på den nye ligning" & VbCrLfMac & VbCrLfMac & "Definitionen kan benyttes i resten af dokumentet, men ikke før." & VbCrLfMac & " Hvis der indsættes en 'slet def:' kommando længere nede i dokumentet kan den ikke benyttes derefter." & VbCrLfMac & VbCrLfMac & "Der kan defineres flere ligninger ved at adskille definitionerne med " & ListSeparator & VbCrLfMac & "f.eks. lig1: c=3*d " & ListSeparator & " lig2: a=2*b"
SA(115, 4) = ""
SA(115, 3) = ""
SA(115, 2) = "Las definiciones de la ecuación se realizan mediante un 'Definir la ecuación :' comandos para definir la nueva ecuación" & VbCrLfMac & VbCrLfMac & "La definición se puede utilizar en el resto del documento, pero no antes." & VbCrLfMac & " Si inserta un comando 'borrar def: ', más abajo, en el documento, no se puede utilizar, ésta definición, a partir de éste punto." & VbCrLfMac & VbCrLfMac & "Puede definir varias ecuaciones, al mismo tiempo, mediante la separación de las definiciones con" & ListSeparator & VbCrLfMac & "por.ej. ec1: c=3*d " & ListSeparator & " ec2: a=2*b"
SA(115, 5) = ""

SA(116, 0) = "Man kan ikke definere en afledet funktion." & VbCrLfMac & "Fjern følgende definition: " & VbCrLfMac
SA(116, 1) = "You can't define a derivative" & VbCrLfMac & "Remove the definition: " & VbCrLfMac
SA(116, 4) = ""
SA(116, 3) = ""
SA(116, 2) = "No puede definir una derivada" & VbCrLfMac & "Remover la definición: " & VbCrLfMac
SA(116, 5) = ""

SA(117, 0) = "Der er en fejl i definitionen: "
SA(117, 1) = "There is an error in the definition: "
SA(117, 4) = ""
SA(117, 3) = ""
SA(117, 2) = "Hay un error en la definición: "
SA(117, 5) = ""

SA(118, 0) = "Den ignoreres. Men ret det så du ikke får denne fejl ved hver beregning."
SA(118, 1) = "It is ignored, but correct the input error to avoid this output error"
SA(118, 4) = ""
SA(118, 3) = ""
SA(118, 2) = "Se ignora, pero debe corregir el error de entrada para evitar el error de salida"
SA(118, 5) = ""

SA(119, 0) = " ...resten fjernet."
SA(119, 1) = " ...rest removed."
SA(119, 4) = ""
SA(119, 3) = ""
SA(119, 2) = "...resto eliminado."
SA(119, 5) = ""

SA(120, 0) = "Indtast definitionen på den nye variabel" & VbCrLfMac & VbCrLfMac & "Definitionen kan benyttes i resten af dokumentet, men ikke før." & VbCrLfMac & "Hvis der indsættes en 'slet def:' kommando længere nede i dokumentet kan den ikke benyttes derefter." & VbCrLfMac & VbCrLfMac & "Du kan defineres flere variable på en gang ved at adskille definitionerne med " & ListSeparator & VbCrLfMac & " f.eks. a:1 " & ListSeparator & " b:2"
SA(120, 1) = "State the definition of the new variable" & VbCrLfMac & VbCrLfMac & "The definition can be used in the rest of the document, but not before." & VbCrLfMac & "If you insert a  'delete def:' command further down in the document it will not be known from that point." & VbCrLfMac & VbCrLfMac & "You can define more variables at a time by separating them with " & ListSeparator & VbCrLfMac & " ex.:  a:1 " & ListSeparator & " b:2"
SA(120, 4) = ""
SA(120, 3) = ""
SA(120, 2) = "Establezca la definición de la nueva variable" & VbCrLfMac & VbCrLfMac & "La definición se puede utilizar en el resto del documento, pero no antes." & VbCrLfMac & "Si usted inserta un comando 'borrar def:'  más abajo, en el documento, no se puede utilizar, ésta definición, a partir de éste punto." & VbCrLfMac & VbCrLfMac & "Puede definir varias variables, al mismo tiempo, mediante la separación de las definiciones con" & ListSeparator & VbCrLfMac & " ej.:  a:1 " & ListSeparator & " b:2"
SA(120, 5) = ""

SA(121, 0) = "Ny variabel"
SA(121, 1) = "New variable"
SA(121, 4) = ""
SA(121, 3) = ""
SA(121, 2) = "Nueva variable"
SA(121, 5) = ""

SA(122, 0) = "Indtast definitionen på den nye funktion" & VbCrLfMac & VbCrLfMac & "Definitionen kan benyttes i resten af dokumentet, men ikke før." & VbCrLfMac & " Hvis der indsættes en 'slet def:' kommando længere nede i dokumentet kan den ikke benyttes derefter." & VbCrLfMac & VbCrLfMac & "Der kan defineres flere funktioner ved at adskille definitionerne med " & ListSeparator & VbCrLfMac & "f.eks. f(x)=x " & ListSeparator & " g(x)=2x+1"
SA(122, 1) = "State the definition of the new function" & VbCrLfMac & VbCrLfMac & "The definition can be used in the rest of the document, but not before." & VbCrLfMac & "If you insert a  'delete def:' command further down in the document it will not be known from that point." & VbCrLfMac & VbCrLfMac & "You can define more variables at a time by separating them with " & ListSeparator & VbCrLfMac & "ex.: f(x)=x " & ListSeparator & " g(x)=2x+1"
SA(122, 4) = ""
SA(122, 3) = ""
SA(122, 2) = "Establezca la definición de la nueva función" & VbCrLfMac & VbCrLfMac & "La definición se puede utilizar en el resto del documento, pero no antes." & VbCrLfMac & "Si usted inserta un comando 'borrar def:'  más abajo, en el documento, no se puede utilizar, ésta definición, a partir de éste punto." & VbCrLfMac & VbCrLfMac & "Puede definir varias funciones, al mismo tiempo, mediante la separación de las definiciones con" & ListSeparator & VbCrLfMac & "ex.: f(x)=x " & ListSeparator & " g(x)=2x+1"
SA(122, 5) = ""

SA(123, 0) = "Ny funktion"
SA(123, 1) = "New function"
SA(123, 4) = ""
SA(123, 3) = ""
SA(123, 2) = "Nueva función"
SA(123, 5) = ""

SA(124, 0) = "Ny ligning"
SA(124, 1) = "New equation"
SA(124, 4) = ""
SA(124, 3) = ""
SA(124, 2) = "Nueva ecuación"
SA(124, 5) = ""

SA(125, 0) = "definer ligning"
SA(125, 1) = "define equation"
SA(125, 4) = ""
SA(125, 3) = ""
SA(125, 2) = "Definir ecuación"
SA(125, 5) = ""

' dont change
SA(126, 0) = "Definer"
SA(126, 1) = "Define"
SA(126, 4) = "Define"
SA(126, 3) = "Define"
SA(126, 2) = "Define"
SA(126, 5) = "Define"

SA(127, 0) = "Udtrykket er meget langt. Vil du indsætte det i Word?" & VbCrLfMac & "Word kan blive meget langsom. Prøv evt. at løse numerisk istedet." & VbCrLfMac & "Du kan se outputtet herunder:"
SA(127, 1) = "The expression is very long. Do you want to insert it into word?" & VbCrLfMac & "This can affect the speed of Word. Long equations can bring Word to a halt on slow computers. Consider solving using numerical methods" & VbCrLfMac & "You can see the output below:"
SA(127, 4) = ""
SA(127, 3) = ""
SA(127, 2) = "La expresión es muy larga.øPrefiere insertarla en Word?" & VbCrLfMac & "Esto puede afectar la velocidad de Word. Ecuaciones largas pueden llevar a Word a detenerse en computadores lentos. Considere la resolución mdiante métodos numéricos" & VbCrLfMac & "Se puede ver el resultado a continuación: "
SA(127, 5) = ""

SA(128, 0) = "Udtrykket kunne ikke beregnes."
SA(128, 1) = "The expression could not be calculated."
SA(128, 4) = ""
SA(128, 3) = ""
SA(128, 2) = "La expresión no puede calcularse"
SA(128, 5) = ""

SA(129, 0) = "Ligningen havde ingen løsninger indenfor " & VBA.ChrW(8450) & " for variabel "
SA(129, 1) = "The equation had no solutions in the domain " & VBA.ChrW(8450) & " for variable "
SA(129, 4) = ""
SA(129, 3) = ""
SA(129, 2) = "La ecuación no tiene soluciones en el dominio " & VBA.ChrW(8450) & " para la variable "
SA(129, 5) = ""

SA(130, 0) = "Ligningen havde ingen løsninger indenfor " & VBA.ChrW(8477) & " for variabel "
SA(130, 1) = "The equation had no solutions in the domain " & VBA.ChrW(8477) & " for variable "
SA(130, 4) = ""
SA(130, 3) = ""
SA(130, 2) = "La ecuación no tiene soluciones en el dominio " & VBA.ChrW(8477) & " para la variable "
SA(130, 5) = ""

SA(131, 0) = "WordMat kunne desværre ikke løse ligningen."
SA(131, 1) = "WordMat couldn't solve the equation"
SA(131, 4) = ""
SA(131, 3) = ""
SA(131, 2) = "WordMat no puede resolver la ecuación"
SA(131, 5) = ""

SA(132, 0) = "Alle tal er løsning til ligningen"
SA(132, 1) = "All numbers are solutions to the equation"
SA(132, 4) = ""
SA(132, 3) = ""
SA(132, 2) = "Todos los números son solución de la ecuación"
SA(132, 5) = ""

SA(133, 0) = "Du har bedt om at løse ligningen for en variabel der allerede er defineret. Indsæt en 'slet def:' kommando før ligningen."
SA(133, 1) = "You have asked to solve the equation for a variable which is already defined. Insert a 'delete def:' command before the equation."
SA(133, 4) = ""
SA(133, 3) = ""
SA(133, 2) = "Se le consulta resolver la ecuación para una variable que ya está definida. Inserte un comando 'borrar def:' antes de la ecuación."
SA(133, 5) = ""

SA(134, 0) = "Ligningssystemet løses for "
SA(134, 1) = "The system of equations is solved for "
SA(134, 4) = ""
SA(134, 3) = ""
SA(134, 2) = "El sistema de ecuaciones está resuelto para "
SA(134, 5) = ""

SA(135, 0) = " vha. " & AppNavn & "'s 'Løs Ligninger' funktion, "
SA(135, 1) = " by " & AppNavn & "'s 'solve equation' function, "
SA(135, 4) = ""
SA(135, 3) = ""
SA(135, 2) = "Mediante" & AppNavn & "'s 'resolver ecuación' función, "
SA(135, 5) = ""

SA(136, 0) = "Maxima kunne ikke løse ligningssystemet analytisk, men har angivet et omskrevet resultat der måske kan arbejdes videre med."
SA(136, 1) = "Maxima could not solve the system of equations analytically, but has stated a rearranged result."
SA(136, 4) = ""
SA(136, 3) = ""
SA(136, 2) = "Maxima no podía resolver el sistema de ecuaciones analíticamente, pero se ha creado un resultado reescrito que podría funcionar."
SA(136, 5) = ""

SA(137, 0) = "Ligningssystemet havde ingen løsninger indenfor " & VBA.ChrW(8450) & " for variablene "
SA(137, 1) = "The system of equations had no solutions in the domain " & VBA.ChrW(8450) & " for variables "
SA(137, 4) = ""
SA(137, 3) = ""
SA(137, 2) = "El sistema de ecuaciones no tiene soluciones en el dominio " & VBA.ChrW(8450) & " para las variables "
SA(137, 5) = ""

SA(138, 0) = "Ligningssystemet havde ingen løsninger indenfor " & VBA.ChrW(8477) & " for variablene "
SA(138, 1) = "The system of equations had no solutions in the domain " & VBA.ChrW(8477) & " for variables "
SA(138, 4) = ""
SA(138, 3) = ""
SA(138, 2) = "El sistema de ecuaciones no tiene soluciones en el dominio " & VBA.ChrW(8450) & " para las variables "
SA(138, 5) = ""

SA(139, 0) = "Der skal være et ligmed tegn i en ligning."
SA(139, 1) = "There must be an equal sign in an equation"
SA(139, 4) = ""
SA(139, 3) = ""
SA(139, 2) = "Tiene que haber un signo igual en una ecuación"
SA(139, 5) = ""

SA(140, 0) = "Ligningssystemet kunne ikke løses analytisk. Det var for komplekst. Der kan dog være løsninger som kan findes vha. numeriske metoder."
SA(140, 1) = "The system of equations could not be solved analytically. You might be able to find solutions using numerical methods"
SA(140, 4) = ""
SA(140, 3) = ""
SA(140, 2) = "El sistema de ecuaciones no pudo resolverse analíticamente. Usted podría encontrar soluciones usando métodos numéricos"
SA(140, 5) = ""

SA(141, 0) = "Der er ikke umiddelbart noget = tegn i ligningen. Vil du fortsætte med at løse den? Udtrykket sættes lig 0"
SA(141, 1) = "There is no = sign in the equation. If you continue the expression will be set =0"
SA(141, 4) = ""
SA(141, 3) = ""
SA(141, 2) = "No se encuentra el signo = en la ecuación. Si continúa la expresión se igualará a cero"
SA(141, 5) = ""

SA(142, 0) = "Følgende variable elimineres fra ligningssystemet: "
SA(142, 1) = "The following variables are eliminated from the equation system: "
SA(142, 4) = ""
SA(142, 3) = ""
SA(142, 2) = "Las siguientes variables serán eliminadas del sistema de ecuaciones"
SA(142, 5) = ""

SA(143, 0) = " vha. " & AppNavn & "'s 'Eliminer variabel' funktion, "
SA(143, 1) = " by " & AppNavn & " 's 'Eliminate variable' function, '"
SA(143, 4) = ""
SA(143, 3) = ""
SA(143, 2) = "Mediante" & AppNavn & " 's 'Elimina variable' función, '"
SA(143, 5) = ""

SA(144, 0) = "sand"
SA(144, 1) = "true"
SA(144, 4) = ""
SA(144, 3) = ""
SA(144, 2) = "Verdadero"
SA(144, 5) = ""

SA(145, 0) = "falsk"
SA(145, 1) = "false"
SA(145, 4) = ""
SA(145, 3) = ""
SA(145, 2) = "Falso"
SA(145, 5) = ""

SA(146, 0) = "ukendt"
SA(146, 1) = "unknown"
SA(146, 4) = ""
SA(146, 3) = ""
SA(146, 2) = "Desconocido"
SA(146, 5) = ""

SA(147, 0) = " vha. " & AppNavn & "'s. 'Løs ligninger numerisk'."
SA(147, 1) = " by " & AppNavn & "'s. 'Solve equation using numerical methods'."
SA(147, 4) = ""
SA(147, 3) = ""
SA(147, 2) = "Mediante" & AppNavn & "'s. 'Resolver ecuación usando métodos numéricos'."
SA(147, 5) = ""

SA(148, 0) = "Der blev fundet en løsning. Der kan dog sagtens være flere løsninger."
SA(148, 1) = "A solution was found. More solutions can exist."
SA(148, 4) = ""
SA(148, 3) = ""
SA(148, 2) = "Se encontró una solución. Pueden existir más soluciones"
SA(148, 5) = ""

SA(149, 0) = "Marker kun et udtryk"
SA(149, 1) = "Only select 1 expression"
SA(149, 4) = ""
SA(149, 3) = ""
SA(149, 2) = "Seleccione sólo una expresión"
SA(149, 5) = ""

SA(150, 0) = "Udtrykket omskrives vha. " & AppNavn & " ved brug af metoderne: "
SA(150, 1) = "The expression is simplified by " & AppNavn & " using the methods: "
SA(150, 4) = ""
SA(150, 3) = ""
SA(150, 2) = "La expresión está simplificada por " & AppNavn & " usando el método: "
SA(150, 5) = ""

SA(151, 0) = "Automatisk reduktion, "
SA(151, 1) = "Automatic simplify, "
SA(151, 4) = ""
SA(151, 3) = ""
SA(151, 2) = "Simplificación automática"
SA(151, 5) = ""

SA(152, 0) = "Trigonometrisk reduktion, "
SA(152, 1) = "Trigonometric identities"
SA(152, 4) = ""
SA(152, 3) = ""
SA(152, 2) = "Identidades trigonométricas"
SA(152, 5) = ""

SA(153, 0) = "Udtrykket kunne ikke reduceres"
SA(153, 1) = "The expression could not be simplfied"
SA(153, 4) = ""
SA(153, 3) = ""
SA(153, 2) = "La expresión no puede simplificarse"
SA(153, 5) = ""

SA(154, 0) = "Udtrykket reduceres vha. " & AppNavn & ". "
SA(154, 1) = "The expression is simplifed by " & AppNavn & ". "
SA(154, 4) = ""
SA(154, 3) = ""
SA(154, 2) = "La expresión está simplificada por " & AppNavn & ". "
SA(154, 5) = ""

SA(155, 0) = "Udtrykket faktoriseres vha. " & AppNavn & ". "
SA(155, 1) = "The expression is factored by " & AppNavn & ". "
SA(155, 4) = ""
SA(155, 3) = ""
SA(155, 2) = "La expresión está factorizada por " & AppNavn & ". "
SA(155, 5) = ""

SA(156, 0) = "Udtrykket kunne ikke faktoriseres"
SA(156, 1) = "The expression could not be factored"
SA(156, 4) = ""
SA(156, 3) = ""
SA(156, 2) = "La expresión no puede factorizarse"
SA(156, 5) = ""

SA(157, 0) = "Udtrykket udvides vha. " & AppNavn & ". "
SA(157, 1) = "The expression is expanded by " & AppNavn & ". "
SA(157, 4) = ""
SA(157, 3) = ""
SA(157, 2) = "La expresión está expandida por " & AppNavn & ". "
SA(157, 5) = ""

SA(158, 0) = "Udtrykket kunne ikke udvides"
SA(158, 1) = "The expression could not be expanded"
SA(158, 4) = ""
SA(158, 3) = ""
SA(158, 2) = "La expresión no puede expandirse"
SA(158, 5) = ""

SA(159, 0) = "Du skal slå enheder fra for at kunne løse differentialligninger"
SA(159, 1) = "You have to disable units in order to solve differential equations"
SA(159, 4) = ""
SA(159, 3) = ""
SA(159, 2) = "Debe desactivar unidades con el fin de resolver las ecuaciones diferenciales"
SA(159, 5) = ""

SA(160, 0) = "Differentialligningen kunne ikke løses"
SA(160, 1) = "The differential equation could not be solved"
SA(160, 4) = ""
SA(160, 3) = ""
SA(160, 2) = "Las ecuaciones diferenciales no se pueden resolver"
SA(160, 5) = ""

SA(161, 0) = "Har du glemt et gangetegn?" & VbCrLfMac & "Der indsættes ikke automatisk gangetegn mellem variable eller mellem variable og parantesstart" & VbCrLfMac & " xy -> x*y" & VbCrLfMac & "y(x+1)  -> y*(x+1)" & VbCrLfMac & "rsagen er at det er tvetydigt. Variable kan være på mere end 1 bogstav og y(x) kan være en funktion."""
SA(161, 1) = "Did you forget a multiplication sign?" & VbCrLfMac & "multiplication signs are not inserted automatically between variables or between variables and start bracket" & VbCrLfMac & " xy -> x*y" & VbCrLfMac & "y(x+1)  -> y*(x+1)" & VbCrLfMac & "Because these situations are ambiguous. Variables can be more than 1 letter and y(x) can be a function."
SA(161, 4) = ""
SA(161, 3) = ""
SA(161, 2) = "øOlvidó un signo de multiplicación?" & VbCrLfMac & "Los signos de multiplicación no se insertan automáticamente entre variables o entre variables y paréntesis de inicio" & VbCrLfMac & " xy -> x*y" & VbCrLfMac & "y(x+1)  -> y*(x+1)" & VbCrLfMac & "Estas situaciones son ambiguas. Las variables pueden quedar de más de una letra y también y(x) puede verse como una función."
SA(161, 5) = ""

SA(162, 0) = "Integrationsgrænserne skal være reelle. Måske mangler du en antagelse."
SA(162, 1) = "Limits of integration must be reel. Maybe you are missing an assumption."
SA(162, 4) = ""
SA(162, 3) = ""
SA(162, 2) = "Los límites de integración deben ser reales. Es posible que necesite un supuesto"
SA(162, 5) = ""

SA(163, 0) = "Udsagnets sandhedsværdi bestemmes vha. " & AppNavn & ". "
SA(163, 1) = "The expression is tested for truth by " & AppNavn & ". "
SA(163, 4) = ""
SA(163, 3) = ""
SA(163, 2) = "El valor de verdad para la expresión se revisó mediante " & AppNavn & ". "
SA(163, 5) = ""

SA(164, 0) = "Resultatet er baseret på numerisk sammenligning"
SA(164, 1) = "The result is based on numerical comparison"
SA(164, 4) = ""
SA(164, 3) = ""
SA(164, 2) = "El resultado se basa en comparación numérica"
SA(164, 5) = ""

SA(165, 0) = "Det kunne ikke afgøres om udsagnet var sand/falsk"
SA(165, 1) = "WordMat was unable to determine whether the expression was true or false"
SA(165, 2) = "WordMat no pudo determinar si la expresión era verdadera o falsa"
SA(165, 3) = ""
SA(165, 4) = ""
SA(165, 5) = ""

SA(166, 0) = "Enheder fra"
SA(166, 1) = "Units off"
SA(166, 2) = ""
SA(166, 3) = ""
SA(166, 4) = ""
SA(166, 5) = ""

SA(167, 0) = "Vælg evt. outputenheder"
SA(167, 1) = "Choose output units"
SA(167, 2) = ""
SA(167, 3) = ""
SA(167, 4) = ""
SA(167, 5) = ""

SA(168, 0) = "Outputenheder"
SA(168, 1) = "Output units"
SA(168, 2) = ""
SA(168, 3) = ""
SA(168, 4) = ""
SA(168, 5) = ""

SA(169, 0) = "Parenteserne mathcher ikke"
SA(169, 1) = "Mismatched brackets"
SA(169, 2) = ""
SA(169, 3) = ""
SA(169, 4) = ""
SA(169, 5) = ""

SA(170, 0) = "Automatisk backup?"
SA(170, 1) = "Automatic backup?"
SA(170, 2) = ""
SA(170, 3) = ""
SA(170, 4) = ""
SA(170, 5) = ""

SA(171, 0) = "Hvor mange backup-filer skal der max gemmes?"
SA(171, 1) = "How many backup-files to store?"
SA(171, 2) = ""
SA(171, 3) = ""
SA(171, 4) = ""
SA(171, 5) = ""

SA(172, 0) = "Tid mellem hver backup"
SA(172, 1) = "Time between backups"
SA(172, 2) = ""
SA(172, 3) = ""
SA(172, 4) = ""
SA(172, 5) = ""

SA(173, 0) = "Word laver ikke en automatisk backup, men opretter kun gendannelsesdata som måske kan bruges til at gendanne en tabt fil." & vbLf & vbLf & " WordMat kan tage automatiske Backup af hele dokumentet hver gang der beregnes og gemme det i flere versioner, så du altid kan gå tilbage og finde noget der virker. (Backup foregår hurtigst hvis dokumentet ikke ligger på OneDrive)" & vbLf & vbLf & "Backup-filerne gemmes i mappen: dokumenter/WordMat-backup"
SA(173, 1) = "Word doesn't do automatic backup, but saves data which possibly can be used to recreate a lost file" & vbLf & vbLf & "WordMat can perform complete backups when you do a calculation and save the document in multiple versions. (Backups are much faster if the document is not in OneDrive) .The backup files are stored in the folder: my documents/WordMat-backup"
SA(173, 2) = ""
SA(173, 3) = ""
SA(173, 4) = ""
SA(173, 5) = ""

SA(174, 0) = "Se backup-filer"
SA(174, 1) = "Open backup folder"
SA(174, 2) = ""
SA(174, 3) = ""
SA(174, 4) = ""
SA(174, 5) = ""

SA(175, 0) = "Spørg"
SA(175, 1) = "Ask"
SA(175, 2) = ""
SA(175, 3) = ""
SA(175, 4) = ""
SA(175, 5) = ""

SA(176, 0) = "Altid"
SA(176, 1) = "Always"
SA(176, 2) = ""
SA(176, 3) = ""
SA(176, 4) = ""
SA(176, 5) = ""

SA(177, 0) = "Aldrig"
SA(177, 1) = "Never"
SA(177, 2) = ""
SA(177, 3) = ""
SA(177, 4) = ""
SA(177, 5) = ""

SA(178, 0) = "Der skete en fejl i forbindelse med backup"
SA(178, 1) = "An error occured when trying to backup"
SA(178, 2) = ""
SA(178, 3) = ""
SA(178, 4) = ""
SA(178, 5) = ""

SA(179, 0) = "Vil du aktivere WordMats automatiske backup-system?" & vbLf & vbLf & "(Se mere under indstillinger | Backup)"
SA(179, 1) = "Do you want to activate WordMats automatic backup-system" & vbLf & vbLf & "(For more info go to settings | Backup)"
SA(179, 2) = ""
SA(179, 3) = ""
SA(179, 4) = ""
SA(179, 5) = ""

SA(180, 0) = "Omskriv"
SA(180, 1) = "Rearrange"
SA(180, 2) = ""

SA(181, 0) = "Lav"
SA(181, 1) = "Low"
SA(181, 2) = ""

SA(182, 0) = "Normal"
SA(182, 1) = "Normal"
SA(182, 2) = ""

SA(183, 0) = "Høj"
SA(183, 1) = "High"
SA(183, 2) = ""

SA(184, 0) = "Meget høj"
SA(184, 1) = "Very high"
SA(184, 2) = ""

SA(185, 0) = "Super (langsom)"
SA(185, 1) = "Super (slow)"
SA(185, 2) = ""

SA(186, 0) = "Forskrift"
SA(186, 1) = "Equation"
SA(186, 2) = ""

SA(187, 0) = "Vent?"
SA(187, 1) = "Wait?"
SA(187, 2) = ""

SA(188, 0) = "Maks"
SA(188, 1) = "Max"
SA(188, 2) = ""

SA(189, 0) = "Kun linjer"
SA(189, 1) = "Only lines"
SA(189, 2) = ""

SA(190, 0) = "Vis akser"
SA(190, 1) = "Show axes"
SA(190, 2) = ""

SA(191, 0) = "Kvalitet"
SA(191, 1) = "Quality"
SA(191, 2) = ""

SA(192, 0) = "Påvirker rotationshastighed"
SA(192, 1) = "Affects rotation speed"
SA(192, 2) = ""

SA(193, 0) = "Når figuren kommer frem kan du ændre synsvinklen ved at holde venstre museknap nede og trække rundt. " & vbLf & "Højreklik på menulinjen, vælg options og 'copy to clipboard' for at kopiere billede. " & vbLf & "Du skal lukke vinduet for at komme tilbage til Word."
SA(193, 1) = "When the picture appears you can change the angle of view by pressing and holding the left mousebutton while dragging. Rightclick the menu line, choose options and 'copy to clipboard' to copy the image. Close the window to return to Word."
SA(193, 2) = ""

SA(194, 0) = "Omdrejningslegeme"
SA(194, 1) = "Solid of revolution"
SA(194, 2) = ""

SA(194, 0) = "Omdrejningslegeme"
SA(194, 1) = "Solid of revolution"
SA(194, 2) = ""

SA(195, 0) = "Indtast koordinater til punkter der skal markeres med rød streg på grafen. Brug semikolon eller tabulator hvis du vil bruge komma som decimalseparator. Du kan anvende definitioner til at aflæse punkter på grafen. Du kan også bruge knappen indsæt, så kan du klikke på grafen hvor der skal markeres." & VbCrLfMac & "eksempel:" & VbCrLfMac & VbCrLfMac & "1,2" & VbCrLfMac & "1,2 ; 3,4" & VbCrLfMac & "2,f(2)"
SA(195, 1) = "Enter coordinates for points to be marked by red dotted lines on the graph. Examples " & VbCrLfMac & "eksempel:" & VbCrLfMac & VbCrLfMac & "1,2" & VbCrLfMac & "1,2 ; 3,4" & VbCrLfMac & "2,f(2)" & VbCrLfMac & " It is also possible to click on the graph to get the coordinate if you use the insert button."
SA(195, 2) = ""

SA(196, 0) = "Her kan du indsætte et anden datasæt af punkter, som vil have en anden farve på grafen. Du kan ikke lave regression på disse, men ved tryk på indsæt kan du placerede punkter på grafen ved at trykke. De kan så evt. forbindes med streger til en figur."
SA(196, 1) = "Also for inserting dotted lines marking points on the graph, but with another color than points in column 1. You cannot perform regression on these point. You can create all sorts of shapes by connecting the dots."
SA(196, 2) = ""

SA(197, 0) = "Vælg tabel med tal inden der trykkes plot, så indsættes punkterne her automatisk. et sæt koordinater pr. linje. Koordinaterne adskilles ad komma, semikolon eller tabulator. Hvis du bruger komma som decimalseparator skal du bruge semikolon eller tabulator. Koordinater kan kopieres ind fra Excel.  Eksempler på notation: " & VbCrLfMac & "1.2 , 3.4" & VbCrLfMac & "3,4 ; 4,8" & VbCrLfMac & "5,6" & vbTab & "7.8"
SA(197, 1) = "Select a table with coordinate data before pressing plot to fill these into this textbox automatically. One set of coordinates pr. line. One set of coordinates pr. line. Coordinates must be separated by comma, semicolon or tab. Coordinates can be copied from Excel. Examples: " & VbCrLfMac & "1.2 , 3.4" & VbCrLfMac & "3,4 ; 4,8" & VbCrLfMac & "5,6" & vbTab & "7.8"
SA(197, 2) = ""

SA(198, 0) = "Variable skal være x og y"
SA(198, 1) = "Variables must be x and y"
SA(198, 2) = ""

SA(199, 0) = "Indsæt cirkel"
SA(199, 1) = "Insert circle"
SA(199, 2) = ""

SA(200, 0) = "Indsæt linje"
SA(200, 1) = "Insert line"
SA(200, 2) = ""

SA(201, 0) = "Markerede punkter"
SA(201, 1) = "Marked points"
SA(201, 2) = ""

SA(202, 0) = "Regression"
SA(202, 1) = "Regression"
SA(202, 2) = ""

SA(203, 0) = "Ja"
SA(203, 1) = "Yes"
SA(203, 2) = ""

SA(204, 0) = "Nej"
SA(204, 1) = "No"
SA(204, 2) = ""

SA(205, 0) = "Husk dette valg"
SA(205, 1) = "Remember this choice"
SA(205, 2) = ""

SA(206, 0) = "Når WordMats backup funktion er slået til vil WordMat tage en backup af dokumentet før hver beregning. Der tages ikke backup hvis sidste backup er foretaget indenfor 5 minutter. Der gemmes 10 backup-udgaver, hvorefter de overskrives." & vbLf & vbLf & "Du kan ændre disse indstillinger under: indstillinger | Backup. Det er også her du kan finde dine backup-filer, hvis det bliver nødvendigt."
SA(206, 1) = "When WordMats backup function is enabled a backup will be performed when a calculation is performed. A backup is only taken every 5 minutes. 10 backup copies are taken before they are overwritten." & vbLf & vbLf & " You can change these settings under settings | Backup."
SA(206, 2) = ""

SA(207, 0) = "Du kan altid ændre dette valg under indstillinger | Backup"
SA(207, 1) = "You can always change this choice under settings | Backup"
SA(207, 2) = ""

SA(208, 0) = "Backup"
SA(208, 1) = "Backup"
SA(208, 2) = ""

SA(209, 0) = "Vinkelsummen er over 180"
SA(209, 1) = "sum of the angles is above 180"
SA(209, 2) = ""

SA(210, 0) = "Du skal indtaste 3 sider/vinkler. Dog ikke 3 vinkler. Mindst 1 side skal være kendt."
SA(210, 1) = "You must enter 3 sides/angles. Not 3 angles though. At least 1 side must be known."
SA(210, 2) = ""

SA(211, 0) = "Vinkelsummen er ikke 180"
SA(211, 1) = "sum of the angles is not 180"
SA(211, 2) = ""

SA(212, 0) = "Du har indtastet for mange sider/vinkler."
SA(212, 1) = "You entered too many angles/sides"
SA(212, 2) = ""

SA(213, 0) = "Du skal indtaste præcis 3 sider/vinkler. Dog ikke 3 vinkler. Mindst 1 side skal være kendt."
SA(213, 1) = "You must enter EXACTLY 3 sides/angles. Not 3 angles though. At least 1 side must be known."
SA(213, 2) = ""

SA(214, 0) = "Mindst en side skal være kendt. 3 vinkler er ikke nok."
SA(214, 1) = ""
SA(214, 2) = ""

SA(215, 0) = "Vinkel"
SA(215, 1) = "Angle"
SA(215, 2) = ""

SA(216, 0) = "findes vha. vinkelsum = 180" & ChrW(176) & " i en trekant"
SA(216, 1) = "found using sum of angles = 180 " & ChrW(176) & " in a triangle"
SA(216, 2) = ""

SA(217, 0) = "Længden af siden"
SA(217, 1) = "The length of side"
SA(217, 2) = ""

SA(218, 0) = "findes vha. pythagoras"
SA(218, 1) = "found using pythagoras"
SA(218, 2) = ""

SA(219, 0) = "Længden af siderne"
SA(219, 1) = "The length of the sides"
SA(219, 2) = ""

SA(220, 0) = "findes vha. tangens"
SA(220, 1) = "found using tan"
SA(220, 2) = ""

SA(221, 0) = "findes vha. sinus"
SA(221, 1) = "found using sine"
SA(221, 2) = ""

SA(222, 0) = "findes vha. cosinus"
SA(222, 1) = "found using cosine"
SA(222, 2) = ""

SA(223, 0) = "findes vha. cosinusrelationer"
SA(223, 1) = "found using law of cosine"
SA(223, 2) = ""

SA(224, 0) = "findes vha. sinusrelationer"
SA(224, 1) = "found using law of sine"
SA(224, 2) = ""

SA(225, 0) = "2. gradsligningen løses for"
SA(225, 1) = "The quadratic equation is solved for"
SA(225, 2) = ""

SA(226, 0) = "2. gradsligningen har to løsninger:"
SA(226, 1) = "This quadratic equation has two solutions: "
SA(226, 2) = ""

SA(227, 0) = "Men den anden løsninger er negativ, og vi kan jo ikke have negative sidelængder."
SA(227, 1) = "But the other solution is negative, and a length cannot be negative."
SA(227, 2) = ""

SA(228, 0) = "Da den cosinusrelation vi løste ovenfor for "
SA(228, 1) = "Since the solution for "
SA(228, 2) = ""

SA(229, 0) = "havde to løsninger og begge var positive er der altså to mulige trekanter. Vi fortsætter nu med at beregne de resterende sider på baggrund af den anden løsning på samme måde."
SA(229, 1) = "had two solutions and both solutions were positive this means there are two possible triangles. We now conitnue to calculate the remaining sides based on the 2nd solution using the same method."
SA(229, 2) = ""

SA(230, 0) = "Indsæt symbol"
SA(230, 1) = "Insert symbol"
SA(230, 2) = ""

SA(231, 0) = "Græske bogstaver"
SA(231, 1) = "Greek letters"
SA(231, 2) = ""

SA(232, 0) = "Andre tegn"
SA(232, 1) = "Other symbols"
SA(232, 2) = ""

SA(233, 0) = "Indsæt"
SA(233, 1) = "Insert"
SA(233, 2) = ""

SA(234, 0) = "Symbol"
SA(234, 1) = "Symbol"
SA(234, 2) = ""

SA(235, 0) = "Løsning af ligning med grafiske og numeriske metoder"
SA(235, 1) = "Solution of equation using graphical and numerical methods"
SA(235, 2) = ""

SA(236, 0) = "Startgæt"
SA(236, 1) = "Initial guess"
SA(236, 2) = ""

SA(237, 0) = "Intervalmetoden"
SA(237, 1) = "Interval method"
SA(237, 2) = ""

SA(238, 0) = "start"
SA(238, 1) = "start"
SA(238, 2) = ""

SA(239, 0) = "slut"
SA(239, 1) = "end"
SA(239, 2) = ""

SA(240, 0) = "Højresiden og venstresiden plottet. " & vbLf & "Venstresiden er sort. Højresiden grøn. Løsninger til ligningen er skæringspunkterne."
SA(240, 1) = "Right side and left side plottet. " & vbLf & "Left side is black. Right side is green."
SA(240, 2) = ""

SA(241, 0) = "Grafisk løsning"
SA(241, 1) = "Graphical solution"
SA(241, 2) = ""

SA(242, 0) = "Zoom også ved at dobbeltklike eller trække firkant"
SA(242, 1) = "Also zoom dy double clicking or dragging a box"
SA(242, 2) = ""

SA(243, 0) = "Vis Graf i stort vindue"
SA(243, 1) = "Show graph in a separate window"
SA(243, 2) = ""

SA(244, 0) = "Du må maks. vælge "
SA(244, 1) = "You may choose a max of"
SA(244, 2) = ""

SA(245, 0) = "Variable"
SA(245, 1) = "Variables"
SA(245, 2) = ""

SA(246, 0) = "Du skal vælge"
SA(246, 1) = "You must choose"
SA(246, 2) = ""

SA(247, 0) = "Vælg maks. "
SA(247, 1) = "Choose max"
SA(247, 2) = ""

SA(248, 0) = "Eller indtast selv maks."
SA(248, 1) = "Or enter manually max"
SA(248, 2) = ""

SA(249, 0) = "variable adskilt af komma"
SA(249, 1) = "variables separated by comma"
SA(249, 2) = ""

SA(250, 0) = "Vælg præcis"
SA(250, 1) = "Choose exactly"
SA(250, 2) = ""

SA(251, 0) = "Eller indtast selv"
SA(251, 1) = "Or enter manually"
SA(251, 2) = ""

SA(252, 0) = "Du kan ikke både løse ligningen for"
SA(252, 1) = "You cant solve the equation for"
SA(252, 2) = ""

SA(253, 0) = "og samtidigt definere den."
SA(253, 1) = "and define it at the same time"
SA(253, 2) = ""

SA(254, 0) = "Du forsøger at løse for en variabel som er en enhed. Det vil formentlig resultere i en fejl. Overvej at bruge et andet variabelnavn eller slå enheder fra."
SA(254, 1) = "You attempting to solve for a variable which is a unit. This will most likely cause an error. Consider using another variable name or disable units."
SA(254, 2) = ""

SA(255, 0) = "jord"
SA(255, 1) = "earth"
SA(255, 2) = ""

SA(256, 0) = "vand"
SA(256, 1) = "water"
SA(256, 2) = ""

SA(257, 0) = "kobber"
SA(257, 1) = "copper"
SA(257, 2) = ""

SA(258, 0) = "jern"
SA(258, 1) = "iron"
SA(258, 2) = ""

SA(259, 0) = "måne"
SA(259, 1) = "moon"
SA(259, 2) = ""

SA(260, 0) = "sol"
SA(260, 1) = "sun"
SA(260, 2) = ""

SA(261, 0) = "Fysiske konstanter"
SA(261, 1) = "Physics constants"
SA(261, 2) = ""

SA(262, 0) = "Enheder"
SA(262, 1) = "Units"
SA(262, 2) = ""

SA(263, 0) = "Tabelværdier - varme"
SA(263, 1) = "Table values - heat"
SA(263, 2) = ""

SA(264, 0) = "Tabelværdi - kernefysik"
SA(264, 1) = "Table values - nuclear physics"
SA(264, 2) = ""

SA(265, 0) = "Astronomi"
SA(265, 1) = "Astronomy"
SA(265, 2) = ""

SA(266, 0) = "Lysets fart i vacuum"
SA(266, 1) = "Speed of light in vacuum"
SA(266, 2) = ""

SA(267, 0) = "Plancks konstant"
SA(267, 1) = "Plancks constant"
SA(267, 2) = ""

SA(268, 0) = "Plancks konstant i Elektron Volt"
SA(268, 1) = "Plancks constant in electron volt"
SA(268, 2) = ""

SA(269, 0) = "Gravitationskonstanten"
SA(269, 1) = "Gravitational constant"
SA(269, 2) = ""

SA(270, 0) = "Tyngdeacceleration ved jordens overflade"
SA(270, 1) = "Free fall of acceleration at the surface of earth"
SA(270, 2) = ""

SA(271, 0) = " Elementarladningen"
SA(271, 1) = "Elementary charge"
SA(271, 2) = ""

SA(272, 0) = "Avogadros konstant"
SA(272, 1) = "Avogadros number"
SA(272, 2) = ""

SA(273, 0) = "Gaskonstanten"
SA(273, 1) = "Gas constant"
SA(273, 2) = ""

SA(274, 0) = "Boltzmanns konstant"
SA(274, 1) = "Boltzmanns constant"
SA(274, 2) = ""

SA(275, 0) = "Vakuumpermeabiliteten"
SA(275, 1) = "Vacuumpermeability"
SA(275, 2) = ""

SA(276, 0) = "Vakuumpermittiviteten"
SA(276, 1) = "Vacuumpermittivity"
SA(276, 2) = ""

SA(277, 0) = "Coulombs konstant"
SA(277, 1) = "Coulombs constant"
SA(277, 2) = ""

SA(278, 0) = "Stefan-Boltzmanns konstant"
SA(278, 1) = "Stefan-Boltzmanns constant"
SA(278, 2) = ""

SA(279, 0) = "atommasseenhed"
SA(279, 1) = "atomic mass unit"
SA(279, 2) = ""

SA(280, 0) = "atommasseenhed i  MeV"
SA(280, 1) = "atomic mass unit in MeV"
SA(280, 2) = ""

SA(281, 0) = "specifik varmekapacitet for"
SA(281, 1) = "specific heat capacity of "
SA(281, 2) = ""

SA(282, 0) = " elektronens masse"
SA(282, 1) = "electron mass"
SA(282, 2) = ""

SA(283, 0) = "i kg"
SA(283, 1) = "in kg"
SA(283, 2) = ""

SA(284, 0) = "protonens masse"
SA(284, 1) = "proton mass"
SA(284, 2) = ""

SA(285, 0) = "neutronens masse"
SA(285, 1) = "neutron mass"
SA(285, 2) = ""

SA(286, 0) = "Jordens masse"
SA(286, 1) = "earth mass"
SA(286, 2) = ""

SA(287, 0) = "Jordens middelradius"
SA(287, 1) = "earth mean radius"
SA(287, 2) = ""

SA(288, 0) = "Afstand mellem jorden og solen (Astronomisk enhed)"
SA(288, 1) = "Distance between earth and sun (Astronomical Unit)"
SA(288, 2) = ""

SA(289, 0) = "Solens masse"
SA(289, 1) = "Mass of the sun"
SA(289, 2) = ""

SA(290, 0) = "Månens masse"
SA(290, 1) = "Mass of the moon"
SA(290, 2) = ""

SA(291, 0) = "Månens middelradius"
SA(291, 1) = "Moon mean radius"
SA(291, 2) = ""

SA(292, 0) = "GeoGebra er ikke installeret. Vil du installere GeoGebra eller vil du køre GeoGebra i en browser?"
SA(292, 1) = "GeoGebra is not installed. Do you want to install GeoGebra or run GeoGebra from a Browser?"
SA(292, 2) = ""

SA(293, 0) = "Bedste måde at få GeoGebra til at virke sammen med WordMat på. Husk dog selv at se efter opdateringer."
SA(293, 1) = "Best way to use GeoGebra but requires an installation."
SA(293, 2) = ""

SA(294, 0) = "Udmærket midlertidig løsning. Giver altid nyeste version. "
SA(294, 1) = "Good temporary solution"
SA(294, 2) = ""

SA(295, 0) = "Installer GeoGebra"
SA(295, 1) = "Install GeoGebra"
SA(295, 2) = ""

SA(296, 0) = "Kør i browser"
SA(296, 1) = "Run in browser"
SA(296, 2) = ""

SA(297, 0) = "eller Evt. randbetingelser"
SA(297, 1) = "or boundary conditions"
SA(297, 2) = ""

SA(298, 0) = "Etiket"
SA(298, 1) = "Label"
SA(298, 2) = ""

SA(299, 0) = "Indtast tekst og klik dernæst på grafen hvor teksten skal placeres"
SA(299, 1) = "Enter text and then click on the graph to place it"
SA(299, 2) = ""

SA(300, 0) = "Der er en fejl i forskrift"
SA(300, 1) = "There is an error in equation"
SA(300, 2) = ""

SA(301, 0) = "Du er først nødt til at sætte ymin og ymax og trykke opdater."
SA(301, 1) = "You must set ymin and ymax first - then press update"
SA(301, 2) = ""

SA(302, 0) = "3D Grafer"
SA(302, 1) = "3D Graphs"
SA(302, 2) = ""

SA(303, 0) = "Vis udtryk"
SA(303, 1) = "Show expression"
SA(303, 2) = ""

SA(304, 0) = "Gennemsigtig"
SA(304, 1) = "Transparent"
SA(304, 2) = ""

SA(305, 0) = "Fladefarver"
SA(305, 1) = "Color of surfaces"
SA(305, 2) = ""

SA(306, 0) = "Indsæt Maxima kommando i Word"
SA(306, 1) = "Insert Maxima command into Word"
SA(306, 2) = ""

SA(307, 0) = "Linjer og flader"
SA(307, 1) = "Lines and surfaces"
SA(307, 2) = ""

SA(308, 0) = "Kun linjer"
SA(308, 1) = "Only lines"
SA(308, 2) = ""

SA(309, 0) = "Kun flader"
SA(309, 1) = "Only surfaces"
SA(309, 2) = ""

SA(310, 0) = "Linjer og flader"
SA(310, 1) = "Lines and surfaces"
SA(310, 2) = ""

SA(311, 0) = "Variable skal være x, y og z"
SA(311, 1) = "Variables must be x, y and z"
SA(311, 2) = ""

SA(312, 0) = "Indsæt kugle"
SA(312, 1) = "Insert sphere"
SA(312, 2) = ""

SA(313, 0) = "Indsæt plan"
SA(313, 1) = "Insert plane"
SA(313, 2) = ""

SA(314, 0) = "Ny vektor"
SA(314, 1) = "New vector"
SA(314, 2) = ""

SA(315, 0) = "Syntaks" & vbLf & "(startpunkt)(vektor)" & vbLf & "(startx , starty , startz)(dx , dy , dz)" & vbLf & "En vektor pr. linje" & vbLf & "Hvis startpunkt ikke angives antages (0,0,0)" & vbLf & "Eksempler:" & vbLf & "(1,1,2)(1,2,3)     " & vbLf & "  tegner vektor (1,2,3) fra punkt (1,1,2)" & vbLf & "" & vbLf & "(1,-2,3)" & vbLf & "  tegner vektor (1,-2,3) fra punkt (0,0,0)" & vbLf & "" & vbLf & "decimaltal eksempler:" & vbLf & "(1,5 ; 3 ; 4,5)" & vbLf & "(1.5 , 3 , 4.5)"
SA(315, 1) = "Syntax" & vbLf & "(startx , starty , startz)(dx , dy , dz)" & vbLf & "One vector pr. line" & vbLf & "Examples:" & vbLf & "(1,1,2)(1,2,3)     " & vbLf & " draws vector (1,2,3) from point (1,1,2)" & vbLf & "" & vbLf & "(1,-2,3)" & vbLf & " draws vector (1,-2,3) from point (0,0,0)" & vbLf & "" & vbLf & "example with decimals:" & vbLf & "(1.5 , 3 , 4.5)"
SA(315, 2) = ""

SA(316, 0) = "Punktstørrelse"
SA(316, 1) = "Point size"
SA(316, 2) = ""

SA(317, 0) = "Indtast punkter således" & vbLf & "" & vbLf & "(1,1;2,5;1,2)" & vbLf & "(2.1,3.4,8.8)"
SA(317, 1) = "Syntax for entering points" & vbLf & "" & vbLf & "(1,2,1)" & vbLf & "(2.5,3.6,8.7)"
SA(317, 2) = ""

SA(318, 0) = "I grafvinduet kan man rotere koordinatsystemet ved at holde venstremuseknap nede og trække rundt." & vbLf & "" & vbLf & "Hvis objektet ser strukket ud kan man ændre grafvinduet så det passer." & vbLf & "" & vbLf & "For at indsætte billedet i Word, skal man højreklikke på menulinjen, vælge 'options' og dernæst 'copy to clipboard'. Så lukkes grafvinduet og så kan billedet indsættes i Word." & vbLf & "" & vbLf & "Bemærk at grafvinduet skal lukkes for at komme tilbage til Word." & vbLf & "" & vbLf & "Antal punkter bestemmer hvor pæn figuren bliver, men påvirker samtidigt hastigheden hvormed der kan roteres."
SA(318, 1) = "In the graph window you can rotate the coordinate system by pressing left mouse button while dragging." & vbLf & "" & vbLf & "If the object looks stretched you can change the size of the graph window for a better fit." & vbLf & "" & vbLf & "To insert the image in Word, right click the menu line, choose  'options' and then 'copy to clipboard'. Then the graph window is closed and the image can be inserted into Word." & vbLf & "" & vbLf & "Notice that the graph window must be closed to be able return to Word." & vbLf & "" & vbLf & "The higher quality of the image the slower rotate speed."
SA(318, 2) = ""

SA(319, 0) = "Forskrifter"
SA(319, 1) = "Rules"
SA(319, 2) = ""

SA(320, 0) = "Parameterfremstillinger"
SA(320, 1) = "Parametric plot"
SA(320, 2) = ""

SA(321, 0) = "Gul/rød/lilla"
SA(321, 1) = "Yellow/red/purple"
SA(321, 2) = ""

SA(322, 0) = "Blå"
SA(322, 1) = "Blue"
SA(322, 2) = ""

SA(323, 0) = "Brun"
SA(323, 1) = "Brown"
SA(323, 2) = ""

SA(324, 0) = "Grå"
SA(324, 1) = "Grey"
SA(324, 2) = ""

SA(325, 0) = "Genvejen   Alt + M   indsætter nyt matematikfelt"
SA(325, 1) = "Shortcut          Alt + M         inserts new math "
SA(325, 2) = ""

SA(326, 0) = "Genvejen   AltGr + Enter   beregner"
SA(326, 1) = "Shortcut        AltGr + Enter       calculates "
SA(326, 2) = ""

SA(327, 0) = "Genvejen   Alt + L    løser ligning"
SA(327, 1) = "Shortcut          Alt + E         Solves equation "
SA(327, 2) = ""

SA(328, 0) = "Genvejen   Alt + i   åbner indstillinger"
SA(328, 1) = "Shortcut          Alt + S         opens settings"
SA(328, 2) = ""

SA(329, 0) = "Genvejen   Alt + r   henter forrige resultater"
SA(329, 1) = "Shortcut          Alt + R         Inserts previous result"
SA(329, 2) = ""

SA(330, 0) = "Genvejen   Alt + E   slår enheder til/fra"
SA(330, 1) = "Shortcut          Alt + U         Turns on unit support"
SA(330, 2) = ""

SA(331, 0) = "sænket skrift: "
SA(331, 1) = "subscript: "
SA(331, 2) = ""

SA(332, 0) = "\lambda giver det lille græske bogstav "
SA(332, 1) = "\lambda is short for the small greek letter lambda"
SA(332, 2) = ""

SA(333, 0) = "Genvejen  Alt + P  viser grafen for en forskrift"
SA(333, 1) = "Shortcut          Alt + P         Plots a graph"
SA(333, 2) = ""

SA(334, 0) = "Du kan selv gemme ændringer i formelsamlingerne og figurdokumentet."
SA(334, 1) = "You can save your changes to the formulae documents"
SA(334, 2) = ""

SA(335, 0) = "Excelark kan både indsættes indlejret i dokumentet eller åbnes i Excel"
SA(335, 1) = "Excel sheets can be opened separately or inserted embedded"
SA(335, 2) = ""

SA(336, 0) = "Hurtig og bedre retning af opgaver? Prøv WordMark - www.eduap.com"
SA(336, 1) = "Marking of handins? Try WordMark  www.eduap.com"
SA(336, 2) = ""

SA(337, 0) = "Du har ialt foretaget " & AntalB & " beregninger med WordMat"
SA(337, 1) = "You have done a total of " & AntalB & " calculations with WordMat"
SA(337, 2) = ""

SA(338, 0) = "Under definitioner kan du definere fysiske konstanter og tabelværdier"
SA(338, 1) = "Under definitions you can define physical constants and table values."
SA(338, 2) = ""

SA(339, 0) = "Word kører hurtigere i kladdevisning, specielt for ligninger."
SA(339, 1) = "Word runs faster in draft view - when equations are involved"
SA(339, 2) = ""

SA(340, 0) = "Indsæt dine egne beskrivelser i Exceldiagrammer vha. Indsæt/tekstboks"
SA(340, 1) = "Insert you own descriptions of Excel diagrams using insert textbox"
SA(340, 2) = ""

SA(341, 0) = "Genvejen   Alt + N   skifter mellem auto,eksakt,num"
SA(341, 1) = "Shortcut          Alt + N         Switches between auto,exact,num"
SA(341, 2) = ""

SA(342, 0) = "En funktion med bogstav-konstanter får automatisk skydere i GeoGebra"
SA(342, 1) = "A function with letters as constant gets sliders in GeoGebra"
SA(342, 2) = ""

SA(343, 0) = "Outputenheder må kun være rene enheder. Dvs. ingen sammensatte enheder som f.eks. km/timer. Hvis du vil have f.eks. km/timer, så indtast: km,timer"
SA(343, 1) = "Output units may not be combined units.  km/s is not allowed. Then state: km,s"
SA(343, 2) = ""

SA(344, 0) = "Du har allerede den nyeste version af"
SA(344, 1) = "You already have the most recent version of"
SA(344, 2) = ""

SA(345, 0) = "Vælg et ligningsnr."
SA(345, 1) = "Select an equation number"
SA(345, 2) = ""

SA(346, 0) = "Hvilket nr. skal denne ligning sættes til?" & VbCrLfMac & "De efterfølgende ligninger får opdateres deres numre så det passer."
SA(346, 1) = "What number to set this equation to?" & VbCrLfMac & "Any following equations will be updated as well."
SA(346, 2) = ""

SA(347, 0) = "Sæt ligningsnummer"
SA(347, 1) = "Set equation number"
SA(347, 2) = ""

SA(348, 0) = "Et ligningssektion nr. vil blive indsat. Typisk indsætter man denne som et kapitel eller sektions nr. i en overskrift. Nummeret kan dog også indsættes skjult. Skal nummeret vises? (Svar nej for at skjule det)"
SA(348, 1) = "An equation section number will be inserted. Typically this is inserted as a chapter/section number in a heading. But it can also be inserted hidden. Do you want to show the number? (Answer No to hide it)"
SA(348, 2) = ""

SA(349, 0) = "Indtast signifikansniveauet i procent"
SA(349, 1) = "Enter significance level in percentage"
SA(349, 2) = ""

SA(350, 0) = "Signifikansniveau"
SA(350, 1) = "Significance level"
SA(350, 2) = ""

SA(351, 0) = "Indtast dimensionen af krydstabellen" & VbCrLfMac & "rækker x kolonner"
SA(351, 1) = "Enter dimensions of the table" & VbCrLfMac & "rows x columns"
SA(351, 2) = ""

SA(352, 0) = "Chi i anden test for sammenhæng"
SA(352, 1) = "Chi square test for independence"
SA(352, 2) = ""

SA(353, 0) = "Der udføres " & VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & " - test for sammenhæng vha. WordMat"
SA(353, 1) = " " & VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & " - test of indepence performed using WordMat"
SA(353, 2) = ""

SA(354, 0) = "P-værdi"
SA(354, 1) = "P-value"
SA(354, 2) = ""

SA(355, 0) = "På signifikansniveau"
SA(355, 1) = "Using significance level of"
SA(355, 2) = ""

SA(356, 0) = "% kan 0-hypotesen, at der er ikke er sammenhæng, forkastes. Dvs. Der er en sammenhæng mellem de to variable."
SA(356, 1) = "% the 0-hypothesis can be rejected. Hence there is dependence between the two variables."
SA(356, 2) = ""

SA(357, 0) = "% kan 0-hypotesen ikke forkastes. Der er derfor ikke belæg for at sige at der er en sammenhæng mellem de to variable."
SA(357, 1) = "% the 0-hypothesis cannot be rejected. Hence there is no evidence to support dependence between the two variables."
SA(357, 2) = ""

SA(358, 0) = "Der er celler med værdier mindre end 5 eller den samlede sum af alle celler er mindre end 50. " & ChrW(&H3C7) & ChrW(&HB2) & "-fordelingen er derfor en dårlig approksimation i dette tilfælde og resultatet upålideligt. I nogle tilfælde kan datamaterialet gøres brugbart ved at lave en anden gruppering."
SA(358, 1) = "There are cells with value less than 5 or the sum of all cells is less than 50. The " & ChrW(&H3C7) & ChrW(&HB2) & "-distribution is then a bad approximation in this case and the result unreliable. In some cases the data can be made useful by grouping the numbers differently."
SA(358, 2) = ""

SA(359, 0) = "Chifordeling"
SA(359, 1) = "Chidistribution"
SA(359, 2) = ""

SA(360, 0) = "Frihedsgrader"
SA(360, 1) = "Degrees of freedom"
SA(360, 2) = ""

SA(361, 0) = "Aflæs"
SA(361, 1) = "Read"
SA(361, 2) = ""

SA(362, 0) = "Åbner GeoGebra"
SA(362, 1) = "Opening GeoGebra"
SA(362, 2) = ""

SA(363, 0) = "Indtast værdi for"
SA(363, 1) = "Enter value for"
SA(363, 2) = ""

SA(364, 0) = "Den kan justeres efterfølgende"
SA(364, 1) = "You can adjust it later"
SA(364, 2) = ""

SA(365, 0) = "Definer konstant"
SA(365, 1) = "Define constant"
SA(365, 2) = ""

SA(366, 0) = "Du har ikke installeret Graph. Det er et lille gratis program der er nemt installeret. Det laver fine grafer, punkter mm. som nemt kan integreres i Word. Vil du hente og installere Graph?"
SA(366, 1) = "Graph is not installed. It is a small free program which is easily installed. Do you want to install it now?"
SA(366, 2) = ""

SA(367, 0) = "Den kan justeres efterfølgende under brugerdefinerede funktioner i Graph. Der kan tilmed laves en animation hvor værdien ændres løbende, under Beregn/Animer."
SA(367, 1) = "You can adjust it later under user defined functions in Graph"
SA(367, 2) = ""

SA(368, 0) = "Plotning af grafer med GeoGebra og Graph kræver at x er den uafhængige variabel." & VbCrLfMac & "Følgende udtryk indeholder ikke x, så du kan nu angive en anden uafhængig variabel."
SA(368, 1) = "Plotting of graphs using GeoGeoGebra and Graph requires x to be the independent variable. " & VbCrLfMac & "The following expressions are without x. You can now state a new variable to use as x."
SA(368, 2) = ""

SA(369, 0) = "Uafhængig variabel"
SA(369, 1) = "Independent variable"
SA(369, 2) = ""

SA(370, 0) = "Denne funktion kræver at GeoGebra installeres"
SA(370, 1) = "This function requires GeoGebra to be installed"
SA(370, 2) = ""

SA(371, 0) = "Åbner Graph"
SA(371, 1) = "Opening Graph"
SA(371, 2) = ""

SA(372, 0) = "Indsætter indlejret"
SA(372, 1) = "Inserting embedded"
SA(372, 2) = ""

SA(373, 0) = "Indsætning af indlejrede objekter kan tage tid. Dobbeltklik på objektet for at redigere det."
SA(373, 1) = "Inserting embedded object can take some time. Double on embedded object to open them."
SA(373, 2) = ""

SA(374, 0) = "Vil du plotte"
SA(374, 1) = "Do you want to plot"
SA(374, 2) = ""

SA(375, 0) = "Indsæt"
SA(375, 1) = "Insert"
SA(375, 2) = ""

SA(376, 0) = "Ligningssystemer af uligheder understøttes ikke. Løs som ligninger eller løs grafisk vha. Graph eller GeoGebra"
SA(376, 1) = "Systems of equations of inequalities are not supported yet. Solve as an equation or solve graphically using Graph or GeoGebra."
SA(376, 2) = ""

SA(377, 0) = "Eliminer variable"
SA(377, 1) = "Eliminate variables"
SA(377, 2) = ""

SA(378, 0) = "med værdierne"
SA(378, 1) = "with the values"
SA(378, 2) = ""

SA(379, 0) = "Indtast startgæt på værdier. Standardgæt er 1"
SA(379, 1) = "Enter startguess. Standard guess value is 1"
SA(379, 2) = ""

SA(380, 0) = "Variable og gæt"
SA(380, 1) = "Variables and guess"
SA(380, 2) = ""

SA(381, 0) = "Ligningssystemet løses numerisk for"
SA(381, 1) = "The system of equations is solved for "
SA(381, 2) = ""

SA(382, 0) = "med startgæt"
SA(382, 1) = "with start guess"
SA(382, 2) = ""

SA(383, 0) = "Ligningssystemet kunne ikke løses"
SA(383, 1) = "The system of equations could not be solved"
SA(383, 2) = ""

SA(384, 0) = "Angav du den rigtige variabel?"
SA(384, 1) = "Did you state the correct variable?"
SA(384, 2) = ""

SA(385, 0) = "Der blev ikke fundet en løsning. Intervalendepunkterne skal være på hver side af een løsning."
SA(385, 1) = "A solution was not found. Interval start and end point must be on both sides of a solution"
SA(385, 2) = ""

SA(386, 0) = "Newtons metode anvendes med startgæt"
SA(386, 1) = "Newtons method applied with start guess"
SA(386, 2) = ""

SA(387, 0) = "Intervalmetoden anvendes i intervallet"
SA(387, 1) = "Interval method applied in the interval"
SA(387, 2) = ""

SA(388, 0) = "Det lykkedes ikke at finde en løsning"
SA(388, 1) = "A solution was not found"
SA(388, 2) = ""

SA(389, 0) = "Ingen løsning"
SA(389, 1) = "No solution"
SA(389, 2) = ""

SA(390, 0) = "Differentialligningen løses vha. " & AppNavn & "'s 'Løs differentialligning' funktion"
SA(390, 1) = "The differential equation is solved by " & AppNavn & "'s 'Solve differential equation' function"
SA(390, 2) = ""

SA(391, 0) = "med randbetingelserne"
SA(391, 1) = "with boundary conditions"
SA(391, 2) = ""

SA(392, 0) = "med startbetingelserne"
SA(392, 1) = "with start conditions"
SA(392, 2) = ""

SA(393, 0) = "med startbetingelsen"
SA(393, 1) = "with start condition"
SA(393, 2) = ""

SA(394, 0) = "Du bliver nødt til at genstarte Word nu, ellers låser Word."
SA(394, 1) = "You have to restart Word. If you dont Word will crash."
SA(394, 2) = ""

SA(395, 0) = "Der skete en fejl i forsøget på at indsætte en formel. Prøv igen et andet sted."
SA(395, 1) = "An error occurred attempting to insert a formula. Try again another place."
SA(395, 2) = ""

SA(396, 0) = "Beregner det markerede eller hele udtrykket tilbage til nærmeste ligmed tegn"
SA(396, 1) = "Calculates the selected expression, or back to the nearest equal sign."
SA(396, 2) = ""

SA(397, 0) = "Løser ligning"
SA(397, 1) = "Solves equation"
SA(397, 2) = ""

SA(398, 0) = "Antal frihedsgrader" & VbCrLfMac & "n giver helt generelt udtryk"
SA(398, 1) = "No of degrees of freedom" & VbCrLfMac & "state n for a generel expression."
SA(398, 2) = ""

SA(399, 0) = "Frekvensfordeling med"
SA(399, 1) = "Frequency distribution with"
SA(399, 2) = ""

SA(400, 0) = "Observationer"
SA(400, 1) = "Observations"
SA(400, 2) = ""

SA(401, 0) = "Forventede"
SA(401, 1) = "Expected"
SA(401, 2) = ""

SA(402, 0) = "I alt"
SA(402, 1) = "Total"
SA(402, 2) = ""

SA(403, 0) = "P-værdi"
SA(403, 1) = "P-value"
SA(403, 2) = ""

SA(404, 0) = "Der er en sammenhæng"
SA(404, 1) = "There is a correlation"
SA(404, 2) = ""

SA(405, 0) = "Der er ikke en sammenhæng"
SA(405, 1) = "There is no correlation"
SA(405, 2) = ""

SA(406, 0) = "Den samlede sum af alle celler er mindre end 50. " & ChrW(&H3C7) & ChrW(&HB2) & "-fordelingen er derfor en dårlig approksimation i dette tilfælde og resultatet upålideligt."
SA(406, 1) = "The sum of all cells is less than 50. " & ChrW(&H3C7) & ChrW(&HB2) & "-distribution is then a poor approximation and the result unreliable."
SA(406, 2) = ""

SA(407, 0) = "Q-værdi"
SA(407, 1) = "Q-value"
SA(407, 2) = ""

SA(408, 0) = "Kritisk Q-værdi"
SA(408, 1) = "Critical Q-value"
SA(408, 2) = ""

SA(409, 0) = "Du har ikke installeret MikTex. Du sendes nu til hjemmesiden hvor MikTex kan downloades. Efter installationen er det vigtigt at computeren genstartes. Når MikTex køres første gang skal der installeres nogle pakker."
SA(409, 1) = "MikTex is not installed. You will now be send to the download page. The computer must be restarted after the installation. The first time you run MikTex from WordMat a series of packages will be installed."
SA(409, 2) = ""

SA(410, 0) = "Denne funktion understøttes ikke på Mac"
SA(410, 1) = "This function is not supported on Mac"
SA(410, 2) = ""

SA(411, 0) = "om 20s stoppes automatisk."
SA(411, 1) = "Stops automatically in 20s"
SA(411, 2) = ""

SA(412, 0) = "om 15s stoppes automatisk."
SA(412, 1) = "Stops automatically in 15s"
SA(412, 2) = ""

SA(413, 0) = "om 10s stoppes automatisk."
SA(413, 1) = "Stops automatically in 10s"
SA(413, 2) = ""

SA(414, 0) = "om 5s stoppes automatisk."
SA(414, 1) = "Stops automatically in 5s"
SA(414, 2) = ""

SA(415, 0) = "Det lykkedes ikke. Stopper."
SA(415, 1) = "Unsuccesful. Stopping."
SA(415, 2) = ""

SA(416, 0) = "Du har indtastet en vektorligning som egentlig er et ligningssystem. Indtast de "
SA(416, 1) = "You have entered a vector equation which is actually a system of equations. Enter the"
SA(416, 2) = ""

SA(417, 0) = "variable der skal løses for adskilt ad komma."
SA(417, 1) = "variables to solve for, separated by comma."
SA(417, 2) = ""

SA(418, 0) = "Hov du skulle indtaste"
SA(418, 1) = "You have to enter "
SA(418, 2) = ""

SA(419, 0) = "variable adskilt ad komma"
SA(419, 1) = "variables separated by comma"
SA(419, 2) = ""

SA(420, 0) = "En af ligninger er ikke defineret indenfor de valgte grænser. Prøv at udvide området"
SA(420, 1) = "One of the equations is not defined within the limits of the plot area. Try expanding the plot area"
SA(420, 2) = ""

SA(421, 0) = "Der indgår et eller flere " & ChrW(177) & " symboler i udtrykket. Hvordan skal det forståes?"
SA(421, 1) = "One or more " & ChrW(177) & " symbols appear in the expression. How to interpret?"
SA(421, 2) = ""

SA(422, 0) = "Plus eller minus"
SA(422, 1) = "Plus or minus"
SA(422, 2) = ""

SA(423, 0) = "Input ikke forstået. Der indsættes plus."
SA(423, 1) = "Input not understood. A plus is inserted."
SA(423, 2) = ""

SA(424, 0) = "Der er et problem med de numeriske tegn"
SA(424, 1) = "There is a problem with the numeric signs"
SA(424, 2) = ""

SA(425, 0) = "Notationen for sumtegnet er ikke korrekt. Den nedre grænse skal skrives f.eks. i=1"
SA(425, 1) = "The notation for the sum is not correct. The syntax for the lower limit is: i=1  (or similar)"
SA(425, 2) = ""

SA(426, 0) = "Notationen for produkttegnet er ikke korrekt. Den nedre grænse skal skrives f.eks. i=1"
SA(426, 1) = "The notation for the product is not correct. The syntax for the lower limit is: i=1  (or similar)"
SA(426, 2) = ""

SA(427, 0) = "Der er problemer med syntaxen. Der skal bruges et dx (eller lign) som afslutning af et integrale. Måske har du brugt d i anden sammenhæng. Brug da differential d (skriv: \dd)"
SA(427, 1) = "There are syntax problems. You must end an integral with dx (if x is the variable). Maybe you have used d in another context. Then use a differential d (type: \dd)"
SA(427, 2) = ""

SA(428, 0) = "Der er problemer med syntaxen. Der skal bruges to dx (eller lign) som afslutning af et dobbletintegrale. Måske har du brugt d i anden sammenhæng. Brug da differential d \dd"
SA(428, 1) = "There are syntax problems. A double integral must end with two dx's (depending on the variables to integrate for). Maybe you have used d in another context. Then use a differential d (type: \dd)"
SA(428, 2) = ""

SA(429, 0) = "Du kan ikke bruge dobbeltintegraltegnet til bestemte integraler. Brug to enkeltintegraler."
SA(429, 1) = "You cant use the double integralsign for definite integrals. Use two integral signs"
SA(429, 2) = ""

SA(430, 0) = "Fejl! Der er problemer med syntaxen. Notation for differentiering er ikke korrekt."
SA(430, 1) = "Error! There are syntax problems. Notation for differentiation is not correct."
SA(430, 2) = ""

SA(431, 0) = "ikke defineret"
SA(431, 1) = "not defined"
SA(431, 2) = ""

SA(432, 0) = "ikke defineret, men begrænset"
SA(432, 1) = "not defined but limited"
SA(432, 2) = ""

SA(433, 0) = "Ikke defineret indenfor"
SA(433, 1) = "Not defined within"
SA(433, 2) = ""

SA(434, 0) = "er reserveret og kan ikke defineres. Definitionen springes over."
SA(434, 1) = "is reserved and cant be defined. The definition is skipped."
SA(434, 2) = ""



LoadSprogArray2

End Sub

Sub LoadSprogArray2()

SA(435, 0) = "S=B" & ChrW(183) & "(1+r)"
SA(435, 1) = "S=B" & ChrW(183) & "(1+r)"
SA(435, 2) = ""

SA(436, 0) = "K=K" & ChrW(&H2092) & ChrW(183) & "(1+r)" & ChrW(&H207F) & "     Kapitalfremskrivningsformel"
SA(436, 1) = "A=P" & ChrW(183) & "(1+r/n)^nt      compound interest calculation"
SA(436, 2) = ""

SA(437, 0) = "y=a" & ChrW(183) & "x+b                Lineær ligning"
SA(437, 1) = "y=m" & ChrW(183) & "x+c                Linear equation"
SA(437, 2) = ""

SA(438, 0) = "a=(y" & ChrW(&H2082) & "-y" & ChrW(&H2081) & ")/(x" & ChrW(&H2082) & "-x" & ChrW(&H2081) & ")     Hældningskoefficient"
SA(438, 1) = "a=(y" & ChrW(&H2082) & "-y" & ChrW(&H2081) & ")/(x" & ChrW(&H2082) & "-x" & ChrW(&H2081) & ")     gradient (slope)"
SA(438, 2) = ""

SA(439, 0) = "y=a" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+y" & ChrW(&H2080) & "         Lineær ligning ud fra punkt (x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ") og a"
SA(439, 1) = "y=m" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+y" & ChrW(&H2080) & "         Linear equation from a point (x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ") and the gradient m"
SA(439, 2) = ""

SA(440, 0) = "y=f'(x" & ChrW(&H2080) & ")" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+f(x" & ChrW(&H2080) & ")     Tangent til f(x) til x=x" & ChrW(&H2080) & ""
SA(440, 1) = "y=f'(x" & ChrW(&H2080) & ")" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+f(x" & ChrW(&H2080) & ")     Tangent to f(x) at x=x" & ChrW(&H2080) & ""
SA(440, 2) = ""

SA(441, 0) = "y=b" & ChrW(183) & "a" & ChrW(&H2E3) & "                  Eksponentiel funktion"
SA(441, 1) = "y=b" & ChrW(183) & "a" & ChrW(&H2E3) & "                  Exponentiel function"
SA(441, 2) = ""

SA(442, 0) = "y=b" & ChrW(183) & "e" & ChrW(&H1D4F) & ChrW(&H2E3) & "                  Eksponentiel funktion"
SA(442, 1) = "y=b" & ChrW(183) & "e" & ChrW(&H1D4F) & ChrW(&H2E3) & "                  Exponentiel function"
SA(442, 2) = ""

SA(443, 0) = "y=b" & ChrW(183) & "2^(x/T" & ChrW(&H2082) & ")                  Eksponentiel funktion"
SA(443, 1) = "y=b" & ChrW(183) & "2^(x/T" & ChrW(&H2082) & ")                  Exponentiel function"
SA(443, 2) = ""

SA(444, 0) = "y=b" & ChrW(183) & ChrW(189) & "^(x/T" & ChrW(189) & ")                  Eksponentiel funktion"
SA(444, 1) = "y=b" & ChrW(183) & ChrW(189) & "^(x/T" & ChrW(189) & ")                  Exponentiel function"
SA(444, 2) = ""

SA(445, 0) = "a=(x" & ChrW(&H2082) & "-x" & ChrW(&H2081) & ")" & ChrW(&H221A) & "(y" & ChrW(&H2082) & "/y" & ChrW(&H2081) & ")   Beregning af a ud fra to kendte punkter (x" & ChrW(&H2081) & ",y" & ChrW(&H2081) & ") og (x" & ChrW(&H2082) & ",y" & ChrW(&H2082) & ")"
SA(445, 1) = "a=(x" & ChrW(&H2082) & "-x" & ChrW(&H2081) & ")" & ChrW(&H221A) & "(y" & ChrW(&H2082) & "/y" & ChrW(&H2081) & ")   calculation of a from to known points (x" & ChrW(&H2081) & ",y" & ChrW(&H2081) & ") og (x" & ChrW(&H2082) & ",y" & ChrW(&H2082) & ")"
SA(445, 2) = ""

SA(446, 0) = "T" & ChrW(&H2082) & "=ln(2)/ln(a)=ln(2)/k         Fordoblingskonstant"
SA(446, 1) = "T" & ChrW(&H2082) & "=ln(2)/ln(a)=ln(2)/k         Double time"
SA(446, 2) = ""

SA(447, 0) = "T" & ChrW(189) & "=ln(" & ChrW(189) & ")/ln(a)=ln(" & ChrW(189) & ")/k             Halveringskonstant"
SA(447, 1) = "T" & ChrW(189) & "=ln(" & ChrW(189) & ")/ln(a)=ln(" & ChrW(189) & ")/k             Half life"
SA(447, 2) = ""

SA(448, 0) = "y=b" & ChrW(183) & "x" & ChrW(&HAA) & "                          potensfunktion ligning"
SA(448, 1) = "y=b" & ChrW(183) & "x" & ChrW(&HAA) & "                          Power function"
SA(448, 2) = ""

SA(449, 0) = "a=log(y" & ChrW(&H2082) & "/y" & ChrW(&H2081) & ")/log(x" & ChrW(&H2082) & "/x" & ChrW(&H2081) & ")   Beregning af a ud fra to kendte punkter (x" & ChrW(&H2081) & ",y" & ChrW(&H2081) & ") og (x" & ChrW(&H2082) & ",y" & ChrW(&H2082) & ")"
SA(449, 1) = "a=log(y" & ChrW(&H2082) & "/y" & ChrW(&H2081) & ")/log(x" & ChrW(&H2082) & "/x" & ChrW(&H2081) & ")   Calculation of a from two known points (x" & ChrW(&H2081) & ",y" & ChrW(&H2081) & ") og (x" & ChrW(&H2082) & ",y" & ChrW(&H2082) & ")"
SA(449, 2) = ""

SA(450, 0) = "(1+r" & ChrW(&H1D67) & ")=(1+r" & ChrW(&H2093) & ")" & ChrW(&HAA) & "    "
SA(450, 1) = "(1+r" & ChrW(&H1D67) & ")=(1+r" & ChrW(&H2093) & ")" & ChrW(&HAA) & "    "
SA(450, 2) = ""

SA(451, 0) = "x" & ChrW(&H1D7C) & "=-b/2a  ,  y" & ChrW(&H1D7C) & "=-d/4a    Toppunktets koordinater"
SA(451, 1) = "x" & ChrW(&H1D7C) & "=-b/2a  ,  y" & ChrW(&H1D7C) & "=-d/4a    extremum coordinates"
SA(451, 2) = ""

SA(452, 0) = "a/sin(A)=b/sin(B)          Sinus-relation"
SA(452, 1) = "a/sin(A)=b/sin(B)          Law of sines"
SA(452, 2) = ""

SA(453, 0) = "c" & ChrW(&HB2) & "=a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "-2" & ChrW(183) & "a" & ChrW(183) & "b" & ChrW(183) & "cos(C)    Cosinus-relation"
SA(453, 1) = "c" & ChrW(&HB2) & "=a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "-2" & ChrW(183) & "a" & ChrW(183) & "b" & ChrW(183) & "cos(C)    Law of cosines"
SA(453, 2) = ""

SA(454, 0) = "T=" & ChrW(189) & ChrW(183) & "a" & ChrW(183) & "b" & ChrW(183) & "sin(C)           Areal af trekant givet vinkel og to sider omkring"
SA(454, 1) = "T=" & ChrW(189) & ChrW(183) & "a" & ChrW(183) & "b" & ChrW(183) & "sin(C)           Area of triangle from known angle and two neighbouring sides"
SA(454, 2) = ""

SA(455, 0) = "K(n,r)" & ChrW(&H2261) & "n!/(r!" & ChrW(183) & "(n-r)!)    Binomialkoefficient"
SA(455, 1) = "K(n,r)" & ChrW(&H2261) & "n!/(r!" & ChrW(183) & "(n-r)!)    Binomialcoefficient"
SA(455, 2) = ""

SA(456, 0) = "p(r)=K(n,r)" & ChrW(183) & "p" & ChrW(&H2B3) & ChrW(183) & "(1-p)" & ChrW(&H207F) & ChrW(&H207B) & ChrW(&H2B3) & "   Frekvensfunktion"
SA(456, 1) = "p(r)=K(n,r)" & ChrW(183) & "p" & ChrW(&H2B3) & ChrW(183) & "(1-p)" & ChrW(&H207F) & ChrW(&H207B) & ChrW(&H2B3) & "   Frequency function"
SA(456, 2) = ""

SA(457, 0) = "P(m)=" & ChrW(&H3A3) & "K(n,r)" & ChrW(183) & "p" & ChrW(&H2B3) & ChrW(183) & "(1-p)" & ChrW(&H207F) & ChrW(&H207B) & ChrW(&H2B3) & "   Kumuleret"
SA(457, 1) = "P(m)=" & ChrW(&H3A3) & "K(n,r)" & ChrW(183) & "p" & ChrW(&H2B3) & ChrW(183) & "(1-p)" & ChrW(&H207F) & ChrW(&H207B) & ChrW(&H2B3) & "   Cumulative"
SA(457, 2) = ""

SA(458, 0) = "" & ChrW(&H3BC) & "=n" & ChrW(183) & "p    Middelværdi"
SA(458, 1) = "" & ChrW(&H3BC) & "=n" & ChrW(183) & "p    Mean value"
SA(458, 2) = ""

SA(459, 0) = "" & ChrW(&H3C3) & "=" & ChrW(&H221A) & "(n" & ChrW(183) & "p" & ChrW(183) & "(1-p))   Spredning"
SA(459, 1) = "" & ChrW(&H3C3) & "=" & ChrW(&H221A) & "(n" & ChrW(183) & "p" & ChrW(183) & "(1-p))   Standard deviation"
SA(459, 2) = ""

SA(460, 0) = "f(x)=1/" & ChrW(&H221A) & "(2" & ChrW(&H3C0) & "" & ChrW(&H3C3) & ")" & ChrW(183) & "e^(-" & ChrW(189) & "(x-" & ChrW(&H3BC) & "/" & ChrW(&H3C3) & ")" & ChrW(&HB2) & ")   frekvensfunktion"
SA(460, 1) = "f(x)=1/" & ChrW(&H221A) & "(2" & ChrW(&H3C0) & "" & ChrW(&H3C3) & ")" & ChrW(183) & "e^(-" & ChrW(189) & "(x-" & ChrW(&H3BC) & "/" & ChrW(&H3C3) & ")" & ChrW(&HB2) & ")   frequency function"
SA(460, 2) = ""

SA(461, 0) = "F(x)=" & ChrW(&H222B) & "1/" & ChrW(&H221A) & "(2" & ChrW(&H3C0) & "" & ChrW(&H3C3) & ")" & ChrW(183) & "e^(-" & ChrW(189) & "(x-" & ChrW(&H3BC) & "/" & ChrW(&H3C3) & ")" & ChrW(&HB2) & ")   Kumuleret frekvensfunktion"
SA(461, 1) = "F(x)=" & ChrW(&H222B) & "1/" & ChrW(&H221A) & "(2" & ChrW(&H3C0) & "" & ChrW(&H3C3) & ")" & ChrW(183) & "e^(-" & ChrW(189) & "(x-" & ChrW(&H3BC) & "/" & ChrW(&H3C3) & ")" & ChrW(&HB2) & ")   Cumulative frequency function"
SA(461, 2) = ""

SA(462, 0) = "p(x)=k" & ChrW(183) & "x" & ChrW(&H207F) & "" & ChrW(189) & "" & ChrW(&HB2) & ChrW(&H207B) & ChrW(&HB9) & ChrW(183) & "e" & ChrW(&H207B) & ChrW(&H2E3) & "" & ChrW(189) & "" & ChrW(&HB2) & "  frekvensfunktion med frihedsgrad n"
SA(462, 1) = "p(x)=k" & ChrW(183) & "x" & ChrW(&H207F) & "" & ChrW(189) & "" & ChrW(&HB2) & ChrW(&H207B) & ChrW(&HB9) & ChrW(183) & "e" & ChrW(&H207B) & ChrW(&H2E3) & "" & ChrW(189) & "" & ChrW(&HB2) & "  frequency function with n degrees of freedom"
SA(462, 2) = ""

SA(463, 0) = "V=" & ChrW(&H3C0) & ChrW(&H222B) & "f(x)" & ChrW(&HB2) & "dx     Rumfang af omdrejningslegeme"
SA(463, 1) = "V=" & ChrW(&H3C0) & ChrW(&H222B) & "f(x)" & ChrW(&HB2) & "dx     Volume of revolution"
SA(463, 2) = ""

SA(464, 0) = "s=" & ChrW(&H222B) & "" & ChrW(&H221A) & "1+(f'(x))" & ChrW(&HB2) & "dx     Kurvelængde af f(x) i interval a-b"
SA(464, 1) = "s=" & ChrW(&H222B) & "" & ChrW(&H221A) & "1+(f'(x))" & ChrW(&HB2) & "dx     Length of curve of f(x) in interval a-b"
SA(464, 2) = ""

SA(465, 0) = "1/(b-a)" & ChrW(&H222B) & "f(x)dx     Middelværdi af f(x) i interval a-b"
SA(465, 1) = "1/(b-a)" & ChrW(&H222B) & "f(x)dx     Mean value of f(x) in interval a-b"
SA(465, 2) = ""

SA(466, 0) = "a" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+b" & ChrW(183) & "(y-y" & ChrW(&H2080) & ")=0     Ligning for en linje"
SA(466, 1) = "a" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+b" & ChrW(183) & "(y-y" & ChrW(&H2080) & ")=0     Equation of a line"
SA(466, 2) = ""

SA(467, 0) = "(x,y)=(x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ")+t" & ChrW(183) & "(r" & ChrW(&H2081) & ",r" & ChrW(&H2082) & ")     parameterfremstilling for en linje"
SA(467, 1) = "(x,y)=(x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ")+t" & ChrW(183) & "(r" & ChrW(&H2081) & ",r" & ChrW(&H2082) & ")     parametric equation of a straight line"
SA(467, 2) = ""

SA(468, 0) = "cos(V)=a" & ChrW(183) & "b/(|a||b|)     Vinkel mellem vektorer"
SA(468, 1) = "cos(V)=a" & ChrW(183) & "b/(|a||b|)     Angle between vectors"
SA(468, 2) = ""

SA(469, 0) = "dist(P,l)=|a" & ChrW(183) & "x" & ChrW(&H2081) & "+b" & ChrW(183) & "y" & ChrW(&H2081) & "+c|/" & ChrW(&H221A) & "a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "     Afstand fra punkt til linje"
SA(469, 1) = "dist(P,l)=|a" & ChrW(183) & "x" & ChrW(&H2081) & "+b" & ChrW(183) & "y" & ChrW(&H2081) & "+c|/" & ChrW(&H221A) & "a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "    Distance from point to line"
SA(469, 2) = ""

SA(470, 0) = "b_a=a" & ChrW(183) & "b/|a|" & ChrW(&HB2) & ChrW(183) & "a     Projektion af vektor b på vektor a"
SA(470, 1) = "b_a=a" & ChrW(183) & "b/|a|" & ChrW(&HB2) & ChrW(183) & "a     Projection of vektor b on vektor a"
SA(470, 2) = ""

SA(471, 0) = "(x-x" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(y-y" & ChrW(&H2080) & ")" & ChrW(&HB2) & "=r" & ChrW(&HB2) & "     Cirklens ligning"
SA(471, 1) = "(x-x" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(y-y" & ChrW(&H2080) & ")" & ChrW(&HB2) & "=r" & ChrW(&HB2) & "     Equation of a circle"
SA(471, 2) = ""

SA(472, 0) = "(x,y,z)=(x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ",z" & ChrW(&H2080) & ")+t" & ChrW(183) & "(r" & ChrW(&H2081) & ",r" & ChrW(&H2082) & ",r" & ChrW(&H2083) & ")     parameterfremstilling for en linje"
SA(472, 1) = "(x,y,z)=(x" & ChrW(&H2080) & ",y" & ChrW(&H2080) & ",z" & ChrW(&H2080) & ")+t" & ChrW(183) & "(r" & ChrW(&H2081) & ",r" & ChrW(&H2082) & ",r" & ChrW(&H2083) & ")     parametric equation of a line"
SA(472, 2) = ""

SA(473, 0) = "cos(V)=a" & ChrW(183) & "b/(|a||b|)     Vinkel mellem vektorer"
SA(473, 1) = "cos(V)=a" & ChrW(183) & "b/(|a||b|)     Angle between vectors"
SA(473, 2) = ""

SA(474, 0) = "b_a=a" & ChrW(183) & "b/|a|" & ChrW(&HB2) & "" & ChrW(183) & "a     Projektion af vektor b på vektor a"
SA(474, 1) = "b_a=a" & ChrW(183) & "b/|a|" & ChrW(&HB2) & "" & ChrW(183) & "a     Projection of vektor b on vektor a"
SA(474, 2) = ""

SA(475, 0) = "dist(P,l)=|r x P" & ChrW(&H2080) & "P|/r     afstand fra punkt til linje"
SA(475, 1) = "dist(P,l)=|r x P" & ChrW(&H2080) & "P|/r     Distance from point to line"
SA(475, 2) = ""

SA(476, 0) = "n" & ChrW(183) & "(x-x" & ChrW(&H2080) & ",y-y" & ChrW(&H2080) & ",z-z" & ChrW(&H2080) & ")     ligning for plan"
SA(476, 1) = "n" & ChrW(183) & "(x-x" & ChrW(&H2080) & ",y-y" & ChrW(&H2080) & ",z-z" & ChrW(&H2080) & ")     Equation of a plane"
SA(476, 2) = ""

SA(477, 0) = " a" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+b" & ChrW(183) & "(y-y" & ChrW(&H2080) & ")+c" & ChrW(183) & "(z-z" & ChrW(&H2080) & ")=0     ligning for plan"
SA(477, 1) = " a" & ChrW(183) & "(x-x" & ChrW(&H2080) & ")+b" & ChrW(183) & "(y-y" & ChrW(&H2080) & ")+c" & ChrW(183) & "(z-z" & ChrW(&H2080) & ")=0     Equation of a plane"
SA(477, 2) = ""

SA(478, 0) = "dist(P," & ChrW(&H3B1) & ")=|n-(x" & ChrW(&H2081) & "-x" & ChrW(&H2080) & ",y" & ChrW(&H2081) & "-y" & ChrW(&H2080) & ",z" & ChrW(&H2081) & "-z" & ChrW(&H2080) & ")|/|n|     Afstand fra punkt til plan"
SA(478, 1) = "dist(P," & ChrW(&H3B1) & ")=|n-(x" & ChrW(&H2081) & "-x" & ChrW(&H2080) & ",y" & ChrW(&H2081) & "-y" & ChrW(&H2080) & ",z" & ChrW(&H2081) & "-z" & ChrW(&H2080) & ")|/|n|     Distance from point to plane"
SA(478, 2) = ""

SA(479, 0) = "dist(P," & ChrW(&H3B1) & ")=|a" & ChrW(183) & "x" & ChrW(&H2081) & "+b" & ChrW(183) & "y" & ChrW(&H2081) & "+c" & ChrW(183) & "z" & ChrW(&H2081) & "+d)|/" & ChrW(&H221A) & "(a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "+c" & ChrW(&HB2) & ")     Afstand fra punkt til plan"
SA(479, 1) = "dist(P," & ChrW(&H3B1) & ")=|a" & ChrW(183) & "x" & ChrW(&H2081) & "+b" & ChrW(183) & "y" & ChrW(&H2081) & "+c" & ChrW(183) & "z" & ChrW(&H2081) & "+d)|/" & ChrW(&H221A) & "(a" & ChrW(&HB2) & "+b" & ChrW(&HB2) & "+c" & ChrW(&HB2) & ")     Distance from point to plane"
SA(479, 2) = ""

SA(480, 0) = "(x-x" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(y-y" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(z-z" & ChrW(&H2080) & ")" & ChrW(&HB2) & "=r" & ChrW(&HB2) & "     Kuglens ligning"
SA(480, 1) = "(x-x" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(y-y" & ChrW(&H2080) & ")" & ChrW(&HB2) & "+(z-z" & ChrW(&H2080) & ")" & ChrW(&HB2) & "=r" & ChrW(&HB2) & "     Equation of a sphere"
SA(480, 2) = ""

SA(481, 0) = "Binomial"
SA(481, 1) = "Binomial"
SA(481, 2) = ""

SA(482, 0) = "Normalford"
SA(482, 1) = "Normal distrib"
SA(482, 2) = ""

SA(483, 0) = ChrW(&H3C7) & ChrW(&HB2) & "-fordeling"
SA(483, 1) = ChrW(&H3C7) & ChrW(&HB2) & "-distribution"
SA(483, 2) = ""

SA(484, 0) = "Plangeometri"
SA(484, 1) = "2D Geometry"
SA(484, 2) = ""

SA(485, 0) = "Rumgeometri"
SA(485, 1) = "3D Geometry"
SA(485, 2) = ""

SA(486, 0) = "Åbner matematikformelsamling"
SA(486, 1) = "Opens math formula booklet"
SA(486, 2) = ""

SA(487, 0) = "Åbner fysikformelsamling"
SA(487, 1) = "Opens physics formula booklet"
SA(487, 2) = ""

SA(488, 0) = "Åbner kemiformelsamling"
SA(488, 1) = "Opens chemistry formula booklet"
SA(488, 2) = ""

SA(489, 0) = "Automatisk valg af eksakt beregning og numerisk værdi."
SA(489, 1) = "Automatic exact or numeric result"
SA(489, 2) = ""

SA(490, 0) = "Automatisk valg af eksakt beregning og numerisk værdi. Ved beregn angives ofte begge resultater."
SA(490, 1) = "WordMat determines automatically wheather to return exact or numeric result or both."
SA(490, 2) = ""

SA(491, 0) = "Eksakte resultater"
SA(491, 1) = "Exact results"
SA(491, 2) = ""

SA(492, 0) = "Alle resultater angives som eksakte resultater. Dvs. så vidt muligt undgåes numeriske approksimationer. Det kan dog ikke garanteres."
SA(492, 1) = "All results returned as exact result. Sometimes a numeric approximation cannot be avoided though."
SA(492, 2) = ""

SA(493, 0) = "Numeriske resultater"
SA(493, 1) = "Numeric results"
SA(493, 2) = ""

SA(494, 0) = "Alle resultater angives som et decimaltal."
SA(494, 1) = "All results returned as decimal numbers"
SA(494, 2) = ""

SA(495, 0) = "Antal betydende cifre"
SA(495, 1) = "Significant figures"
SA(495, 2) = ""

SA(496, 0) = "Angiver hvor mange betydende cifre som resultater skal angives med."
SA(496, 1) = "Determines the number of significant figures on the results"
SA(496, 2) = ""

SA(497, 0) = "Radianer"
SA(497, 1) = "Radians"
SA(497, 2) = ""

SA(498, 0) = "Når trykket ind regnes i radianer istedet for grader."
SA(498, 1) = "When depressed all trigonometric calculations are in radians otherwise degrees."
SA(498, 2) = ""

SA(499, 0) = "Videnskabelig notation"
SA(499, 1) = "Scientific notation"
SA(499, 2) = ""

SA(500, 0) = "Når trykket ind angives resultater i videnskabelig notation dvs. fx 2,3" & ChrW(183) & "10^4"
SA(500, 1) = "When depressed all results are returned in scientific notation. Example: 2.3" & ChrW(183) & "10^4"
SA(500, 2) = ""

SA(501, 0) = "Indstillinger (Alt+i)"
SA(501, 1) = "Settings (Alt+S)"
SA(501, 2) = ""

SA(502, 0) = "Indstil diverse indstillinger."
SA(502, 1) = "Change all sorts of settings."
SA(502, 2) = ""

SA(503, 0) = "Beregn (AltGr+Enter)"
SA(503, 1) = "Calculate (AltGr+Enter or Alt+C)"
SA(503, 2) = ""

SA(504, 0) = "Udregner det markerede udtryk. Hvis cursoren placeres i et udtryk beregnes hele udtrykket. Hvis der er et =tegn i udtrykket beregnes kun tilbage til =tegnet."
SA(504, 1) = "Calculates the selected expression, or back to the nearest equal sign. "
SA(504, 2) = ""

SA(505, 0) = "Maxima Kommando"
SA(505, 1) = "Maxima command"
SA(505, 2) = ""

SA(506, 0) = "Sender kommando ufiltreret til Maxima. Resultatet bliver dog formatteret."
SA(506, 1) = "Sends command unfiltered to maxima. The result is formatted though."
SA(506, 2) = ""

SA(507, 0) = "Løs ligning(er) (Alt+L)"
SA(507, 1) = "Solves equation(s) (Alt+E)"
SA(507, 2) = ""

SA(508, 0) = "Løser en ligning eller et ligningssystem. Placer cursoren i ligningen. Vælg flere ligninger for at løse ligningsystem. "
SA(508, 1) = "Solves an equation or system of equations. Place the cursor in the equation. Select more equations to solve a system of equations"
SA(508, 2) = ""

SA(509, 0) = "Løs vha. numeriske metoder"
SA(509, 1) = "Solve numerically"
SA(509, 2) = ""

SA(510, 0) = "Løs ligning vha. numeriske metoder. Der er flere metoder at vælge mellem."
SA(510, 1) = "Solve equation using numerical methods."
SA(510, 2) = ""

SA(511, 0) = "Eliminer variabel"
SA(511, 1) = "Eliminate variables"
SA(511, 2) = ""

SA(512, 0) = "Eliminerer en eller flere variable fra et ligningsystem."
SA(512, 1) = "Eliminate one or more variables from a system of equations."
SA(512, 2) = ""

SA(513, 0) = "Test om udsagn er sandt"
SA(513, 1) = "Test if true"
SA(513, 2) = ""

SA(514, 0) = "Tester om udsagn er sandt, falsk eller ukendt"
SA(514, 1) = "Test if equation is true"
SA(514, 2) = ""

SA(515, 0) = "Løs Differentialligning"
SA(515, 1) = "Solve differential equation"
SA(515, 2) = ""

SA(516, 0) = "Løser differentialligning. Kan indtastes på forskellige måder.  f.eks.  y'=y , y''=y eller f'(x)=f(x)"
SA(516, 1) = "Solve differential equation. Enter be entered using different notation. Examples:  y'=y , y''=y or f'(x)=f(x)"
SA(516, 2) = ""

SA(517, 0) = "Løs koblede Differentialligninger numerisk"
SA(517, 1) = "Solve system of differential equations numerically"
SA(517, 2) = ""

SA(518, 0) = "Skal stå på formen y'=..."
SA(518, 1) = "Must be in the form y'=?"
SA(518, 2) = ""

SA(519, 0) = "Vis definitioner"
SA(519, 1) = "Show definitions"
SA(519, 2) = ""

SA(520, 0) = "Viser hvilke definitioner der er gældende på dette sted i dokumentet. "
SA(520, 1) = "Shows definitions active at this position in the document"
SA(520, 2) = ""

SA(521, 0) = "Slet definitioner (Alt+S)"
SA(521, 1) = "Delete definitions (Alt+F)"
SA(521, 2) = ""

SA(522, 0) = "Indsætter 'slet definitioner:' kommando i dokument. Alle definitioner placeret før dette sted i dokumentet kan ikke bruges efter dette punkt."
SA(522, 1) = "Inserts 'Delete def:' command in the document. All definitions placed before this command will not be valid in the rest of the document."
SA(522, 2) = ""

SA(523, 0) = "Definer (Alt+D)"
SA(523, 1) = "Define (Alt+D)"
SA(523, 2) = ""

SA(524, 0) = "Indsætter definition i dokumentet som gælder i resten af dokumentet."
SA(524, 1) = "Inserts definition which is valid in the rest of the document"
SA(524, 2) = ""

SA(525, 0) = "Definer fysiske konstanter"
SA(525, 1) = "Define physical constants"
SA(525, 2) = ""

SA(526, 0) = "Vælg blandt de mest brugte fysiske konstanter og få dem defineret i dokumentet så de kan indgå i beregninger"
SA(526, 1) = "Choose among the most used constants and have them defined in the document for use in calculations."
SA(526, 2) = ""

SA(527, 0) = "Omskriv"
SA(527, 1) = "Reduce"
SA(527, 2) = ""

SA(528, 0) = "Omskriver udtrykket på forskellige måder."
SA(528, 1) = "Reduce an expression using different methods"
SA(528, 2) = ""

SA(529, 0) = "Reducer"
SA(529, 1) = "Simplify"
SA(529, 2) = ""

SA(530, 0) = "Hvis muligt skrives udtrykket om så det bliver reduceret mindst muligt. Forsøg evt. omskriv med specifikke metoder  hvis det ikke virker."
SA(530, 1) = "Automatic attempt at simplifying expression. Use reduce and choose specific methods if it fails."
SA(530, 2) = ""

SA(531, 0) = "Faktoriser"
SA(531, 1) = "Factorize"
SA(531, 2) = ""

SA(532, 0) = "Hvis muligt sættes udenfor parantes ved brug af gængse regler som f.eks. kvadratsætningerne"
SA(532, 1) = "Factorizes an expression"
SA(532, 2) = ""

SA(533, 0) = "Ganger paranteser ud"
SA(533, 1) = "Expand"
SA(533, 2) = ""

SA(534, 0) = "Alle paranteser i det markerede udtryk bliver ganget ud."
SA(534, 1) = "All parenthesis are expanded"
SA(534, 2) = ""

SA(535, 0) = "Infinitesimalregning"
SA(535, 1) = "Calculus"
SA(535, 2) = ""

SA(536, 0) = "Man kan også differentiere og integrere vha. differentialregningsnotation som fx f'(x) og dy/dx - se mere i manualen"
SA(536, 1) = "Functions to differentiate and integrate. You also do calculus using standard notation such as f'(x) and dy/dx - see the manual."
SA(536, 2) = ""

SA(537, 0) = "Differentier"
SA(537, 1) = "Differentiate"
SA(537, 2) = ""

SA(538, 0) = "Marker udtryk der skal differentieres, eller placer cursoren i udtrykket."
SA(538, 1) = "Select the expression to differentiate or place the cursor in the expression"
SA(538, 2) = ""

SA(539, 0) = "Stamfunktion"
SA(539, 1) = "Integrate"
SA(539, 2) = ""

SA(540, 0) = "Finder stamfunktion til markerede udtryk (ubestemte integrale)."
SA(540, 1) = "Determines the indefinite integral of function"
SA(540, 2) = ""

SA(541, 0) = "Wolfram Alpha"
SA(541, 1) = "Wolfram Alpha"
SA(541, 2) = ""

SA(542, 0) = "Åbner hjemmesiden Wolfram Alpha, med det valgte udtryk som input."
SA(542, 1) = "Opens Wolfram Alpha with the selected expression as input"
SA(542, 2) = ""

SA(543, 0) = "Løs ligning"
SA(543, 1) = "Solve Equation"
SA(543, 2) = ""

SA(544, 0) = "Løs ligning vha. Wolfram Alpha"
SA(544, 1) = "Solve equation using Wolffram Alpha"
SA(544, 2) = ""

SA(545, 0) = "Vis graf (Alt+P)"
SA(545, 1) = "Plot graph (Alt+P)"
SA(545, 2) = ""

SA(546, 0) = "Viser grafen for de(t) valgte funktioner/ligninger og punker. Grafen kan efterfølgende indsættes i dokumentet. Der understøttes forskellige grafprogrammer. Vælg din favorit i indstillinger, eller vælg ved at udvide denne knap"
SA(546, 1) = "Shows the graph for the selected function(s)/equation(s) and points. The graph can then be inserted into the document. Several graph programs are supported. Choose your favorite under settings."
SA(546, 2) = ""

SA(547, 0) = "Indsæt graf vha. programmet GnuPlot"
SA(547, 1) = "Plot using GnuPlot"
SA(547, 2) = ""

SA(548, 0) = "Vis graf vha. GnuPlot. Gnuplot understøtter mange funktioner og kan lave en pæn graf. Man kan efterfølgende dobbeltklikke på grafen i dokumentet for at redigere det."
SA(548, 1) = "Show graph using GnuPlot. GnuPlot supports many functions and creates a nice looking graph. You can doubleclick the graph in the document to edit it again."
SA(548, 2) = ""

SA(549, 0) = "Indsæt graf vha. programmet Graph"
SA(549, 1) = "Plot using Graph"
SA(549, 2) = ""

SA(550, 0) = "Indsætter Graph-objekt i Word-dokumentet, som man redigeres ved dobbeltklikke på grafobjektet. Definitioner, forskrifter og punkter overføres fra Word til Graph. Graph er et meget hurtigt program og har mange gode funktioner. Ulempen er at det ikke understøttes på Mac og at graferne ikke vises i så høj opløsning."
SA(550, 1) = "Inserts embedded Graph-object in Word, which can be edited by doubleclicking the graph. Definitions, functions, equations and points are transferred from Word to Graph. Graph is a very fast program with many functions. The downside is that is not supported on Mac and the graphs are shown in low resolution."
SA(550, 2) = ""

SA(551, 0) = "GeoGebra"
SA(551, 1) = "GeoGebra"
SA(551, 2) = ""

SA(552, 0) = "Åbner GeoGebra og overfører den valgte funktion samt definitioner. GeoGebra er det grafprogram der giver de fleste muligheder, og en rigtig pæn graf, men der er nogle begrænsninger: Det er ret langtsomt at åbne og billedet overføres ikke automatisk til Word. Det skal kopieres ind manuelt, via rediger-menuen og 'kopier tegningen'.Grafen kan ikke redigeres igen, når den er indsat."
SA(552, 1) = "Opens GeoGebra and transfers the selected function(s)/equation(s) to GeoGebra, as well as any definitions. GeoGebra is the most feature rich graph program. It also creates a high quality graph. Downsides are: It is slow to open and the graph is not automatically transfered to Word. It has to copied manually to Word using the edit-menu/'Copy Graphics view to clipboard'. The graph cannot be edited again once inserted."
SA(552, 2) = ""

SA(553, 0) = "Indsæt Excel-graf"
SA(553, 1) = "Excel graph"
SA(553, 2) = ""

SA(554, 0) = "Indsætter indlejret Excel ark i Word-dokumentet, som er designet til visning af grafer og punkter. Man kan dobbeltklikke på excelobjektet for at redigere det igen. Det fungerer ikke så hurtigt som nogle af de andre løsninger og har ikke så mange funktioner. Fordelene er: Understøtter læring i Excel og er speciel designet til nem plotning af punkter. Det kan derfor være et program som finder nytte i samarbejde med andre fag specielt indenfor naturvidenskab."
SA(554, 1) = "Inserts embedded Excel-sheet designed to show graphs and points. Doubleclick the excelsheets to edit it again. It is not as fast and feature rich as some of the other graph-solutions. Advantages are: Supports learning Excel and is designed for easy plotting of points. It is a good choice for collaboration with other subject such as science subjects."
SA(554, 2) = ""

SA(555, 0) = "Retn.felt"
SA(555, 1) = "Direction field"
SA(555, 2) = ""

SA(556, 0) = "Viser hældningsfelt for første ordens differentialligning og giver mulighed for at at plotte løsningskurver."
SA(556, 1) = "Shows direction field and plots solution curves for first order differential equation."
SA(556, 2) = ""

SA(557, 0) = "Indsæt GeoGebra-fil"
SA(557, 1) = "Insert GeoGebra file"
SA(557, 2) = ""

SA(558, 0) = "Indsætter indlejret tom GeoGebra-fil, som kan redigeres og gemmes direkte i Word. Filen vises som ikon. Billeder skal stadig kopieres over."
SA(558, 1) = "Inserts empty GeoGebra file, which can be edited and saved from within Word. The file is shown as an icon. Graph images must be copied into the document."
SA(558, 2) = ""

SA(559, 0) = "3D plot"
SA(559, 1) = "3D plot"
SA(559, 2) = ""

SA(560, 0) = "Vis forskrift, ligning, punkter, vektorer og parameterfremstillinger i rummet."
SA(560, 1) = "Plot equations, points, vectors and parametric plots, in 3D."
SA(560, 2) = ""

SA(561, 0) = "Omdrejningslegeme"
SA(561, 1) = "Solid of revolution"
SA(561, 2) = ""

SA(562, 0) = "Viser omdrejningslegeme for angiven forskrift. Dvs. Grafen for forskriften roteres 360 grader om x-aksen, så der opstår 3D objekt."
SA(562, 1) = "Shows solid of revolution for a function. The graph of the function is rotated 360 degrees around the x-axis creating a 3D object."
SA(562, 2) = ""

SA(563, 0) = "Statistik"
SA(563, 1) = "Statistics"
SA(563, 2) = ""

SA(564, 0) = "Åbner regneark der bl.a. kan deskriptiv statistik"
SA(564, 1) = "Opens excel workbook designed for one variable statistics. Look at the sheets to see all the functions."
SA(564, 2) = ""

SA(565, 0) = "Ugrupperede Observationer"
SA(565, 1) = "Grouped observations"
SA(565, 2) = ""

SA(566, 0) = "Excelark til behandling af ugrupperede data. Beregner frekvenser, kvartilsæt mm og generer pindediagram, prikdiagram, trappediagram og boksplot."
SA(566, 1) = "Excel sheet for ungrouped data. Calculates frequencies, mean, quartiles etc and generates relevant diagrams like box and whiskers diagram"
SA(566, 2) = ""

SA(567, 0) = "Grupperede Observationer"
SA(567, 1) = "Grouped observations"
SA(567, 2) = ""

SA(568, 0) = "Excelark til beregning af frekvenser, kumulerede frekvenser, kvartilsæt og visning af histogram, boksplot og sumkurve."
SA(568, 1) = "Excel sheet for grouped data. Calculates frequencies, mean etc and generates relevant diagrams like histogram and box and whiskers diagram"
SA(568, 2) = ""

SA(569, 0) = "Pindediagram"
SA(569, 1) = "Bar chart"
SA(569, 2) = ""

SA(570, 0) = "Excelark med forberedt pindediagram"
SA(570, 1) = "Excel sheet for creating bar charts"
SA(570, 2) = ""

SA(571, 0) = "Histogram"
SA(571, 1) = "Histogram"
SA(571, 2) = ""

SA(572, 0) = "Excelark beregnet til at plotte flere histogrammer i samme diagram."
SA(572, 1) = "Excel sheet to plot multiple histograms in one diagram."
SA(572, 2) = ""

SA(573, 0) = "Trappediagram"
SA(573, 1) = "Stair case diagram"
SA(573, 2) = ""

SA(574, 0) = "Excelark der kan tegnes flere trappediagrammer i samme diagram til sammenligning"
SA(574, 1) = ""
SA(574, 2) = ""

SA(575, 0) = "Sumkurve"
SA(575, 1) = "Cumulative frequency diagram"
SA(575, 2) = ""

SA(576, 0) = "Excelark der kan tegnes flere sumkurver i samme diagram til sammenligning"
SA(576, 1) = ""
SA(576, 2) = ""

SA(577, 0) = "Boksplot"
SA(577, 1) = "Box and whisker plot"
SA(577, 2) = ""

SA(578, 0) = "Excelark der kan plotte flere boksplot i samme diagram"
SA(578, 1) = "Excel sheet to plot multiple box and whisker plots in one diagram."
SA(578, 2) = ""

SA(579, 0) = "Regression"
SA(579, 1) = "Regression"
SA(579, 2) = ""

SA(580, 0) = "Indsæt en almindelig tabel med punkterne. Marker punkterne og vælg regressionstype. Tabellen kan både være vandret eller lodret. Hvis der kun er to punkter antages den at være lodret."
SA(580, 1) = "First insert a table with the points. Then place the cursor in the table and choose regression type. The table can be vertical or horizontal. If there is only two points it is assumed to be vertical."
SA(580, 2) = ""

SA(581, 0) = "Indsæt tabel "
SA(581, 1) = "Insert table"
SA(581, 2) = ""

SA(582, 0) = "Indsætter tabel med to kolonner beregnet til indsætning af punkter som der kan udføres regression på."
SA(582, 1) = "Inserts table with two columns for entering points for regression."
SA(582, 2) = ""

SA(583, 0) = "Fordelinger"
SA(583, 1) = "Distributions"
SA(583, 2) = ""

SA(584, 0) = "Excelark med grafer over udvalgte fordelinger. Arkene er dynamiske så de kan plotte graferne for alle parametre og lave de mest gængse beregninger."
SA(584, 1) = "Excel sheet with graphs of distribution. The diagrams are dynamic and can plot distributions for all parameters."
SA(584, 2) = ""

SA(585, 0) = "Binomial-Fordeling"
SA(585, 1) = "Binomial distribution"
SA(585, 2) = ""

SA(586, 0) = "Excel-ark med graf over binomial-Fordeling. På arket vælges basissandsynlighed og antalsparameter"
SA(586, 1) = "Excel sheet that plot binomial distribution. On the sheet choose p and n"
SA(586, 2) = ""

SA(587, 0) = "Normal-Fordeling"
SA(587, 1) = "Normal distribution"
SA(587, 2) = ""

SA(588, 0) = "Excel-ark med graf over Normal-Fordeling. På arket vælges middelværdi og spredning."
SA(588, 1) = "Excel sheet that plots normal distribution. Choose mean and standard deviation on the sheets"
SA(588, 2) = ""

SA(589, 0) = ChrW(&H3C7) & ChrW(&HB2) & "-Fordeling"
SA(589, 1) = ChrW(&H3C7) & ChrW(&HB2) & "-distribution"
SA(589, 2) = ""

SA(590, 0) = "Excel-ark med graf over " & ChrW(&H3C7) & ChrW(&HB2) & "-Fordeling. På arket vælges antal frihedsgrader og værdier kan aflæses."
SA(590, 1) = "Excel sheet that plots " & ChrW(&H3C7) & ChrW(&HB2) & "-distribution. On the sheet choose degrees of freedom."
SA(590, 2) = ""

SA(591, 0) = "Test"
SA(591, 1) = "Test"
SA(591, 2) = ""

SA(592, 0) = "Forskellige statistiske test"
SA(592, 1) = "Statistical tests"
SA(592, 2) = ""

SA(593, 0) = "Binomialtest"
SA(593, 1) = "Binomialtest"
SA(593, 2) = ""

SA(594, 0) = "Excelark hvor man kan udføre binomialtest. Samme ark som der er binomialfordeling på."
SA(594, 1) = "Excel sheet for binomial test. Same Excel workbook as Binomial distribution, but another sheet."
SA(594, 2) = ""

SA(595, 0) = "" & ChrW(&H3C7) & ChrW(&HB2) & "-test"
SA(595, 1) = "" & ChrW(&H3C7) & ChrW(&HB2) & "-test"
SA(595, 2) = ""

SA(596, 0) = "Udfører " & ChrW(&H3C7) & ChrW(&HB2) & "-test på n x n krydstabel. Marker tallene i krydstabellen."
SA(596, 1) = "Performs " & ChrW(&H3C7) & ChrW(&HB2) & "-test on n x n table. Select the numbers in the table."
SA(596, 2) = ""

SA(597, 0) = "Goodness of fit"
SA(597, 1) = "Goodness of fit"
SA(597, 2) = ""

SA(598, 0) = "Udfører goodness of fit test. (Chi i anden test) For om to fordelinger er ens."
SA(598, 1) = "Performs goodness of fit test. (Chi square test) to see if two distributions are the same."
SA(598, 2) = ""

SA(599, 0) = "Simulering"
SA(599, 1) = "Statistical simulation"
SA(599, 2) = ""

SA(600, 0) = "Ark der kan permutere(omrøre) og gemme en række teststørrelser."
SA(600, 1) = "Excel sheet with added functionality that can permute and save result. Includes example of use."
SA(600, 2) = ""

SA(601, 0) = "Ny ligning (Alt+M)"
SA(601, 1) = "New equation (Alt+M)"
SA(601, 2) = ""

SA(602, 0) = "Indsætter nyt matematikfelt."
SA(602, 1) = "Insert the math input box"
SA(602, 2) = ""

SA(603, 0) = "Ny nummeret ligning (Alt+M Alt+M)"
SA(603, 1) = "New numbered equation (Alt+M Alt+M)"
SA(603, 2) = ""

SA(604, 0) = "Indsætter nyt matematikfelt der automatisk bliver nummeret. Tast dobbelt Alt+M for at indsætte nummeret ligning. Dette skal dog først aktiveres under indstillinger."
SA(604, 1) = "Inserts new automatically numbered equation. Doubletap Alt+M to insert numbered equation. This must be activated under settings | numbered equations."
SA(604, 2) = ""

SA(605, 0) = "Ny nummeret ligning der kan refereres"
SA(605, 1) = "New named numbered equation that can be referenced."
SA(605, 2) = ""

SA(606, 0) = "Indsætter nyt matematikfelt der automatisk bliver nummeret og som kan refereres."
SA(606, 1) = "Inserts new numbered equation that you can name for easy referencing later. If the number of the equation later changes the reference changes as well."
SA(606, 2) = ""

SA(607, 0) = "Indsæt reference til ligning"
SA(607, 1) = "Insert reference to equation. "
SA(607, 2) = ""

SA(608, 0) = "Indsætter reference til ligning ud fra tidligere angivet navn"
SA(608, 1) = "Insert the equation number of a previously named equation. If the number of the equation later changes the reference changes as well."
SA(608, 2) = ""

SA(609, 0) = "Sæt ligning til bestemt nr."
SA(609, 1) = "Set equation to fixed number"
SA(609, 2) = ""

SA(610, 0) = "En ligning der har fået et bestemt nr kan sættes til et specifikt nr."
SA(610, 1) = "An equation already numbered can be set to a specific number"
SA(610, 2) = ""

SA(611, 0) = "Indsæt en ny ligningssektion."
SA(611, 1) = "Inserts a new equation section. "
SA(611, 2) = ""

SA(612, 0) = " Dette vil øge det første nr. i ligningsnummereringen. Hvis forrige er (2.5) så bliver den næste (3.1). Kan indsættes i overskrifter som sektionsnumre. Kan dog også indsættes skjult."
SA(612, 1) = "Will forward the first number in equation numbering. if previous is (2.5) then the next will be (3.1). Can be inserted into headings as heading number or inserted hidden."
SA(612, 2) = ""

SA(613, 0) = "Opdater alle ligningsnnumre i dokumentet."
SA(613, 1) = "Update all equation numbers in document"
SA(613, 2) = ""

SA(614, 0) = "Sker automatisk når en ny ligning indsættes, men det kan være nødvendigt at trykke på opdater-knappen når en ligning er blevet slettet."
SA(614, 1) = "Typically needs to be run when equations have been deleted. Inserting new numbered equations will force an update though."
SA(614, 2) = ""

SA(615, 0) = "Symboler"
SA(615, 1) = "Symbols"
SA(615, 2) = ""

SA(616, 0) = "Forskellige metoder til manipulation af symboler."
SA(616, 1) = "Different methods for manipulating symbols"
SA(616, 2) = ""

SA(617, 0) = "Erstat til prikgangetegn"
SA(617, 1) = "Replace multiplication signs"
SA(617, 2) = ""

SA(618, 0) = "Erstatter alle * gangetegn med almindelige prikgangetegn " & ChrW(183) & ""
SA(618, 1) = "replace all * multiplications signs with " & ChrW(183) & ""
SA(618, 2) = ""

SA(619, 0) = "Konverter ligninger til Latex"
SA(619, 1) = "Convert equations to LaTex. (Shortcut alt+T converts a single equation)"
SA(619, 2) = ""

SA(620, 0) = "Giver mulighed for at konvertere alle eller udvalgte ligninger i dokumentet til Latex. Giver også mulighed for at at generere Latex fil som pdf, hvis MikTex er installeret. En enkelt ligning kan konverteres til LaTex med genvejen Alt+T."
SA(620, 1) = "Convert some or all equations to LaTex in a document. Can also use MikTex to generate a complete LaTex document."
SA(620, 2) = ""

SA(621, 0) = "Konverter til Latex (Alt+T)"
SA(621, 1) = "Convert to LaTex (Alt+T)"
SA(621, 2) = ""

SA(622, 0) = "Konverterer en ligning til Latex. Omslutningstegnene vælges på hovedsiden for latex-konvertering. Brug også genvejen Alt+T."
SA(622, 1) = "Converts one equation to LaTex. Embed characters can be chosen on the LaTex form. Alt+T shortcut also convert one eqaution."
SA(622, 2) = ""

SA(623, 0) = "Erstat til stjernegangetegn"
SA(623, 1) = "Replace to start multiplication"
SA(623, 2) = ""

SA(624, 0) = "Erstatter alle " & ChrW(183) & " gangetegn til * gangetegn."
SA(624, 1) = "Replaces all " & ChrW(183) & " multiplication signs to * multiplication signs."
SA(624, 2) = ""

SA(625, 0) = "Punktum som decimalseparator"
SA(625, 1) = "Dot as decimalseparator"
SA(625, 2) = ""

SA(626, 0) = "Erstatter alle , (komma) i matematikfelter med . (punktum) og alle ;(semikolon) med ,(komma)"
SA(626, 1) = "Replaces all , (commas) in math inputboxes with . (dot) and all ;(semicolon) with ,(commas)"
SA(626, 2) = ""

SA(627, 0) = "Komma som decimalseparator"
SA(627, 1) = "Comma as decimal separator"
SA(627, 2) = ""

SA(628, 0) = "Erstatter alle . (punktum) i matematikfelter med ,(komma) og alle ,(komma) med ;(semikolon)"
SA(628, 1) = "Replaces all . (dots) in math input boxeswith ,(commas) and all ,(commas) with ;(semicolons)"
SA(628, 2) = ""

SA(629, 0) = "Figurer"
SA(629, 1) = "Math shapes, diagrams & images"
SA(629, 2) = ""

SA(630, 0) = "Åbner dokument med figurer som kan kopieres ind i det dokument. Trekanter, Monotoniforhold mm."
SA(630, 1) = "Opens document with typical Math shapes, diagrams etc. These can then be copied into your document and manipulated to your desire."
SA(630, 2) = ""

SA(631, 0) = "Tegn"
SA(631, 1) = "Draw"
SA(631, 2) = ""

SA(632, 0) = "Åbner Paint og indsætter det tegnede i dokumentet."
SA(632, 1) = "Opens Paint for drawing. The result is automatically inserted into Word."
SA(632, 2) = ""

SA(633, 0) = "Tabel"
SA(633, 1) = "Table"
SA(633, 2) = ""

SA(634, 0) = "WordMat kan manipulere data i lister. Fx kan man nemt lægge et tal til en masse tal i en liste. Typisk har man dog data i tabeller i Word, men her er funktioner der kan omdanne lister til tabeller og omvendt."
SA(634, 1) = "WordMat can manipulate data in list. Example: You can easily add a number to all elements in a list. Data is often in a table in Word though. Here are functions to convert list til tables and reverse."
SA(634, 2) = ""

SA(635, 0) = "Indsæt Excelark"
SA(635, 1) = "Insert Excel sheet"
SA(635, 2) = ""

SA(636, 0) = "Indsætter indlejret Excelark i dokumentet. Kan åbnes ved at dobbeltklikke."
SA(636, 1) = "Inserts empty embedded Excel sheet. Double click to open and edit."
SA(636, 2) = ""

SA(637, 0) = "Konverter til liste"
SA(637, 1) = "Convert to list"
SA(637, 2) = ""

SA(638, 0) = "Konverterer en Tabel til en liste i et matematikfelt, som så kan manipuleres i WordMat. Eksempel: [1, 2, 3]+1 = [2, 3, 4] Se manual for mere info"
SA(638, 1) = "Convert a table to a list. This list can then manipulated using WordMat. Example: [1, 2, 3]+1 = [2, 3, 4] Se manual for more info"
SA(638, 2) = ""

SA(639, 0) = "Konverter til Tabel"
SA(639, 1) = "Convert to table"
SA(639, 2) = ""

SA(640, 0) = "Konverterer en liste i et matematikfelt eller bare skrevet, til en tabel. f.eks.  (1 , 2),(3 , 4),(5 , 6)"
SA(640, 1) = "Converts a list to a table. The list can be in a math box or written in the document and selected such as (1 , 2),(3 , 4),(5 , 6)"
SA(640, 2) = ""

SA(641, 0) = "Trekantsløser"
SA(641, 1) = "Triangle solver"
SA(641, 2) = ""

SA(642, 0) = "Indtast 3 sider/vinkler og få beregnet de resterende sider/vinkler samt indsat trekant med de rigtige dimensioner."
SA(642, 1) = "Enter 3 sides/angles. The triangle solver computes the remaining sides/angles and generates an image of the triangle."
SA(642, 2) = ""

SA(643, 0) = "Word manual"
SA(643, 1) = "Word Manual"
SA(643, 2) = ""

SA(644, 0) = "Åbner dokument med hjælp."
SA(644, 1) = "Opens manual as Word document"
SA(644, 2) = ""

SA(645, 0) = "Webmanual"
SA(645, 1) = "Web manual"
SA(645, 2) = ""

SA(646, 0) = "Åbner Hjemmesidemanual. Har bedre navigation og videoer end Word-manualen, men ikke komplet endnu."
SA(646, 1) = "Opens the online web manual. Has better navigation and videos than the Word manual, but a work in progress."
SA(646, 2) = ""

SA(647, 0) = "Maxima manual"
SA(647, 1) = "Maxima manual"
SA(647, 2) = ""

SA(648, 0) = "Maxima er det CAS-program der udfører mange af de matematiske funktioner i WordMat. Maxima har indbygget en masse funktioner som ikke umiddelbart er synlige i WordMat, men de kan udføres hvis man kender funktionsnavnene. Her kan du åbne en oversigt over maxima-kommandoer."
SA(648, 1) = "Maxima is the CAS-program that does most of the math calculations in WordMat. Maxima has a lot of builtin functions which are not immediate visible from the WordMat user interface, but can be executed if you know the function names. This shows the Maxima help where you can see documentation for all the maxima functions."
SA(648, 2) = ""

SA(649, 0) = "Om"
SA(649, 1) = "About"
SA(649, 2) = ""

SA(650, 0) = "Hvem lavede WordMat?"
SA(650, 1) = "Who made WordMat?"
SA(650, 2) = ""

SA(651, 0) = "Check for opdateringer"
SA(651, 1) = "Check for updates"
SA(651, 2) = ""

SA(652, 0) = "Check automatisk på nettet om der er kommet opdatering til WordMat."
SA(652, 1) = "Automatically check for updates."
SA(652, 2) = ""

SA(653, 0) = "Genvejsoversigt"
SA(653, 1) = "Keyboard shortcuts"
SA(653, 2) = ""

SA(654, 0) = "Oversigt over hvilke tastaturgenveje der er i WordMat og genveje til matematikfelter i Word."
SA(654, 1) = "Shows a list of keyboard shortcut to activate WordMat commands and shortcuts for entering math in the math input boxes in Word."
SA(654, 2) = ""

SA(655, 0) = "LaTex skabelon"
SA(655, 1) = "LaTex template"
SA(655, 2) = ""

SA(656, 0) = "LaTex skabelon"
SA(656, 1) = "LaTex template"
SA(656, 2) = ""

SA(657, 0) = "Åbner Worddokument der er en næsten tro kopi af LaTex dokument. "
SA(657, 1) = "Opens a Word document which is styled like a LaTex document, but as easy to write as a Word document."
SA(657, 2) = ""


SA(658, 0) = "Præcisionen er sat meget højt. Beregningstiden kan blive meget lang eller fejle. Prøv evt. igen med en lavere indstilling."
SA(658, 1) = "Precisionen is set very high. The calculation time can be very long or fail. If so try again with a lower setting."
SA(658, 2) = ""

SA(659, 0) = "WordMat har en udmærket funktion til konvertering til LaTex, men der er også mulighed for at benytte den som er indbygget i Maxima"
SA(659, 1) = "WordMat convert to LaTex very well, but there is also the option to use Maxima to convert to LaTex."
SA(659, 2) = ""

SA(660, 0) = "Enheder skal normalt ikke skrives med kursiv. Med denne funktion forsøger WordMat at identificere enheder og indsætter Latex-kode der skriver enheder i normal skrift"
SA(660, 1) = "Unit are normally not written in italic as variables. With this function WordMat tries to identify units and insert LaTex code to write units in normal font."
SA(660, 2) = ""

SA(661, 0) = "Luk"
SA(661, 1) = "Close"
SA(661, 2) = ""

SA(662, 0) = "Genererer rigtigt LaTex-dokument vha. MikTex, som skal være installeret. Bør kun bruges til simple dokumenter da ikke alt er understøttet."
SA(662, 1) = "Generates a true LaTex-document using MikTex, which must be installed. Only use for simple document as not all document objects are supported."
SA(662, 2) = ""

SA(663, 0) = "Viser hvordan LaTex outputtet vil blive forstået"
SA(663, 1) = "Shows how the LaTex output is interpreted"
SA(663, 2) = ""

SA(664, 0) = "Generer LaTexkoden og åbner i TexWorks så koden kan gennemgås og evt. justeres"
SA(664, 1) = "Generates LaTex code and opens in TexWorks where you can make adjustments."
SA(664, 2) = ""

SA(665, 0) = "Det eksisterende worddokument, "
SA(665, 1) = "The existing Word document"
SA(665, 2) = ""

SA(666, 0) = " ,overskrives. Er du sikker på du vil forsætte?"
SA(666, 1) = "is overwritten. Are you sure you wish to proceed?"
SA(666, 2) = ""

SA(667, 0) = "Nulstil Worddokument"
SA(667, 1) = "Reset Word document"
SA(667, 2) = ""

SA(668, 0) = "Dit Worddokument"
SA(668, 1) = "Your Word document"
SA(668, 2) = ""

SA(669, 0) = "er nu nulstillet"
SA(669, 1) = "is now reset."
SA(669, 2) = ""

SA(670, 0) = "Gennemført"
SA(670, 1) = "Completed"
SA(670, 2) = ""

SA(671, 0) = "Du skal genstarte Word før end sprogændringen træder i kraft alle steder"
SA(671, 1) = "You must restart Word for the language change to take effect everywhere."
SA(671, 2) = ""

SA(672, 0) = "Kan ikke tage log af værdier <=0"
SA(672, 1) = "Cant take log of values <= 0"
SA(672, 2) = ""

SA(673, 0) = "Regner"
SA(673, 1) = "Calculating"
SA(673, 2) = ""

SA(674, 0) = "Jeg regner"
SA(674, 1) = "I'm calculating"
SA(674, 2) = ""

SA(675, 0) = "Denne metode kræver et interval omkring en løsning. Der må kun være en løsning i intervallet. Se på grafen for at finde intervallet. Fordelen ved denne metode er at man ved hvilken løsningen der findes."
SA(675, 1) = "This method requires an interval where a solution is within the interval. Only one solution must be in the interval. Look at the graph to indentify the solutions."
SA(675, 2) = ""

SA(676, 0) = "Om intervalmetoden"
SA(676, 1) = "About interval method"
SA(676, 2) = ""

SA(677, 0) = "Newton-Raphson metoden kræver et startgæt på en løsning til ligningen. Det er ikke garanteret at der findes en løsning selvom der er en. Der findes max. 1 løsning. Kombiner evt. metoden med grafisk løsning for at finde startgæt."
SA(677, 1) = "Newton-Raphson method requires an inital guess at a solution. The method is not guaranteed to find a solution if there is one.Only 1 solution is found. Look at the graph to identify an initial guess."
SA(677, 2) = ""

SA(678, 0) = "Om Newton-Raphson"
SA(678, 1) = "About Newton-Raphson"
SA(678, 2) = ""

SA(679, 0) = "Der kan ikke tages Backup før end du har gemt dit dokument!"
SA(679, 1) = "You must save your document before Backup can be perfomed!"
SA(679, 2) = ""

SA(680, 0) = "Brug IKKE dette dokument, og luk det ikke. WordMat har brug for det til mellemregninger. Alt hvad der skrives i dokumentet vil blive slettet."
SA(680, 1) = "Do NOT edit this document or close or it. WordMat needs it for calculations. Anything you enter here will be deleted."
SA(680, 2) = ""

SA(681, 0) = "Problemer med filen 'normal.dotm', kan ødelægge tastaturgenvejene i WordMat. Problemet løses ved at slette filen, da der automatisk genereres en ny når Word genstarter. Filen kan dog ikke slettes mens Word er åben så du skal selv gøre det. " & vbCrLf & vbCrLf & "- Når du trykker OK åbnes den mappe hvor filen findes. " & vbCrLf & "- Gem dit arbejde og luk Word " & vbCrLf & "- Slet normal.dotm " & vbCrLf & "- Genstart Word og åben og luk WordMat indstillinger"
SA(681, 1) = "Problems with the file 'normal.dotm', can cause problems with keyboard shortcuts in WordMat. The problems can be resolved by deleting the file, since a new default file is automatically generated when Word restarts. The file cannot be deleted while Word is open. You must do it manually." & vbCrLf & vbCrLf & "- Press OK to open the folder where Normal.dotm is located. " & vbCrLf & "- Save your Work and close Word " & vbCrLf & "- delete normal.dotm " & vbCrLf & "- Open Word and open and close WordMat settings"
SA(681, 2) = ""

SA(682, 0) = "Beregning foretaget vha. GeoGebra"
SA(682, 1) = "Calculation performed by GeoGebra"
SA(682, 2) = ""

SA(683, 0) = "pil"
SA(683, 1) = "arrow"
SA(683, 2) = ""

SA(684, 0) = "Starter Safari"
SA(684, 1) = "Opening Safari"
SA(684, 2) = ""

SA(685, 0) = "Problemet kunne ikke løses. Vil du prøve med en anden metode eller CAS-motor?"
SA(685, 1) = "The problem copuld not be solved. Do you want to try another method or CAS-engine?"
SA(685, 2) = ""

SA(686, 0) = "CAS motor"
SA(686, 1) = "CAS engine"
SA(686, 2) = ""

SA(687, 0) = ChrW(&H70) & ChrW(&H302) & ChrW(177) & "2" & ChrW(183) & ChrW(&H221A) & "(" & ChrW(&H70) & ChrW(&H302) & ChrW(183) & "(1-" & ChrW(&H70) & ChrW(&H302) & ")/n)       Usikkerhed til 95% konfidensinterval"
SA(687, 1) = ChrW(&H70) & ChrW(&H302) & ChrW(177) & "2" & ChrW(183) & ChrW(&H221A) & "(" & ChrW(&H70) & ChrW(&H302) & ChrW(183) & "(1-" & ChrW(&H70) & ChrW(&H302) & ")/n)       Uncertainty"
SA(687, 2) = ""

'Problemet kunne ikke løses. Vil du prøve med en anden CAS motor?
End Sub

' i xml må der ikke være label attribut når der er  callback.
' prepare for new language by find/replace Case 3: s="" ->getLabel Case 3: s=""   vbrlf   Case 4: s=""

Public Property Get RibFormulae() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Formelsamling"
Case 3: s = "Formulario"
Case 4: s = ""
Case Else: s = "Formulae"
End Select
RibFormulae = s
End Property
Public Property Get RibMath() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Mat"
Case 3: s = "Matemáticas"
Case 4: s = ""
Case Else: s = "Math"
End Select
RibMath = s
End Property
Public Property Get RibPhysics() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fys"
Case 3: s = "Física"
Case 4: s = ""
Case Else: s = "Phys"
End Select
RibPhysics = s
End Property
Public Property Get RibChemistry() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Kemi"
Case 3: s = "Química"
Case 4: s = ""
Case Else: s = "Chem"
End Select
RibChemistry = s
End Property
Public Property Get RibPercentage() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Procent"
Case 3: s = "Porcentaje"
Case 4: s = ""
Case Else: s = "Percentage"
End Select
RibPercentage = s
End Property
Public Property Get RibFunctions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Funktioner"
Case 3: s = "Funciones"
Case 4: s = ""
Case Else: s = "Functions"
End Select
RibFunctions = s
End Property
Public Property Get RibGeometry() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Geometri"
Case 3: s = "Geometría"
Case 4: s = ""
Case Else: s = "Geometry"
End Select
RibGeometry = s
End Property
Public Property Get RibProbabilityShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sands."
Case 3: s = "Probabilidades"
Case 4: s = ""
Case Else: s = "Prob."
End Select
RibProbabilityShort = s
End Property
Public Property Get RibInfinitesimalShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "inf."
Case 3: s = "Infinitesimal"
Case 4: s = ""
Case Else: s = "inf."
End Select
RibInfinitesimalShort = s
End Property
Public Property Get RibVector() As String
Dim s As String
Select Case SprogNr
Case 1: s = "vektor"
Case 3: s = "Vector"
Case 4: s = ""
Case Else: s = "vector"
End Select
RibVector = s
End Property

Public Property Get RibSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indstillinger"
Case 3: s = "Configuración"
Case 4: s = ""
Case Else: s = "Settings"
End Select
RibSettings = s
End Property
Public Property Get RibSettingsShort() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indst."
Case 3: s = "Configuración"
Case 4: s = ""
Case Else: s = "Settings"
End Select
RibSettingsShort = s
End Property
Public Property Get RibSciNot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "vid"
Case 3: s = "NCi" 'Notación Científica
Case 4: s = ""
Case Else: s = "sci"
End Select
RibSciNot = s
End Property

Public Property Get RibBeregn() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Beregn"
Case 3: s = "Calcular"
Case 4: s = ""
Case Else: s = "Calculate"
End Select
RibBeregn = s
End Property

Public Property Get RibSolve() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs Ligning(er)"
Case 3: s = "Resolver ecuación(es)"
Case 4: s = ""
Case Else: s = "Solve equation(s)"
End Select
RibSolve = s
End Property

Public Property Get RibMaximaCommand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Maxima-kommando"
Case 3: s = "Maxima command"
Case 4: s = ""
Case Else: s = "Maxima command"
End Select
RibMaximaCommand = s
End Property
Public Property Get RibSolveNum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs ligning(er) vha. numerisk metode"
Case 3: s = "Resolver ecuación(es) usando métodos numéricos"
Case 4: s = ""
Case Else: s = "Solve equation(s) using numerical method"
End Select
RibSolveNum = s
End Property
Public Property Get RibEliminate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eliminer variable"
Case 3: s = "Eliminar variables"
Case 4: s = ""
Case Else: s = "Eliminate variables"
End Select
RibEliminate = s
End Property
Public Property Get TestTF() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Test sand/falsk"
Case 3: s = "Prueba Verdadero/falso"
Case 4: s = ""
Case Else: s = "Test true/false"
End Select
TestTF = s
End Property
Public Property Get RibSolveDE() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs differentialligning"
Case 3: s = "Resolver ecuación diferencial"
Case 4: s = ""
Case Else: s = "Solve differential equation"
End Select
RibSolveDE = s
End Property
Public Property Get RibDefinitions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
RibDefinitions = s
End Property
Public Property Get RibDefDelete() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Slet definitioner"
Case 3: s = "Borrar definiciones"
Case 4: s = ""
Case Else: s = "Delete definitions"
End Select
RibDefDelete = s
End Property
Public Property Get RibDefVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer variabel"
Case 3: s = "Definir variable"
Case 4: s = ""
Case Else: s = "Define variable"
End Select
RibDefVariable = s
End Property
Public Property Get RibDefFunction() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer funktion"
Case 3: s = "Definir función"
Case 4: s = ""
Case Else: s = "Define function"
End Select
RibDefFunction = s
End Property
Public Property Get RibDefEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer ligning"
Case 3: s = "Definir ecuación"
Case 4: s = ""
Case Else: s = "Define equation"
End Select
RibDefEquation = s
End Property
Public Property Get RibDefConstants() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definer fysiske konstanter"
Case 3: s = "Definir constantes físicas"
Case 4: s = ""
Case Else: s = "Define physical constants"
End Select
RibDefConstants = s
End Property

Public Property Get RibReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omskriv"
Case 3: s = "Reducir"
Case 4: s = ""
Case Else: s = "Reduce"
End Select
RibReduce = s
End Property
Public Property Get RibInfinitesimal() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Infinites."
Case 3: s = "Infinitos"
Case 4: s = ""
Case Else: s = "Infinites."
End Select
RibInfinitesimal = s
End Property
Public Property Get RibDif() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Differentier"
Case 3: s = "Diferenciar"
Case 4: s = ""
Case Else: s = "Differentiate"
End Select
RibDif = s
End Property
Public Property Get RibIntegrate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Stamfunktion"
Case 3: s = "Integrar"
Case 4: s = ""
Case Else: s = "Integrate"
End Select
RibIntegrate = s
End Property

Public Property Get RibPlotting() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graftegning"
Case 3: s = "Graficar"
Case 4: s = ""
Case Else: s = "Plotting"
End Select
RibPlotting = s
End Property
Public Property Get RibShowGraph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis Graf"
Case 3: s = "Mostrar gráfico"
Case 4: s = ""
Case Else: s = "Show Graph"
End Select
RibShowGraph = s
End Property
Public Property Get RibDirectionField() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Hældningsfelt"
Case 3: s = "Dirección del campo"
Case 4: s = ""
Case Else: s = "Direction field"     ' May be Field's direction?  Like magnetic field's direction...or vectorial field's direction
End Select
RibDirectionField = s
End Property
Public Property Get RibInsertGeoGebra() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt GeoGebra fil"
Case 3: s = "Insertar GeoGebra archivo"
Case 4: s = ""
Case Else: s = "Insert GeoGebra file"
End Select
RibInsertGeoGebra = s
End Property
Public Property Get Rib3DPlot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "3D Plot"
Case 3: s = "Gráfico 3D"
Case 4: s = ""
Case Else: s = "3D Plot"
End Select
Rib3DPlot = s
End Property
Public Property Get Rib3DRotate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omdrejningslegeme"
Case 3: s = "Volumen de revolución - Gráfico 3D rota en el eje x"
Case 4: s = ""
Case Else: s = "Volume of revolution - 3D Rotate Graph about x-axis"
End Select
Rib3DRotate = s
End Property
Public Property Get RibStatistics() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Statistik"
Case 3: s = "Estadística"
Case 4: s = ""
Case Else: s = "Statistics"
End Select
RibStatistics = s
End Property
Public Property Get RibUgrup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ugrupperede observationer"
Case 3: s = "Observaciones no agrupadas"
Case 4: s = ""
Case Else: s = "Ungrouped observations"
End Select
RibUgrup = s
End Property
Public Property Get RibGrup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grupperede observationer"
Case 3: s = "Observaciones agrupadas"
Case 4: s = ""
Case Else: s = "Grouped observations"
End Select
RibGrup = s
End Property
Public Property Get RibStickChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Pindediagram"
Case 3: s = "Gráfico de Barras"
Case 4: s = ""
Case Else: s = "Stick chart"
End Select
RibStickChart = s
End Property
Public Property Get RibHistogram() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Histogram"
Case 3: s = "Histograma"
Case 4: s = ""
Case Else: s = "Histogram"
End Select
RibHistogram = s
End Property
Public Property Get RibStepChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trappediagram"
Case 3: s = "Gráfico escalonado"
Case 4: s = ""
Case Else: s = "Step chart"
End Select
RibStepChart = s
End Property
Public Property Get RibCumChart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sumkurve"
Case 3: s = "Gráfico acumulativo"
Case 4: s = ""
Case Else: s = "Cumulative chart"
End Select
RibCumChart = s
End Property
Public Property Get RibBoxPlot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Boksplot"
Case 3: s = "Gráfico de cajas y conectores"
Case 4: s = ""
Case Else: s = "Box and whiskers plot"
End Select
RibBoxPlot = s
End Property
Public Property Get RibStatProb() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Statistik/sands."
Case 3: s = "Probabilidades y Estadística"
Case 4: s = ""
Case Else: s = "Statistics/prob."
End Select
RibStatProb = s
End Property
Public Property Get RibRegression() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Regression"
Case 3: s = "Regresión"
Case 4: s = ""
Case Else: s = "Regression"
End Select
RibRegression = s
End Property
Public Property Get RibInsertTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt tabel"
Case 3: s = "Insertar tabla"
Case 4: s = ""
Case Else: s = "Insert table"
End Select
RibInsertTable = s
End Property
Public Property Get RibLinRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Lineær"
Case 3: s = "Lineal"
Case 4: s = ""
Case Else: s = "Linear"
End Select
RibLinRegr = s
End Property
Public Property Get RibExpRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eksponentiel"
Case 3: s = "Exponencial"
Case 4: s = ""
Case Else: s = "Exponential"
End Select
RibExpRegr = s
End Property
Public Property Get RibPowRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Potens"
Case 3: s = "Potencias"
Case 4: s = ""
Case Else: s = "Power"
End Select
RibPowRegr = s
End Property
Public Property Get RibPolRegr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Polynomisk"
Case 3: s = "Polinomios"
Case 4: s = ""
Case Else: s = "Polynomial"
End Select
RibPolRegr = s
End Property
Public Property Get RibDistributions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fordelinger"
Case 3: s = "Distribuciones"
Case 4: s = ""
Case Else: s = "Distributions"
End Select
RibDistributions = s
End Property
Public Property Get RibBinomDist() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Binomialfordeling"
Case 3: s = "Distribución binomial"
Case 4: s = ""
Case Else: s = "Binomial distribution"
End Select
RibBinomDist = s
End Property
Public Property Get RibNormalDist() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Normalfordeling"
Case 3: s = "Distribución Normal"
Case 4: s = ""
Case Else: s = "Normal distribution"
End Select
RibNormalDist = s
End Property
Public Property Get RibChi2Dist() As String
Dim s As String
Select Case SprogNr
Case 1: s = VBA.ChrW(967) & VBA.ChrW(178) & "-fordeling"
Case 3: s = "Distribución " & VBA.ChrW(967) & VBA.ChrW(178)
Case 4: s = ""
Case Else: s = VBA.ChrW(967) & VBA.ChrW(178) & " distribution"
End Select
RibChi2Dist = s
End Property
Public Property Get RibTest() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Test"
Case 3: s = "Prueba"
Case 4: s = ""
Case Else: s = "Test"
End Select
RibTest = s
End Property
Public Property Get RibBinomTest() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Binomialtest"
Case 3: s = "Prueba binomial"
Case 4: s = ""
Case Else: s = "Binomial test"
End Select
RibBinomTest = s
End Property
Public Property Get RibChi2Test() As String
Dim s As String
Select Case SprogNr
Case 1: s = VBA.ChrW(967) & VBA.ChrW(178) & "-test for sammenhæng"
Case 3: s = "Prueba " & VBA.ChrW(967) & VBA.ChrW(178) & ""
Case 4: s = ""
Case Else: s = VBA.ChrW(967) & VBA.ChrW(178) & " test for independence"
End Select
RibChi2Test = s
End Property
Public Property Get RibSimulation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Simulering"
Case 3: s = "Simulación"
Case 4: s = ""
Case Else: s = "Simulation"
End Select
RibSimulation = s
End Property
Public Property Get RibGroup() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grupp" & VBA.ChrW(233) & "r"
Case 3: s = "Grupo" & VBA.ChrW(233) & "r"""
Case 4: s = ""
Case Else: s = "Group"
End Select
RibGroup = s
End Property
Public Property Get RibDiverse() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Diverse"
Case 3: s = "Otros"          ' The real translations is "Diversos" but in the latin academic mind is well accepted 'Others'
Case 4: s = ""
Case Else: s = "Others"
End Select
RibDiverse = s
End Property
Public Property Get RibNewEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ny ligning"
Case 3: s = "Nueva ecuación"
Case 4: s = ""
Case Else: s = "New Equation"
End Select
RibNewEquation = s
End Property
Public Property Get RibSymbols() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Symboler"
Case 3: s = "Símbolos"
Case 4: s = ""
Case Else: s = "Symbols"
End Select
RibSymbols = s
End Property
Public Property Get RibMultsign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Prik gangetegn"
Case 3: s = "Signo de multiplicación"
Case 4: s = ""
Case Else: s = "Dot multiplication sign"
End Select
RibMultsign = s
End Property
Public Property Get RibFigurs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Figurer"
Case 3: s = "Figuras"
Case 4: s = ""
Case Else: s = "Shapes"
End Select
RibFigurs = s
End Property
Public Property Get RibDraw() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tegn"
Case 3: s = "Dibujar"
Case 4: s = ""
Case Else: s = "Draw"
End Select
RibDraw = s
End Property
Public Property Get RibTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tabel"
Case 3: s = "Tabla"
Case 4: s = ""
Case Else: s = "Table"
End Select
RibTable = s
End Property
Public Property Get RibInsertExcel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Excel"
Case 3: s = "Insertar Excel"
Case 4: s = ""
Case Else: s = "Insert Excel"
End Select
RibInsertExcel = s
End Property
Public Property Get RibTableToList() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Tabel " & VBA.ChrW(8594) & " Liste"
Case 3: s = "Tabla " & VBA.ChrW(8594) & " Lista "
Case 4: s = ""
Case Else: s = "Table " & VBA.ChrW(8594) & " List"
End Select
RibTableToList = s
End Property
Public Property Get RibListToTable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Liste " & VBA.ChrW(8594) & " Tabel"
Case 3: s = ""
Case 4: s = ""
Case Else: s = "List " & VBA.ChrW(8594) & " Table"
End Select
RibListToTable = s
End Property
Public Property Get RibTriangle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trekant"
Case 3: s = "Triángulo"
Case 4: s = ""
Case Else: s = "Triangle"
End Select
RibTriangle = s
End Property

''''''''''''''''''''''''''''''''''''
''''''''''''''Forms'''''''''''''''''
''''''''''''''''''''''''''''''''''''

''''''''''''''MaximaSettings'''''''''
Public Property Get CASSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "CAS indstillinger"
Case 3: s = "Configuración de CAS"
Case 4: s = ""
Case Else: s = "CAS settings"
End Select
CASSettings = s
End Property
Public Property Get Notation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Notation"
Case 3: s = "Notación"
Case 4: s = ""
Case Else: s = "Notation"
End Select
Notation = s
End Property
Public Property Get Graph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graf"
Case 3: s = "Gráfico"
Case 4: s = ""
Case Else: s = "Graph"
End Select
Graph = s
End Property
Public Property Get Advanced() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Avanceret"
Case 3: s = "Avanzado"
Case 4: s = ""
Case Else: s = "Advanced"
End Select
Advanced = s
End Property
Public Property Get ScientificNotation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Videnskabelig notation"
Case 3: s = "Notación científica"
Case 4: s = ""
Case Else: s = "Scientific notation"
End Select
ScientificNotation = s
End Property
Public Property Get Complex() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Komplekse tal"
Case 3: s = "Números complejos"
Case 4: s = ""
Case Else: s = "Complex numbers"
End Select
Complex = s
End Property
Public Property Get SignificantFigures() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Betydende cifre"
Case 3: s = "Cifras significativas"
Case 4: s = ""
Case Else: s = "Significant figures"
End Select
SignificantFigures = s
End Property
Public Property Get Units() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Enheder"
Case 3: s = "Unidades"
Case 4: s = ""
Case Else: s = "Units"
End Select
Units = s
End Property
Public Property Get PolarNotation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Polær notation output"
Case 3: s = "Salida en notación polar"
Case 4: s = ""
Case Else: s = "Polar notation output"
End Select
PolarNotation = s
End Property
Public Property Get Degrees() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Grader"
Case 3: s = "Grados"
Case 4: s = ""
Case Else: s = "Degrees"
End Select
Degrees = s
End Property
Public Property Get Radians() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Radianer"
Case 3: s = "Radianes"
Case 4: s = ""
Case Else: s = "Radians"
End Select
Radians = s
End Property
Public Property Get AngleUnit() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vinkelenhed"
Case 3: s = "Unidad angular"
Case 4: s = ""
Case Else: s = "Angle unit"
End Select
AngleUnit = s
End Property
Public Property Get ShowSolutionConditions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis løsningsbetingelser (beta)"
Case 3: s = "Ver condiciones de solución"
Case 4: s = ""
Case Else: s = "Show solution conditions (beta)"
End Select
ShowSolutionConditions = s
End Property
Public Property Get Exact() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eksakt"
Case 3: s = "Exacto"
Case 4: s = ""
Case Else: s = "Exact"
End Select
Exact = s
End Property
Public Property Get Numeric() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Numerisk"
Case 3: s = "Numérico"
Case 4: s = ""
Case Else: s = "Numeric"
End Select
Numeric = s
End Property
Public Property Get Auto() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Automatisk"
Case 3: s = "Automático"
Case 4: s = ""
Case Else: s = "Automatic"
End Select
Auto = s
End Property
Public Property Get HighPrecision() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Meget høj præcision (beta)"
Case 3: s = "Muy alta precisión"
Case 4: s = ""
Case Else: s = "Very high precision (beta)"
End Select
HighPrecision = s
End Property
Public Property Get InsertExplanation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt forklaring"
Case 3: s = "Insertar explicación"
Case 4: s = ""
Case Else: s = "Insert explanation"
End Select
InsertExplanation = s
End Property
Public Property Get InsertMaximaCommand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Maxima kommando"
Case 3: s = "Insertar comando Maxima"
Case 4: s = ""
Case Else: s = "Insert Maxima command"
End Select
InsertMaximaCommand = s
End Property
Public Property Get Separators() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Separatorer"
Case 3: s = "Separadores"
Case 4: s = ""
Case Else: s = "Separators"
End Select
Separators = s
End Property
Public Property Get List() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Liste"
Case 3: s = "Lista"
Case 4: s = ""
Case Else: s = "List"
End Select
List = s
End Property
Public Property Get Multsign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gangetegn"
Case 3: s = "Signos de multiplicación"
Case 4: s = ""
Case Else: s = "Mult.sign"
End Select
Multsign = s
End Property
Public Property Get Dot() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Prik"
Case 3: s = "Punto"
Case 4: s = ""
Case Else: s = "Dot"
End Select
Dot = s
End Property
Public Property Get SolutionType() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsningsangivelse"
Case 3: s = "Solución tipo"
Case 4: s = ""
Case Else: s = "Solution type"
End Select
SolutionType = s
End Property
Public Property Get SolutionBoolean() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Boolsk"
Case 3: s = "Booleano"
Case 4: s = ""
Case Else: s = "Boolean"
End Select
SolutionBoolean = s
End Property
Public Property Get SolutionSet() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsningsmængde"
Case 3: s = "Conjunto de soluciones"
Case 4: s = ""
Case Else: s = "Solution Set"
End Select
SolutionSet = s
End Property
Public Property Get IndexSubscript() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sænket skrift"
Case 3: s = "Subíndice"
Case 4: s = ""
Case Else: s = "Subscript"
End Select
IndexSubscript = s
End Property
Public Property Get JustText() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Er bare tekst som er del af variabelnavn"
Case 3: s = "Es simplemente el texto que forma parte del nombre de la variable"
Case 4: s = ""
Case Else: s = "Is just text part of variable name"
End Select
JustText = s
End Property
Public Property Get IsIndex() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Er liste/matrix indeks"
Case 3: s = "Indice de lista / matriz"
Case 4: s = ""
Case Else: s = "Is list/matrix index"
End Select
IsIndex = s
End Property
Public Property Get StandardGraph() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Standard grafprogram"
Case 3: s = "Programa de gráficos"
Case 4: s = ""
Case Else: s = "Standard graph program"
End Select
StandardGraph = s
End Property
Public Property Get ExcelEmbed() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt Excelark indlejret"
Case 3: s = "Insertar hoja Excel"
Case 4: s = ""
Case Else: s = "Embed Excel sheets"
End Select
ExcelEmbed = s
End Property
Public Property Get OutputUnits() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Output enheder"
Case 3: s = "Unidades de salida"
Case 4: s = ""
Case Else: s = "Output units"
End Select
OutputUnits = s
End Property
Public Property Get UnitHelp() As String
Dim s As String
Select Case SprogNr
Case 1
s = "Angiv hvilke enheder som skal være i output adskilt ad komma" & VbCrLfMac & "eksempler:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ, Timer" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "Hvis intet angives returneres alle enheder som base SI-enheder eller afledte SI-enheder uden præfiks."
Case 3: s = "Establece unidades de salida " & VbCrLfMac & " separadas por coma ejemplos:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "Si no se indican las unidades se devuelve en unidades básicas del SI o unidades del SI derivadas sin prefijo."
Case 4: s = ""
Case Else:
s = "State output units separated by comma" & VbCrLfMac & "examples:" & VbCrLfMac & VbCrLfMac
s = s & "cm , kJ" & VbCrLfMac
s = s & "eV , nm, mV" & VbCrLfMac
s = s & "aJ , AU, Bq" & VbCrLfMac & VbCrLfMac
s = s & "If nothing is stated units are returned in basic SI-units or SI-derived units without prefix."
End Select
UnitHelp = s
End Property
Public Property Get UnitExamples() As String
Dim s As String
Select Case SprogNr
Case 1:
s = "Eksempler på understøttede enheder:" & VbCrLfMac
s = s & "(Se manual for flere)" & VbCrLfMac & VbCrLfMac
s = s & "Længde: cm , AU, ly, pc" & VbCrLfMac
s = s & "Masse: u , G, mg" & VbCrLfMac
s = s & "Tid: min , Timer, dage, år" & VbCrLfMac
s = s & "Frekvens: Hz , Bq" & VbCrLfMac
s = s & "Energi: kWh , eV, aJ" & VbCrLfMac
s = s & "Tryk: Bar , mmHg, atm, torr" & VbCrLfMac
Case 3: s = "Ejemplos de unidades compatibles:" & VbCrLfMac
s = s & "(Para mayor información consulte el manual)" & VbCrLfMac & VbCrLfMac
s = s & "Longitud: cm , AU, ly, pc" & VbCrLfMac
s = s & "Masa: u , G, mg" & VbCrLfMac
s = s & "Frecuencia: Hz , Bq" & VbCrLfMac
s = s & "Energía: kWh , eV, aJ" & VbCrLfMac
s = s & "Presión: Bar , mmHg, atm, torr" & VbCrLfMac
Case 4: s = ""
Case Else:
s = "Examples of supported units:" & VbCrLfMac
s = s & "(See manual for more)" & VbCrLfMac & VbCrLfMac
s = s & "Length: cm , AU, ly, pc" & VbCrLfMac
s = s & "Mass: u , G, mg" & VbCrLfMac
s = s & "Frequency: Hz , Bq" & VbCrLfMac
s = s & "Energy: kWh , eV, aJ" & VbCrLfMac
s = s & "Pressure: Bar , mmHg, atm, torr" & VbCrLfMac
End Select
UnitExamples = s
End Property
Public Property Get TrigEquations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trignonometriske ligninger"
Case 3: s = "Ecuaciones trigonométricas"
Case 4: s = ""
Case Else: s = "Trigonometric equations"
End Select
TrigEquations = s
End Property
Public Property Get AllSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle løsninger"
Case 3: s = "Todas las soluciones"
Case 4: s = ""
Case Else: s = "All solutions"
End Select
AllSolutions = s
End Property
Public Property Get OnlyOneSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Kun 1 løsning"
Case 3: s = "Sólo una solución"
Case 4: s = ""
Case Else: s = "Only 1 solution"
End Select
OnlyOneSolution = s
End Property
Public Property Get AutoStart() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Start automatisk med Word"
Case 3: s = "Iniciar automáticamente con Word"
Case 4: s = ""
Case Else: s = "Start automatic with Word"
End Select
AutoStart = s
End Property
Public Property Get AutoUpdate() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Check automatisk for opdateringer"
Case 3: s = "Comprobar actualizaciones automáticamente"
Case 4: s = ""
Case Else: s = "Automatically check for updates"
End Select
AutoUpdate = s
End Property
Public Property Get GenerateShortcuts() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Opret tastaturgenveje"
Case 3: s = "Crear atajos de teclado"
Case 4: s = ""
Case Else: s = "Keyboard shortcuts"
End Select
GenerateShortcuts = s
End Property
Public Property Get RestartMaxima() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genstart Maxima"
Case 3: s = "Reiniciar Maxima"
Case 4: s = ""
Case Else: s = "Restart Maxima"
End Select
RestartMaxima = s
End Property
Public Property Get RestartWordMat() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genstart WordMat"
Case 3: s = "Reiniciar WordMat"
Case 4: s = ""
Case Else: s = "Restart WordMat"
End Select
RestartWordMat = s
End Property
Public Property Get CalculationCount() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Antal beregninger"
Case 3: s = "Número de cálculos"
Case 4: s = ""
Case Else: s = "No. of calculations"
End Select
CalculationCount = s
End Property
Public Property Get ResetFigurDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil figurdokument"
Case 3: s = "Restablecer documento figura"
Case 4: s = ""
Case Else: s = "Reset Figur doc"
End Select
ResetFigurDoc = s
End Property
Public Property Get ResetMathDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil Matdokument"
Case 3: s = "Restablecer documento matemático"
Case 4: s = ""
Case Else: s = "Reset math doc"
End Select
ResetMathDoc = s
End Property
Public Property Get ResetPhysicsDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil fysikdokument"
Case 3: s = "Restablecer documento físico"
Case 4: s = ""
Case Else: s = "Reset physics doc"
End Select
ResetPhysicsDoc = s
End Property
Public Property Get ResetChemistryDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil kemidokument"
Case 3: s = "Restablecer documento químico"
Case 4: s = ""
Case Else: s = "Reset chemistry doc"
End Select
ResetChemistryDoc = s
End Property
Public Property Get SignDepends() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Resultatet af beregningen afhænger af fortegnet på følgende"
Case 3: s = "El resultado del cálculo depende del signo de"
Case 4: s = ""
Case Else: s = "The result of the calculation depends on the sign of"
End Select
SignDepends = s
End Property
Public Property Get StateSign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Angiv fortegn"
Case 3: s = "Signo de estado"
Case 4: s = ""
Case Else: s = "State sign"
End Select
StateSign = s
End Property
Public Property Get DifferentialEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Differentialligning"
Case 3: s = "Ecuación Diferencial"
Case 4: s = ""
Case Else: s = "Differential equation"
End Select
DifferentialEquation = s
End Property
Public Property Get SolveDE() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løsning af differentialligning"
Case 3: s = "Resolviendo ecuación diferencial"
Case 4: s = ""
Case Else: s = "Solving Differential equation"
End Select
SolveDE = s
End Property
Public Property Get IndepVar() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Uafh. var"
Case 3: s = "Variable independiente"
Case 4: s = ""
Case Else: s = "Indep. var"
End Select
IndepVar = s
End Property
Public Property Get DepVar() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Afh. var/funktion"
Case 3: s = "Variable Dependiente"
Case 4: s = ""
Case Else: s = "Dep. var/function"
End Select
DepVar = s
End Property
Public Property Get StartCond() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Startbetingelse"
Case 3: s = "Condición inicial"
Case 4: s = ""
Case Else: s = "Initial condition"
End Select
StartCond = s
End Property
Public Property Get MaximaError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Maxima fejlmeddelse"
Case 3: s = "Error de Maxima"
Case 4: s = ""
Case Else: s = "Maxima error"
End Select
MaximaError = s
End Property
Public Property Get SyntaxError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Syntaksfejl"
Case 3: s = "Error de sintaxis"
Case 4: s = ""
Case Else: s = "Syntax error"
End Select
SyntaxError = s
End Property
Public Property Get SyntaxErrorLong() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en syntaksfejl. Dvs. WordMat har ikke forstået dit input."
Case 3: s = "Error de sintaxis. WordMat no entiende su entrada"
Case 4: s = ""
Case Else: s = "Syntax error. WordMat did not understand your input."
End Select
SyntaxErrorLong = s
End Property
Public Property Get Missing() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der mangler en"
Case 3: s = "Está perdiendo una"
Case 4: s = ""
Case Else: s = "You're missing a"
End Select
Missing = s
End Property
Public Property Get TooMany() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er for mange"
Case 3: s = "Hay demasiados"
Case 4: s = ""
Case Else: s = "There is too many"
End Select
TooMany = s
End Property
Public Property Get LispError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en fejl i programafviklingen."
Case 3: s = "Ocurrió un error en el programa"
Case 4: s = ""
Case Else: s = "An error occured in the code."
End Select
LispError = s
End Property
Public Property Get DivisionByZero() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er sket en division med 0."
Case 3: s = "División por cero"
Case 4: s = ""
Case Else: s = "Division by zero"
End Select
DivisionByZero = s
End Property
Public Property Get DefError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der skete en fejl. Det skyldes formentlig en fejl i en af dine definitioner." & VbCrLfMac & VbCrLfMac & "Måske kan du se ud af nedenstående hvor den er."
Case 3: s = "Ocurrió un error. El error, probablemente, lo causó una de sus definiciones." & VbCrLfMac & VbCrLfMac & "Tal vez usted puede detectar el error de la información abajo."
Case 4: s = ""
Case Else: s = "An errror occured. The error was probably caused by one of your definitions." & VbCrLfMac & VbCrLfMac & "Maybe you can spot the error from the info below."
End Select
DefError = s
End Property
Public Property Get DefError3() As String
Dim s As String
Select Case SprogNr
'Case 1: s = "Du har en fejl i en af dine definitioner! " & VbCrLfMac & " kig på listen over definitioner nedenfor og se om du kan lokalisere fejlen. Alternativt indsæt en  'slet definitioner' kommando og se om det så virker. "
Case 1: s = "Du har en skrevet noget som WordMat ikke kan forstå."
Case 3: s = "Hay un error en una de sus definiciones! " & VbCrLfMac & " Vea la lista de definiciones a continuación y compruebe si puede localizar el error. También puede tratar de insertar el comando 'delete def' para depurar."
Case 4: s = ""
Case Else: s = "You have written something which WordMat does not understand."
'Case Else: s = "There is an error in one of your definitions! " & VbCrLfMac & " Check the list of definitions below and check if you can locate the error. You could also try to insert a  'delete def' command to debug."
End Select
DefError3 = s
End Property
Public Property Get DefError2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Bemærk at en syntaksfejl godt kan skyldes en fejl i en af dine definitioner, men det er ikke sikkert. Kig på listen over definitioner nedenfor og se om du kan lokalisere fejlen. Alternativt kan du prøve at indsætte en 'slet definitioner' kommando og se om det så virker."
Case 3: s = "Un error de sintaxis puede haber sido causado por un error en una de sus definiciones. Vea la lista de definiciones abajo y verifique si puede localizar el error. También puede tratar de insertar el comando 'delete def' para depurar."
Case 4: s = ""
Case Else: s = "A syntax error could be caused by an error in one of your definitions. Check the list of definitions below and check if you can locate an error.  You could also try to insert a  'delete def' command to debug"
End Select
DefError2 = s
End Property
Public Property Get IllegalSymbol() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er et ulovligt tegn der er placeret ved"
Case 3: s = "Hay un carácter ilegal que se encuentra en"
Case 4: s = ""
Case Else: s = "An illegal character has been placed by"
End Select
IllegalSymbol = s
End Property
Public Property Get SolveEquation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs Ligning"
Case 3: s = "Resolver ecuación"
Case 4: s = ""
Case Else: s = "Solve equation"
End Select
SolveEquation = s
End Property
Public Property Get ChooseVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vælg variabel"
Case 3: s = "Escoger variable"
Case 4: s = ""
Case Else: s = "Choose variable"
End Select
ChooseVariable = s
End Property
Public Property Get WriteVariable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eller skriv selv:"
Case 3: s = "O establezca variable:"
Case 4: s = ""
Case Else: s = "Or state variable:"
End Select
WriteVariable = s
End Property
Public Property Get WriteVariables() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Eller skriv selv adskilt ad komma"
Case 3: s = "O establezca variables separadas por coma"
Case 4: s = ""
Case Else: s = "Or state variables separated by comma"
End Select
WriteVariables = s
End Property
Public Property Get PresentDefs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gældende definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
PresentDefs = s
End Property
Public Property Get TempDefs() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Midl. definitioner/antagelser"
Case 3: s = "Definiciones/suposiciones temporales"
Case 4: s = ""
Case Else: s = "Temp. definitions/assumptions"
End Select
TempDefs = s
End Property
Public Property Get UnitWarning() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Advarsel: Enheder er slået til." & VbCrLfMac & "Pas på med variable der er enheder fx A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
Case 3: s = "Advertencia: Las unidades están encendidas. " & VbCrLfMac & "Cuidado con las variables que son unidades como: A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
Case 4: s = ""
Case Else: s = "Warning: Units are turned on." & VbCrLfMac & "Watch out for variables which are units like: A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s"
End Select
UnitWarning = s
End Property
Public Property Get SolveSystem() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs ligningssystem"
Case 3: s = "Resolver sistemas de ecuaciones"
Case 4: s = ""
Case Else: s = "Solve system of equations"
End Select
SolveSystem = s
End Property
Public Property Get ChooseVariables() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vælg variable"
Case 3: s = "Escoger variables"
Case 4: s = ""
Case Else: s = "Choose variables"
End Select
ChooseVariables = s
End Property
Public Property Get WordMatStarting() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat starter op"
Case 3: s = "WordMat está iniciando"
Case 4: s = ""
Case Else: s = "WordMat is starting"
End Select
WordMatStarting = s
End Property
Public Property Get PleaseWait() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent venligst"
Case 3: s = "Por favor, espere"
Case 4: s = ""
Case Else: s = "Please wait"
End Select
PleaseWait = s
End Property
Public Property Get TriangleSolver() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Trekantsløser"
Case 3: s = "Solucionario de triángulos"
Case 4: s = ""
Case Else: s = "Triangle solver"
End Select
TriangleSolver = s
End Property
Public Property Get RightAngled() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Retvinklet"
Case 3: s = "Triángulos rectángulos"
Case 4: s = ""
Case Else: s = "Right angled"
End Select
RightAngled = s
End Property
Public Property Get AnyTriangle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vilkårlig"
Case 3: s = "Cualquier triángulo"
Case 4: s = ""
Case Else: s = "Any"
End Select
AnyTriangle = s
End Property
Public Property Get right() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ret"
Case 3: s = "Recto"
Case 4: s = ""
Case Else: s = "Right"
End Select
right = s
End Property
Public Property Get Naming() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Navngivning"
Case 3: s = "Denominación"
Case 4: s = ""
Case Else: s = "Naming"
End Select
Naming = s
End Property
Public Property Get Manuel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manuel"
Case 3: s = "Manuel"
Case 4: s = ""
Case Else: s = "Manuel"
End Select
Manuel = s
End Property
Public Property Get AngleNaming1() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vinkler store, sider små bogstaver"
Case 3: s = "Angulos - mayúsculas, lados - minúsculas"
Case 4: s = ""
Case Else: s = "Angles capital, sides lowercase letters"
End Select
AngleNaming1 = s
End Property
Public Property Get AngleNaming2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sider navngives efter hjørner AB"
Case 3: s = "Los lados se nombran a partir del ángulo AB"
Case 4: s = ""
Case Else: s = "Sides named by corners AB"
End Select
AngleNaming2 = s
End Property
Public Property Get InsertNumbers() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt tal på figuren i Word"
Case 3: s = "Inserte números en el triángulo en Word"
Case 4: s = ""
Case Else: s = "Insert numbers on the triangle in Word"
End Select
InsertNumbers = s
End Property
Public Property Get ShowCalculations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vis mellemregninger i Word"
Case 3: s = "Mostrar cálculos en Word"
Case 4: s = ""
Case Else: s = "Show calculations in Word"
End Select
ShowCalculations = s
End Property
Public Property Get TriangleSolverExplanation1() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt de kendte sider og vinkler så beregnes de resterende sider og vinkler, og indsættes på figur i Word." & VbCrLfMac & VbCrLfMac & "Det er også muligt at ændre betegnelserne på figuren. bare klik på bogstaverne."
Case 3: s = "Insertar los lados y ángulos conocidos. los restantes lados y ángulos serán calculados e insertados en Word." & VbCrLfMac & VbCrLfMac & "Es posible cambiar los nombres de los lados y ángulos en el triángulo. Simplemente haga clic en las letras."
Case 4: s = ""
Case Else: s = "Insert the known sides and angles. The remaining sides and angles will then be calculated, and inserted in Word." & VbCrLfMac & VbCrLfMac & "It is possible to change the names of sides and angles on the triangle. Just click the letters."
End Select
TriangleSolverExplanation1 = s
End Property
Public Property Get TriangleSolverExplanation2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Figuren viser ikke de korrekte størrelser, men det gør den når den indsættes i Word."
Case 3: s = "La figura no muestra el tamaño correcto, pero lo hace cuando se inserta en Word."
Case 4: s = ""
Case Else: s = "The triangle is not to size. But a correctly scaled triangle will be inserted into Word."
End Select
TriangleSolverExplanation2 = s
End Property
Public Property Get TriangleSolverExplanation3() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat's trekantsløser anvendes med input: "
Case 3: s = "El solucionario de triángulos de WordMat se usa con la entrada:"
Case 4: s = ""
Case Else: s = "WordMat's trianglesolver is used with input: "
End Select
TriangleSolverExplanation3 = s
End Property
Public Property Get TS2Solutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Bemærk der er to mulige trekanter ud fra dine oplysninger"
Case 3: s = "Con su información se pueden construir dos triángulos"
Case 4: s = ""
Case Else: s = "Two triangles can be constructed from your information"
End Select
TS2Solutions = s
End Property
Public Property Get TS2Solutions2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "To Trekanter"
Case 3: s = "Dos triángulos"
Case 4: s = ""
Case Else: s = "Two triangles"
End Select
TS2Solutions2 = s
End Property
Public Property Get TSNoSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der findes ikke en trekant der kan konstrueres ud fra dine oplysninger."   'May be this line has no correct speech: danish says not the same in english but the property is TSNoSolution then the spanish will be that.
Case 3: s = "Con su información no se puede resolver el triángulo"
Case 4: s = ""
Case Else: s = "No triangle can be found from your information"
End Select
TSNoSolution = s
End Property
Public Property Get TSMissingInfo() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der er ikke nok information til at beregne de resterende sider og vinkler."
Case 3: s = "Se requiere más información para calcular los restantes lados y ángulos"
Case 4: s = ""
Case Else: s = "More information is needed to calculate the remaining sides and angles"
End Select
TSMissingInfo = s
End Property
Public Property Get TSInfoOK() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle sider og vinkler kan nu findes."
Case 3: s = "Se han calculado todos los lados y ángulos."
Case 4: s = ""
Case Else: s = "All sides and angles can now be calculated."
End Select
TSInfoOK = s
End Property
Public Property Get NumericSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Numerisk løsning"
Case 3: s = "Solución numérica"
Case 4: s = ""
Case Else: s = "Numeric solution"
End Select
NumericSolution = s
End Property
Public Property Get InsertSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt løsning(er)"
Case 3: s = "Insertar solución"
Case 4: s = ""
Case Else: s = "Insert solution(s)"
End Select
InsertSolutions = s
End Property
Public Property Get SeekGraphically() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Søg flere løsninger grafisk"
Case 3: s = "Buscar soluciones múltiples gráficamente"
Case 4: s = ""
Case Else: s = "Seek more solutions graphically"
End Select
SeekGraphically = s
End Property
Public Property Get InsertCASresult() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Indsæt omskrevet resultat"
Case 3: s = "Inserte resultado reducido"
Case 4: s = ""
Case Else: s = "Insert reduced result"
End Select
InsertCASresult = s
End Property
Public Property Get NoAnalyticResult() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen kunne desværre ikke løses analytisk"
Case 3: s = "No se puede resolver la ecuación analíticamente"
Case 4: s = ""
Case Else: s = "The equation could not be solved analytically"
End Select
NoAnalyticResult = s
End Property
Public Property Get FoundNumericSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der blev fundet løsninger vha. numeriske metoder, men der findes muligvis flere."
Case 3: s = "Se han encontrado soluciones utilizando métodos numéricos, pero posiblemente hay más."
Case 4: s = ""
Case Else: s = "Solutions where found using numeric methods. More solutions may exist though"
End Select
FoundNumericSolutions = s
End Property
Public Property Get NoNumericSolution() As String
Dim s As String
Select Case SprogNr
Case 1: s = "WordMat kunne ikke løse ligningen analytisk, men har angivet et omskrevet resultat der måske kan arbejdes videre med."
Case 3: s = "WordMat no podía resolver la ecuación analíticamente, pero ha establecido un resultado reducido que podría funcionar."
Case 4: s = ""
Case Else: s = "WordMat could not solve the equation analytically. A reduced result has been returned though."
End Select
NoNumericSolution = s
End Property
Public Property Get SearchingSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent der søges efter løsninger ..."
Case 3: s = "Espere, buscando soluciones ..."
Case 4: s = ""
Case Else: s = "Wait, searching for solutions ..."
End Select
SearchingSolutions = s
End Property
Public Property Get SearchingMoreSolutions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent der søges efter flere løsninger ..."
Case 3: s = "Espere, buscando más soluciones ..."
Case 4: s = ""
Case Else: s = "Wait, searching for more solutions ..."
End Select
SearchingMoreSolutions = s
End Property
Public Property Get NoNumericSolution2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der blev ikke fundet nogen løsninger vha. numeriske metoder, men der findes måske løsninger som du kan søge efter grafisk."
Case 3: s = "No se encontró ninguna solución utilizando métodos numéricos, pero, tal vez, haya soluciones que se pueden buscar gráficamente."
Case 4: s = ""
Case Else: s = "No solutions where found using numerical method, but you could try searching graphically for solutions."
End Select
NoNumericSolution2 = s
End Property
Public Property Get TempSettings() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Midl. indstillinger"
Case 3: s = "Configuración temporal"
Case 4: s = ""
Case Else: s = "Temp. settings"
End Select
TempSettings = s
End Property
Public Property Get AutoReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Automatisk reducering"
Case 3: s = "Reducción automática"
Case 4: s = ""
Case Else: s = "Auto reduce"
End Select
AutoReduce = s
End Property
Public Property Get AutoReduceExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Forsøg at reducere udtrykket til kortest mulig form vha. flere forskellige metoder."
Case 3: s = "Intenta reducir la expresión a la forma más breve posible utilizando diferentes métodos"
Case 4: s = ""
Case Else: s = "Tries to reduce the expression to shortest possible form using different methods."
End Select
AutoReduceExplain = s
End Property
Public Property Get FactorExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sæt uden for parantes" & VbCrLfMac & "Forsøger at omdanne udtrykket til produkter"
Case 3: s = "Intenta factorizar la expresión"
Case 4: s = ""
Case Else: s = "Tries to factor the expression"
End Select
FactorExplain = s
End Property
Public Property Get ExpandExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Gang ind i paranteser." & VbCrLfMac & "Alle paranteser ganges ud."
Case 3: s = "Expande la expresión"
Case 4: s = ""
Case Else: s = "Expands the expression."
End Select
ExpandExplain = s
End Property
Public Property Get Rationalize() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Omskriv til brøker"
Case 3: s = "Racionaliza"
Case 4: s = ""
Case Else: s = "Rationalize"
End Select
Rationalize = s
End Property
Public Property Get RationalizeExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Alle decimaltal omskrives til brøker."
Case 3: s = "Convierte todos los números decimales a fracciones"
Case 4: s = ""
Case Else: s = "Converts all decimal numbers to fractions"
End Select
RationalizeExplain = s
End Property
Public Property Get TrigReduce() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Reducer trigonometrisk"
Case 3: s = "Reduce trigonometricamente"
Case 4: s = ""
Case Else: s = "Reduce trigonometrically"
End Select
TrigReduce = s
End Property
Public Property Get TrigReduceExplain() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Udtrykket forsøges omskrevet vha. trigonometriske identiteter."
Case 3: s = "Reduce la expresión usando identidades trigonométricas"
Case 4: s = ""
Case Else: s = "Reduce the expression using trigonometric identities."
End Select
TrigReduceExplain = s
End Property
Public Property Get OKSame() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK, samme linje"
Case 3: s = "OK, la misma línea"
Case 4: s = ""
Case Else: s = "OK, same line"
End Select
OKSame = s
End Property
Public Property Get OKNew() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK, ny linje"
Case 3: s = "OK, una nueva línea"
Case 4: s = ""
Case Else: s = "OK, new line"
End Select
OKNew = s
End Property
Public Property Get PlotCaption() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Graftegning mm. i planen"
Case 3: s = "Graficar en el plano"
Case 4: s = ""
Case Else: s = "2D plotting"
End Select
PlotCaption = s
End Property
Public Property Get ResetAll() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil alt"
Case 3: s = "Reiniciar todo"
Case 4: s = ""
Case Else: s = "Reset all"
End Select
ResetAll = s
End Property
Public Property Get AxisTitle() As String
Dim s As String
Select Case SprogNr
Case 1: s = "aksetitel"
Case 3: s = "Título del eje"
Case 4: s = ""
Case Else: s = "axis title"
End Select
AxisTitle = s
End Property
Public Property Get LineType() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Linjetype"
Case 3: s = "Tipo de línea"
Case 4: s = ""
Case Else: s = "Line type"
End Select
LineType = s
End Property
Public Property Get EquationMissingError() As String
Dim s As String
Select Case SprogNr
Case 1: s = "For at løse en ligning skal du først indtaste en ligning vha. indsæt ligning i Word. Lad cursoren stå i ligningen og tryk igen på løs ligning. Marker flere ligninger for at løse flere ligninger med flere ubekendte."
Case 3: s = "Para resolver una ecuación debe, primero, insertar una ecuación en Word. Ponga el cursor en la ecuación y presione 'Resolver ecuación(es)' otra vez. Marque más ecuaciones para resolver un sistema de ecuaciones."
Case 4: s = ""
Case Else: s = "To solve an equation you must first insert an equation into Word. Place the cursor in the equation and press 'Solve equation(s)' again. Mark more equations to solve a system of equations"
End Select
EquationMissingError = s
End Property
Public Property Get EquationMissingError2() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ingen ligning valgt"
Case 3: s = "No hay ecuación seleccionada"
Case 4: s = ""
Case Else: s = "No equation selected"
End Select
EquationMissingError2 = s
End Property


''''' Any ''''''''''''''

Public Property Get Simplify() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Reducer"
Case 3: s = "Simplificar"
Case 4: s = ""
Case Else: s = "Simplify"
End Select
Simplify = s
End Property
Public Property Get Factor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Faktoriser"
Case 3: s = "Factorizar"
Case 4: s = ""
Case Else: s = "Factor"
End Select
Factor = s
End Property
Public Property Get Expand() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Udvid"
Case 3: s = "Expandir"
Case 4: s = ""
Case Else: s = "Expand"
End Select
Expand = s
End Property
Public Property Get Help() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Hjælp"
Case 3: s = "Ayuda"
Case 4: s = ""
Case Else: s = "Help"
End Select
Help = s
End Property
Public Property Get Manual() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manual"
Case 3: s = "Manual"
Case 4: s = ""
Case Else: s = "Manual"
End Select
Manual = s
End Property
Public Property Get ManualDoc() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Manual som Worddokument"
Case 3: s = "Manual como documento Word"
Case 4: s = ""
Case Else: s = "Manual as Word document"
End Select
ManualDoc = s
End Property
Public Property Get ManualOnline() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Online hjemmeside-Manual"
Case 3: s = "Manual como página web on-line"
Case 4: s = ""
Case Else: s = "Online webpage Manual"
End Select
ManualOnline = s
End Property
Public Property Get About() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Om"
Case 3: s = "Acerca de"
Case 4: s = ""
Case Else: s = "About"
End Select
About = s
End Property
Public Property Get Update() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Opdater"
Case 3: s = "Actualizar"
Case 4: s = ""
Case Else: s = "Update"
End Select
Update = s
End Property
Public Property Get Shortcuts() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Genveje"
Case 3: s = "Atajos"
Case 4: s = ""
Case Else: s = "Shortcuts"
End Select
Shortcuts = s
End Property
Public Property Get Clear() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil"
Case 3: s = "Borrar"
Case 4: s = ""
Case Else: s = "Clear"
End Select
Clear = s
End Property
Public Property Get Logarithm() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Logaritme"
Case 3: s = "Logaritmo"
Case 4: s = ""
Case Else: s = "Logarithm"
End Select
Logarithm = s
End Property
Public Property Get Language() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Sprog (beta)"
Case 3: s = "Lenguaje (beta)"
Case 4: s = ""
Case Else: s = "Language (beta)"
End Select
Language = s
End Property
Public Property Get Sign() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fortegn"
Case 3: s = "Signo"
Case 4: s = ""
Case Else: s = "Sign"
End Select
Sign = s
End Property
Public Property Get Positive() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Positiv"
Case 3: s = "Positivo"
Case 4: s = ""
Case Else: s = "Positive"
End Select
Positive = s
End Property
Public Property Get Negative() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Negativ"
Case 3: s = "Negativo"
Case 4: s = ""
Case Else: s = "Negative"
End Select
Negative = s
End Property
Public Property Get Zero() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nul"
Case 3: s = "Cero"
Case 4: s = ""
Case Else: s = "Zero"
End Select
Zero = s
End Property
Public Property Get Error() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Fejl"
Case 3: s = "Error"
Case 4: s = ""
Case Else: s = "Error"
End Select
Error = s
End Property
Public Property Get Definitions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Definitioner"
Case 3: s = "Definiciones"
Case 4: s = ""
Case Else: s = "Definitions"
End Select
Definitions = s
End Property
Public Property Get OK() As String
Dim s As String
Select Case SprogNr
Case 1: s = "OK"
Case 3: s = "OK"
Case 4: s = ""
Case Else: s = "OK"
End Select
OK = s
End Property
Public Property Get Cancel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Annuller"
Case 3: s = ""
Case 4: s = ""
Case Else: s = "Cancel"
End Select
Cancel = s
End Property
Public Property Get Wait() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vent"
Case 3: s = "Cancelar"
Case 4: s = ""
Case Else: s = "Wait"
End Select
Wait = s
End Property
Public Property Get Activity() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Aktivitet"
Case 3: s = "Esperar"
Case 4: s = ""
Case Else: s = "Activity"
End Select
Activity = s
End Property
Public Property Get StopLabel() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Stop"
Case 3: s = "Parar"
Case 4: s = ""
Case Else: s = "Stop"
End Select
StopLabel = s
End Property


Public Property Get EquationSolvedFor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen løses for "
Case 3: s = "La ecuación se resuelve para "
Case 4: s = ""
Case Else: s = "The equation is solved for "
End Select
EquationSolvedFor = s
End Property
Public Property Get EquationSolvedNumFor() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligningen løses numerisk for "
Case 3: s = "La ecuación se resuelve numéricamente para "
Case 4: s = ""
Case Else: s = "The equation is solved numerically for "
End Select
EquationSolvedNumFor = s
End Property
Public Property Get ByCAS() As String
Dim s As String
Select Case SprogNr
Case 1: s = " vha. " & AppNavn & ". "
Case 3: s = "por" & AppNavn & ". "
Case 4: s = ""
Case Else: s = " by " & AppNavn & ". "
End Select
ByCAS = s
End Property
Public Property Get Functions() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Forskrifter"
Case 3: s = "Funciones"
Case 4: s = ""
Case Else: s = "Functions"
End Select
Functions = s
End Property
Public Property Get Equation() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligning"
Case 3: s = "Ecuación"
Case 4: s = ""
Case Else: s = "Equation"
End Select
Equation = s
End Property
Public Property Get Equations() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Ligninger"
Case 3: s = "Ecuaciones"
Case 4: s = ""
Case Else: s = "Equations"
End Select
Equations = s
End Property
Public Property Get Points() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Punkter"
Case 3: s = "Puntos"
Case 4: s = ""
Case Else: s = "Points"
End Select
Points = s
End Property
Public Property Get Vectors() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Vektorer"
Case 3: s = "Vectores"
Case 4: s = ""
Case Else: s = "Vectors"
End Select
Vectors = s
End Property
Public Property Get Title() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Titel"
Case 3: s = "Título"
Case 4: s = ""
Case Else: s = "Title"
End Select
Title = s
End Property
Public Property Get Reset() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Nulstil"
Case 3: s = "Reiniciar"
Case 4: s = ""
Case Else: s = "Reset"
End Select
Reset = s
End Property
Public Property Get RetryNum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Komplicerede beregninger der udføres eksakt kan tage meget lang tid. Vil du istedet prøve at regne numerisk?"
Case 3: s = "Cálculos complicados se realizan exactamente y puede tomar un tiempo mulargo. øPrefiere intentar numéricamente?"
Case 4: s = ""
Case Else: s = "Complex calculation done exactly can take a very long time. Do you want to retry numerically?"
End Select
RetryNum = s
End Property




Public Property Get dAsDiffChr() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Brug alm. d som differential d"
Case 3: s = "Use la letra d como el diferencial d"
Case 4: s = ""
Case Else: s = "Use regular d as differential d"
End Select
dAsDiffChr = s
End Property

Public Property Get dAsDiffChrHelp() As String
Dim s As String
Select Case SprogNr
Case 1: s = "dy/dx notation kan normalt kun bruges vha. et specielt differential d (\dd) eller symbolet for partielt afledte (\partial). Hvis det ikke forventes at d/dx kan forveksles med et andet udtryk er denne indstilling en fordel idet det er nemmere at taste normalt d"
Case 3: s = "La notación dy/dx se usa normalmente representada por una diferencial especial d (\dd) o por el símbolo para las derivadas parciales (\partial). Si usted cree que no se confundirá con otras expresiones puede utilizar esta opción, ya que es más rápido escribir dy / dx usando d"
Case 4: s = ""
Case Else: s = "dy/dx notation can normally only be used with a special differential d (\dd) or the symbol for partial derivatives (\partial). If you dont expect d/dx to be confused with other expression, you can use this setting since it is faster to type dy/dx using regular d."
End Select
dAsDiffChrHelp = s
End Property

Public Property Get RegrUser() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Brugerdefineret regression"
Case 3: s = "Regresión definida por el usuario"
Case 4: s = ""
Case Else: s = "User defined regression"
End Select
RegrUser = s
End Property

Public Property Get RibSolveDEnum() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Løs koblede differentialligning(er) numerisk"
Case 3: s = "Resuelve sistemas de ecuaciones diferenciales"
Case 4: s = ""
Case Else: s = "Solve system of differential equation(s) numerically"
End Select
RibSolveDEnum = s
End Property

Public Property Get ErrorGeneral() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Der skete en fejl"
Case 3: s = "Ha ocurrido un error"
Case 4: s = ""
Case Else: s = "An error occured"
End Select
ErrorGeneral = s
End Property

Public Property Get Variable() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Variabel"
Case 3: s = "Variable"
Case 4: s = ""
Case Else: s = "Variable"
End Select
Variable = s
End Property

Public Property Get Warning() As String
Dim s As String
Select Case SprogNr
Case 1: s = "Advarsel"
Case 3: s = "Advertencia"
Case 4: s = ""
Case Else: s = "Warning"
End Select
Warning = s
End Property


