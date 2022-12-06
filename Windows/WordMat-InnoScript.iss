; Inno script for creating WordMat installer

#define MyAppName "WordMat"
#define MyAppVersion "1.25.3"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{301A8257-D5EF-48B4-AAC2-E86700DDA6FE}
AppName={#MyAppName}
AppVersion= {#MyAppVersion}
AppVerName={#MyAppName} v. {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}.0
AppMutex={#MyAppName}Mutex
AppPublisher=Eduap
AppPublisherURL=http://www.eduap.com/
AppSupportURL=http://www.eduap.com/
AppUpdatesURL=http://www.eduap.com/
AppCopyright=EDUAP
DefaultDirName={pf}\{#MyAppName}\
DisableDirPage=yes
DisableReadyMemo=yes
DisableReadyPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
SetupIconFile=Installer-files\WordMatIkon.ico
WizardImageFile=Installer-files\WordMatInstallImage.bmp
WizardImageStretch=no
;SourceDir=.
OutputDir=.
OutputBaseFilename={#MyAppName}
LicenseFile=Installer-files\License_en.txt
Compression=lzma
SolidCompression=yes
ChangesAssociations=yes
ShowLanguageDialog=no
CloseApplications=force
SignToolRetryCount=1
SignToolRetryDelay=2000
SignToolRunMinimized=yes
SignTool=SignToolExe
SignedUninstaller=yes 
; Command: $qC:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe$q sign /f $qC:\Users\mikae\OneDrive\Dokumenter\Eduap\Code signing certificate\EDUAP-certificate.pfx$q /p MyPassword /t http://timestamp.digicert.com /fd SHA256 $f

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl" ; LicenseFile: Installer-files\License_en.txt
Name: "da"; MessagesFile: compiler:Languages\Danish.isl  ; LicenseFile: Installer-files\License_da.txt
Name: "sp"; MessagesFile: compiler:Languages\Spanish.isl ; LicenseFile: Installer-files\License_en.txt

[Messages]
da.WelcomeLabel2=Denne guide installerer [name/ver] på computeren.%n%n Det anbefales at lukke alle kørende programmer.
en.WelcomeLabel2=This guide installs [name/ver].%n%n It is recommended to close all programs.
sp.WelcomeLabel2=Esta guía le instalará [name / ver] en su equipo.% N% n Se recomienda cerrar todas las aplicaciones que se ejecutan.

da.FinishedLabel=Du finder {#MyAppName} som en værktøjslinje i Word.
en.FinishedLabel=You will find {#MyAppName} as a Ribbon in Word.
sp.FinishedLabel=Ahora dispone de una nueva barra de herramientas en Word.

[CustomMessages]
da.UnInstallLabel=Du har allerede installeret {#MyAppName}. Bekræft først at du vil have den gamle version afinstalleret først.
en.UnInstallLabel=You have already installed {#MyAppName}. Confirm that you want to uninstall the old version first.
sp.UnInstallLabel=Ya ha instalado {#MyAppName}. Confirme que primero va a desinstalar la versión antigua.

da.SettingsLabel=Indstillinger
en.SettingsLabel=Settings
sp.SettingsLabel=Ajustes

da.SettingsLabel2=(Disse kan senere ændres under indstillinger/avanceret i WordMat)
en.SettingsLabel2=(These settings can be changed under settings / advanced in WordMat)
sp.SettingsLabel2=(Éstos ajustes se pueden cambiar más adelante en la configuración / avanzados en WordMat)

da.SettingsLabel3=Tryk på næste når du har valgt.
en.SettingsLabel3=Press next when you have made your choice.
sp.SettingsLabel3=Pulse Siguiente cuando haya seleccionado.

da.AutoStartCheckbox=Start automatisk WordMat i baggrunden med Word.
en.AutoStartCheckbox=Automatically start WordMat when Word starts.
sp.AutoStartCheckbox=Inicie WorMat automáticamente cuando inicie Word.

da.AutoUpdateCheckbox=Check Automatisk for opdateringer
en.AutoUpdateCheckbox=Automatically check for updates
sp.AutoUpdateCheckbox=Revise actualizaciones automáticamente.

da.ClickRunLabel=Du har installeret en klik og kør version af Office. Programmet installeres kun for denne bruger.'
en.ClickRunLabel=You have a clickn run version of Office. The program is only installed for the current user.
sp.ClickRunLabel=Usted ha accedido a una versión de Office en funcionamiento. El programa está instalado solamente para el usuario actual.

da.OverwriteDocs=De redigerbare Word-filer med figurer og formelsamlinger eksisterer allerede. Vil du overskrive dem?%n (ellers beholdes de gamle).%n De kan altid nulstilles under indstillinger i WordMat
en.OverwriteDocs=The Word documents with figurs and formulae already exists. Do you want to overwrite them? %n(If not you will keep the old).%n You can always reset these documents in the settings in WordMat.
sp.OverwriteDocs=El documento Word con figuras y fórmulas ya existe. Desea reemplazarlo? %n (Si pulsa no, mantendrá el antiguo).%n Siempre puede restablecer estos documentos en los ajustes en WordMat. 

da.CloseWord=Gem og luk alle åbne Word-dokumenter nu.%n%n Tryk dernæst OK for at fortsætte installationen. Evt. åbne Word-dokumenter lukkes så automatisk. %n%n (Hvis du allerede har lukket Word, er det måske nødvendigt at genstarte) 
en.CloseWord=Save and close all open Word documents now.%n%n Then press OK to continue the installation. Any open Word documents are then automatically closed. %n%n (If Word is already closed, you might need to restart your computer)
sp.CloseWord=Guarde y cierre todos los documentos de Word abiertos.%n%n A continuación, pulse Aceptar para continuar la instalación. Todos los documentos de Word abiertos se cerrarán automáticamente. %n%n (Si Word ya está cerrado, es posible que deba reiniciar el equipo)

da.CloseWord2=Word blev tvunget til at lukke.%n%n Tryk OK for at fortsætte installationen.'
en.CloseWord2=Word was forced to close. %n%n Press OK to continues the installation.
sp.CloseWord2=Word fue forzado a cerrarse. %n%n Pulse Aceptar para continuar la instalación

da.MakroDeakt=Makroer er deaktiverede i
en.MakroDeakt=Macros are disabled in 
sp.MakroDeakt=Las macros se desactivan en

da.MakroDeakt2=Det anbefales at ændre denne indstilling, da WordMat ellers ikke vil virke. Tryk OK for ændre indstillingen nu til aktivering af makroer med meddelelse, eller anuller for at fortsætte uden ændring. 
en.MakroDeakt2=It is recommended to change this setting, otherwise WordMat will not function. Press OK to change the setting to 'macros enabled with message', or cancel to continue without any change.
sp.MakroDeakt2=Se recomienda cambiar esta configuración, de lo contrario WordMat no funcionará. Pulse Aceptar para cambiar el ajuste a 'macros habilitados con el mensaje', o en Cancelar para continuar sin ningún cambio.

da.NoWord=Det ser ikke ud til at du har installeret Microsoft Word, og WordMat vil ikke fungere uden Word. Vil du alligevel fortsætte installationen?
en.NoWord=Microsoft Word is not installed. WordMat will not function without it. Do you wish to continue?
sp.NoWord=Word no está instalado. La instalación finalizará

da.NoExcel=Du har ikke installeret Excel 2007, 2010, 2013 eller 2016. Installationen afsluttes.
en.NoExcel=Excel 2007, 2010, 2013 or 2016 is not installed. The installation will terminate.
sp.NoExcel=Excel 2007, 2010, 2013 o 2016 no está instalado. La instalación finalizará

da.NotAdmin=Du skal være Administrator for at installere WordMat.%n%n Installationen afsluttes.
en.NotAdmin=You need to be administrator to install WordMat.%n%n The installation will terminate.
sp.NotAdmin=Debe ser administrador para instalar WordMat.%n%n la instalación finalizará.

da.NoDotNet=WordMat kræver Microsoft .NET Framework 4.0%n%n Opdateringen fra Microsoft vil nu blive hentet og installeret, men installationen vil tage noget tid (15-30min) og kræve internetforbindelse.%n%n Denne installation vil ikke skulle køres ved fremtidige opdateringer af WordMat.
en.NoDotNet=WordMat requires Microsoft .NET Framework 4.0 %n%n The update from Microsoft will now be downloaded and installed, but the installation will take some time (>15 min) and you need internet connection.%n%n This installation will not be required on future updates of WordMat. 
sp.NoDotNet=WordMat requiere Microsoft.NET Framework 4.0 %n%n Ahora se descargará e instalará la actualización de Microsoft, pero la instalación se llevará algún tiempo (> 15 min) y se necesita conexión a Internet.%n%n Esta instalación no se requerirá en las futuras actualizaciones de WordMat.

da.InstallError=Der skete en fejl under installationen
en.InstallError=An error occured during installation
sp.InstallError=Ha ocurrido un error durante la instalación

da.UninstallOK=Den eksisterende version af {#MyAppName} blev afinstalleret korrekt. Nu fortsætter installationen af den nye version.
en.UninstallOK=The existing version of {#MyAppName} was uninstalled correct. The installation of the new version will now continue.
sp.UninstallOK=La versión existente de {#MyAppName} se ha desinstalado correctamente. Ahora continuará la instalación de la nueva versión.

da.UnInstallError=Der skete en fejl under afinstallationen, men installationen af den nye version fortsætter.
en.UnInstallError=An error occured during the uninstallation, but the installation of the new version will continue.
sp.UnInstallError=Ha ocurrido un error durante la desinstalación, pero la instalación de la nueva versión continuará.

da.InstallThisUser=Opret {#MyAppName} menu i Word for denne bruger
en.InstallThisUser=Create {#MyAppName} ribbon in Word for this user
sp.InstallThisUser=Crear cinta {#MyAppName} en Word para este usuario

da.ReactivateWordMat=Reaktiver WordMat
en.ReactivateWordMat=Reactivate WordMat
sp.ReactivateWordMat=Reactivar WorMat

da.QuickWordMat=WordMat lommeregner
en.QuickWordMat=Quick WordMat
sp.QuickWordMat=WordMat rápido

da.InstallFor=Installer for:
en.InstallFor=Install for:
sp.InstallFor=Instalar para:

da.AllUsers=Alle brugere
en.AllUsers=All users
sp.AllUsers=Para todos los usuarios

da.ThisUser=Kun for denne bruger
en.ThisUser=Only for this user
sp.ThisUser=Sólo para este usuario

 ;ExpandConstant('{cm:InstallFor}')
   
[Files]                    
;Source: WordMat.dotm; DestDir: {%appdata}\Microsoft\Word\{code:StartFolder2007}; Flags: ignoreversion restartreplace
;Source: WordMat.dotm; DestDir: {%appdata}\Microsoft\Word\{code:StartFolder2010}; Flags: ignoreversion restartreplace
Source: WordMat.dotm; DestDir: {code:VUserStartupFolder12}; Check: VOffice12Installed ; tasks: installerbruger; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VUserStartupFolder14}; Check: VOffice14Installed ; tasks: installerbruger; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VUserStartupFolder15}; Check: VOffice15Installed ; tasks: installerbruger; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VUserStartupFolder16}; Check: VOffice16Installed ; tasks: installerbruger; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VStartupFolderAll12}; Check: VOffice12Installed ; tasks: installeralle; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VStartupFolderAll14}; Check: VOffice14Installed ; tasks: installeralle; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VStartupFolderAll15}; Check: VOffice15Installed ; tasks: installeralle; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {code:VStartupFolderAll16}; Check: VOffice16Installed ; tasks: installeralle; Flags: ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: WordMat.dotm; DestDir: {app}; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: Other\WordMatLommeregner.docm; DestDir: {app}; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: ExternalPrograms\SetupGraph-4.4.2.exe; DestDir: {app}; Components: Graph ; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete deleteafterinstall
Source: ExternalPrograms\GeoGebra-Windows-Installer-5-0-723-0.exe; DestDir: {app}; Components: GeoGebra ; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete deleteafterinstall
;Source: GeoGebra-Windows-Installer-4-4-10-0.msi; DestDir: {app}; Components: GeoGebra ; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete deleteafterinstall
Source: ..\Shared\WordDocs\FormelSamling.dotx; DestDir: {%appdata}\Microsoft\Document Building Blocks\; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly
Source: Excelfiles\*; DestDir: {app}\ExcelFiles; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly
Source: Images\*; DestDir: {app}\Images; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly
Source: ExternalPrograms\GeoGebra.ggb; DestDir: {app}; Flags: ignoreversion overwritereadonly replacesameversion uninsremovereadonly
;Source: GeoGebraFiler\geogebra_thumbnail.png; DestDir: {app}\GeoGebraFiler\; Flags: ignoreversion
Source: ExternalPrograms\graphtemplate.grf; DestDir: {app}; Flags: ignoreversion overwritereadonly uninsremovereadonly  
;Source: Other\OpretWordMenu.vbs; DestDir: {app}; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: Other\ReaktiverWordMat.vbs; DestDir: {app}; Flags: ignoreversion overwritereadonly uninsremovereadonly
; These files are in a folder a level up from the Windows folder as they are shared with the Mac version
Source: ..\Shared\WordDocs\WordMatManual.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\WordMatManual_english.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer_english.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer_english.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer_spansk.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\Figurer_spansk.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler_spansk.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler_spansk.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler_english.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\MatFormler_english.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\FysikFormler.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\FysikFormler.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\FysikFormler_spansk.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\FysikFormler_spansk.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\KemiFormler.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\KemiFormler.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\KemiFormler_spansk.docx; DestDir: {app}\WordDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\KemiFormler_spansk.docx; DestDir: {%appdata}\WordMat\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion uninsneveruninstall overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\LatexWordTemplate.dotx; DestDir: {app}\WordDocs\; Check: VOverWriteDocs; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\geogebra-math-apps\*; DestDir: {app}\geogebra-math-apps\; Check: VOverWriteDocs; Flags: recursesubdirs ignoreversion restartreplace overwritereadonly replacesameversion uninsremovereadonly uninsrestartdelete
Source: MathMenu.dll\MathMenu.dll; DestDir: {app}; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: MathMenu.dll\MathMenu.tlb; DestDir: {app}; Flags: ignoreversion regtypelib overwritereadonly uninsremovereadonly

;fjernet da overgået til geogebra silent install /S
;Source: GeoGebra\*; DestDir: {app}\GeoGebra; Flags: ignoreversion overwritereadonly uninsremovereadonly
;Source: GeoGebra\unsigned\*; DestDir: {app}\GeoGebra\unsigned; Flags: ignoreversion overwritereadonly uninsremovereadonly

Source: ExternalPrograms\dotNetFx40_Client_setup.exe; DestDir: {tmp}; Flags: dontcopy

Source: ExternalPrograms\Maxima-5.45.1\*; DestDir: {app}\Maxima-5.45.1; Flags: ignoreversion recursesubdirs overwritereadonly uninsremovereadonly; AfterInstall: ReplacePathInBat(); 
;Source: Maxima-5.23.2\*; DestDir: {app}\Maxima-5.23.2; Flags: ignoreversion
;Source: Maxima-5.23.2\bin\*; DestDir: {app}\Maxima-5.23.2\bin; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\gnuplot\*; DestDir: {app}\Maxima-5.23.2\gnuplot; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\include\*; DestDir: {app}\Maxima-5.23.2\include; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\info\*; DestDir: {app}\Maxima-5.23.2\info; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\lib\*; DestDir: {app}\Maxima-5.23.2\lib; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\libexec\*; DestDir: {app}\Maxima-5.23.2\libexec; Flags: ignoreversion recursesubdirs
;Source: Maxima-5.23.2\share\*; DestDir: {app}\Maxima-5.23.2\share; Flags: ignoreversion recursesubdirs
                                      
;egne mac-filer
;Source: mac-files\*; DestDir: {app}\Maxima-5.25.1-gcl\share\maxima\5.25.1\share; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\Maxima-files\SolveReal.mac; DestDir: {app}\Maxima-5.45.1\share\maxima\5.45.1\share\contrib; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\Maxima-files\WordMatUnitAddon.mac; DestDir: {app}\Maxima-5.45.1\share\maxima\5.45.1\share\contrib; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\Maxima-files\unit.mac; DestDir: {app}\Maxima-5.45.1\share\maxima\5.45.1\share\contrib\unit; Flags: ignoreversion overwritereadonly uninsremovereadonly

;Source: C:\Users\Mikael\Documents\Eduap\MathMenu\MathMenu.dll; DestDir: {app}; Flags: ignoreversion gacinstall ; StrongAssemblyName: "MathMenu, Version=1.0.0.1, Culture=neutral, PublicKeyToken=89e88f8548e16f37, ProcessorArchitecture=MSIL"

;Latex   Skabelonen bliver dog ikke installeret korrekt for forskellige sprog hvis ikke Word 2013
Source: ..\Shared\Fonts\lmroman12-regular.otf; DestDir: {fonts};FontInstall: "LM Roman 12"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\latinmodern-math.otf; DestDir: {fonts};FontInstall: "Latin Modern Math"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\STIXMath-Regular.otf; DestDir: {fonts};FontInstall: "STIX Math"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\STIX-Regular.otf; DestDir: {fonts};FontInstall: "STIX"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\STIX-Bold.otf; DestDir: {fonts};FontInstall: "STIX"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\STIX-Italic.otf; DestDir: {fonts};FontInstall: "STIX"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\STIX-BoldItalic.otf; DestDir: {fonts};FontInstall: "STIX"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-math.otf; DestDir: {fonts};FontInstall: "XITS Math"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-mathbold.otf; DestDir: {fonts};FontInstall: "XITS Math"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-regular.otf; DestDir: {fonts};FontInstall: "XITS"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-italic.otf; DestDir: {fonts};FontInstall: "XITS"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-bold.otf; DestDir: {fonts};FontInstall: "XITS"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\Fonts\xits-bolditalic.otf; DestDir: {fonts};FontInstall: "XITS"; Flags: fontisnttruetype onlyifdoesntexist uninsneveruninstall
Source: ..\Shared\WordDocs\LatexWordTemplate.dotx; DestDir: {reg:HKCU\Software\Microsoft\Office\12.0\Word\Options,PersonalTemplates|{userdocs}\Custom Office templates}; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\LatexWordTemplate.dotx; DestDir: {reg:HKCU\Software\Microsoft\Office\14.0\Word\Options,PersonalTemplates|{userdocs}\Custom Office templates}; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\LatexWordTemplate.dotx; DestDir: {reg:HKCU\Software\Microsoft\Office\15.0\Word\Options,PersonalTemplates|{userdocs}\Custom Office templates}; Flags: ignoreversion overwritereadonly uninsremovereadonly
;Source: ..\Shared\WordDocs\LatexWordTemplate.dotx; DestDir: {app}; Flags: ignoreversion overwritereadonly uninsremovereadonly
Source: ..\Shared\WordDocs\BibWord_Stylesheet.xsl; DestDir: {code:WordInstallFolder|15}\Bibliography\Style; Flags: ignoreversion overwritereadonly uninsremovereadonly
        
Source: WebViewWrap\*; DestDir: {app}\WebViewWrap; Flags: ignoreversion overwritereadonly uninsremovereadonly recursesubdirs

[Components]
Name: "GeoGebra"; Description: "Installs GeoGebra compatible with WordMat"; Types: full
Name: "Graph"; Description: "Installs Graph 4.4.2"; Types: full

[Tasks]
Name: installeralle; Description: {cm:AllUsers}; GroupDescription: {cm:InstallFor};  Flags: exclusive
Name: installerbruger; Description: {cm:ThisUser}; GroupDescription: {cm:InstallFor};  Flags: exclusive unchecked
;Name: Taskcheckforupdates; Description: "Check Automatisk for opdateringer"; GroupDescription: "Indstillinger  (kan ændres under indstillinger/avanceret)";  Flags: unchecked
;Name: TaskAutoStart; Description: "Start automatisk WordMat i baggrunden med Word"; GroupDescription: "Indstillinger  (kan ændres under indstillinger/avanceret)";  Flags: unchecked

[Registry]
; trusted locations
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\Location10"; ValueType: string; ValueName: "Path"; ValueData: {code:VStartupFolderAll16} ;Check: VWord16Installed ;Tasks: installeralle; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\Location11"; ValueType: string; ValueName: "Path"; ValueData: {code:VUserStartupFolder16} ;Check: VWord16Installed ;Tasks: installerbruger; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location12"; ValueType: string; ValueName: "Path"; ValueData: "{pf}\WordMat\" ;Check: VExcel16Installed; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location12"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: 1 ;Check: VExcel16Installed;Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location10"; ValueType: string; ValueName: "Path"; ValueData: {code:VStartupFolderAll15} ;Check: VWord15Installed ;Tasks: installeralle; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location11"; ValueType: string; ValueName: "Path"; ValueData: {code:VUserStartupFolder15} ;Check: VWord15Installed ;Tasks: installerbruger; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\Location12"; ValueType: string; ValueName: "Path"; ValueData: "{pf}\WordMat\" ;Check: VExcel15Installed; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\Location12"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: 1 ;Check: VExcel15Installed;Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location10"; ValueType: string; ValueName: "Path"; ValueData: {code:VStartupFolderAll14} ;Check: VWord14Installed ;Tasks: installeralle; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location11"; ValueType: string; ValueName: "Path"; ValueData: {code:VUserStartupFolder14} ;Check: VWord14Installed ;Tasks: installerbruger; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Excel\Security\Trusted Locations\Location12"; ValueType: string; ValueName: "Path"; ValueData: "{pf}\WordMat\" ;Check: VExcel14Installed; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Excel\Security\Trusted Locations\Location12"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: 1 ;Check: VExcel14Installed;Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location10"; ValueType: string; ValueName: "Path"; ValueData: {code:VStartupFolderAll12} ;Check: VWord12Installed ;Tasks: installeralle;Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location11"; ValueType: string; ValueName: "Path"; ValueData: {code:VUserStartupFolder12} ;Check: VWord12Installed ;Tasks: installerbruger; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Excel\Security\Trusted Locations\Location12"; ValueType: string; ValueName: "Path"; ValueData: "{pf}\WordMat\" ;Check: VExcel12Installed ;Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Excel\Security\Trusted Locations\Location12"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: 1 ;Check: VExcel12Installed ;Flags: uninsdeletekey
; macros with warning must be enabled. But apparently not neccesary when you have trusted locations.
; Yes it is nessesary 27/6-11 error resolved at user by setting this 
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Word\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('12','Word') 
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Word\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('14','Word')
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('15','Word')
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('16','Word')
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Excel\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('12','Excel')
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Excel\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('14','Excel')
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Excel\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('15','Excel')
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\"; ValueType: dword; ValueName: "VBAWarnings"; ValueData: 2 ;Check: IsMacrosDisabled('16','Excel')
; These must be deleted if exist, since they will force 2003 mode
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Options\"; ValueType: none; ValueName:"DefaultFormat"; ValueData: none ; Flags: deletekey dontcreatekey uninsdeletekey 
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Options\"; ValueType: none; ValueName:"DefaultFormat"; ValueData: none ; Flags: deletekey dontcreatekey uninsdeletekey 
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Word\Options\"; ValueType: none; ValueName:"DefaultFormat"; ValueData: none ; Flags: deletekey dontcreatekey uninsdeletekey 
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Word\Options\"; ValueType: none; ValueName:"CompatMode"; ValueData: none ; Flags: deletekey dontcreatekey uninsdeletekey 
; Dont disable. Måske er det ikke nok at kalde den WordMat. Den skal måske have hele stien
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList\"; ValueType: dword; ValueName:"WordMat"; ValueData: 1 
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList\"; ValueType: dword; ValueName:"WordMat.dotm"; ValueData: 1 
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList\"; ValueType: dword; ValueName:"C:\Program Files\Microsoft Office\Root\Office16\STARTUP\WordMat.dotm"; ValueData: 1 


; WordMat settings
Root: HKCU; Subkey: "Software\WordMat"; Flags: 
Root: HKCU; Subkey: "Software\WordMat\Settings"; Flags: 
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "AllTrig"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Exact"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Separator"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "LogOutput"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Complex"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Units"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "SolveBoolOrSet"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "SigFig"; ValueData: 7 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "VidNot"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Radians"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Forklaring"; ValueData: 1 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "MaximaCommand"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "ExcelEmbed"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Gangetegn"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: string; ValueName: "OutUnits"; ValueData: "" ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Index"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "AntalBeregninger"; ValueData: 0 ; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "AutoStart"; ValueData: 0 ; Flags: uninsdeletekey
;Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "AutoStart"; ValueData: 1 ; Flags: uninsdeletekey ; Tasks: TaskAutoStart
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "CheckForUpdate"; ValueData: 0 ; Flags: uninsdeletekey
;Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "CheckForUpdate"; ValueData: 1 ; Flags: uninsdeletekey ; Tasks: Taskcheckforupdates;
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "RibbonPointer"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "BigFloat"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "ShowAssum"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "PolarOutput"; ValueData: 0 ; Flags: uninsdeletekey

Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "GraphApp"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Language"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "dAsDiffChr"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: string; ValueName: "LatexStart"; ValueData: "" ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: string; ValueName: "LatexSlut"; ValueData: "" ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "LatexUnits"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "ConvertTexWithMaxima"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "EqNumPlacement"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "EqNumType"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "EqAskRef"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "Backup"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "BackupNo"; ValueData: 0 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "BackupMaxNo"; ValueData: 20 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\WordMat\Settings"; ValueType: dword; ValueName: "BackupTime"; ValueData: 5 ; Flags: uninsdeletekey


; GeoGebra fileassociation
;Root: HKCR; Subkey: ".ggb"; ValueType: string; ValueName: ""; ValueData: "WordMatGeoGebra"; Flags: uninsdeletevalue createvalueifdoesntexist
;".ggb" is the extension we're associating. "MyProgramFile" is the internal name for the file type as stored in the registry. Make sure you use a unique name for this so you don't inadvertently overwrite another application's registry key.
;Root: HKCR; Subkey: "WordMatGeoGebra"; ValueType: string; ValueName: ""; ValueData: "GeoGebra"; Flags: uninsdeletekey
;"My Program File" above is the name for the file type as shown in Explorer.
;Root: HKCR; Subkey: "WordMatGeoGebra\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\GeoGebra\GeoGebra.exe,0"
;"DefaultIcon" is the registry key that specifies the filename containing the icon to associate with the file type. ",0" tells Explorer to use the first icon from MYPROG.EXE. (",1" would mean the second icon.)
;Root: HKCR; Subkey: "WordMatGeoGebra\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\GeoGebra\GeoGebra.exe"" ""%1"""
;shell\open\command" is the registry key that specifies the program to execute when a file of the type is double-clicked in Explorer. The surrounding quotes are in the command line so it handles long filenames correctly.
Root: HKCR; Subkey: ".ggb"; ValueType: string; ValueName: ""; ValueData: "GeoGebraWM"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "GeoGebraWM"; ValueType: string; ValueName: ""; ValueData: "GeoGebra 5"; Flags: uninsdeletekey 
Root: HKCR; Subkey: "GeoGebraWM\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{pf32}\GeoGebra 5.0\geogebra.exe,0" 
Root: HKCR; Subkey: "GeoGebraWM\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{pf32}\GeoGebra 5.0\geogebra.exe"" ""%1""" 

;user template
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Options"; ValueType: string; ValueName: "PersonalTemplates"; ValueData: "{userdocs}\Custom Office templates" ; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Options"; ValueType: string; ValueName: "PersonalTemplates"; ValueData: "{userdocs}\Custom Office templates" ; Flags: createvalueifdoesntexist

[Run]
;Filename:"{dotnet40}\RegAsm.exe"; Parameters: " ""{app}\MathMenu.dll"" /codebase" ;WorkingDir: "{app}"; StatusMsg: "Registrerer MathMenu.dll ..."; Flags: runhidden;
Filename:"{code:DotNetFolder}\RegAsm.exe"; Parameters: " ""{app}\MathMenu.dll"" /codebase" ;WorkingDir: "{app}"; StatusMsg: "Registrerer MathMenu.dll ..."; Flags: runhidden;
Filename: {app}\SetupGraph-4.4.2.exe; Parameters: "/verysilent /SUPPRESSMSGBOXES"; StatusMsg: "Installerer Graph..."; check:InstallGraph; Components: Graph ; Flags: nowait
Filename: {app}\GeoGebra-Windows-Installer-5-0-723-0.exe; Parameters: "/S";check:InstallGeoGebra; StatusMsg: "Installerer GeoGebra...";  Components: GeoGebra ;
;Filename: msiexec; Parameters: "/i ""{app}\GeoGebra-Windows-Installer-4-4-10-0.msi"" /quiet";check:InstallGeoGebra; StatusMsg: "Installerer GeoGebra..."; Components: GeoGebra ;
;Filename: javaws ; Parameters: "-import -system -silent -association -shortcut http://www.geogebra.org/webstart/geogebra.jnlp"; StatusMsg: "Installerer GeoGebra...";
Filename:"{code:DotNetFolder}\RegAsm.exe"; Parameters: " ""{app}\WebViewWrap\WebViewWrap.dll"" /codebase" ;WorkingDir: "{app}"; StatusMsg: "Registrerer WebViewWrap.dll ..."; Flags: runhidden;


[UninstallRun]
;Filename:"{dotnet40}\RegAsm.exe"; Parameters: " ""{app}\MathMenu.dll"" /u " ; WorkingDir: "{app}"; StatusMsg: "Unregistering controls ..."; Flags: runhidden;
Filename:"{code:DotNetFolder}\RegAsm.exe"; Parameters: " ""{app}\MathMenu.dll"" /u " ; WorkingDir: "{app}"; StatusMsg: "Unregistering controls ..."; Flags: runhidden;
Filename:"{code:DotNetFolder}\RegAsm.exe"; Parameters: " ""{app}\WebViewWrap\WebViewWrap.dll"" /u " ; WorkingDir: "{app}"; StatusMsg: "Unregistering controls ..."; Flags: runhidden;

[UninstallDelete]
;Type: files; Name: "{code:VUserStartupFolder12}\MathMenu.dotm"
;Type: files; Name: "{code:VUserStartupFolder14}\MathMenu.dotm"
Type: files; Name: "{code:VUserStartupFolder12}\WordMat.dotm"
Type: files; Name: "{code:VUserStartupFolder14}\WordMat.dotm"
Type: files; Name: "{code:VUserStartupFolder15}\WordMat.dotm"
Type: files; Name: "{code:VUserStartupFolder16}\WordMat.dotm"
Type: files; Name: "{code:VStartupFolderAll12}\WordMat.dotm"
Type: files; Name: "{code:VStartupFolderAll14}\WordMat.dotm"
Type: files; Name: "{code:VStartupFolderAll15}\WordMat.dotm"
Type: files; Name: "{code:VStartupFolderAll16}\WordMat.dotm"
Type: filesandordirs; Name: "{app}\WebViewWrap"
Type: filesandordirs; Name: "{app}\geogebra-math-apps"
Type: filesandordirs; Name: "{app}\Maxima*"
Type: dirifempty; Name: "{app}\WordDocs"
Type: dirifempty; Name: "{app}"
;Type: files; Name: "{code:VUserStartupFolder12}\WordMat2.dotm"
;Type: files; Name: "{code:VUserStartupFolder14}\WordMat2.dotm"
;Type: files; Name: "{code:VStartupFolderAll12}\WordMat2.dotm"
;Type: files; Name: "{code:VStartupFolderAll14}\WordMat2.dotm"
                   

[Icons]
Name: "{group}\{cm:InstallThisUser}"; Filename: "{app}\OpretWordMenu.vbs"; WorkingDir: "{app}"
Name: "{group}\{cm:ReactivateWordMat}"; Filename: "{app}\ReaktiverWordMat.vbs"; WorkingDir: "{app}"
;Name: "{group}\GeoGebra"; Filename: "{app}\GeoGebra\geogebra.exe"; WorkingDir: "{app}" 
;Name: "{group}\GeoGebra"; Filename: "javaws.exe";Parameters: "-system http://www.geogebra.org/webstart/geogebra.jnlp"; WorkingDir: "{app}"
Name: "{group}\{cm:QuickWordMat}"; Filename: "{app}\WordMatLommeregner.docm"; WorkingDir: "{app}"
Name: {group}\{cm:ProgramOnTheWeb,WordMat}; Filename: http://www.eduap.com/
Name: {group}\{cm:UninstallProgram,WordMat}; Filename: {uninstallexe}
;Name: "{userdesktop}\{cm:QuickWordMat}"; Filename: "{app}\WordMatLommeregner.docm"; WorkingDir: "{app}"
 
[Code]
var
  ClickToRun: Boolean;
  ClickToRun15: Boolean;
  ClickToRun16: Boolean;
  Word12Installed:Boolean;
  Word14Installed:Boolean;
  Word15Installed:Boolean;
  Word16Installed:Boolean;
  Excel12Installed:Boolean;
  Excel14Installed:Boolean;
  Excel15Installed:Boolean;
  Excel16Installed:Boolean;
  Office12Installed:Boolean;
  Office14Installed:Boolean;
  Office15Installed:Boolean;
  Office16Installed:Boolean;
  Office12Folder:string;
  Office14Folder:string;
  Office15Folder:string;
  Office16Folder:string;
  UserStartupFolder12:string;
  UserStartupFolder14:string;
  UserStartupFolder15:string;
  UserStartupFolder16:string;
  OverWriteDocs:Boolean;
  AutoStart:Cardinal;
  CheckForUpdate:Cardinal;
  AntalBeregninger:Cardinal;
  GangeTegn:Cardinal;
  ExcelEmbed:Cardinal;
  MaximaCommand:Cardinal;
  Forklaring:Cardinal;
  Radians:Cardinal;
  VidNot:Cardinal;
  SigFig:Cardinal;
  SolveBoolOrSet:Cardinal;
  Units:Cardinal;
  Complex:Cardinal;
  LogOutput:Cardinal;
  Separator:Cardinal;
  Exact:Cardinal;
  AllTrig:Cardinal;
  OutUnits:String;
  mIndex:Cardinal;
  BigFloat:Cardinal;
  ShowAssum:Cardinal;
  PolarOutput:Cardinal;

  GraphApp:Cardinal;
  Language:Cardinal;
  dAsDiffChr:Cardinal;
  LatexStart:String;
  LatexSlut:String;
  LatexUnits:Cardinal;
  ConvertTexWithMaxima:Cardinal;
  EqNumPlacement:Cardinal;
  EqNumType:Cardinal;
  EqAskRef:Cardinal;
  Backup:Cardinal;
  BackupNo:Cardinal;
  BackupMaxNo:Cardinal;
  BackupTime:Cardinal;

/////////////////////////////////////////////////////////////////////

function VWord12Installed():Boolean;
begin
  Result:=Word12Installed;
end;
function VWord14Installed():Boolean;
begin
  Result:=Word14Installed;
end;
function VWord15Installed():Boolean;
begin
  Result:=Word15Installed;
end;
function VWord16Installed():Boolean;
begin
  Result:=Word16Installed;
end;
function VExcel12Installed():Boolean;
begin
  Result:=Excel12Installed;
end;
function VExcel14Installed():Boolean;
begin
  Result:=Excel14Installed;
end;
function VExcel15Installed():Boolean;
begin
  Result:=Excel15Installed;
end;
function VExcel16Installed():Boolean;
begin
  Result:=Excel16Installed;
end;
function VOffice12Installed():Boolean;
begin
  Result:=Office12Installed;
end;
function VOffice14Installed():Boolean;
begin
  Result:=Office14Installed;
end;
function VOffice15Installed():Boolean;
begin
  Result:=Office15Installed;
end;
function VOffice16Installed():Boolean;
begin
  Result:=Office16Installed;
end;
function VOverWriteDocs():Boolean;
begin
  Result:=OverwriteDocs;
end;
function VOffice12Folder(Param:String):String;
begin
  Result:=Office12Folder;
end;
function VOffice14Folder(Param:String):String;
begin
  Result:=Office14Folder;
end;
function VOffice15Folder(Param:String):String;
begin
  Result:=Office15Folder;
end;
function VOffice16Folder(Param:String):String;
begin
  Result:=Office16Folder;
end;
function VUserStartupFolder12(Param:String):String;
begin
  Result:=UserStartupFolder12 ;
end;
function VUserStartupFolder14(Param:String):String;
begin
  Result:=UserStartupFolder14;
end;
function VUserStartupFolder15(Param:String):String;
begin
  Result:=UserStartupFolder15;
end;
function VUserStartupFolder16(Param:String):String;
begin
  Result:=UserStartupFolder16;
end;
function VStartupFolderAll12(Param:String):String;
begin
  Result:=Office12Folder + 'STARTUP';
end;     
function VStartupFolderAll14(Param:String):String;
begin
  if ClickToRun then
    Result:= UserStartupFolder14
  else
    Result:=Office14Folder + 'STARTUP';
end;
function VStartupFolderAll15(Param:String):String;
begin
  if ClickToRun15 then
    Result:= UserStartupFolder15
  else
    Result:=Office15Folder + 'STARTUP';
end;
function VStartupFolderAll16(Param:String):String;
begin
  if ClickToRun16 then
    Result:= UserStartupFolder16
  else
    Result:=Office16Folder + 'STARTUP';
end;
/////////////////////////////////////////////////////////////////////  custom wizard side til slut der spr om indstillinger 
//var
//  ActionPage: TInputOptionWizardPage;

//procedure InitializeWizard;
//begin
//  ActionPage := CreateInputOptionPage(wpReady,
//    ExpandConstant('{cm:SettingsLabel}'), ExpandConstant('{cm:SettingsLabel2}'),
//    ExpandConstant('{cm:SettingsLabel3}'),
//    False, False);
//  CheckForUpdate:=0;          
//  ActionPage.Add(ExpandConstant('{cm:AutoUpdateCheckbox}'));
//
//  if CheckForUpdate=0 then
//    ActionPage.Values[0] := false
//  else
//    ActionPage.Values[0] := true;
//end;

//function NextButtonClick(CurPageID: Integer): Boolean;
//begin
//  Result := True;
//  if CurPageID = ActionPage.ID then begin 
//    if ActionPage.Values[0] then
//      CheckForUpdate:=1
//    else
//      CheckForUpdate:=0;
//  end;
//end;


/////////////////
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
//  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{{301A8257-D5EF-48B4-AAC2-E86700DDA6FE}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

/////////////////////////////////////////////////////////////////////
function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

function testtask(): String;
begin
  Result := 'unchecked'
end;
/////////////////////////////////////////////////////////////////////
function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
  Res: Integer;
begin
// Return Values:
// 1 - uninstall string is empty
// 2 - error executing the UnInstallString
// 3 - successfully executed the UnInstallString
// 4 - uninstall cancelled

  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    if WizardSilent() then
      Res:=IDOK
    else
     Res:=MsgBox(ExpandConstant('{cm:UnInstallLabel}'), mbInformation, MB_OKCANCEL);
     if Res=IDOK then
     begin
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','AntalBeregninger', AntalBeregninger);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','AutoStart', AutoStart);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','CheckForUpdate', CheckForUpdate);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','AllTrig', AllTrig);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Complex', Complex);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Exact', Exact);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ExcelEmbed', ExcelEmbed);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Forklaring', Forklaring);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Gangetegn', Gangetegn);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LogOutput', LogOutput);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','MaximaCommand', MaximaCommand);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Radians', Radians);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Separator', Separator);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','SigFig', SigFig );
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','SolveBoolOrSet', SolveBoolOrSet);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Units', Units);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','VidNot', VidNot);
      RegQueryStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','OutUnits', OutUnits);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Index', mIndex);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BigFloat', BigFloat);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ShowAssum', ShowAssum);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','PolarOutput', PolarOutput);

      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','GraphApp', GraphApp);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Language', Language);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','dAsDiffChr', dAsDiffChr);
      RegQueryStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexStart', LatexStart);
      RegQueryStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexSlut', LatexSlut);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexUnits', LatexUnits);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ConvertTexWithMaxima', ConvertTexWithMaxima);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqNumPlacement', EqNumPlacement);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqNumType', EqNumType);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqAskRef', EqAskRef);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Backup', Backup);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupNo', BackupNo);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupMaxNo', BackupMaxNo);
      RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupTime', BackupTime);
      if (BackupMaxNo=0) and (BackupTime=0) then
      begin
        BackupMaxNo:=10;
        BackupTime:=5;
      end;

    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
    end
    else
      Result := 4;
  end else
    Result := 1;
end;

/////////////////////////////////////////////////////////////////////
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////
// Checks wheather the installation is click-to-run
function CheckClickToRun():Boolean;
var 
regk:string;
Begin
  RegQueryStringValue(HKLM, 'Software\Microsoft\Office\14.0\Common\InstallRoot\Virtual','VirtualWord',regk) ;
  if regk='' then
  begin
      if RegKeyExists(HKLM, 'Software\Microsoft\Office\14.0\CVH') then
         regk:='OK';
  end;
  // check evt. også om filen c:\programmer\Common files\Microsoft shared\Virtulization handler\CVH.EXE  findes og ikke installeret normalt
  if regk<>'' then
  begin
    Result:= True;
    MsgBox(ExpandConstant('{cm:ClickRunLabel}'),mbInformation, MB_OK);
  end
  else
  begin
    Result:= False;
  end;  
end;
function CheckClickToRun15():Boolean;
var 
regk:string;
Begin
  RegQueryStringValue(HKLM, 'Software\Microsoft\Office\15.0\Common\InstallRoot\Virtual','VirtualWord',regk) ;
  if regk='' then
  begin
       if RegKeyExists(HKLM, 'Software\Microsoft\Office\15.0\CVH')=True then
         regk:='OK';
  end;
  if regk='' then
  begin
       if RegKeyExists(HKLM, 'Software\Microsoft\Office\15.0\Common\InstallRoot\Virtual\VirtualOutlook')=True then
         regk:='OK';
  end;
  // check evt. også om filen c:\programmer\Common files\Microsoft shared\Virtulization handler\CVH.EXE  findes og ikke installeret normalt
  if regk<>'' then
  begin
    Result:= True;
    MsgBox(ExpandConstant('{cm:ClickRunLabel}'),mbInformation, MB_OK);
  end
  else
  begin
    Result:= False;
  end;  
end;
function CheckClickToRun16():Boolean;
var 
regk:string;
Begin
  RegQueryStringValue(HKLM, 'Software\Microsoft\Office\16.0\Common\InstallRoot\Virtual','VirtualWord',regk) ;
  if regk='' then
  begin
       if RegKeyExists(HKLM, 'Software\Microsoft\Office\16.0\CVH')=True then
         regk:='OK';
  end;
  if regk='' then
  begin
       if RegKeyExists(HKLM, 'Software\Microsoft\Office\16.0\Common\InstallRoot\Virtual\VirtualOutlook')=True then
         regk:='OK';
  end;
  // check evt. også om filen c:\programmer\Common files\Microsoft shared\Virtulization handler\CVH.EXE  findes og ikke installeret normalt
  if regk<>'' then
  begin
    Result:= True;
    MsgBox(ExpandConstant('{cm:ClickRunLabel}'),mbInformation, MB_OK);
  end
  else
  begin
    Result:= False;
  end;  
end;
//////////////////////////////////////////
function BoolToString(Value: Boolean): String; 
begin 
  if Value then Result := 'True' else Result := 'False'; 
end; 
////////////////////////////////////////////
function OverWriteFigurer():Boolean;
begin
   Result:=true;
   if not(WizardSilent()) then
   begin
   if FileExists(ExpandConstant('{%appdata}\WordMat\Figurer.docx')) then
     begin
       if MsgBox(ExpandConstant('{cm:OverwriteDocs}'), mbInformation, MB_YESNO)=IDNO then
         Result:=false;
     end;
   end;
end;

///////////////////////////////////////////
// default office word folder
function DefaultWordInstallFolder(Param: String): String;
var
foldername: String;
begin
  RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe','Path', foldername);
//  if not(DirExists(foldername)) then
//    RegQueryStringValue(HKEY_CLASSES_ROOT, 'Applications\Winword.exe\shell\edit\command','', foldername);   //giver ikke sti men også winword.exe
  Result := foldername;
end;

//////////////////////////////////////////////////////////
function WordInstallFolder(Param: String): String;
var
foldername: String;
begin
  RegQueryStringValue(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Office\' + Param + '.0\Word\InstallRoot','Path', foldername);
  if not(FileExists(foldername + 'winword.exe')) then
    RegQueryStringValue(HKEY_CURRENT_USER, 'Software\Microsoft\Office\' + Param + '.0\Word\Options','PROGRAMDIR', foldername);

  if not(FileExists(foldername + 'winword.exe')) then
      RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe','Path', foldername);
  
    if not(FileExists(foldername + 'winword.exe')) then
      try
      foldername := ExpandConstant('{pf64}\Microsoft Office\Office' + Param + '\');
      except
      end;

    if not(FileExists(foldername + 'winword.exe')) then
      try
      foldername := ExpandConstant('{pf64}\Microsoft Office\root\Office' + Param + '\');
      except
      end;

    if not(FileExists(foldername + 'winword.exe')) then
      try
      foldername := ExpandConstant('{pf32}\Microsoft Office\root\Office' + Param + '\');
      except
      end;
  
  if not(FileExists(foldername + 'winword.exe')) then
      foldername := ExpandConstant('{pf32}\Microsoft Office\Office' + Param + '\');
  
    Result := foldername;
end;

function StartupFolder(OfficeVersion: String): String;
var
sf: String;
sf2: String;
begin
  RegQueryStringValue(HKEY_CURRENT_USER, 'Software\Microsoft\Office\' + OfficeVersion + '.0\Common\General','Startup', sf);
  if sf = '' then
    sf := 'START';
  sf2 := ExpandConstant('{%appdata}\Microsoft\Word\' + sf);

  if DirExists(sf2) then
    Result := sf2
  else 
  begin
    sf2 := ExpandConstant('{%appdata}\Microsoft\Word\STARTUP');
    if DirExists(sf2) then
      Result := sf2
    else
//      Result := ExpandConstant(WordInstallFolder(OfficeVersion) + 'STARTUP') // ved klik og kør duer denne ikke 
      Result := ExpandConstant('{%appdata}\Microsoft\Word\' + sf )
    end;
end;

const
  // Constants for GetBinaryType return values.
  SCS_32BIT_BINARY = 0;
  SCS_64BIT_BINARY = 6;
  // There are other values that GetBinaryType can return, but we're
  // not interested in them.

// Declare Win32 function  
function GetBinaryType(lpApplicationName: AnsiString; var lpBinaryType: Integer): Boolean;
external 'GetBinaryTypeA@kernel32.dll stdcall';

function Is64BitWordFromRegisteredExe(): Boolean;
var
  wordPath: String;
  binaryType: Integer;
begin
  Result := False; // Default value - assume 32-bit unless proven otherwise.
  // RegQueryStringValue second param is '' to get the (default) value for the key
  // with no sub-key name, as described at
  // http://stackoverflow.com/questions/913938/
  if IsWin64() and RegQueryStringValue(HKEY_LOCAL_MACHINE,
      'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe',
      '', wordPath) then begin
    // We've got the path to Word.
    try
      if GetBinaryType(wordPath, binaryType) then begin
        Result := (binaryType = SCS_64BIT_BINARY);
      end;
    except
      // Ignore - better just to assume it's 32-bit than to let the installation
      // fail.  This could fail because the GetBinaryType function is not
      // available.  I understand it's only available in Windows 2000
      // Professional onwards.
    end;
  end;
end;

function Is64bit(): boolean;
var
  key: string;
  success: boolean;
  bitness: string;
begin
  Result := false;
  Result := Is64BitWordFromRegisteredExe();

  key := 'SOFTWARE\Microsoft\Office\14.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;
    
  key := 'SOFTWARE\Microsoft\Office\15.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;

  key := 'SOFTWARE\Microsoft\Office\16.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;

  key := 'SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;

  key := 'SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;

  key := 'SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook';
  RegQueryStringValue(HKLM, key, 'Bitness', bitness); 
  if bitness='x64' then
    Result := true;

end;


function DotNetFolder(Param: string):string;
begin
  if Is64bit()=true then
        Result := ExpandConstant('{dotnet4064}' )
  else
        Result := ExpandConstant('{dotnet4032}' )
end;

function StartupFolder12(Param: string): string;
begin
  Result := StartupFolder('12');
end;
function StartupFolder14(Param: string): string;
begin
  Result := StartupFolder('14');
end;
function StartupFolder15(Param: string): string;
begin
  Result := StartupFolder('15');
end;
function StartupFolder16(Param: string): string;
begin
  Result := StartupFolder('16');
end;

// For alle brugere
function StartupFolderAll(OfficeVersion: String): String;
var
clsid: String;
wordfolder: String;
begin
 //    RegQueryStringValue(HKEY_CLASSES_ROOT, 'Word.Application.' + OfficeVersion + '\CLSID','', clsid);
    if (OfficeVersion='14') or (OfficeVersion='15') or (OfficeVersion='16') then
         if ClickToRun then
            Result:=StartupFolder('14')
        else if clicktorun15 then
            Result:=StartupFolder('15')
        else if clicktorun16 then
            Result:=StartupFolder('16')
        else
    begin
//    RegQueryStringValue(HKEY_CURRENT_USER, 'Software\Microsoft\Office\' + OfficeVersion + '.0\Word\Options','PROGRAMDIR', wordfolder);
    wordfolder:=WordInstallFolder(OfficeVersion);
    wordfolder:=wordfolder + 'STARTUP'
    if not(DirExists(wordfolder)) then begin
        try
        Result:=ExpandConstant('{pf64}\Microsoft Office\Office' + OfficeVersion + '\STARTUP')
        except
        end;
        if not(DirExists(Result)) then
        Result:=ExpandConstant('{pf}\Microsoft Office\Office' + OfficeVersion + '\STARTUP')
        end
    else
        Result:=wordfolder
    end;
//    if not(DirExists(Result)) then
//         Result:= WordInstallFolder('') + '\STARTUP';
end;

function StartupFolder12a(Param: string): string;
begin
  Result := StartupFolderAll('12');
end;
function StartupFolder14a(Param: string): string;
begin
  Result := StartupFolderAll('14');
end;
function StartupFolder15a(Param: string): string;
begin
  Result := StartupFolderAll('15');
end;
function StartupFolder16a(Param: string): string;
begin
  Result := StartupFolderAll('16');
end;

// er office installeret?  IsWordInstalled('12') eller IsWordInstalled('14')
function IsWordInstalled(version: string): boolean;
begin         
       if FileExists(WordInstallFolder(version) + 'winword.exe') then
          Result := True      
       else if version='14' then
          Result:= ClickToRun
       else if version='15' then
          Result:= ClickToRun15
       else if version='16' then
          Result:= ClickToRun16
       else 
          Result:=false;
end;

// er office installeret?  IsWordInstalled('12') eller IsWordInstalled('14')
function IsExcelInstalled(version: string): boolean;
begin         
       if FileExists(WordInstallFolder(version) + 'excel.exe') then
          Result := True      
      else if version='14' then
          Result:= ClickToRun
      else if version='15' then
          Result:= ClickToRun15
      else if version='16' then
          Result:= ClickToRun16
       else 
          Result:=false;
end;


//  RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe','Path', sWordVar);
//  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Word.Application\CurVer', '', sWordVar)
//  if sWordVar ='Word.Application.' + version then
//    Result := True
//  else
//    Result := False;


function CloseWord():boolean;
var
  winHwnd: longint;
  retVal: boolean;
  indC: integer;
begin
   Result := True;
   try
    winHwnd := FindWindowByClassName('OpusApp');
   except
   end;
   if winHwnd <> 0 then begin
   if MsgBox(ExpandConstant('{cm:CloseWord}'), mbInformation, MB_OKCANCEL)=IDOK then begin
     result:=true;  // even if WOrd is not succesfully close, continue
     while (winHwnd <> 0) and (indC<5) do begin     
       indC:=indC + 1;
       try
       winHwnd := FindWindowByClassName('OpusApp');
       except
       end;
       if winHwnd <> 0 then begin
         retVal:=postmessage(winHwnd,18,0,0); // 18=WM_CLOSE
//       MsgBox(ExpandConstant('{cm:CloseWord2}'), mbInformation, MB_OK);
         try
         winHwnd := FindWindowByClassName('OpusApp');
         except
         end;
        end; //if
      end; //while
   end else begin            
     result:=false;  // ændret v. 1.18 for at installation stadig fortsætter
   end; //if
   end;   
end;

function IsDotNetDetected(version: string; service: cardinal): boolean;
// Indicates whether the specified version and service pack of the .NET Framework is installed.
//
// version -- Specify one of these strings for the required .NET Framework version:
//    'v1.1.4322'     .NET Framework 1.1
//    'v2.0.50727'    .NET Framework 2.0
//    'v3.0'          .NET Framework 3.0
//    'v3.5'          .NET Framework 3.5
//    'v4\Client'     .NET Framework 4.0 Client Profile
//    'v4\Full'       .NET Framework 4.0 Full Installation
//
// service -- Specify any non-negative integer for the required service pack level:
//    0               No service packs required
//    1, 2, etc.      Service pack 1, 2, etc. required
var
    key: string;
    install, serviceCount: cardinal;
    success: boolean;
begin
    key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\' + version;
    // .NET 3.0 uses value InstallSuccess in subkey Setup
    if Pos('v3.0', version) = 1 then begin
        success := RegQueryDWordValue(HKLM, key + '\Setup', 'InstallSuccess', install);
    end else begin
        success := RegQueryDWordValue(HKLM, key, 'Install', install);
    end;
    // .NET 4.0 uses value Servicing instead of SP
    if Pos('v4', version) = 1 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Servicing', serviceCount);
    end else begin
        success := success and RegQueryDWordValue(HKLM, key, 'SP', serviceCount);
    end;
    result := success and (install = 1) and (serviceCount >= service);
end;

function IsDotNet40Detected(): boolean;
begin
  Result := IsDotNetDetected('v4\Client', 0);
end;


// Checks if macros are disabled if not they will be enabled in the registry section then
// tilsyneladende ikke nødvendigt når trusted locations bruges, men hvis excelfiler gemmes andre steder giver det problemer
// er åbenbart alligevel nødvendigt da fejl fundet hos bruger blev løst ved at sætte denne.
// IsMacrosDisabled("12","Word")
function IsMacrosDisabled(version:string;offapp:string):Boolean;
var
vbawarn:Cardinal;
begin
Result:=False;
if IsWordInstalled(version) then
if RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\Microsoft\Office\' + version + '.0\' + offapp + '\Security\',
     'VBAWarnings', vbawarn) then
  begin
    // Successfully read the value      vbawarn=2 er makroer med meddelelse  vbawarn=1 er aktiver alle makroer  vbawarn=4 er deaktiver alle makroer
    if (vbawarn=4) or (vbawarn=3) then
       if MsgBox(ExpandConstant('{cm:MakroDeakt} ') + offapp + #13#13 + ExpandConstant('{cm:MakroDeakt2}'), mbInformation, MB_OKCANCEL)=IDOK then 
            Result:=True;
  end
end;

procedure ShowParameters() ;
begin
    MsgBox('Installationsparametre:'#13#13
    'ClickToRun: ' + BoolToString(ClickToRun) + ''#13   
    'Office12Folder:' + Office12Folder + ''#13
    'Office14Folder:' + Office14Folder + ''#13
    'UserStartupFolder12:' + UserStartupFolder12 + ''#13
    'UserStartupFolder14:' + UserStartupFolder14 + ''#13
    'Word12Installed: ' + BoolToString(Word12Installed) + ''#13   
    'Word14Installed: ' + BoolToString(Word14Installed) + ''#13   
    'Excel12Installed: ' + BoolToString(Excel12Installed) + ''#13   
    'Excel14Installed: ' + BoolToString(Excel14Installed) + ''#13   
     , mbInformation, MB_OK);
end;

//==================================================================
{ Allows for standard command line parsing assuming a key/value organization }
function GetCommandlineParam (inParam: String):String;
var
  LoopVar : Integer;
  BreakLoop : Boolean;
begin
  { Init the variable to known values }
  LoopVar :=0;
  Result := '';
  BreakLoop := False;

  { Loop through the passed in arry to find the parameter }
  while ( (LoopVar < ParamCount) and
          (not BreakLoop) ) do
  begin
    { Determine if the looked for parameter is the next value }
    if ( (ParamStr(LoopVar) = inParam) and
         ( (LoopVar+1) <= ParamCount )) then
    begin
      { Set the return result equal to the next command line parameter }
      Result := ParamStr(LoopVar+1);

      { Break the loop }
      BreakLoop := True;
    end;

    { Increment the loop variable }
    LoopVar := LoopVar + 1;
  end;
end;
function ExistCommandlineParam (inParam: String):Boolean;
var
  LoopVar : Integer;
  BreakLoop : Boolean;
begin
  { Init the variable to known values }
  LoopVar :=0;
  Result := False;
  BreakLoop := False;

  { Loop through the passed in arry to find the parameter }
  while ( (LoopVar <= ParamCount) and
          (not BreakLoop) ) do
  begin
    { Determine if the looked for parameter is the next value }
//          MsgBox(ParamStr(LoopVar), mbInformation, MB_OK);

    if ( ParamStr(LoopVar) = inParam) then
    begin
      { Set the return result equal to the next command line parameter }
      Result := True;

      { Break the loop }
      BreakLoop := True;
    end;

    { Increment the loop variable }
    LoopVar := LoopVar + 1;
  end;
end;

function InstallGeogebra():Boolean;
begin
  if (ExistCommandlineParam('/NoGeoGebra')) then
    Result:=false
  else
    Result:=true;
end;
function InstallGraph():Boolean;
begin
  if (ExistCommandlineParam('/NoGraph')) then
    Result:=false
  else
    Result:=true;
end;


function InitializeSetup(): Boolean;
var 
  ResultCode:Integer;
  uninsval: Integer;
begin
//RegWriteStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings\','Test', 'Hej Dennis');
//          MsgBox(GetCmdTail, mbInformation, MB_OK);

  AutoStart:=0;
  CheckForUpdate:=0;
  AntalBeregninger:=0;

  ClickToRun:=CheckClickToRun();
  Office12Folder:=WordInstallFolder('12');
  Office14Folder:=WordInstallFolder('14');
  Office15Folder:=WordInstallFolder('15');
  Office16Folder:=WordInstallFolder('16');
  UserStartupFolder12:=StartupFolder('12');
  UserStartupFolder14:=StartupFolder('14');
  UserStartupFolder15:=StartupFolder('15');
  UserStartupFolder16:=StartupFolder('16');
  Word12Installed:=IsWordInstalled('12');
  Word14Installed:=IsWordInstalled('14');
  Word15Installed:=IsWordInstalled('15');
  Word16Installed:=IsWordInstalled('16');
  Excel12Installed:=IsExcelInstalled('12');
  Excel14Installed:=IsExcelInstalled('14');
  Excel15Installed:=IsExcelInstalled('15');
  Excel16Installed:=IsExcelInstalled('16');
  Office12Installed:=Word12Installed and Excel12Installed;
  Office14Installed:=Word14Installed and Excel14Installed;
  Office15Installed:=Word15Installed and Excel15Installed;
  Office16Installed:=Word16Installed and Excel16Installed;
  OverWriteDocs:=OverWriteFigurer();

  //ShowParameters();



  Result := True;
  if not (Word12Installed or Word14Installed or Word15Installed or Word16Installed) then begin
    if MsgBox(ExpandConstant('{cm:NoWord}'), mbInformation, MB_YESNO) = IDYES then
      begin
      Result := True;
      Office16Installed:=true;
      Word16Installed:=true;
      Excel16Installed:=true;
      end
    else
      Result := False;
  end
  else if not (Excel12Installed or Excel14Installed or Excel15Installed or Excel16Installed) then begin 
    if MsgBox(ExpandConstant('{cm:NoExcel}'), mbInformation, MB_YESNO) = IDYES then
      begin
      Result := True;
      Excel16Installed:=true;
      end
    else
      Result := False; 
  end
  else if (not IsAdminLoggedOn()) then
     begin
                MsgBox(ExpandConstant('{cm:NotAdmin}'), mbInformation, MB_OK); 
                  Result:=False;
     end
  else if not CloseWord() then
    Result := False
  else begin
    
    if not IsDotNetDetected('v4\Client', 0) then begin
      MsgBox(ExpandConstant('{cm:NoDotNet}'), mbInformation, MB_OK);
 //     ShellExec('open','explorer.exe','http://download.microsoft.com/download/1/B/E/1BE39E79-7E39-46A3-96FF-047F95396215/dotNetFx40_Full_setup.exe','', SW_SHOW, ewNoWait, ErrorCode);
 ////ShellExec('open','explorer.exe','http://www.update.microsoft.com','', SW_SHOW, ewNoWait, ErrorCode);
     ExtractTemporaryFile('dotNetFx40_Client_setup.exe');

    if Exec(ExpandConstant('{tmp}\dotNetFx40_Client_setup.exe'), '', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
                    begin

                        // handle success if necessary; ResultCode contains the exit code
                        if not (ResultCode = 0) then
                            begin
                                Result := false;
                            end
                    end
                    else
                        begin
                            // handle failure if necessary; ResultCode contains the error code
                             MsgBox(ExpandConstant('{cm:InstallError}'), mbInformation, MB_OK);
                            Result := false;
                        end
                    end;
         end; // for adminloggedon

     if Result=true then 
     begin
       uninsval := UnInstallOldVersion()
      if not(WizardSilent()) then
      begin  
       if uninsval=3 then 
           MsgBox(ExpandConstant('{cm:UnInstallOK}'), mbInformation, MB_OK)
       else if uninsval=2 then
           MsgBox(ExpandConstant('{cm:UnInstallError}'), mbInformation, MB_OK)
       else if uninsval=4 then
           Result := False;
     end;
    end;
     end;


procedure DeinitializeSetup();
begin
//     if not FileExists(Office12Folder + 'STARTUP\WordMat.dotm') then
  //         if not FileExists(Office14Folder + 'STARTUP\WordMat.dotm') then
    //            if not FileExists(Office12Folder + 'STARTUP\WordMat.dotm') then
    if AntalBeregninger>0 then
    begin
    RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings\', 'AntalBeregninger', AntalBeregninger);  // ikke slet alle beregninger ved nyinstallation
//    if IsTaskSelected('AutoStart') then
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings\', 'AutoStart', AutoStart);
//    if IsTaskSelected('CheckForUpdate') then
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings\', 'CheckForUpdate', CheckForUpdate);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','AllTrig', AllTrig);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Complex', Complex);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Exact', Exact);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ExcelEmbed', ExcelEmbed);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Forklaring', Forklaring);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Gangetegn', Gangetegn);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LogOutput', LogOutput);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','MaximaCommand', MaximaCommand);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Radians', Radians);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Separator', Separator);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','SigFig', SigFig );
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','SolveBoolOrSet', SolveBoolOrSet);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Units', Units);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','VidNot', VidNot);
      RegWriteStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','OutUnits', OutUnits);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Index', mIndex);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BigFloat', BigFloat);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ShowAssum', ShowAssum);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','PolarOutput', PolarOutput);

      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','GraphApp', GraphApp);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Language', Language);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','dAsDiffChr', dAsDiffChr);
      RegWriteStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexStart', LatexStart);
      RegWriteStringValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexSlut', LatexSlut);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','LatexUnits', LatexUnits);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','ConvertTexWithMaxima', ConvertTexWithMaxima);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqNumPlacement', EqNumPlacement);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqNumType', EqNumType);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','EqAskRef', EqAskRef);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','Backup', Backup);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupNo', BackupNo);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupMaxNo', BackupMaxNo);
      RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\WordMat\Settings','BackupTime', BackupTime);

     end;
end;

function InitializeUninstall(): Boolean;
begin
  ClickToRun:=CheckClickToRun();
  Office12Folder:=WordInstallFolder('12');
  Office14Folder:=WordInstallFolder('14');
  Office15Folder:=WordInstallFolder('15');
  Office16Folder:=WordInstallFolder('16');
  UserStartupFolder12:=StartupFolder('12');
  UserStartupFolder14:=StartupFolder('14');
  UserStartupFolder15:=StartupFolder('15');
  UserStartupFolder16:=StartupFolder('16');
  Result := CloseWord();
end;

{ Check if a path contains spaces.  If it does, convert it to
  the equivalent short path }
function PathWithoutSpaces( strIn: String): String;
begin
  if (Pos(' ',strIn) = 0) then
    Result := strIn
  else
    Result := GetShortName(strIn);
end;
{ Based on code from Inno Setup Extensions Knowledge Base
  Article 14 - How to replace a line in a text file
  http://www13.brinkster.com/vincenzog/isxart.asp?idart=14
  Author: Stefan Bracke }

{ Note: Functions called by AfterInstall can have maximum of one argument }

procedure ReplaceApp(MaximaPath: String);
var
  iLineCounter : Integer;
  a_strTextfile : TArrayOfString;
  strFilename : String;
  strFind : String;
  strNewLine : String;
  strApp : String;
begin
  strApp := ExpandConstant('{app}');
{  strFilename := strApp + '\Maxima-5.30.0\bin\maxima.bat';}
  strFilename := strApp + '\' + MaximaPath;
  strFind := 'set maxima_prefix';
  strApp := strApp + '\Maxima-5.45.1\';
  strNewLine := 'set maxima_prefix='+PathWithoutSpaces(strApp);
  
  { Load textfile into string array }
  LoadStringsFromFile(strFilename, a_strTextfile);
  
  { Search through all textlines for given text, only replace first}
  for iLineCounter := 0 to GetArrayLength(a_strTextfile)-1 do
    begin
      { Overwrite textline when text searched for is part of it }
      if (Pos(strFind, a_strTextfile[iLineCounter]) > 0) then
        a_strTextfile[iLineCounter] := strNewLine;
        Break;
    end;

  { Save string array to textfile (overwrite, no append!) }
  SaveStringsToFile(strFilename, a_strTextfile, False);

end;

// this I believe is deprecated
procedure ReplacePathInBat();
begin
//  ReplaceApp('Maxima-5.45.1\bin\maxima.bat');
//  ReplaceApp('Maxima-5.45.1\bin\maximaunit.bat');  
end;

