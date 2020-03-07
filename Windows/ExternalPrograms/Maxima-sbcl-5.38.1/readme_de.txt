Dies ist die Windows-LIESMICH-Datei.


In dieser Distribution enthaltende Bin�rdateien
-----------------------------------------------

Die Windows-Version des Maxima-Pakets enth�lt Bin�rdateien 
aus anderen Open-Source-Projekten, die ebenfalls auf Sourceforge 
gewartet werden.


gcc: 

gcc.exe, cc1.exe und die Dateien in den lib/gcc-lib und include/ 
Unterverzeichnissen stammen aus der mingw-Version des gcc.  Diese
ist erh�ltlich unter http://prdownloads.sf.net/mingw/
     

binutils:

as.exe stammt vom mingw (http://www.mingw.org/)-Port der binutils
erh�ltlich unter http://prdownloads.sf.net/mingw/


gnuplot:

Die Dateien wgnuplot.exe, wgnuplot.hlp und wgnuplot.mnu stammen 
aus der Windows-Distribution von gnuplot unter 
http://gnuplot.sourceforge.net


wxMaxima:

Die Dateien im wxMaxima-Unterverzeichnis stammen aus der Windows-
Distribution von wxMaxima erh�ltlich unter 
http://wxmaxima.sourceforge.net


Maxima-GUI und Firewall
-----------------------

Manchmal kann die Maxima-GUI (xmaxima oder wxMaxima) Maxima nicht 
starten, gibt Zeit�berschreitungsmeldungen aus oder erh�lt auf 
Maxima-Kommandos keine Antworten. Sehr wahrscheinlich ist dieses 
Problem dann durch die Firewall und/oder eine Antivirus-Software 
verursacht. Die Maxima-GUI kommuniziert mit der Berechnungseinheit 
�ber ein Socket. Antivirus- und/oder Firewall-Progamme erkennen 
das und versuchen eventuell, dies zu blocken 
(da einige sch�dliche Programme ebenfalls Sockets �ffnen). 

Zur L�sung des Problems:

1.  Versuchen Sie, die Systemsteuerung der Antivirus- und/oder 
    Firewall-Software zu finden.

2.  Finden Sie die Maxima-GUI auf der Liste der geblockten Programme 
    und deaktivieren Sie die Blockierung der GUI. Das GUI-Programm 
    wird eventuell als "Tcl/Tk" aufgef�hrt 
    (der Name des GUI-Werkzeugsatzes f�r xmaxima).


Datenausf�hrungsverhinderung (DEP)
----------------------------------

In einigen F�llen funktioniert weder Maxima GUI noch 
Kommandozeilen-Maxima (maxima.bat startet und beendet sich sofort wieder).
Sehr wahrscheinlich h�ngt das Problem mit Windwos DEP zusammen.
Einige Lisp-Implementierungen f�hren Code in Datenbereichen des 
Arbeitsspeichers aus. Windows DEP bemerkt und blockiert dies (weil
einige sch�dliche Programme ebenfalls Code in Datenbereichen ausf�hren).

L�sung:

   F�gen Sie den vollen Programmpfad von Maxima (z.B. 
   C:\Programme\Maxima-5.12.0\lib\maxima\5.12.0\binary-gcl\maxima.exe)
   in die Liste von DEP-Ausnahmen ein  
   (Systemsteuerung -> System ->  Erweitert -> Systemleistung "Einstellung" 
    -> Datenausf�hrungsverhinderung)


Zu wenig Umgebungsspeicher
--------------------------

Wenn Sie versuchen, Maxima auf einer Windows 9x-Maschine laufen 
zu lassen, erhalten sie eventuell die Fehlermeldung 

    Zu wenig Umgebungsspeicher (Out of environment space)

Der folgende Hinweis des Artikels 230205 der Microsoft Knowledge Base 
(http://support.microsoft.com/support/kb/articles/Q230/2/05.ASP)
k�nnte Ihnen helfen.


URSACHE
 
Dieses Problem kann auftreten, wenn Sie nicht �ber ausreichend 
Arbeitsspeicher in der MS-DOS-Umgebung verf�gen, um eine 
Umgebungsvariable festzulegen. 

L�SUNG 

Erh�hen Sie den f�r MS-DOS-Programme verf�gbaren vorgegebenen 
Umgebungsspeicherplatz, um dieses Problem zu beheben. Wenden Sie 
zu diesem Zweck eine der folgenden Methoden an. 


�ndern Sie f�r alle MS-DOS-Programme die Umgebung

Bearbeiten Sie den Shell-Befehl in der Datei Config.sys, um den 
vorgegebenen Umgebungsspeicherplatz f�r alle MS-DOS-Programme zu erh�hen, 
die in Windows ausgef�hrt werden. Gehen Sie hierzu folgenderma�en vor:

1. Klicken Sie auf Start und klicken Sie dann auf Ausf�hren.  

2. Geben Sie in dem Feld �ffnen Sysedit ein und klicken Sie dann auf OK.  

3. Klicken Sie auf das Fenster Config.sys.  

4. Geben Sie am Anfang der Shell= -Zeile REM ein und dr�cken anschlie�end 
   die LEERTASTE. 
   Sollte die Shell= -Zeile nicht vorhanden sein, fahren Sie mit dem 
   n�chsten Schritt fort.  

5. Dr�cken Sie die Taste POS1.  

6. Um eine neue Shell= -Zeile zu erstellen, geben Sie die folgende Zeile 
   ein und dr�cken anschlie�end die EINGABETASTE: 

      SHELL=C:\COMMAND.COM /E:4096 /P 

7. Klicken Sie in dem Men� Datei auf Speichern.  

8. Klicken Sie in dem Men� Datei auf Beenden.  

9. Starten Sie den Computer erneut.  


�ndern Sie die Umgebung f�r ein bestimmtes MS-DOS-Programm 

Gehen Sie folgenderma�en vor, um den vorgegebenen Umgebungsspeicherplatz nur 
f�r maxima.bat zu erh�hen: 

1. Klicken Sie mit der rechten Maustaste auf das maxima.bat-Programmsymbol 
   und klicken Sie dann auf Eigenschaft.  

2. Klicken Sie auf die Registerkarte Speicher.  

3. Geben Sie in dem Feld Anf�nglicher Umgebungsspeicher die Zahl von 
   Kilobytes (KB), die f�r das Programm ben�tigt werden, ein und 
   klicken Sie dann auf OK. 

Hinweis: 4096 KB ist die maximale Gr��e des Speichers, die Sie f�r ein 
MS-DOS-Programm reservieren k�nnen.  



