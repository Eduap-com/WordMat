Ceci est la version Windows du fichier README.

Fichiers binaires inclus dans cette distribution.
-------------------------------------------------

La version pour Windows de ce paquage de Maxima inclut
des fichiers binaires d'autres projets Open Source �galement
disponible sur Sourceforge.

gcc: 

gcc.exe, cc1.exe et les fichiers dans les sous r�pertoires
lib/gcc-lib et include/ sont issue de la version mingw de gcc. 
Lequel est disponible � l'adresse http://prdownloads.sf.net/mingw/
     

binutils:

as.exe viens du portage mingw (http://www.mingw.org/) de binutils
disponible � l'adresse http://prdownloads.sf.net/mingw/


gnuplot:

Les fichiers wgnuplot.exe, wgnuplot.hlp et wgnuplot.mnu sont issus
de la distribution Windows de gnuplot disponible sur
http://gnuplot.sourceforge.net


wxMaxima:

Les fichiers dans le sou r�pertoire wxMaxima sont la distribution 
Windows de wxMaxima disponible sur http://wxmaxima.sourceforge.net


Interface graphique utilisateur (GUI) Maxima et pare-feux
---------------------------------------------------------

Parfois l'interface graphique utilisateur de Maxima (xmaxima ou wxMaxima) 
ne peut pas lancer Maxima ou produit un message de d�passement de temps
ou ne re�ois pas de r�ponse pour les commandes Maxima.
Le probl�me est fort probablement provoqu� par un pare-feux ou un
logiciel anti-virus.
La GUI de Maxima communique avec le moteur de calcul � travers une socket.
Le programme d'anti-virus et/ou le pare-feux voie cela et peut essayer de bloquer 
cette communication. (�tant donn�e que certain programmes malicieux ouvrent
des sockets)

Pour r�soudre le probl�me:

1.  Essayez de trouver le panneau de contr�le de l'anti-virus et/ou du pare-feux.

2.  S�lectionnez la GUI Maxima dans la liste des programmes bloqu�s
     et d�sactivez le blocage pour ce programme. Le programme de la
     GUI pourrait appara�tre comme "Tcl/Tk" (c'est le nom de la boite � outils de
     la GUI pour xmaxima)


Pr�vention de l'Ex�cution des Donn�es (PED) 
------------------------------------------- 

Il arrive que non seulement l'interface graphique de Maxima mais m�me Maxima 
en ligne de commandes ne fonctionne pas (maxima.bat se lance et quitte 
aussit�t).� Le probl�me vient tr�s probablement de la PED de Windows. Quelques 
impl�mentations de LISP ex�cutent du code dans des secteurs de donn�es de la 
m�moire. La PED de Windows le voit et le bloque (car certains programmes 
malveillants ex�cutent aussi du code dans des secteurs de donn�es). 

Solution : 

Inclure le chemin complet du programme ex�cutable Maxima, par exemple : 
C:\Program Files\Maxima-5.12.0\lib\maxima\5.12.0\binary-gcl\maxima.exe 
dans la liste des exceptions PED 
(Panneau de configuration - > Syst�me - > Avanc� - > Performances - > PED) 


Espace d'environnement insuffisant
----------------------------------

Lorsque vous essayez d'ex�cuter Maxima sur une machine Windows 9x,
le message d'erreur suivant risque de s'afficher  :

    Espace d'environnement insuffisant (Out of environment space)

Le conseil suivant de la Base de Connaissance Microsoft, Article 230205
(http://support.microsoft.com/support/kb/articles/Q230/2/05.ASP) 
peut �tre une aide utile.

CAUSE

Ce probl�me peut se produire si vous ne disposez pas, dans l'environnement 
MS-DOS, de m�moire suffisante pour d�finir une variable d'environnement. 

RESOLUTION

Pour r�soudre ce probl�me, augmentez l'espace par d�faut disponible pour les
programmes dans l'environnement MS-DOS. Pour cela, appliquez l'une des 
m�thodes suivantes.

Modifiez l'environnement pour tous les programmes MS-DOS

Pour augmenter l'espace d'environnement par d�faut pour tous les programmes 
MS-DOS en cours d'ex�cution dans Windows, modifiez la commande Shell dans 
le fichier Config.sys. Pour cela suivez cette proc�dure : 

1.  Cliquez sur D�marrer et puis cliquez sur Ex�cuter. 

2.  Dans la zone Ouvrir, tapez sysedit et puis cliquez sur OK. 

3.  Cliquez sur la fen�tre Config.sys. 

4.  Au d�but de la ligne Shell= line, entrez REM, et pressez la bare d'espace 

      si la ligne Shell= line n'existe pas passez � l'�tape suivante. 

5.  Appuyez sur la touche  HOME. 

6.  Tapez la ligne suivante pour cr�er une nouvelle commande Shell= line, et
    presser ENTER: 

      SHELL=C:\COMMAND.COM /E:4096 /P 

7.  Dans le menu Fichier, cliquez sur Enregistrer. 

8.  Dans le menu Fichier, cliquez sur Quitter. 

9.  Red�marrer l'ordinateur. 


Modifiez l'environnement pour un programme MS-DOS sp�cifique

Pour augmenter l'espace d'environnement par d�faut pour un programme MS-DOS
sp�cifique, proc�dez comme suit : 

1.  Cliquez avec le bouton droit sur l'ic�ne du programme maxima.bat et puis cliquez sur Propri�t�s. 

2.  Cliquez sur l'onglet M�moire. 

3.  Dans la zone Environnement initial, tapez le nombre de kilo-octets (Ko) que le 
      programme requiert et puis cliquez sur OK. 

REMARQUE : 4096 KB repr�sente le montant maximal de m�moire que vous pouvez 
allouer pour un programme MS-DOS.

