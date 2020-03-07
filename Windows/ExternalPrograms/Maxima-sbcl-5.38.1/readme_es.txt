Fichero LEAME para Windows

Ficheros binarios incluidos en esta distribuci�n
------------------------------------------------

La versi�n de Windows del paquete Maxima incluye ficheros
binarios procedentes de otros proyectos de C�digo Abierto,
tambi�n alojados en Sourceforge.


gcc: 

gcc.exe, cc1.exe y los ficheros de los subdirectorios lib/gcc-lib
y include/ proceden de la versi�n mingw de gcc, que pueden
conseguirse en http://prdownloads.sf.net/mingw/
     

binutils:

as.exe procede de la versi�n espec�fica de binutils de mingw
(http://www.mingw.org/), la cual puede a su vez conseguirse
en http://prdownloads.sf.net/mingw/


gnuplot:

Los ficheros wgnuplot.exe, wgnuplot.hlp y wgnuplot.mnu proceden de
la distribuci�n para Windows de gnuplot, http://gnuplot.sourceforge.net


wxMaxima:

Los ficheros del sudirectorio wxMaxima proceden de la distribuci�n
para Windows de wxMaxima, alojado en http://wxmaxima.sourceforge.net


Entornos gr�ficos de Maxima y cortafuegos
-----------------------------------------

En ocasiones, los entornos gr�ficos de Maxima (xmaxima o wxMaxima)
no pueden iniciar una sesi�n de Maxima, o emiten mensajes relativos
al exceso de tiempo de espera o, simplemente, no obtienen respuesta
alguna de las instrucciones de Maxima. Es probable que la causa de
estos problemas se encuentre en los antivirus y/o cortafuegos que
haya instalados en el sistema. Los entornos gr�ficos conectan con
Maxima a trav�s de sockets, los cuales pueden verse bloqueados por
los antivirus y/o cortafuegos; esto es as� porque algunos programas
da�inos tambi�n utilizan sockets.

Para resolver el problema:

1.  Int�ntese encontrar el panel de control del antivirus y/o cortafuegos.

2.  B�squese el entorno gr�fico de Maxima en la lista de programas
    bloqueados y desact�vese el bloqueo. El entorno gr�fico puede aparecer
    con el nombre "Tcl/Tk", que es el nombre de la herramienta gr�fica de
    xmaxima.


Prevenci�n en la ejecuci�n de datos (DEP)
-----------------------------------------

Algunas veces no s�lo la interfaz gr�fica de Maxima sino tambi�n la
interfaz en l�nea de comandos pueden no trabajar  (maxima.bat inicia y
termina inmediatamente). Probablemente el problema es causado por el sistema
DEP de Windows. Algunas implementaciones de Lisp ejecutan c�digo en �reas
de datos en memoria. El sistema DEP de Windows bloquea estos c�digos
(debido a que algunos programas maliciiosos ejecutan c�digo en �stas
�reas tambi�n)

Soluci�n:

Incluya la ruta completa del ejecutable de Maxima, por ejemplo:
C:\Program Files\Maxima-5.12.0\lib\maxima\5.12.0\binary-gcl\maxima.exe
en la lista de excepciones del sistema DEP:
(Panel de Control -> Sistema -> Avanzado -> Desempe�o -> DEP)


Mensaje de fuera de espacio de entorno
--------------------------------------

Cuando se intenta ejecutar Maxima en una plataforma Windows 9x,
es posible recibir el siguiente mensaje de error:

            Fuera de espacio de entorno 

El art�culo 230205 de la Microsoft Knowledge Base
(http://support.microsoft.com/support/kb/articles/Q230/2/05.ASP)
puede ser de utilidad en estos casos.

CAUSA

Este problema puede ocurrir si no tiene suficiente memoria en el entorno
MS-DOS para establecer una variable de entorno.

SOLUCI�N

Para resolver este problema, se debe aumentar el espacio predeterminado
de entorno disponible para programas de MS-DOS. Para ello, util�cese uno
de los m�todos siguientes.


Modificaci�n del entorno para todos los programas de MS-DOS.

Para aumentar el espacio predeterminado de entorno para todos los programas
de MS-DOS que se ejecutan en Windows, modif�quese el comando Shell en el 
archivo Config.sys; para ello deben seguirse los siguientes pasos:

1. H�gase clic en Inicio y a continuaci�n otro clic en Ejecutar.

2. En el cuadro Abrir, escr�base sysedit y a continuaci�n clic en Aceptar.

3. H�gase clic en la ventana Config.sys.

4. Al comienzo del rengl�n Shell=, escr�base REM y a continuaci�n presi�nese la 
   BARRA ESPACIADORA.

   Si no existe el rengl�n Shell=, contin�ese con el paso siguiente.

5. P�lsese la tecla INICIO.

6. Escr�base el siguiente c�digo para crear un nuevo rengl�n Shell=
   y a continuaci�n presi�nese ENTRAR:

      SHELL=C:\COMMAND.COM /E:4096 /P

7. En el men� Archivo, clic en Guardar.

8. En el men� Archivo, clic en Salir.

9. Rein�ciese el equipo.



Modificaci�n del entorno para un programa espec�fico de MS-DOS

A fin de aumentar el espacio predeterminado de entorno
para maxima.bat, deben seguirse los siguientes pasos:

1. H�gase clic con el bot�n secundario en el icono de maxima.bat y
   a continuaci�n clic en Propiedades.

2. H�gase clic en la ficha Memoria.

3. En el cuadro Entorno inicial, escr�base el n�mero de kilobytes (KB)
   que necesita el programa y a continuaci�n clic en Aceptar. 


NOTA: La cantidad m�xima de memoria que se puede asignar a
      un programa de MS-DOS es 4096 KB.
