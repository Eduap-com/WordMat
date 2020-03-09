# Maxima-files

Mac files are package files which can be loaded into Maxima with new functions.
The files listed below are the changes made to the original maxima installation, which are identical on Windows and Mac.
The Windows installer copies these files automatically to the Maxima-installation
On Mac they must manually be placed in the correct path when they are changed.


**Solvereal.mac**
Is the main file which is loaded into the original maxima.core.
It adds many functions which WordMat uses.

Any changes to this file requies the generation of a new maxima.core for the maxima installation on both windows and mac.
It must be placed in the following folder: *Maxima-sbcl-5.38.1\share\maxima\5.38.1\share\contrib*
It can also be used by other Maxima users. just load(solvereal)

**WordMatunitaddon.mac**
This file is part of the maximaunit.core on windows.
on mac it is loaded when units are turned on.
It mainly loads unit.mac which makes maxima able to calculate with units. 

**Unit.mac**
This file is change from the original file, to add more units.
It is placed in *Maxima-sbcl-5.38.1\share\maxima\5.38.1\share\contrib\unit* by the Windows installer


**Other changes to the maxima installation are:**
- maxima.core file is changed on both Windows and Mac, but must be build separately (*Maxima-sbcl-5.38.1\lib\maxima\5.38.1\binary-sbcl\maxima.core*)
- maximaunit.core is added on Windows
- On Windows *Maxima-sbcl-5.38.1\bin\maxima.bat* has been modified to fit the path in the WordMat folder
- On Mac *Contents/Resources/maxima.sh* has been modified to fit the path to maxima.app
