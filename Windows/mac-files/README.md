#Mac files

Mac files are package files which can be loaded into Maxima with new functions.

**Solvereal.mac**
Is the main file which is loaded into the original maxima.core.
It adds many functions which WordMat uses.

Any changes to this file requies the generation of a new maxima.core for the maxima installation on both windows and mac.
It must be placed in the following folder: Maxima-sbcl-5.38.1\share\maxima\5.38.1\share\contrib
It can also be used by other Maxima users. just load(solvereal)

**WordMatunitaddon.mac**
This file is part of the maximaunit.core on windows.
on mac it is loaded when units are turned on.
It mainly loads unit.mac which makes maxima able to calculate with units. 