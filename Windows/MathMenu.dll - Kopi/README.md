#Mathmenu dll library
This folder holds the Mathmenu.dll library, and source code

The installer uses these two files:
- MathMenu.dll
- MathMenu.tlb

Which must be in the root of MathMenu.dll

The mathmenu.dll library is referenced from WordMat.dotm VBA code.
The purpose of the library is to communicate safely between Word and Maxima. It is an API.
To edit the source code you need Visual Studio.
If you build a new dll and tlb-file, you must place them in the root of this folder.
The newly build files will be in the release folder.

The code has not been changed for many years.