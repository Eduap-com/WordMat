# External programs
This folder holds standalone programs / installers for other open source programs which are distributed as part of WordMat, and which WordMat can communicate rudimentary.

**GeoGebra**
Can easily be updated. It must be GeoGebra version 5. The filename must be changes a couple of places in the inno script when a new version is introduced.
GeoGebra.ggb is an empty file WordMat uses.

**Graph**
Has not had an update i many years.
graphtemplate.grf is an empty file WordMat uses.

**Maxima**
Is not easy to update. The version must match the Mac-version to ensure similar behaior. This version is the most recent version which is build in a way that works on both Mac and Windows.
Also new versions of maxima could potentially require changes to solvereal.mac and wordmat.dotm

This maxima installation is changed from the original.
- unit.mac is replaced in Maxima-5.45.1\share\maxima\5.45.1\share\contrib\unit by the file in shared/Maxima-files
- WordMat/shared/Maxima-files/solvereal.mac added to contrib
- WordMat/shared/Maxima-files/Wordmatunitaddon.mac added to contrib
- maxima.bat changed in Maxima-5.45.1  (A section setting maxima-prefix is copied in)
- maximaunit.bat added     (copy of maxima.bat with maxima.core changed to maximaunit.core in the sbcl.exe line)
- maxima.core is changed in Maxima-5.45.1\lib\maxima\5.45.1\binary-sbcl    (see 'how to build WordMat.md' on how to create the new core)
- maximaunit.core is added

**dotnet...**
.net installer used by the installer if the computer does not have .net. 4.0
Probably depracated as this is standard for many years.