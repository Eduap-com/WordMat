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
- unit.mac is changed in Maxima-sbcl-5.38.1\share\maxima\5.38.1\share\contrib\unit
- maxima.core is changed in Maxima-sbcl-5.38.1\lib\maxima\5.38.1\binary-sbcl
- maximaunit.core is added
- maxima.bat changed in Maxima-sbcl-5.38.1
- maximaunit.bat added
- solvereal.mac added to contrib
- Wordmatunitaddon.mac added to contrib

**dotnet...**
.net installer used by the installer if the computer does not have .net. 4.0
Probably depracated as this is standard for many years.