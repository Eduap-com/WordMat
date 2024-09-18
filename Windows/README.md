# Windows-folder
This folder contains all the files required to install WordMat for Windows.

In the root folder:
WordMat.dotm   - Contains the main code in Visual Basic for Applications (VBA) minus VBAmodul
WordMatDev.dotm   - devolper version. Rename to WordMat.dotm and remove VBAmodul + ThisDocument content
Right click the file and choose open. Then press alt+F11 to open the VBE and see the code.

##How to edit on Mac and Windows
A dotm-file edited on Windows can be opened on Mac without problems. If a dotm-file is edited on mac it will have scrambled speciel characters when opened on Windows again in the code. Example characters that are scrambled: øæåáéóÀ 
If this has happened, just run the macro: ConvertVBECharactersFromMacToWindows  (alt+F8)
How to convert this file to Mac:
•	Make a copy of WordMat.dotm called WordMatMac.dotm 
•	(Deprecated: Open WordMatMac.dotm and Run ‘ReplaceToNonUnicode’ on Windows (Alt+F8))
•	Synk to Mac 
•	Open WordMatMac.dotm on a Mac
•	(Deprecated: Run ’ReplaceToUnicode’ (fn+Alt+F8))
•	Check if the code compiles on Mac (fn+Alt+F11)
•	Reset all keyboard shortcuts. Tools ´customize keyboard
•	Run ‘GenerateKeyboardShortcuts’
•	Check if equations are set to unicode and NOT latex
Do not upgrade this file to a newer Word-version. It will break menu compatibility.
How to work with Git-hub and VBA-modules
The idea is to always end by exporting all vba-modules to the folder ‘VBA-modules’. Whenever you start to code, start by importing. The advantage of this method is that Git can then track changes in the exported modules, and multiple people can work on the VBA-code. It cannot do that within a word-file.
Functions in VBAmodul:
•	ImportAllModules
Imports all forms, classes and modules from subfolder ‘VBA-modules’. Removes all existing modules before import, including VBAmodul
•	ExportAllModules
Exports all forms, classes and modules from current project to subfolder ‘VBA-modules’
(including VBAmodul) All current content in folder is deleted before export. A logfile with date is added to the folder.
•	RemoveAllModules
Removes all forms, classes and modules from current project, except VBAmodul.bas.
The work process is:
1.	You have a WordMat.dotm file you are working on in folder WordMat/Windows
2.	When done coding run ‘ExportAllModules’ (Click Alt+F8)
3.	Push to github origin
4.	When resuming coding start by pulling from github
5.	Open your WordMat.dotm file and start by running ‘ImportAllModules’. Click ok to remove existing.
(Import and export can also be done using MZ-tools as an alternative to the VBA-modules in this file. To use MZ-tools right click the project in VBA-ide and choose ‘Import Files….’)
If you make changes to the Ribbon make sure this file is updated.
For the VBA-modules to work you need:
•	Add reference 'Microsoft Visual Basic for Applications Extensibility 5.3' in VBA-Ide tools/references
•	Settings | Trust center | Settings for macros | always trust VBA project object model


