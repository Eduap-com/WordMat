# Windows-folder
This folder contains all the files required to build WordMat for Windows.
Run the inno setup script: ‘WordMat-InnoScript’ to create the WordMat.exe file

You must install Inno Setup to build and edit the script-file: [Inno Setup](https://jrsoftware.org/isinfo.php) or [Inno script studio](https://www.kymoto.org/products/inno-script-studio/)

In the root folder:
WordMat.dotm   - Contains the main code in Visual Basic for Applications (VBA)
Right click the file and choose open. Then press alt+F11 to open the VBE and see the code.

Normally the exe-file is codesigned before release

##How to work with Git-hub and VBA-modules
The idea is to always end by exporting all vba-modules to the folder ‘VBA-modules’. Whenever you start to code, start by importing. The advantage of this method is that Git can then track changes in the exported modules. It cannot do that within a word-file.
Functions in VBAmodul:
•	ImportAllModules
Imports all modules from subfolder ‘VBA-modules’ except VBAmodul.bas
Overwrites existing modules. You are prompted to remove all existing.
•	ExportAllModules
Exports all forms, classes and modules from current project to subfolder ‘VBA-modules’
(including VBAmodul) All current content in folder is deleted before export
•	RemoveAllModules
Removes all forms, class and modules from current project, except VBAmodul.bas.
Can be used to create a new WordMat-clean from an existing WordMat file
The work process is:
1.	You have a WordMat.dotm file you are working on in folder WordMat/Windows
2.	When done coding run ‘ExportAllModules’
3.	Push to github
4.	When resuming coding start by pulling from github
5.	Open your WordMat.dotm file and start by running ‘ImportAllModules’. Click ok to remove existing.
(Import and export can also be done using MZ-tools as an alternative to the VBA-modules in this file. To use MZ-tools right click the project in VBA-ide and choose ‘Import Files….’)
If you make changes to the menus make sure this file is updated.
For the VBA-modules to work you need:
•	Add reference 'Microsoft Visual Basic for Applications Extensibility 5.3' in VBA-Ide tools/references
•	Settings | Trust center | Settings for macros | always trust VBA project object model

