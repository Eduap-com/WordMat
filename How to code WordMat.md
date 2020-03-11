# How to code WordMat

## Intro
This document describes the structure of the WordMat code and how to go about changing or adding to the code.

*Mikael Samsøe Sørensen*

## WordMat.dotm
Holds the main code as Visual Basic for Application (VBA) code.
Word documents can have macros (VBA code) attached to the document. Word documents with VBA code have the extension .dotm or .dotx in stead of .docx.
Since the WordMat.dotm file is a Word template double clicking the file will create a copy of the original file, to edit the original file right click and open.
Once the file is open you can click Alt+F11 to open the Visual Basic development Environment (VBE).
You can also go to **Files | settings | Customize ribbon** and put a checkmark next to developer. Which will add a new menu to the ribbon in Word, where you can also open the VBE by clicking *Visual Basic*.
From this menu you can also launch macros (VBA functions). You can also launch Macros by pressing alt+F8.
In the VBE there are 
- Forms
- Modules
- Class modules

Here is a description of some the most important functions in these modules:

**Module: Maxima**


**Class module:** *Cmaxima*
Holds all the 

**Class module:** *MaximaProcess*
*Function:* 


### WordMat Ribbon
The WordMat ribbon can be edited using the *Custom UI editor* or the *Ribbonx-editor*
Download both here: http://www.rondebruin.nl/win/s2/win001.htm
or ribbonx-editor here: https://github.com/fernandreu/office-ribbonx-editor
The ribbon is made from xml-code linked to events in the VBA code. The events are placed in the *RibbonSubs module*. 
Events are triggered when the buttons are pressed but also on creation to get the language specific text.

## Maxima changes

### SolveReal.mac
When you want to change this code the best approach is the paste it all into WxMaxima, where you can test the functions, and make changes. When done copy/paste back to the .mac file.

## MathMenu.dll
The mathmenu.dll library is referenced from WordMat.dotm VBA code.
```
Set GetMaxProc = CreateObject("MaximaProcessClass")
```

Both these files are required:
- MathMenu.dll
- MathMenu.tlb
The installer registers the library with Windows

The purpose of the library is to communicate safely between Word and Maxima. It is an API.
To edit the source code written in c# you need Visual Studio to open the project. The newly build files will be in the release folder.
The code has not been changed for many years.
It is basically

## MaximaConnectionDylib
Code for the dylib API library written in c to communicate between Word and Maxima on Mac.
Use xcode to edit.

The installer places the library in
*/Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/*
It must be this library for Word to have the correct privilige

The library MaximaConnection.dylib is referenced from the VBA code in WordMatMac.dotm in the class module *MaximaProces* in this manner:
```
Private Declare PtrSafe Function StartMaxima Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" () As LongPtr
```
There must be a line for each of the functions in the library: SetMaximaPath(string), StartMaxima, CloseMaxima, IsMaximaRunning, SendToMaxima(string), ReadOutputIfAny(string)

