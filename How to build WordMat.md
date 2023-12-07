# How to build WordMat

## Intro
Download all the source files to your computer from GitHub. You must then proceed to the Windows or Mac section to build the installer files separately.

*Mikael Samsøe Sørensen*

## Windows
You must install Inno Setup to build and edit the script-file: [Inno Setup](https://jrsoftware.org/isinfo.php) or [Inno script studio](https://www.kymoto.org/products/inno-script-studio/)
To build do the following:
1. Open *WordMat-InnoScript.iss* file
2. Click Projects | Compile  (Alt+F9)
3. Wait until compiler says complete
4. A new *WordMat.exe* file is now created in the Windows folder

Be aware that if you are building a new version there are a number of things you must observe. See section *New release checklist*

## Mac
First install [Packages](http://s.sudre.free.fr/Software/Packages/about.html) (The program that generates the installer-file)

Some paths in the pkgproj file probably needs to be set manually on a new computer. A work in progress.

To build do the following
1. Open WordMat.pkgproj
2. Click Build / Build in the menu
3. Wait until the build log reads: **Build succeded**
4. A new *WordMat.pkg* file is created in the same folder

Be aware that if you are building a new version there are a number of things you must observe, depending on where changes have been made in the code.
See section *New release checklist*

If the build fails because of permission issues a reboot often helps.

# Creating the MaximaWM folder for Mac
A special folder called 'MaximaWM' with all the needed files to run maxima via WordMat must be used.
This must be done for intel and M1 Macs separately. These folders are already made, but in the case they break because of a new Mac update:
1. Install Xcode
2. Install Macports.  macports.org. 
    If already installed just update 'sudo port selfupdate' A new MacOS might require a new install though
3. Install Maxima through Macports 'sudo port install maxima'. (Also see instructions in dmg download from maxima.sourceforge.io if you have problems)
4. (Install gnuplot via macport 'sudo port install gnuplot') 
5. (Install gnu core utils using macport 'sudo port install coreutils')
6. From /opt/local/bin find the folowing files:
sbcl
gtimeout
(gnuplot)
(maxima)
and copy to MaximaWM/maxima/bin
7. Rename gtimeout to 'Maximatimeout'
8. Start maxima:
    Open terminal
    cd to MaximaWM by dragging the path from finder
    './maxima/bin/maxima'
    create maxima.core and maximaunit.core(see description below)
9. Place the core-files in MaximaWM/maxima/lib/maxima/5.47.0/binary-sbcl/

**Structure of MaximaWM**
maxima/bin      contains all executables(sbcl,Maximatimeout)
maxima/info     Must be present for maxima to run
maxima/lib/maxima/5.47.0/binary-sbcl/   maxima.core and maximaunit.core files
maxima/share    Contains math-files that can be loaded into Maxima.
root            maxima.sh used by WordMat to run Maxima (and maybe other scripts to test maxima)

# New release checklist
Whenever a new release is compiled the following checklist must be followed
1. If any changes have been made to SolveReal.mac or WordMatunitAddon.mac a new maxima.core and/or maximaunit.core must be built for both Windows and Mac. See section *Buidling a new maxima.core*
2. Update the version number in WordMat.dotm, the inno script and Package script
3. Open WordMat.dotm and run VBA functions: ‘GenerateKeyboardShortcuts’
*(Make sure to add reference 'Microsoft Visual Basic for Applications Extensibility 5.3')*
*This macros will only run if you have set Word to trust the VBA project object model*
   *Files | Settings | Trust Center | Trust Center Settings | Macro Settings | Trust access to the VBA project object model *
4. If any changes have been made to WordMat.dotm a new WordMatMac.dotm must be created. See section *Creating WordMatMac.dotm for Mac*
5. If any changes has been made to Excelfiles in Windows it must be copied to the corresponding Mac-folder. To avoid special character problems. The copied excel files must first be opened on windows. Then the VBA function 'replacetononunicode' must be run. Then it must opened on Mac, where you run 'replacetounicode'
6. Download a new GeoGebra 5 installer and place it in Externalprograms. The filenames must be changed in the inno script file to reflect the new file.
7. Download a new version of GeoGebra Math apps bundle: https://wiki.geogebra.org/en/Reference:GeoGebra_Apps_Embedding and place in Shared
8. Build the Windows version
9. Codesign the installer (Project owner only)
10. Build the Mac M1 and intel versions
11. Create a new release on GitHub and upload the 3 installers

## Buidling a new maxima.core
The default installation of Maxima loads the compiled *maxima.core* file from this location:
*Maxima-sbcl-5.47.0\lib\maxima\5.47.0\binary-sbcl\\*
It is however posible to load maxima, make some changes and save a new maxima.core file.
This makes it much faster to start up Maxima, instead of having to load all special function and settings upon startup.

### Compiling maxima.core on Windows
- In the following use a normal maxima 5.45.1 installation from the programs folder and not from the GitHub repository
- Copy the most recent versions of solvereal.mac and WordMatunitaddon.mac to *Maxima-sbcl-5.47.0\share\maxima\5.47.0\share\contrib*
- run *Maxima-sbcl-5.47.0\bin\maxima.bat* by right clicking and choosing *'run as administrator'*
If you fail to run as administrator you will recieve a permission denied error in the last step.
- Run the following commands in the command window (You can copy/paste all in one go)
```
load(solvereal)$
load(draw)$
solvereal(x^2=9,x);
:lisp(sb-vm::set-floating-point-modes :traps nil)
linenum:-1;
:lisp (sb-ext:save-lisp-and-die "maxima.core" :toplevel #'cl-user::run)
```

- The command window will now close if everything went fine
- The new *maxima.core* file will be placed in *C:\Windows\System32* or possibly in the same folder as maxima.bat
- Copy *C:\Windows\System32\maxima.core* to *Maxima-sbcl-5.47.0\lib\maxima\5.47.0\binary-sbcl* in the GitHub repository

- Again run *Maxima-sbcl-5.47.0\bin\maxima.bat* as administrator
- Run the following commands in the command window
```
load(solvereal)$
load(draw)$
load(WordMatUnitAddon)$
solvereal(x^2=9,x);
:lisp(sb-vm::set-floating-point-modes :traps nil)
linenum:-1;
:lisp (sb-ext:save-lisp-and-die "maximaunit.core" :toplevel #'cl-user::run)
```

- Copy *C:\Windows\System32\maximaunit.core* to *Maxima-sbcl-5.47.0\lib\maxima\5.47.0\binary-sbcl\\*

**Explanation of commands**
>*:lisp(sb-vm::set-floating-point-modes :traps nil)*  is a command which ensures that Maxima does not crash on overflow errors.

>*linenum:-1;*  resets the linenumbering

>*:lisp (sb-ext:save-lisp-and-die "maxima.core" :toplevel #'cl-user::run)*     saves the core file and exits

The version of Maxima used is the most recent version which I could get to work on both Windows and Mac. Both version are SBCL compiled versions, hence they should behave almost identical.

### Compiling maxima.core on Mac
Must be done separately on intel and M1
- Ensure you have a functioning MaximaWM folder. (See above)
- Copy the most recent versions of solvereal.mac and WordMatunitaddon.mac From Shared/Maxima-files to *MaximaWM/maxima/share/maxima/5.47.0/share/contrib/*
- Copy unit.mac to the subfolder unit of the contrib-folder
- Run maxima
    Open terminal 
    cd to MaximaWM
    ./maxima/bin/maxima   eller 'sh maximarun.sh'
- Run the following commands in the terminal window (Make sure to adjust the path in the last line to your user)
```
*Open solvereal.mac and copy/paste content to terminal*
load(draw)$
solvereal(x^2=9,x);
gnuplot_command:"/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot";
set_plot_option([gnuplot_term, aqua])$
:lisp(sb-vm::set-floating-point-modes :traps nil)
linenum:-1;
:lisp (sb-ext:save-lisp-and-die "/users/mikael/maxima.core" :toplevel #'cl-user::run)
```

- The command window will now close if everything went fine
- The new *maxima.core* file will be placed in *Users/youruser* or same dir as maxima
- Copy *Users/youruser/maxima.core* to  *MaximaWM/maxima/lib/maxima/5.47.0/binary-sbcl/* in the Mac folder in the GitHub repository
- The new maxima.core file can be run using the command 'sbcl --core maxima.core'
    This is handled by maxima.sh
    You can test the new core by cd to MaximaWM. Then run 'sh maximatext.sh'
    
Repeat the process to generate the maximaunit.core, but also copy the content from WordMatUnitaddon.mac into the terminal, after copying from solvereal.mac
```
*Open solvereal.mac and copy/paste content to terminal*
*Open WordMatUnitAddon.mac and copy/paste content to terminal*
load(draw)$
solvereal(x^2=9,x);
gnuplot_command:"/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot";
set_plot_option([gnuplot_term, aqua])$
:lisp(sb-vm::set-floating-point-modes :traps nil)
linenum:-1;
:lisp (sb-ext:save-lisp-and-die "/users/mikael/maximaunit.core" :toplevel #'cl-user::run)
```

## Creating WordMatMac.dotm for Mac
Whenever there are changes to the VBA code in WordMat.dotm a new Mac-version of the same file must be prepared. This file is called *WordMatMac.dotm*
Changes in the VBA code must always take place in the Windows version, because Mac can mess up the file. 
WordMatMac.dotm is basically a copy of WordMat.dotm. Any changes in the code between Windows and Mac are handles within the VBA code using compiler constants.
However there are a couple of caveats:
- Keyboard shortcuts will not work from Windows to Mac
- Special characters will mess up in the code, menus etc
  This second problem was resolved at some point by Microsoft or Apple, but has been reintroduced since version v.1.19.
  
To avoid this, follow this procedure:
1. On Windows: Copy WordMat.dotm and rename to WordMatMac.dotm
2. open WordMatMac on the mac and run vba-function ‘ReplacetoExtendedASCII’ & ‘GenerateKeyboardShortcuts’ (Press Alt+F8)
   (The macro replaces special charcters by other text sequences like *ae* for æ)
3. Open WordMatMac.dotm on Mac by first opening Word and then choosing the file. (Do not open by clicking the file in Finder)
   Activate macros when prompted
4. In the menu click *Tools | Customize keyboard*. Then click *Reset all* at the bottom and OK.
5. Run the macro *'ReplaceToUnicode'* (Press fn+Alt+F8)
6. Run the macro *'GenerateKeyboardShortcuts'* (Press fn+Alt+F8)
7. Check if the VBA code will compile on Mac by pressing fn+Alt+F11 to open the VBE then click in the menu *Debug | Compile project*
8. Place the *WordMatMac.dotm* file in the root of the Mac folder

## Creating Excel-files
There is a 'Excel-files' folder in both Windows and Mac folders.
Excel-files must be changed/coded on Windows.
To make a 'Mac'-copy do the following:
- Copy the Excel-book
- Open it on a Windows computer and Run Macro 'ReplaceToNonUnicode' 
- On a Mac open the file and run macro 'ReplaceToUnicode'
*(Make sure to add reference 'Microsoft Visual Basic for Applications Extensibility 5.3')*
*This macro will only run if you have set Word to trust the VBA project object model*
   *Files | Settings | Trust Center | Trust Center Settings | Macro Settings | Trust access to the VBA project object model *
   (The macro replaces special charcters by other text sequences like *ae* for æ)

