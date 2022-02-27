# WordMat Mac

This folder contains all files needed to build WordMat for Mac.

To create the installer pkg-file you need to download and install [Packages](http://s.sudre.free.fr/Software/Packages/about.html)
Then open WordMat.pkgproj

Unfortunately you cannot just compile as paths are not correct on a new computer. It is still a work in progress.

Normally all editing of code should be done on windows and then copied to this folder. When the files are opened on mac some characters can mess up like æøå.
Depending on the file it all requires special handling.
**Excel-files**
Excel-files must be changed/coded on Windows.
To make a 'Mac'-copy do the following:
- Copy the Excel-book
- Open it on a Windows computer and Run Macro 'ReplaceToNonUnicode 
- On a Mac open the file and run macro 'ReplaceToUnicode'
*(Make sure to add reference 'Microsoft Visual Basic for Applications Extensibility 5.3')*
*This macro will only run if you have set Word to trust the VBA project object model*
   *Files | Settings | Trust Center | Trust Center Settings | Macro Settings | Trust access to the VBA project object model *
   (The macro replaces special charcters by other text sequences like *ae* for æ)


See more details in [How to build WordMat](https://github.com/Eduap-com/WordMat/blob/master/How%20to%20build%20WordMat.docx)
