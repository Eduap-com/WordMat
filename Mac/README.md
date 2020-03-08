# WordMat Mac

This folder contains all files needed to build WordMat for Mac, except the Maxima.app which is too large to be committed to GitHub (115 MB). 
You can get this file from the most recent installer. It will be placed in the folder: 
*/Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/*
Then place it in the root of the Mac folder.

To create the installer pkg-file you need to download and install [Packages](http://s.sudre.free.fr/Software/Packages/about.html)
Then open WordMat.pkgproj

Unfortunately you cannot just compile as paths are not correct on a new computer. It is still a work in progress.


Normally all editing of code should be done on windows and then copied to this folder. When the files are opened on mac some characters can mess up like æøå.
Depending on the file it all requires special handling.

See more details in [How to build WordMat](https://github.com/Eduap-com/WordMat/blob/master/How%20to%20build%20WordMat.docx)
