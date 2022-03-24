# VBA-modules

These files are forms, modules and class modules exported from WordMat.dotm

They are here as a means of backup, but also because this is the way to go for multiple people to work on the VBA-code in WordMat.dotm, as these files are git supported where as WordMat.dotm is not.

The work process could be as follows:

- Pull git
- Open an empty WordMat.dotm file (cleaned of VBA code)
- Import all the modules (using MZ-tools, or [VBA developer tools](http://vbatools.sourceforge.net))
- code, test, ....
- Export all the modules
- Push git
