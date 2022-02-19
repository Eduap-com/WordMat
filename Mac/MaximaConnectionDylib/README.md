# MaximaConnection dylib

**Note** This lib is depracted from version 1.23. It was not possible to get it working on M1 or Monterey. It has been replaced by a new method where a new maxima process is run every calculation.


Code for the dylib API library written in c to communicate between Word and Maxima.
Use xcode to edit.

The installer places the library in
*/Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/*
It must be this library for Word to have the correct privilige

The library MaximaConnection.dylib is referenced from the VBA code in WordMatMac.dotm

