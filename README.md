# WordMat
[![Downloads](https://img.shields.io/github/downloads/Eduap-com/WordMat/total.svg?style=popout)](https://github.com/Eduap-com/WordMat/releases)
[![Release Version](https://img.shields.io/github/release/Eduap-com/WordMat)](https://github.com/Eduap-com/WordMat/releases/latest)
[![Build Status](https://dev.azure.com/Eduap-com/WordMat/_apis/build/status/CI%20Pipeline?branchName=master&stageName=Build)]
WordMat is open source, released as GNU General public License.

WordMat is written in VBA, c#, c and Maxima script
The installer is created using Inno setup and written in pascal.

Most of the code is VBA and can be found in the Word template WordMat.dotm which is located in the install folder when the application is installed. Open file templatefile in Word and press alt+F11
Some VBA-code is also found in the Excel-files for graph-plotting, statistics etc...

The API library which handles the communication with Maxima is called MathMenu.dll and is written in c# using .Net framework 4.0
The MathMenu sourcecode is found in this folder.

For Mac the API is called LibMaximaConnection.dylib and is written in c using xcode.

Maxima comes with a Maxima-installation which is not standard. Significant changes has been made to it
- new images with a lot of specific setup changes
- Units packages has been altered
- Many new functions has been created for use with WordMat such as Solvereal and differential equation solver.

The inno setup script file is also in the folder.

Mikael Samsøe Sørensen
www.eduap.com

