# How to contribute

Bugs
-----------------
If you find a bug create an [issue](https://github.com/Eduap-com/WordMat/issues) and I will have a look at it.
I cannot solve the issue if I cannot recreate the problem. Hence it is important to write the version of WordMat you are using and if it is Windows or Mac.
If the problem is related to a math expression in Word you must attach a Word-document with the math-expression causing the problem, not just a screenshot.
Some problems are caused by incorrect entering of the math expression which looks correct. I can only identify this from a worddocument by inspecting the linear code.
Be adviced that I work on the project depending on time available and motivation. You cannot always expect swift response, and some minor issues or issues that require a lot of work might not be prioritized. Any help is appreciated :-)

Translations
-----------------
You can contribute by translating WordMat to your language. Create an issue and leave a comment about the tranlastion.
Then open this excelfile: Shared/translations/translations.xlsm
Here you can see the translations of existing languages. Fill in your language in the top row and start translating. Do so for each sheet.
It is ok to do a part translation as any missing tranlations just go to english.
When you are done please commit the file.
The excel-file has builtin VBA code to generate language code for the WordMat.dotm file.
The Excel file does not hold all translations. If you succedd in translating all I will help with the rest.

Coding
-------------
WordMat consist of many different parts and is written in 4 different languages (VBA, c#, c and Lisp)
You dont need to know everything about the structure to contribute.
Warning: Do not change anything in the Windows folder from a Mac and vice versa. Special characters will mess up in text-files. 
Read the 'How to build WordMat' document to get started. But here is a sum up:
The installer is created using Inno setup and written in pascal.

Most of the code is VBA and can be found in the Word template WordMat.dotm which is located in the install folder when the application is installed. Open file templatefile in Word and press alt+F11
Some VBA-code is also found in the Excel-files for graph-plotting, statistics etc...

The API library which handles the communication with Maxima on Windows is called MathMenu.dll and is written in c# using .Net framework 4.0
The MathMenu sourcecode is found in this folder.

Maxima comes with a Maxima-installation which is not standard. Significant changes has been made to it
- new images with a lot of specific setup changes
- Units packages has been altered
- Many new functions has been created for use with WordMat such as Solvereal and differential equation solver.

Pull requests
-------------

Pull requests are welcome, but it is always better if there is no duplication of work:

- If you are working on a bug / enhancement that is already listed as an issue, please
  leave a comment saying that you intent to do so. I can then share my thoughts about
  how to address that issue, assign you to it, etc.
- If there is no issue for it, it is preferable if an issue is created
  beforehand, in case I have some reservations about it

Changes to WordMat.dotm can be a problem as it holds the main code and if several people are working on different versions of the file they cannot be merged.
To circumvent this modules must be imported/exported with each edit. More info in the WordMat.dotm file
Creating a new Excel-template is no problem.


Coding guidelines
-----------------
Try to comment the code:-)
