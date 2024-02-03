# How to install WordMath

This file contains information on how to install WordMat on both Windows and Mac. Please see the apropriate section
These instructions does not contain information on how to modify, compile and build all parts of WordMat. Please see 'How to build WordMat' If you are looking for that.
Also please observe that you are in accordance with the License for the different parts. WordMath is in many regards an aggregate of different components, and some parts must be installed separately.

## Windows

Prerequisites
- Microsoft Windows
- Microsoft .Net Framework
- Microsoft Word
- Microsoft Excel
- WordMat Source code folder

Please observe that if any changes has been made to the Source code, you may need to recompile those parts. (See how to build)

Copy these entire folders: 
*Source code folder           ->  Local system folder*
Windows\ExcelFiles           ->  C:\Program Files (x86)\WordMat\ExcelFiles  
Shared\WordDocs              ->  C:\Program Files (x86)\WordMat\WordDocs
Shared\geogebra-math-apps    ->  C:\Program Files (x86)\WordMat\geogebra-math-apps 
Shared\Fonts                 ->  C:\Program Files (x86)\WordMat\Fonts
Windows\Images               ->  C:\Program Files (x86)\WordMat\Images
Windows\ExternalPrograms\maxima-5.47.0  ->  C:\Program Files (x86)\WordMat\maxima-5.47.0
Windows\WebViewWrap          ->  C:\Program Files (x86)\WordMat\WebViewWrap

Copy these individual files:
Windows\MathMenu.dll\MathMenu.dll  ->    C:\Program Files (x86)\WordMat\
Windows\MathMenu.dll\MathMenu.tlb  ->    C:\Program Files (x86)\WordMat\
Windows\WordMat.dotm               ->    C:\Program Files\Microsoft Office\root\Office16\STARTUP\
Shared\WordDocs\BibWord_Stylesheet.xsl -> C:\Program Files\Microsoft Office\root\Office16\Bibliography\Style

In terminal with administrator priviligeges run 
regasm MathMenu.dll /codebase              (adjust path appropriately)
regasm WebViewWrap.dll /codebase           (adjust path appropriately)

Install the fonts in 'C:\Program Files (x86)\WordMat\Fonts' by right clicking each font

Consider installing GeoGebra 5, from GeoGebra.org website. Please observe License.
Consider installing Graph from padowan.dk.

## Mac

Prerequisites
- MacOS
- Microsoft Word
- Microsoft Excel
- WordMat Source code folder

Change M1 to Intel depending on your system

Copy these entire folders
Mac/External-files/WordMat/M1/MaximaWM    ->    /Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/       
Shared/geogebra-math-apps    ->     /Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/ 
Mac/ExcelFiles               ->     /Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/ 
Shared\WordDocs              ->     /Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/

Copy these files 
Mac/WordMatMac.dotm  ->  /Library/Application support/Microsoft/Office365/User Content/Add-ins/Startup/Word/
Mac/Installer-files/WordMat/WordMatScripts.scpt    ->   ~/Library/Application Scripts/com.microsoft.Word/
Shared/Fonts/latin-modern-math.otf                 ->    /Library/Fonts/
Shared/Fonts/lmroman12-regular.otf                 ->    /Library/Fonts/
Mac/External-files/WordMat/M1/libzstd.1.5.5.dylib  ->   /opt/local/lib/
Mac/External-files/WordMat/M1/libzstd.1.dylib      ->   /opt/local/lib/

Consider Installing Geogebra 5 from geogebra.org
