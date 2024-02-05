#!/bin/sh

Echo
Echo Uninstalling WordMat
Echo
Echo You will have to enter your password to complete this uninstall
Echo '(no text or * will be shown when typing the password but it is registered)'
Echo


#Main Add-in file. Removes WordMat from Word menu
sudo rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMat.dotm
sudo rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMatP.dotm
sudo rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMatP2.dotm


#Main installation
sudo rm -rf /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Add-ins.localized/WordMat/

#Script files for applescripttask
sudo rm ~/Library/Application\ Scripts/com.microsoft.Word/WordMatScripts.scpt

#Temp files and backups files
sudo rm -rf ~/Library/Containers/com.microsoft.Word/Data/WordMat/

#Fonts
sudo rm /Library/fonts/Latinmodern-math.otf
sudo rm /Library/fonts/Lmroman12-regular.otf

#libs
sudo rm /opt/local/lib/Libzstd.1.5.1.dylib
sudo rm /opt/local/lib/Libzstd.1.dylib

#Formuale and Temp files and backups files
#sudo rm -rf ~/Library/Containers/com.microsoft.Word/Data/WordMat/

#rmdir /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Add-ins.localized/WordMat/
#rmdir /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/

sudo rm /Applications/WordMat/InstallScript.sh
sudo rm /Applications/WordMat/WordMat\ installation\ files.docx
sudo rm /Applications/WordMat/UnInstallScript
sudo rmdir  /Applications/WordMat/

Echo
Echo
Echo WordMat is now removed from Word.
Echo You can now close this windows
Echo
