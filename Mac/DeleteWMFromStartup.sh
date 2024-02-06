#!/usr/bin/env zsh
# Copies WordMat.dotm from Windows to Mac folder
# Also Copies WordMatP.dotm and WordMatP2.dotm
# Also copies P-files to the Word startup folder
# Example: ./WordMatCopyFrWin.sh
# Or set to run from Terminal as default

Echo "*******************************************************"
Echo "*This script deletes WordMat.dotm from Startup folder *"
Echo "*******************************************************"

# Scriptfolder. needed if script is run via click. It will only hold the path if run from terminal
# move up one folder level
#MacFolder=${PWD%/*}
#Macfolder=${0:a:h}
MacFolder=$(dirname "$0")
WMFolder=${MacFolder%/*}
#read REPLY\?"Achitecture: "

# Copy WordMat.dotm from Windows to Mac
#cp $WMFolder/Windows/WordMat.dotm $WMFolder/Mac
#cp $WMFolder/Windows/WordMatP.dotm $WMFolder/Mac
#cp $WMFolder/Windows/WordMatP2.dotm $WMFolder/Mac

#cp Windows/WordMat.dotm /Library/Application Support/Microsoft/Office365/User Content.localized/Startup.localized/Word/WordMat.dotm
sudo rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMat.dotm

Echo
Echo "Done"
