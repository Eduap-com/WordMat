#!/usr/bin/env zsh
# Copies WordMat.dotm from Windows to Mac folder
# Also Copies WordMatP.dotm and WordMatP2.dotm
# Also copies P-files to the Word startup folder
# Example: ./WordMatCopyFrWin.sh
# Or set to run from Terminal as default

Echo "******************************************************"
Echo "*This script copies WordMat.dotm and P To Mac folder *"
Echo "******************************************************"

# Scriptfolder. needed if script is run via click. It will only hold the path if run from terminal
# move up one folder level
#MacFolder=${PWD%/*}
#echo Macfolder: $MacFolder
MacFolder=${0:a:h}
MacFolder2=$(dirname "$0")
if [[ ${#MacFolder2} -gt ${#MacFolder} ]]
then
MacFolder=$MacFolder2
fi
echo $MacFolder
WMFolder=${MacFolder%/*}
#read REPLY\?"Achitecture: "

# Copy WordMat.dotm from Windows to Mac
cp $WMFolder/Windows/WordMat.dotm $WMFolder/Mac
cp $WMFolder/Windows/WordMatP.dotm $WMFolder/Mac
cp $WMFolder/Windows/WordMatP2.dotm $WMFolder/Mac

#cp Windows/WordMat.dotm /Library/Application Support/Microsoft/Office365/User Content.localized/Startup.localized/Word/WordMat.dotm
#sudo cp $MacFolder/Windows/WordMatP.dotm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word
#sudo cp $MacFolder/Windows/WordMatP2.dotm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word

Echo
Echo "Done"
# I shell indstillinger kan man sætte shell til at lukke automatisk, så er denne rar
# Åben terminsla. command + ,   åbner indstillinger vælg profil, så shell, så vælg 'Luk hvis shell afsluttede korrekt'
read -s -k '?Press any key to close.'
