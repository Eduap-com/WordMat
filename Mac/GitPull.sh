#!/usr/bin/env zsh
# Copies WordMat.dotm from Windows to Mac folder
# Also Copies WordMatP.dotm and WordMatP2.dotm
# Also copies P-files to the Word startup folder
# Example: ./WordMatCopyFrWin.sh
# Or set to run from Terminal as default

Echo "************************************************************************"
Echo "*This script Pulls WordMat from Git repository and copies to Mac folder*"
Echo "************************************************************************"

# Scriptfolder. needed if script is run via click. It will only hold the path if run from terminal
# move up one folder level
#MacFolder=${PWD%/*}
#Macfolder=${0:a:h}
MacFolder=$(dirname "$0")
#WMFolder=${MacFolder%/*}
cd $MacFolder

git pull

# It must be ./ and not sh otherwise the folders are not resolved correctly in the script
./CopyWordMatFrWin.sh

#sh DeleteWMFromStartup.sh
#sh CopyWordMatPToStartup.sh

Echo
Echo "Done"
# I shell indstillinger kan man sætte shell til at lukke automatisk, så er denne rar
# Åben terminsla. command + ,   åbner indstillinger vælg profil, så shell, så vælg 'Luk hvis shell afsluttede korrekt'
#read -s -k '?Press any key to close.'
