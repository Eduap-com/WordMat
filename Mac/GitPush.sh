#!/usr/bin/env zsh
# Copies WordMat.dotm from Windows to Mac folder
# Also Copies WordMatP.dotm and WordMatP2.dotm
# Also copies P-files to the Word startup folder
# Example: ./WordMatCopyFrWin.sh
# Or set to run from Terminal as default

Echo "*************************************************************"
Echo "*This script Pushes WordMat to Git repository *"
Echo "*************************************************************"

# Scriptfolder. needed if script is run via click. It will only hold the path if run from terminal
# move up one folder level
#MacFolder=${PWD%/*}
#Macfolder=${0:a:h}
MacFolder=$(dirname "$0")
#WMFolder=${MacFolder%/*}
cd $MacFolder
git push origin master

Echo
Echo "Done"
# I shell indstillinger kan man sætte shell til at lukke automatisk, så er denne rar
# Åben terminsla. command + ,   åbner indstillinger vælg profil, så shell, så vælg 'Luk hvis shell afsluttede korrekt'
read -s -k '?Press any key to close.'
