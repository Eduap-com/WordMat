#!/usr/bin/env zsh
# Virker ikke hvis startet med "sh CreateInstaller.sh"
# If parameters are not given you will be prompted

# For at kunne køre script ved dobbeltklik:
#   Omdøb fil fra .sh til .command så virker det
#   eller:
#   chmod a+x CreateInstaller.sh
#   Højreklik og vælg 'Åbn i' og andet
#   Slå til: Alle apps
#   Flueben ved 'Åbn altid i'
#   Hjælpeprogrammer
#   Terminal
#   Åbn

Echo "**************************************************************"
Echo "*This script converts textfiles from utf-8 to MACROMAN       *"
Echo "**************************************************************"

# Path to WordMatP folder
#WMfolder='/users/test/documents/GitHub/WordMatP'
#WMfolder=~/Documents/GitHub/WordMatP
# Scriptfolder. needed if script is run via click. It will only hold the path if run from terminal
Langfolder=${0:a:h}
cd $Langfolder

#iconv -f UTF-8 -t MACROMAN da.csv > da.csv
iconv -f iso-8859-1 -t MACROMAN ../win/da.csv > da.csv
iconv -f iso-8859-1 -t MACROMAN ../win/en.csv > en.csv
iconv -f iso-8859-1 -t MACROMAN ../win/sp.csv > sp.csv

Echo
Echo "Done"
Echo

# I shell indstillinger kan man sætte shell til at lukke automatisk, så er denne rar
# Åben terminal. command + ,   åbner indstillinger vælg profil, så shell, så vælg 'Luk hvis shell afsluttede korrekt'
#read -s -k '?Press any key to close.'
