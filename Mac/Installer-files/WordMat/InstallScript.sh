#!/bin/sh

# Create the temp folder
#mkdir -p ~/Library/Containers/com.microsoft.word/Data/WordMat/

# the scriptfile for applescripttask
mkdir -p ~/Library/Application\ Scripts/com.microsoft.Word/
mv /Applications/WordMat/WordMatScripts.scpt ~/Library/Application\ scripts/com.microsoft.word/

# previously named file from prev version
#rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMatWinMac.dotm
