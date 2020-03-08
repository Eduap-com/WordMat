#!/bin/sh

# now created using installer
# mkdir -p /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/
# mkdir -p /Library/Application\ Support/Microsoft/Office365/User\ content.localized/Startup.localized/Word/
# mv /Applications/WordMat2016/WordMatWinMac.dotm /Library/Application\ support/Microsoft/Office365/User\ content.localized/Startup.localized/Word/
# mv /Applications/WordMat2016/LibMaximaConnection.dylib /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/
# mv /Applications/WordMat2016/maxima.app /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/
# mv /Applications/WordMat2016/gnuplot /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/

# Create the temp folder
#mkdir -p ~/Library/Containers/com.microsoft.word/Data/WordMat/

# the scriptfile for applescripttask
mkdir -p ~/Library/Application\ Scripts/com.microsoft.Word/
mv /Applications/WordMat/WordMatScripts.scpt ~/Library/Application\ scripts/com.microsoft.word/

# previously named file from prev version
rm /Library/Application\ Support/Microsoft/Office365/User\ Content.localized/Startup.localized/Word/WordMatWinMac.dotm

#mkdir -p ~/Library/Containers/com.microsoft.Word/Data/WordMat/Formulae/

#cp -R /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/Formulae/ ~/Library/Containers/com.microsoft.Word/Data/WordMat/Formulae/

#cp  /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/Figurer.docx ~/Library/Containers/com.microsoft.Word/Data/WordMat/

#cp  /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/Figurer_english.docx ~/Library/Containers/com.microsoft.Word/Data/WordMat/

#cp  /Library/Application\ support/Microsoft/Office365/User\ content.localized/Add-ins.localized/WordMat/Figurer_spansk.docx ~/Library/Containers/com.microsoft.Word/Data/WordMat/

