# Requires -Version 5.1
# This script cleans WordMat.dotm for compiled code and more.
# It requires Ribbon Commander to be installed with a license.
# This is a help if WordMat.dotm suddenly crashes upon opening.

# Jeg har tilføjet en VBAclean2 sub til VBAcleaner.docm filen, fordi jeg ikke kunne få det til at virke med at powershell skulle sende en parameter med til VBAcleaner.docm filen.
# Define variables
$filNavn = "WordMat.dotm"
$filArr = $filNavn -split ","
$strPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$vbaCleanerDocm = Join-Path -Path $strPath -ChildPath "VBACleaner.docm"
$wordApp = New-Object -ComObject Word.Application
#$wordApp.Visible = $false

# Open the VBA cleaner document
$document = $wordApp.Documents.Open($vbaCleanerDocm, $false, $true)

# Initialize log string
$fl = "Filename`t`tOldsize`tNewsize`r`n"


foreach ($fil in $filArr) {
    $strFile = Join-Path -Path $strPath -ChildPath $fil
    if (Test-Path $strFile) {
        $oldSize = (Get-Item $strFile).length / 1000
        # Run the CleanVBA macro 
#        Write-Host "Cleaning: " $strFile
#        $wordApp.Run("CleanVBA", $strFile , $false, $true) # virker ikke med de tre parametre, forsøgt [ref]$strFile.value og [ref]$strFile
        $newSize = (Get-Item $strFile).length / 1000
        $fl += "$fil`t$oldSize`t$newSize`r`n"
    }
}

# Denne kører VBAclean på alle 3 filer via speciel CleanVBA2 sub
Write-Host "Cleaning files"
$wordApp.Run("CleanVBA2") 

# Close the document and quit Word
Write-Host "Closing Word"
$document.Close()
$wordApp.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Copy the cleaned files to Mac folder one level up
Copy-Item -Path "$strPath\WordMat.dotm" -Destination "$strPath\..\Mac\WordMat.dotm" -Force
Copy-Item -Path "$strPath\WordMatP.dotm" -Destination "$strPath\..\Mac\WordMatP.dotm" -Force
Copy-Item -Path "$strPath\WordMatP2.dotm" -Destination "$strPath\..\Mac\WordMatP2.dotm" -Force

# Display cleaned files
Write-Host "Following files have been cleaned:"
Write-Host $fl