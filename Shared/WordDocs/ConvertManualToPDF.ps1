$strPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$vbaDocm = Join-Path -Path $strPath -ChildPath "CreatePDF.docm"
Write-Host "Opening Word application"
$wordApp = New-Object -ComObject Word.Application
#$wordApp.Visible = $false

# Open the VBA cleaner document
Write-Host "Opening Word document $vbaDocm"
$document = $wordApp.Documents.Open($vbaDocm, $false, $true)

# Denne kører VBAclean på alle 3 filer via speciel CleanVBA2 sub
Write-Host "Converting document to PDF"
$wordApp.Run("ConvertManualToPDF") 

# Close the document and quit Word
Write-Host "Closing Word"
$document.Close()
$wordApp.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Display cleaned files
Write-Host "Done."
