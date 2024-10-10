# Shows Hash codes for files

# Define variables
$filNavn = "WordMat.dotm,WordMatP.dotm,WordMatP2.dotm" # OBS: Den fil der cleanes står i vbacleaner.docm filen som CleanVBA2
$filArr = $filNavn -split ","

#Get-Filehash -path $filNavn -algorithm SHA512 | fl
#Get-Filehash -path $filNavn -algorithm MD5 | fl



# Loop over files
foreach ($fil in $filArr) {
    # Get file hash
    $hash = Get-Filehash -path $fil -algorithm SHA256
    # Display hash
    Write-Host "File: $fil           Algorithm: $($hash.Algorithm)"
    Write-Host "Hash: $($hash.Hash)"
#    Write-Host "Path: $($hash.Path)"
    Write-Host " "
}

# Display cleaned files
Write-Host "Done"
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")