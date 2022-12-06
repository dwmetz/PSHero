<#
Mal-Hash.ps1 v1.0
https://github.com/dwmetz/PSHero
Author: @dwmetz

Function: This script will generate hashes (MD5, SHA1, SHA256), submit the MD5 to Virus Total, and produce a text file with the results.

Prerequisites:
Internet access is required.
Virus Total API key saved in vt-api.txt

#>
Write-Host -Fore Gray "------------------------------------------------------"
Write-Host -Fore Cyan "       Mal-Hash v1.0" 
Write-Host -Fore Gray "       [ quickly check a file hash against Virus Total ]"
Write-Host -Fore Cyan "       @dwmetz | bakerstreetforensics.com"
Write-Host -Fore Gray "------------------------------------------------------"
Start-Sleep -Seconds 3
write-host " "
$tstamp = (Get-Date -Format "yyyy-MM-dd-HH-mm")
$script:file = Read-Host -Prompt 'enter path and filename'
write-host " "
$script:file > malhash.-t.txt
$apiKey = (Get-Content vt-api.txt)
Get-FileHash -Algorithm MD5 $script:file >> malhash.-t.txt
Get-FileHash -Algorithm SHA1 $script:file >> malhash.-t.txt
Get-FileHash -Algorithm SHA256 $script:file >> malhash.-t.txt
$fileHash = (Get-FileHash $file -Algorithm MD5).Hash
write-host "Submitting MD5 hash $fileHash to Virus Total" -Fore Cyan
Write-host ""
$uri = "https://www.virustotal.com/vtapi/v2/file/report?apikey=$apiKey&resource=$fileHash"
write-host "VIRUS TOTAL RESULTS:" -Fore Cyan
Invoke-RestMethod -Uri $uri
$vtResults = Invoke-RestMethod -Uri $uri
Invoke-RestMethod -Uri $uri | Out-File -FilePath malhash.-t.txt -Append
$vtresults
$vtResults.scans 
$vtResults.scans >> malhash.-t.txt
Get-ChildItem -Filter 'malhash*' -Recurse | Rename-Item -NewName {$_.name -replace '-t', $tstamp }
Write-host ""
Write-host "Displaying results. Details are saved in malhash.$tstamp.txt"
Write-host ""
Start-Sleep -Seconds 10
