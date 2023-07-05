<#
Mal-Hash.ps1 v1.4
https://github.com/dwmetz/Mal-Hash
Author: @dwmetz
Function: This script will generate hashes (MD5, SHA1, SHA256) for a specified file, 
        run strings against the file,
        submit the MD5 to Virus Total, 
        produce a report with the results.
        * Now works on Windows, Mac & Linux!
Prerequisites:
        Internet access is required for VT lookup.
        Virus Total API key saved in vt-api.txt
23-January-2023 ascii art update
6-December-2022 simplified hash output; 
                strings (+8); 
                UTC timestamp in report
                report name change
5-July-2023 Changed default hash value submission to SHA256
#>
Clear-Host
Write-Host ""
Write-Host ""
Write-Host ""
Write-host -Fore Cyan "
    .',;::cccccc:;.                         ...'''''''..'.  
   .;ccclllloooddxc.                   .';clooddoolcc::;:;. 
   .:ccclllloooddxo.               .,coxxxxxdl:,'..         
   'ccccclllooodddd'            .,,'lxkxxxo:'.              
   'ccccclllooodddd'        .,:lxOkl,;oxo,.                 
   ':cccclllooodddo.      .:dkOOOOkkd;''.                   
   .:cccclllooooddo.  ..;lxkOOOOOkkkd;                      
   .;ccccllloooodddc:coxkkkkOOOOOOx:.                       
    'cccclllooooddddxxxxkkkkOOOOx:.                         
     ,ccclllooooddddxxxxxkkkxlc,.                           
      ':llllooooddddxxxxxoc;.                               
       .';:clooddddolc:,..                                  
           ''''''''''                                                                                                                 
"                
Write-Host -Fore Cyan "                    Mal-Hash v1.4" 
Write-Host -Fore Gray "          https://github.com/dwmetz/Mal-Hash"
Write-Host -Fore Gray "          @dwmetz | bakerstreetforensics.com"
Write-Host ""
Write-Host ""
write-host " "
$tstamp = (Get-Date -Format "yyyy-MM-dd-HH-mm")
$script:file = Read-Host -Prompt 'enter path and filename'
write-host " "
$sourcefile = [system.IO.Path]::GetFileName("$script:file")
"SOURCE: $sourcefile" | Out-File -FilePath malhash.-t.txt -Append
" " | Out-File -FilePath malhash.-t.txt -Append
$datetime = Get-Date
$date = $datetime.ToUniversalTime()
"DATE/TIME UTC: $date" | Out-File -FilePath malhash.-t.txt -Append
" " | Out-File -FilePath malhash.-t.txt -Append
$apiKey = (Get-Content vt-api.txt)
$MD5hash = (Get-FileHash $file -Algorithm MD5).Hash 
$SHA1hash = (Get-FileHash $file -Algorithm SHA1).Hash
$SHA256hash = (Get-FileHash $file -Algorithm SHA256).Hash
"** HASHES: **" | Out-File -FilePath malhash.-t.txt -Append
"MD5: $MD5hash" | Out-File -FilePath malhash.-t.txt -Append
"SHA1: $SHA1hash" | Out-File -FilePath malhash.-t.txt -Append
"SHA256: $SHA256hash" | Out-File -FilePath malhash.-t.txt -Append
" " | Out-File -FilePath malhash.-t.txt -Append
"** STRINGS: ** " | Out-File -FilePath malhash.-t.txt -Append
strings -n 8 $script:file | Out-File -FilePath malhash.-t.txt -Append
write-host "STRINGS:" -Fore Cyan
strings -n 8 $script:file
" " | Out-File -FilePath malhash.-t.txt -Append
Write-host ""
# Comment out below to skip VT query (offline analysis)
"** VIRUS TOTAL RESULTS: **" | Out-File -FilePath malhash.-t.txt -Append
$fileHash = (Get-FileHash $file -Algorithm SHA256).Hash
write-host "Submitting SHA256 hash $fileHash to Virus Total" -Fore Cyan
Write-host ""
$uri = "https://www.virustotal.com/vtapi/v2/file/report?apikey=$apiKey&resource=$fileHash"
write-host "VIRUS TOTAL RESULTS:" -Fore Cyan
Invoke-RestMethod -Uri $uri
$vtResults = Invoke-RestMethod -Uri $uri
Invoke-RestMethod -Uri $uri | Out-File -FilePath malhash.-t.txt -Append
$vtresults
$vtResults.scans 
$vtResults.scans | Out-File -FilePath malhash.-t.txt -Append
#>
" " | Out-File -FilePath malhash.-t.txt -Append
"** END REPORT **" | Out-File -FilePath malhash.-t.txt -Append
$report = $sourcefile + "." + $tstamp
Get-ChildItem -Filter 'malhash*' -Recurse | Rename-Item -NewName {$_.name -replace '-t', $report }
Write-host " "
Write-host "Mal-Hash complete. Report saved as malhash.$report.txt" -Fore Cyan