<# MS Teams Security & Compliance Search
author: Doug Metz https://github.com/dwmetz
Note this script requires previous installation of the ExchangeOnlineManagement PowerShell module
See https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps for more information.#>
[string]$user = Read-Host -Prompt 'Exchange Credentials'
Connect-IPPSSession -UserPrincipalName $user
[string]$name = Read-Host -Prompt 'Enter a name for the search'
[string]$email = Read-Host -Prompt 'Enter the users email address'
new-compliancesearch -name $name -ExchangeLocation $email -ContentMatchQuery 'kind=microsoftteams','ItemClass=IPM.Note.Microsoft.Conversation','ItemClass=IPM.Note.Microsoft.Missed','ItemClass=IPM.Note.Microsoft.Conversation.Voice','ItemClass=IPM.Note.Microsoft.Missed.Voice','ItemClass=IPM.SkypeTeams.Message'
Start-ComplianceSearch $name
Get-ComplianceSearch $name
New-ComplianceSearchAction -SearchName $name -Export
Write-Host "Search initiated"-ForegroundColor Blue
Write-Host "Proceed to https://protection.office.com/ to download the results."-ForegroundColor Blue