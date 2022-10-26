<# MS Teams Security & Compliance Search 
version 2.0
https://github.com/dwmetz/Axiom-PowerShell
Author: @dwmetz
Function:
    Collect an O365 mailbox search for MS Teams communications.
    Note this script requires previous installation of the ExchangeOnlineManagement PowerShell module
    See https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps for more information.
    
    25.October.2022 - updated ExchangeOnlineManagement connection, Security & Compliance Center (IPPSSession)
    
    #>
    Import-module ExchangeOnlineManagement
    Connect-ExchangeOnline
    [string]$aname = Read-Host -Prompt 'Enter your account name'
    Connect-IPPSSession -UserPrincipalName $aname
    [string]$name = Read-Host -Prompt 'Enter a name for the search'
    [string]$email = Read-Host -Prompt 'Enter the users email address'
    new-compliancesearch -name $name -ExchangeLocation $email -ContentMatchQuery 'kind=microsoftteams','ItemClass=IPM.Note.Microsoft.Conversation','ItemClass=IPM.Note.Microsoft.Missed','ItemClass=IPM.Note.Microsoft.Conversation.Voice','ItemClass=IPM.Note.Microsoft.Missed.Voice','ItemClass=IPM.SkypeTeams.Message'
    Start-ComplianceSearch $name
    Get-ComplianceSearch $name
    Write-Host "Search initiated"-ForegroundColor Blue
    Write-Host "Proceed to https://protection.office.com/ to download the results."-ForegroundColor Blue