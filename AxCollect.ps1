<# AxCollect.ps1 
version 2.0
https://github.com/dwmetz/Axiom-PowerShell
Author: @dwmetz
Function:
    This script enables investigator access to a users O365 mailbox.
    Script shoud be run by Exchange Administrator to grant mailbox delegation permissions.
    Once enabled, sign in using investigator credentials to perform the cloud collection in Axiom. 

25.October.2022 - update ExchangeOnlineManagement connection

#>#>
Import-module ExchangeOnlineManagement
Connect-ExchangeOnline
$script:axiom_o365 = Read-Host -Prompt 'Enter the used ID of the target (ex.Alice)'
$script:axiom_examiner =Read-Host -Prompt 'Enter the used ID of the examiner (ex.Bob)'
Add-MailboxPermission -Identity "$script:axiom_o365" -User "$script:axiom_examiner" -AccessRights FullAccess -ErrorAction stop
Write-Host -ForegroundColor Green $script:axiom_examiner 'permissions added to' $script:axiom_o365 'mailbox.'