# This script enables investigator access to a users O365 mailbox.
# dwmetz@gmail.com   https://dwmetz.github.io/ @dwmetz
# Once enabled sign in using investigator credentials to Axiom

Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
$script:axiom_o365 = Read-Host -Prompt 'Enter the used ID of the target (ex.R12345)'
$script:axiom_examiner =Read-Host -Prompt 'Enter the used ID of the examiner (ex.R34567)'
Add-MailboxPermission -Identity "$script:axiom_o365" -User "$script:axiom_examiner" -AccessRights FullAccess
Write-Host -ForegroundColor Green $script:axiom_examiner 'permissions added to' $script:axiom_o365 'mailbox.'