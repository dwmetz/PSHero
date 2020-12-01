# PSHero Powershell Menu
# (c)2020 github.com/dwmetz
# All rights to 3rd party scripts remain with the original owners
#
# Note: do a find/replace for D:\PowerShell\PSHero\ and subsititute the path where PSHero scripts are locally stored


function Show-Menu
{
    param (
        [string]$Title = 'My Menu'
    )
    Clear-Host
    Write-Host "=== $Title ==========="-ForegroundColor darkcyan
    Write-Host "=== Logins ============================="-ForegroundColor darkgreen
    Write-Host "LA: Alternate PS Login"
    Write-Host "LO: O365 Admin Login"
    Write-Host "LE: O365 Admin Login [Modern Auth]"
    #Write-Host "LR: Remote PowerShell (alt user)"
	Write-Host "LN: Non-Attrib SSH"
    Write-Host "=== Hosts ============================="-ForegroundColor darkgreen
    Write-Host "HB: Bitlocker Lookup"
    Write-Host "HG: Get Computer Info"
	Write-Host "HA: Host Alive"
    Write-Host "=== Aquisition ======================="-ForegroundColor darkgreen
    Write-Host "AI: IRMemPull Memory Acquistion"
	Write-Host "AA: Axiom Cloud - O365 Connect to Collect"
    Write-Host "=== Email ============================="-ForegroundColor darkgreen
    Write-Host "EX: MX Header Analysis"
	#Write-Host "EL: O365 Litigation Hold (E)"
	Write-Host "ES: SadPhishes - email search (E)"
	#Write-Host "EP: Purge Email Search (E)"
    #Write-Host "=== VirusTotal ========================"-ForegroundColor darkgreen
	#Write-Host "VH: VirusTotal hash lookup"
	#Write-Host "VF: VirusTotal file lookup"
    #Write-Host "=== User =============================="-ForegroundColor darkgreen
	#Write-Host "UL: User Lookup"
    Write-Host "=== Conversion ========================"-ForegroundColor darkgreen
    #Write-Host "CU: Resolve URI"
	Write-Host "CT: Unix time to Human Readable"
    #Write-Host "=== Data Management ==================="-ForegroundColor darkgreen
    #Write-Host "DC: Sync Cases"
    #Write-Host "DK: KAPE Update"
    Write-Host "=== Exit =============================="-ForegroundColor darkgreen
    Write-Host "Q: Press 'Q' to quit."
    Write-Host "=== (E) = Requires Exchange Login ====="	-ForegroundColor darkcyan
}
 do
{
    Show-Menu –Title 'PSHero - PowerShell Menu'
    $input = Read-Host "Please make a selection"
    switch ($input)
    {
        'LA' {               
                C:\Windows\System32\runas.exe /profile /user:USER@DOMAIN.COM "powershell"
            }
            #update /user:USER@DOMAIN.COM to requested credential
        'LO' {
                D:\PowerShell\PSHero\ExchangeOnline.ps1
            }
        'LE' {
                D:\PowerShell\PSHero\ExchangeOnline_new.ps1
            }
        'HB' {
                D:\PowerShell\PSHero\Bitlocker.ps1
            }
        'EX' {
                D:\PowerShell\PSHero\Parse-EmailHeader.ps1
            }
        'AC' {
                D:\PowerShell\PSHero\CloudConnect.ps1
            }
		'AA' {
                D:\PowerShell\PSHero\AxCollect.ps1
            }
        'AD' {
                D:\PowerShell\PSHero\CloudRemove.ps1
            }
        'HA' {
                D:\PowerShell\PSHero\HostAlive.ps1
            }
        'CU' {
                D:\PowerShell\PSHero\resolve-uri.ps1
            }
        'DC' {
                D:\PowerShell\PSHero\SyncCases.ps1
            }
        'AI' {
                cd D:\Temp\IR
                .\Irmempull.ps1
                cd D:\PowerShell\Scripts
            }
        'DK' {
                cd D:\Tools\KAPE\
                .\Get-KAPEUpdate.ps1
                cd D:\PowerShell\Scripts
            }
		'LN' {
                D:\PowerShell\PSHero\NonAttribSSH.ps1
            }
		'LR' {
                D:\PowerShell\PSHero\PSession.ps1
            }
	    'ES' {
                D:\PowerShell\PSHero\SadPhishes.ps1
            }
		'UL' {
                D:\PowerShell\PSHero\GetUser.ps1
            }
        'EB' {
                D:\PowerShell\PSHero\BlockSender.ps1
            }
        'HG' {
                D:\PowerShell\PSHero\GetComputer.ps1
            }
        'ED' {
                D:\PowerShell\PSHero\BlockDomain.ps1
            }
        'CT' {
                D:\PowerShell\PSHero\UnixTime.ps1
            }
        'EI' {
                D:\PowerShell\PSHero\BlockSenderIP.ps1
            }
        'EL' {
                D:\PowerShell\PSHero\LitHold.ps1
            }
        'EP' {
                D:\PowerShell\PSHero\Purge.ps1
            }
        'AK' {
                D:\PowerShell\PSHero\KapeTriage.ps1
            }
        'VH' {
                D:\PowerShell\PSHero\GetVT-hash.ps1
            }
        'VF' {
                D:\PowerShell\PSHero\GetVT-file.ps1
            }

		'x' {
                D:\PowerShell\PSHero\PSHero.ps1
            }
			
        'q' {
                 return
            }
    }
    pause
}
until ($input -eq 'q')