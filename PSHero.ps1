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
    Write-Host "=== Hosts ============================="-ForegroundColor darkgreen
    Write-Host "HB: Bitlocker Lookup"
    Write-Host "HG: Get Computer Info"
	Write-Host "HA: Host Alive"
    Write-Host "=== Aquisition ======================="-ForegroundColor darkgreen
    Write-Host "AI: IRMemPull Memory Acquistion"
	Write-Host "AA: Axiom Cloud - O365 Connect to Collect"
    Write-Host "=== Email ============================="-ForegroundColor darkgreen
    Write-Host "EX: MX Header Analysis"
	Write-Host "ES: SadPhishes - email search (E)"
	Write-Host "=== Conversion ========================"-ForegroundColor darkgreen
	Write-Host "CT: Unix time to Human Readable"
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
            $script:userID = Read-Host -Prompt 'Enter the ID'
    C:\Windows\System32\runas.exe /profile /user:$script:userID "powershell"
        }
        'LO' {
                D:\PowerShell\PSHero\ExchangeOnline.ps1
            }
        'LE' {
                D:\PowerShell\PSHero\Connect-ExchangeOnline.ps1
            }
        'HB' {
                D:\PowerShell\PSHero\Bitlocker.ps1
            }
        'HG' {
                D:\PowerShell\PSHero\GetComputer.ps1
            }
        'HA' {
                D:\PowerShell\PSHero\HostAlive.ps1
            }
        'AI' {
                Set-Location D:\Temp\IR
                .\Irmempull.ps1
                Set-Location D:\PowerShell\Scripts
            }
        'AA' {
                D:\PowerShell\PSHero\AxCollect.ps1
            }
        'EX' {
                D:\PowerShell\PSHero\Parse-EmailHeader.ps1
            }
        'ES' {
                D:\PowerShell\PSHero\SadPhishes.ps1
            }
        'CT' {
                D:\PowerShell\PSHero\UnixTime.ps1
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