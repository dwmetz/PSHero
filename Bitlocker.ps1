#Param(
#  [Parameter(Mandatory=$True,Position=0)]
#   [string]$target
#     
#    )
# This script retrieves the BitLocker recovery password for a computer

[string]$target = Read-Host -Prompt 'Enter the hostname'
# Get Computer Object
$computer = Get-ADComputer -Filter {Name -eq $target}
       
# Get all BitLocker Recovery Keys for that Computer. Note the 'SearchBase' parameter
$BitLockerObjects = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $computer.DistinguishedName -Properties  'msFVE-RecoveryPassword','whenChanged'
$BitLockerObjects | format-list  -Property 'DistinguishedName','msFVE-RecoveryPassword','whenChanged' 
#| select msFVE-RecoveryPassword