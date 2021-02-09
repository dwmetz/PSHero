# This script retrieves the user informationfor a specific AD computer.

$script:computername = Read-Host -Prompt 'Enter the hostname'

Get-ADComputer $script:computername -Properties * | Select-Object Name, Description, LastLogonDate, OperatingSystem, CanonicalName, ManagedBy 

