# ADUserCredentialsCheck.ps1 
# Script by Tim Buntrock 
# Run this script like ####PS C:\admin\Scripts\UserAccountAuthenticationCheck> .\ADUserCredentialsCheck.ps1 .\users.csv#### 
 
# users.csv look like -> 
# samaccountname,password 
# User1,PASSWORD1 
# User2,PASSWORD2 
 
param($UsersCsv) 
 
# specify function to test credentials 
Function Test-ADAuthentication { 
    param($samaccountname,$password) 
    (new-object directoryservices.directoryentry "",$samaccountname,$password).psbase.name -ne $null 
} 
 
# get domain infos 
$section = "search" 
import-module activedirectory 
$domobj = get-addomain 
$domain = $domobj.dnsroot # you can also specify another domain using $domain 
 
# import user data 
$data = import-csv $UsersCsv 
 
# verify all specified credentials, and output valid or invalid 
foreach($rank in $data) { 
$samaccountname = $rank.samaccountname 
$password = $rank.Password 
if (Test-ADAuthentication "$domain\$samaccountname" "$password") { 
    write-host "$samaccountname >> CREDENTIALS VALID" -foregroundcolor "green" 
} else { 
    write-host "$samaccountname >> CREDENTIALS INVALID" -foregroundcolor "red" 
} 
}