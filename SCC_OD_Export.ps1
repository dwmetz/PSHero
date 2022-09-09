<# 

This script will generate a Security and Compliance Search to capture OneDrive for a list of custodians.

This PowerShell script will prompt you for the following information:
    * Your user credentials                                          
    * The pathname for the text file that contains a list of user email addresses
    * The name of the Content Search that will be created
    * The search query string (optional. mastering the search query cmd is a dark art.)
 The script will then:
    * Find the OneDrive for Business site for each user in the text file
    * Create and start a Content Search using the above information

#>

# Get user credentials
if (!$credentials)
{
    $credentials = Get-Credential
}

# Get other required information
$inputfile = read-host "Enter the file name of the text file that contains the email addresses for the users you want to search"
$searchName = Read-Host "Enter the name for the new search"
$searchQuery = Read-Host "Enter the search query you want to use"
$emailAddresses = Get-Content $inputfile | Where-Object {$_ -ne ""}  | ForEach-Object{ $_.Trim() }
# Connect to Office 365
if (!$s -or !$a)
{
    $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid" -Credential $credentials -Authentication Basic -AllowRedirection -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
    $a = Import-PSSession $s -AllowClobber
    if (!$s)
    {
        Write-Error "Could not create PowerShell session."
        return;
    }
}
Write-Host "Getting each user's OneDrive for Business URL"

$urls = @()

foreach($emailAddress in $emailAddresses)
{
    $OneDriveURL = Get-OneDriveUrl -userPrincipalName $emailAddress

    if ($null -ne $OneDriveURL){
        $urls += $OneDriveURL
        Write-Host "-$emailAddress => $OneDriveURL"
    } else {
        Write-Warning "Could not locate OneDrive for $emailAddress"
    }

}
Write-Host "Creating and starting the search"

$search = New-ComplianceSearch -Name $searchName -ExchangeLocation $emailAddresses -SharePointLocation $urls -ContentMatchQuery $searchQuery

# Finally, start the search and then display the status

if($search)
{
    Start-ComplianceSearch $search.Name
    Get-ComplianceSearch $search.Name
} 

