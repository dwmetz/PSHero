 
# Create function to receive UNIX Time Format and return it for the local timezone
 
function Convert-UnixTime {
 
Param(
[Parameter(Mandatory=$true)][int32]$udate
)
 
$Timezone = (Get-TimeZone)
IF($Timezone.SupportsDaylightSavingTime -eq $True){
$TimeAdjust =  ($Timezone.BaseUtcOffset.TotalSeconds + 3600)
}ELSE{
$TimeAdjust = ($Timezone.BaseUtcOffset.TotalSeconds)
}
 
 
    # Adjust time from UTC to local based on offset that was determined before.
    $udate = ($udate + $TimeAdjust)
 
    # Retrieve start of UNIX Format
    $orig = (Get-Date -Year 1970 -Month 1 -Day 1 -hour 0 -Minute 0 -Second 0 -Millisecond 0)
 
    # Return final time
return $orig.AddSeconds($udate)
}

$script:utime = Read-Host -Prompt 'Enter the Unix timestamp (ex.1539138798.368)'

Convert-UnixTime $script:utime
