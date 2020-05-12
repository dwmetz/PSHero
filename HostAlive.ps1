<#
.SYNOPSIS
   Test a network connection to a computer, and notify with an alert when the computer is contacted.
.DESCRIPTION
   This function uses the Test-Connection cmdlette, but adds audio notification when a ping is successful. This allows the 
   function user to walk away, but be notified with an audio alert when a server is available.
.PARAMETER ComputerName
   The name of the server, or computer that is being pinged.
.PARAMETER Voice
   If used, provides a text to voice converted audio alert when a server is available. Default is audio tone.  
.EXAMPLE
   Test-ConnectWithAlert -ComputerName ServerName
.EXAMPLE
   Test-ConnectWithAlert -ComputerName ServerName -Voice   

   original source: https://thescriptlad.com/2013/08/24/continual-ping-with-success-alert/
   updated with input from: https://www.itprotoday.com/powershell/getting-input-and-inputboxes-powershell
#>

Param([string]$ComputerName = $target,[switch]$voice = $true)
[System.Reflection.Assembly]::LoadWithPartialName(‘Microsoft.VisualBasic’) | Out-Null
$ComputerName = [Microsoft.VisualBasic.Interaction]::InputBox(“Enter asset to monitor:”, “Computer or Device”, “ ”).toupper()
function Test-ConnectWithAlert
{

	do {
		(write-host "Testing connection to $ComputerName.")
			}
	until (Test-Connection $ComputerName -quiet)#This is where the ping occurs.

	if(-not $voice)#DEFAULT. This option is used when no voice output is selected.
	{
		#Notify that $computername is pingable.
		Write-Host "$ComputerName is now online." -ForegroundColor Green
		
		#http://scriptolog.blogspot.com/2007/09/playing-sounds-in-powershell.html
		#Play a repeating sound to alert of the computer being pingable.
		$sound = new-Object System.Media.SoundPlayer;
		$sound.SoundLocation="c:\WINDOWS\Media\chimes.wav"; #Many sounds options are available in c:\WINDOWS\Media
		$sound.PlayLooping();#Repeat continually.
		
		#Here is a one-liner version of the above code:
		#$sound = new-object Media.SoundPlayer "c:\WINDOWS\Media\chimes.wav").PlayLooping();
		#I haven't figured out how to stop it once it starts though.

		Write-Host "Press any key to continue ..."
		$dummy = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		$sound.Stop()
	
	}
	else #This option is used if -voice is included in the command line.
	{
		Write-Host "$ComputerName is now online. (CTRL-C to Quit)" -ForegroundColor Green
		#Create the speech object.
		#Tips on this can be found here: http://thescriptlad.com/tag/voice/
		$SpokenAlert = new-object -com SAPI.SpVoice #Make a voice object using the com object.
		#Repeat until acknowledge with CTRL-C.
		do
		{
			($SpokenAlert.Speak("$ComputerName is now on line.", 0 )|out-null),
			(start-sleep -seconds 5),
			($dummy = 0) | Out-Null
		}
		until ($dummy -eq 1)
	}
	
	Write-Host "Script has completed." -ForegroundColor GREEN
}
Test-ConnectWithAlert -ComputerName $target -voice