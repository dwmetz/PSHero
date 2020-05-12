# SADPhishes.ps1
# Exchange 2016 Compliance Search & Destroy Phishing Emails
# Colin Edwards / @EdwardsCP
# Development Started - September 2018
#--------------------------------------------------
# Prerequisites: Script must be run from Exchange Management Shell by a User with the Exchange Discovery Management Role
#
# Basic Usage: 	Execute the script from within EMS
#			The user is prompted to search using various combinations of the Subject, Sender Address, Date Range, and Attachment Names.
#			The user selects to search either All Exchange Locations, or specify the Email Address associated with a specific MailBox or Group
#			The script will create and execute a Compliance Search.
#			The user then has the option to view details of the search results, delete the Items found by the Search, create an eDiscovery Search or Delete the search and return to the mail search options menu.
#
#
# Microsoft's docs say that a Compliance Search will return a max of 500 source mailboxes, and if there are more than 500 mailboxes that contain content that matches the query, the top 500 with the most search results are included in the results.  This means large environments may need to re-run searches.  Look for a future version of this script to be able to loop back through and perform another search if 500 results are returned and then deleted.
#
#=================
#Version 1.0.12
# (shout out to Doug Metz for the feedback and these suggestions!)
# Added a "Sender and Date Range" search option.
# Added option to specify your own Search Name for your Compliance Search instead of having SADPhishes automatically create one based on your Search Criteria
# Added option to set the Description for your Compliance Search
#=================
#Version 1.0.11
# Modified the search that extracts info from a Headers text file so that it would detect UTF-8 encoded subjects and convert them to plain text for executing the search.
#=================
#Version 1.0.10
# Added the option to list all existing Compliance Searches and select one to re-run (RunPreviousComplianceSearch Function)
#=================
#Version 1.0.9
# Bug Fixes - after eDiscovery search was run, needed to change the way the user is prompted about launching the results in the browser, and the workflow after that.
#=================
#Version 1.0.8
# Fixed some bugs with the header text file search options
#=================
#Version 1.0.7
# Reorganized the order of the functions in the script so it reads more easily top to bottom.
# Added option to launch eDiscovery Search results preview in default browser.
# Added search option to extract the Sender, Subject, and Date info from a text file containing Headers from a sample email
# Other minor bug fixes
#=================
#Version 1.0.6
# A ComplianceSearch Name can have a maximum of 200 Characters.  Changed the Search Name building process to use a unique integer at the end instead of a timestamp, so that it's less likely that we'll hit the 200 char limit. Then added some error trapping to prompt the user to specify a name if the one built automatically by SADPhishes was >200.  
#=================
#Version 1.0.5
# Removed HasAttachment:True from ContentMatchQuery because it was causing eDiscovery Searches to fail
# Changed all Variables set by SADPhishes from $VarName to $Script:VarName to fix variable cleanup when returning to the top menu after running a search.
# Menu options back the user out to the main menu and clear out Vars from previous searches instead of requiring that they exit the script and re-run to run another search.
# Added some debugging options to the main menu.
# Other minor bug fixes
#=================
#Version 1.0.4
# Added options to create an Exchange In-place eDiscovery Search from the Compliance Search results.
# The option to execute the eDiscovery search is completely experimental. It knocked Exchange offline during testing. Not recommended in Prod.
#=================
#Version 1.0.3
# Added AttachmentNameOptions and AttachmentNameMenu functions to search for emails with a specific Attachment name. 
# Added an option for Attachment Name to the workflow of all searches
# Added an Attachment Name Only search option.
# Added an option for a Pre-Built Suspicious Attachment Types Search, and new functions in that workflow that don't allow for delete. This is for info-gathering only.
#=================
# Version 1.0.2
# Modified ComplianceSearch function to add a TimeStamp to SearchName to make it unique if an identical search is re-run.
# Modified ThisSearchMailboxCount function to display a warning if the Compliance search returns 500 source mailboxes.
#=================
# Version 1.0.1
# Added ThisSearchMailboxCount function to display the number of mailboxes and a list of email addresses with Compliance Search Hits
# Added ExchangeSearchLocationOptions and ExchangeSearchLocationMenu functions so the user can choose to search all Exchange Locations, or limit the search targets based on the Email Address associated with a Mailbox, Distribution Group, or Mail-Enabled Security Group
#=================




# Who doesn't like gratuitous ascii art?
Function DisplayBanner {
	Write-Host ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:                      (:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:( (:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(  :(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:  (:(:(:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:( (:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(: :(:(:  (:(:(:  (:(  :(:(:(  :(: :(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:( (:(:(:(  :(:  (:(:(:(  :(:  (:(:( (:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(: :(:(:(:(:  (  :(:(:(:(:  (  :(:(:(:  (:(:(<><<><<><<><<><  <<><<><<><<><<><<><<><<><<><  <"
	Write-Host ":(:(: :(:(:(:(:(   (:(:(:(:(:(   (:(:(:(:( (:(:(<><<><<><<>    <>    <><<><<><<><<><<><<><<>   <"
	Write-Host ":(:(: :(:(:(:(  :(   (:(:(:(  :(   (:(:(:( (:(:(<><<><<     <><<><<>     ><<><<><><<><<     <<><"
	Write-Host ":(:(: :(:(:(:  (:(:(:  (:(:  (:(:(: :(:(:( (:(:(<><     <<><<><<><<><<><     <<><><     ><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(<    <<><<><<><<><<><<><<><    ><    ><<><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(<><     <<><<><<><<><<><     <<><><     ><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:           :(:(:(:(:(:( (:(:(<><<><<     <><<><<>     ><<><<><><<><<     <<><"
	Write-Host ":(:(: :(:(:(:(:(:   :(:(:(:(:   :(:(:(:(:( (:(:(<><<><<><<>    <>    <><<><<><<><<><<><<><<>   <"
	Write-Host ":(:(:( (:(:(:(:   :(:(:(:(:(:(:   :(:(:(: :(:(:(<><<><<><<><<><  <<><<><<><<><<><<><<><<><<><  <"
	Write-Host ":(:(:(: :(:(:(   (:(:(:(:(:(:(:(   (:(:( (:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:( (:(:(: :(:(:(:(:(:(:(:(: :(:(: :(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:( (:(:(:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:  (:(:(:(:(:(:(:(:(:(:(:  (:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:                       :(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host "================================================================================================"
	Start-Sleep -m 200
	Write-Host "             _____             _____                  _____  _     _     _                 "
	Write-Host "            / ____|    /\     |  __ \                |  __ \| |   (_)   | |                "
	Write-Host "           | (___     /  \    | |  | |               | |__) | |__  _ ___| |__   ___  ___   "
	Write-Host "            \___ \   / /\ \   | |  | |               |  ___/| '_ \| / __| '_ \ / _ \/ __|  "
	Write-Host "            ____) | / ____ \ _| |__| |               | |    | | | | \__ \ | | |  __/\__ \  "
	Write-Host "   _____   |_____(_)_/    \_(_)_____(_)             _|_|_   |_| |_|_|___/_| |_|\___||___/  "
	Write-Host "  / ____|                   | |          ___       |  __ \          | |                    "
	Write-Host " | (___   ___  __ _ _ __ ___| |__       ( _ )      | |  | | ___  ___| |_ _ __ ___  _   _   "
	Write-Host "  \___ \ / _ \/ _' | '__/ __| '_ \      / _ \/\    | |  | |/ _ \/ __| __| '__/ _ \| | | |  "
	Write-Host "  ____) |  __/ (_| | | | (__| | | |    | (_>  <    | |__| |  __/\__ \ |_| | | (_) | |_| |  "
	Write-Host " |_____/ \___|\__,_|_|  \___|_| |_|     \___/\/    |_____/ \___||___/\__|_|  \___/ \__, |  "
	Write-Host "                                                             _________________________/ |  "
	Write-Host "                                                            |@EdwardsCP v1.0.12 2018___/   "
	Write-Host "================================================================================================"
	Start-Sleep -m 200
	Write-Host "===============================================================" -ForegroundColor Yellow
	Write-Host "== Exchange 2016 Compliance (S)earch (A)nd (D)estroy Phishes ==" -ForegroundColor Yellow
	Write-Host "===============================================================" -ForegroundColor Yellow
	Write-Host "                                                               " -ForegroundColor Yellow
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "------------------------!!!Warning!!!--------------------------" -ForegroundColor Red
	Write-Host "If you use this script to delete emails, there is no automatic " -ForegroundColor Red
	Write-Host "method to undo the removal of those emails.                    " -ForegroundColor Red
	Write-Host "USE AT YOUR OWN RISK!                                          " -ForegroundColor Red
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "---------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "Important Info:" -ForegroundColor Yellow
	Write-Host "Microsoft's Compliance Search can search your entire environment, but will only return a maximum of 500 Mailboxes in the results." -ForegroundColor Yellow
	Write-Host "Microsoft's Compliance Search Purge Action will purge a maximum of 10 Items per Mailbox." -ForegroundColor Yellow
	Write-Host "Microsoft's Compliance Search Purge Action moves items to the user's Recoverable Items folder, and they will remain there based on the Retention Period that is configured for the mailbox." -ForegroundColor Yellow
	Write-Host "Microsoft's Compliance Search results will return Items that were already purged (and are located in the Recoverable Items folder)." -ForegroundColor Yellow
	Write-Host "Please consider those limitations when using SADPhishes." -ForegroundColor Yellow
	Write-Host "---------------------------------------------------------------" -ForegroundColor Yellow
	SearchTypeMenu
	}

#Function for SearchType Menu Options Display
Function SearchTypeOptions {
	Write-Host "SEARCH OPTIONS MENU" -ForegroundColor Green
	Write-Host "What type of search are you going to perform?" -ForegroundColor Yellow
	Write-Host "---------- New Searches ----------"
	Write-Host "    [1] Subject and Sender Address and Date Range"
	Write-Host "    [2] Subject and Date Range"
	Write-Host "    [3] Subject and Sender Address"
	Write-Host "    [4] Subject Only"
	Write-Host "    [5] Sender Address Only (DANGEROUS)"
	Write-Host "    [6] Sender and Date Range"
	Write-Host "    [7] Attachment Name Only"
	Write-Host "    [8] Pre-Built Suspicious Attachment Types Search"
	Write-Host "    [9] Extract Subject and Sender Address from a text file containing EMail Headers"
	Write-Host " "
	Write-Host "--- Execute Existing Searches ---"
	Write-Host "    [10] View and Run an existing Compliance Search"
	#Write-Host "    [11] View and Run an existing eDiscovery Search"
	Write-Host " "
	Write-host "------- Debugging Options -------"
	Write-Host "    [X] gci variable:"
	Write-Host "    [Y] Print SADPhishesVars"
	Write-Host "    [Z] Clear SADPhishesVars"
	Write-Host " "
	Write-Host "------------- Quit --------------"
	Write-Host "    [Q] Quit"
	Write-Host "---------------------------------"
}	
	
#Function for Search Type Menu
Function SearchTypeMenu{
	Do {	
		SearchTypeOptions
		CreateSADPhishesNullVars
		$script:SearchType = Read-Host -Prompt 'Please enter a selection from the menu (1 - 10, X, Y, Z, or Q) and press Enter'
		switch ($script:SearchType){
			'1'{
				$script:Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$script:Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$script:DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$script:DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$script:DateRangeSeparator = ".."
				$script:DateRange = $script:DateStart + $script:DateRangeSeparator + $script:DateEnd
				$script:ContentMatchQuery = "(Received:$script:DateRange) AND (From:$script:Sender) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'2'{
				$script:Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$script:DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$script:DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$script:DateRangeSeparator = ".."
				$script:DateRange = $script:DateStart + $script:DateRangeSeparator + $script:DateEnd
				$script:ContentMatchQuery = "(Received:$script:DateRange) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'3'{
				$script:Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$script:Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$script:ContentMatchQuery = "(From:$script:Sender) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'4'{
				$script:Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$script:ContentMatchQuery = "(Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'5'{
				Do {
					Write-Host "WARNING: Are you sure you want to search based on only Sender Address?" -ForegroundColor Red
					Write-Host "WARNING: This has the potential to return many results and delete many emails." -ForegroundColor Red
					$script:DangerousSearch = Read-Host -Prompt 'After reading the warning above, would you like to proceed? [Y]es or [Q]uit'
					switch ($script:DangerousSearch){
						'Y'{
							$script:Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
							$script:ContentMatchQuery = "(From:$script:Sender)"
							AttachmentNameMenu
						}
						'q'{
							Read-Host -Prompt "Please press Enter to return to the Search Options Menu"
							ClearSADPhishesVars
							SearchTypeMenu
						}
					}
				}
				until ($script:DangerousSearch -eq 'q')
			}
			'6'{
				$script:Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$script:DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$script:DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$script:DateRangeSeparator = ".."
				$script:DateRange = $script:DateStart + $script:DateRangeSeparator + $script:DateEnd
				$script:ContentMatchQuery = "(Received:$script:DateRange) AND (From:$script:Sender)"
				AttachmentNameMenu			
			}
			'7'{
				$script:AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. SADPhishes.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'8'{
				Write-Host "You have chosen to conduct the SADPhishes Pre-Built Suspicious Attachment Types Search." -ForegroundColor Yellow
				Write-Host "This search will return a list of Mailboxes that contain Attachments with specific file extensions." -ForegroundColor Yellow
				Write-Host "This search is a Search-Only option, with no Delete built into the SADPhishes Workflow." -ForegroundColor Yellow
				Write-Host "Take these results and investigate." -ForegroundColor Yellow
				Read-Host -Prompt "After you have read the information about this Suspicious Attachment Search, Press Enter to continue."
				$script:ContentMatchQuery = "((Attachment:'.ade') OR (Attachment:'.adp') OR (Attachment:'.apk') OR (Attachment:'.bas') OR (Attachment:'.bat') OR (Attachment:'.chm') OR (Attachment:'.cmd') OR (Attachment:'.com') OR (Attachment:'.cpl') OR (Attachment:'.dll') OR (Attachment:'.exe') OR (Attachment:'.hta') OR (Attachment:'.inf') OR (Attachment:'.iqy') OR (Attachment:'.jar') OR (Attachment:'.js') OR (Attachment:'.jse') OR (Attachment:'.lnk') OR (Attachment:'.mht') OR (Attachment:'.msc') OR (Attachment:'.msi') OR (Attachment:'.msp') OR (Attachment:'.mst') OR (Attachment:'.ocx') OR (Attachment:'.pif') OR (Attachment:'.pl') OR (Attachment:'.ps1') OR (Attachment:'.reg') OR (Attachment:'.scr') OR (Attachment:'.sct') OR (Attachment:'.shs') OR (Attachment:'.slk') OR (Attachment:'.sys') OR (Attachment:'.vb') OR (Attachment:'.vbe') OR (Attachment:'.vbs') OR (Attachment:'.wsc') OR (Attachment:'.wsf') OR (Attachment:'.wsh'))"
				ExchangeSearchLocationMenu
			}
			'9'{
				Write-Host "You have chosen to have SADPhishes open a Text file containing the Headers" -ForegroundColor Yellow
				Write-Host "from a sample EMail.  Please select the text file to open in the dialog box" -ForegroundColor Yellow
				Write-Host "that will open when you proceed." -ForegroundColor Yellow
				Read-Host -Prompt "After you have read the information above, Press Enter to proceed."
				ParseEmailHeadersFile
			}
			'10'{
				RunPreviousComplianceSearch
			}
			#'11'{
			#	RunPreviousEDiscoverySearch
			#}
			'q'{
				Write-Host "Thanks for using SADPhishes!" -ForegroundColor Yellow
				Exit
			}
			'x'{
			gci variable:
			}
			'y'{
			PrintSADPhishesVars
			}
			'z'{
			ClearSADPhishesVars
			}
		}
	}
	until ($script:SearchType -eq 'q')
}

#Function to Open a Text file containing email headers, parse each line to find the From:, Subject:, and Date: values, and output to the results.
Function ParseEmailHeadersFile{
	$script:EmailHeadersFile = Get-FileName
	$script:EmailHeadersLines = Get-Content $Script:EmailHeadersFile
	Write-Host "=======================================================" -ForegroundColor Yellow
	Foreach ($script:EmailHeadersLine in $script:EmailHeadersLines){
		$Script:FromHeaderMatches = $script:EmailHeadersLine -match '^From:.*<(.*@.*)>$'
		$Script:SubjectHeaderMatches = $script:EmailHeadersLine -match '^Subject: (.*)$'
		$Script:DateHeaderMatches = $script:EmailHeadersLine -match '^Date: (([a-zA-Z][a-zA-Z][a-zA-Z]), (\d{1,2}) ([a-zA-Z][a-zA-Z][a-zA-Z]) (\d{4}).*)$'
		If ($Script:FromHeaderMatches) {
			$Script:Sender = $matches[1]
			Write-Host "SADPhishes found this Sender Address..." -ForegroundColor Yellow
			Write-Host $script:Sender
		}
        
		If ($script:SubjectHeaderMatches){
			$Script:Subject = $matches[1]
			Write-Host "SADPhishes Found this Subject..." -ForegroundColor Yellow
			Write-Host $Script:Subject
			#check to see if the subject is UTF-8 encoded, and extract plain text to use for search if it is.
			If ($Script:Subject -match '^=\?UTF-8\?B\?(.*)\?\=$') {
				$Script:SubjectB64Encode = $matches[1]
				$Script:SubjectB64Decode = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("$Script:SubjectB64Encode"))
				Write-Host "========================================================================="
				Write-Host "The Subject found in the Headers was UTF-8 Encoded instead of plain text."
				Write-Host "Decoding this UTF-8 Encoded Subject..."
				Write-Host $Script:Subject -ForegroundColor Yellow
				Write-host "Base64 encoded string..."
				Write-host $Script:SubjectB64Encode -ForegroundColor Yellow
				Write-Host "Decoded plain text Subject for SADPhishes Search..."
				Write-Host $Script:SubjectB64Decode -ForegroundColor Yellow
				Write-Host "========================================================================="
				$Script:Subject = $Script:SubjectB64Decode
			}
		}
		
        If ($script:DateHeaderMatches){
			$Script:DateFromHeader = $matches[1]
			$Script:DateFromHeaderDayOfWeek = $matches[2]
			$Script:DateFromHeaderDayOfMonth = $matches[3]
			$Script:DateFromHeaderMonth = $matches[4]
			$Script:DateFromHeaderYear = $matches[5]
			If ($Script:DateFromHeaderMonth -eq 'Jan'){
				$Script:DateFromHeaderMonthNum = "1"
			}
			If ($Script:DateFromHeaderMonth -eq 'Feb'){
				$Script:DateFromHeaderMonthNum = "2"
			}
			If ($Script:DateFromHeaderMonth -eq 'Mar'){
				$Script:DateFromHeaderMonthNum = "3"
			}
			If ($Script:DateFromHeaderMonth -eq 'Apr'){
				$Script:DateFromHeaderMonthNum = "4"
			}
			If ($Script:DateFromHeaderMonth -eq 'May'){
				$Script:DateFromHeaderMonthNum = "5"
			}
			If ($Script:DateFromHeaderMonth -eq 'Jun'){
				$Script:DateFromHeaderMonthNum = "6"
			}
			If ($Script:DateFromHeaderMonth -eq 'Jul'){
				$Script:DateFromHeaderMonthNum = "7"
			}
			If ($Script:DateFromHeaderMonth -eq 'Aug'){
				$Script:DateFromHeaderMonthNum = "8"
			}
			If ($Script:DateFromHeaderMonth -eq 'Sep'){
				$Script:DateFromHeaderMonthNum = "9"
			}
			If ($Script:DateFromHeaderMonth -eq 'Oct'){
				$Script:DateFromHeaderMonthNum = "10"
			}
			If ($Script:DateFromHeaderMonth -eq 'Nov'){
				$Script:DateFromHeaderMonthNum = "11"
			}
			If ($Script:DateFromHeaderMonth -eq 'Dec'){
				$Script:DateFromHeaderMonthNum = "12"
			}
		$Script:DateFromHeaderFormatted = "$Script:DateFromHeaderMonthNum" + "/" + "$Script:DateFromHeaderDayOfMonth" + "/" + "$Script:DateFromHeaderYear"
		Write-Host "SADPhishes Found this Date in the Headers..." -ForegroundColor Yellow
		Write-Host $Script:DateFromHeader
		Write-Host $Script:DateFromHeaderFormatted
        }
    }
    Write-Host "=======================================================" -ForegroundColor Yellow
UseParsedEmailHeadersSender	
}

#Function to give the user the option to use the Sender extracted from the headers, or specify their own
Function UseParsedEmailHeadersSender{	
    Do{
        Write-Host "Do you want to use the Sender [$script:sender] as part of your search criteria? [Y]es or [N]o" -ForegroundColor Yellow
        $script:UseSenderFromHeaderFile = Read-Host -Prompt "Please answer the question with Y or N and press Enter to proceed."
        Switch ($script:UseSenderFromHeaderFile){
            'Y'{
                #$Script:Sender already set correctly
				UseParsedEmailHeadersSubject
            }
            'N'{
                $Script:Sender = Read-Host -Prompt "Please enter the exact Sender (From:) address of the Email you would like to search for"
                UseParsedEmailHeadersSubject
            }
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu
			}
        }
    }
    Until ($Script:UseSenderFromHeaderFile -eq 'q')
}

#Function to give the user the option to use the Subject extracted from the headers, or specify their own
Function UseParsedEmailHeadersSubject{	
	Do{
        Write-Host "Do you want to use the Subject [$script:subject] as part of your search criteria? [Y]es or [N]o" -ForegroundColor Yellow
        $script:UseSubjectFromHeaderFile = Read-Host -Prompt "Please answer the question with Y or N and press Enter to proceed."
        Switch ($script:UseSubjectFromHeaderFile){
            'Y'{
                #$Script:Subject already set correctly
				UseParsedEmailHeadersDate
            }
            'N'{
                $Script:Subject = Read-Host -Prompt "Please enter the exact Subject of the Email you would like to search for"
                UseParsedEmailHeadersDate
            }
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu
			}
        }
    }
    Until ($Script:UseSubjectFromHeaderFile -eq 'q')

}	

#Function for UseParsedEmailHeadersDate Menu Options Display
Function UseParsedEmailHeadersDateOptions {
	Write-Host "Email Headers Date Options Menu" -ForegroundColor Green
	Write-Host "This Date was found in the Headers [$script:DateFromHeader]" -ForegroundColor Yellow
	Write-Host "Please select an option from this menu to proceed..."
	Write-Host "[1] Search for emails using only that date"
	Write-Host "[2] Specify your own date range"
	Write-Host "[3] Do not include a date in your search"
}

#Function to give the user the option to use the Date extracted from the headers, or specify their own
Function UseParsedEmailHeadersDate{
	Do{
		UseParsedEmailHeadersDateOptions
		$Script:UseDateFromHeaderFile = Read-Host -Prompt "Please enter a selection from the menu (1, 2, or 3) and press Enter to proceed."
		Switch ($Script:UseDateFromHeaderFile){
			'1'{
				$script:DateStart = $Script:DateFromHeaderFormatted
				$script:DateEnd = $Script:DateFromHeaderFormatted
				$script:DateRangeSeparator = ".."
				$script:DateRange = $script:DateStart + $script:DateRangeSeparator + $script:DateEnd
				$script:ContentMatchQuery = "(Received:$script:DateRange) AND (From:$script:Sender) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'2'{
				$script:DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$script:DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$script:DateRangeSeparator = ".."
				$script:DateRange = $script:DateStart + $script:DateRangeSeparator + $script:DateEnd
				$script:ContentMatchQuery = "(Received:$script:DateRange) AND (From:$script:Sender) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'3'{
				$script:ContentMatchQuery = "(From:$script:Sender) AND (Subject:'$script:Subject')"
				AttachmentNameMenu
			}
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu
			}
		}
	}
	Until ($Script:UseDateFromHeaderFile -eq 'q')
}


#Function for AttachmentName Menu Options Display
Function AttachmentNameOptions {
	Write-Host "ATTACHMENT OPTIONS MENU" -ForegroundColor Green
	Write-Host "Do you want to search for EMails containing an Attachment with a specific File Name?" -ForegroundColor Yellow
	Write-Host "[1] No"
	Write-Host "[2] Yes"
	Write-Host "[Q] Quit and Return to the Search Options Menu"
}

#Function for AttachmentName Menu
Function AttachmentNameMenu {
	Do{
		AttachmentNameOptions
		$script:AttachmentNameSelection = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and Press Enter'
		switch ($script:AttachmentNameSelection){
			'1'{
				ExchangeSearchLocationMenu
			}
			'2'{
				$script:AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. SADPhishes.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu
			}
		}
	}
	until ($script:AttachmentNameSelection -eq 'q')
}

#Function for ExchangeSearchLocation Menu Options Display
Function ExchangeSearchLocationOptions {
	Write-Host ""
	Write-Host "LOCATION OPTIONS MENU" -ForegroundColor Green
	Write-Host "Do you want to search All Mailboxes, or restrict your search to a specific Mailbox, Distribution Group, or Mail-Enabled Security Group?" -ForegroundColor Yellow
	Write-Host "If you restrict your search, you might leave phishes in other places." -ForegroundColor Yellow
	Write-Host "[1] All Mailboxes"
	Write-Host "[2] A specific MailBox, Distribution Group, or Mail-Enabled Security Group"
	Write-Host "[Q] Quit and Return to the Search Options Menu"
}

#Function for ExchangeSearchLocation Menu
Function ExchangeSearchLocationMenu {
	Do {
		ExchangeSearchLocationOptions
		$script:ExchangeSearchLocation = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and press Enter'
		switch ($script:ExchangeSearchLocation){
			'1'{
				$script:ExchangeLocation = "all"
				UserSetSeachNameMenu
			}
			'2'{
				$script:ExchangeLocation = Read-Host -Prompt 'Please enter the EMail Address of the MailBox or Group you would like to search within'
				UserSetSeachNameMenu
			}
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu
			}
		}
	}
	until ($script:SearchType -eq 'q')
}

#Function for UserSetSearchName Menu Options Display
Function UserSetSearchNameOptions{
	Write-Host ""
	Write-Host "USER SPECIFIED SEARCH NAME MENU" -ForegroundColor Green
	Write-Host "Do you want to specify your own name for this search?" -ForegroundColor Yellow
	Write-Host "If you don't need to specify your own name, SADPhishes will automatically create a name based on the search criteria you have specified." -ForegroundColor Yellow
	Write-Host "If you aren't sure what to choose, pick No so you can see how SADPhishes builds Search Names." -ForegroundColor Yellow
	Write-Host "[1] No"
	Write-Host "[2] Yes" 
	Write-Host "[Q] Quit and Return to the Search Options Menu"
}
#Function to allow the user to specify their own Search Name
Function UserSetSeachNameMenu {
	Do {
		UserSetSearchNameOptions
		$Script:UserSetSearchNameChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and press Enter'
		switch ($Script:UserSetSearchNameChoice){
			'1'{
				AddDescriptionMenu
			}
			'2'{
				$Script:SearchName = Read-Host -Prompt 'Please enter a Name for this search'
				AddDescriptionMenu
			}
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu			
			}
		}
	}
	Until ($Script:UserSetSearchNameChoice -eq 'q')
}


#Function for AddDescription Menu Options Display
Function AddDescriptionOptions {
	Write-Host ""
	Write-Host "ADD DESCRIPTION MENU" -ForegroundColor Green
	Write-Host "Do you want to specify a Description for this search?" -ForegroundColor Yellow
	Write-Host "You might want to this to add some additional details or Incident/Tracking #'s to the Search" -ForegroundColor Yellow
	Write-Host "[1] No"
	Write-Host "[2] Yes" 
	Write-Host "[Q] Quit and Return to the Search Options Menu"
}


#Function to allow the user to specify a Description for their Compliance Search
Function AddDescriptionMenu {
	Do {
		AddDescriptionOptions
		$script:AddDescription = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and press Enter'
		switch ($Script:AddDescription){
			'1'{
				ComplianceSearch
			}
			'2'{
				$Script:SearchDescription = Read-Host -Prompt 'Please enter a Description for this search'
				ComplianceSearch
			}
			'q'{
				ClearSADPhishesVars
				SearchTypeMenu			
			}
		}
	}
	until ($Script:AddDescription -eq 'q')
}



#Function to Re-Run a previous Compliance Search
Function RunPreviousComplianceSearch {
	Write-Host "SADPhishes is going to list all of the existing Compliance Searches" -ForegroundColor Yellow
	Write-Host "They will be in the the format '[#] SearchNameHere', where # is the integer you will use to select which search to run." -ForegroundColor Yellow
	Read-Host -Prompt "After reading the information above, Please press Enter to continue."
	#create an empty array for existing ComplianceSearches
	$script:ComplianceSearches = @()
	#set up an Integer to use for tagging each existing Compliance Search with a number
	$I = 1
	$script:ComplianceSearches = Get-ComplianceSearch
	#For every Compliance Search found, add a NoteProperty named Search number, assign it with our integer, and then increase the Integer by 1 so it's ready for the next Compliance Search in the array.
	$Script:ComplianceSearches | %{$_ | Add-Member -NotePropertyName SearchNumber -NotePropertyValue $I -Force; $I++}
	#set the Integer back to 1 so we can display a list of existing Compliance Searches with the SearchNumber in a bracket so it's displayed similar to our other menus.
	$I = 1
		foreach ($script:ComplianceSearch in $script:ComplianceSearches){
			Write-Host [$I] $Script:ComplianceSearch.Name
			$I++
		}
	#after looking through all of the Compliance Searches in the array, decrease the Integer by 1 so that we can display the last used value in the instruction below.
	$I--
	Do {
		$Script:ComplianceSearchNumberSelection = Read-Host -Prompt "Please enter a Search Number from the list above (1 - $I), and Press Enter to continue"
	}
	Until ($Script:ComplianceSearchNumberSelection -ge 1 -and $Script:ComplianceSearchNumberSelection -le $I)
	#set up variables so our ComplianceSearch Function will run
	$Script:SelectedComplianceSearch = $Script:ComplianceSearches | Where {$_.SearchNumber -eq $Script:ComplianceSearchNumberSelection}
	$script:SearchName = $Script:SelectedComplianceSearch.Name
	$Script:ContentMatchQuery = $script:SelectedComplianceSearch.ContentMatchQuery
	$Script:ExchangeLocation = $script:SelectedComplianceSearch.ExchangeLocation
	Write-Host "==========================================================================="
	Write-Host "Re-Running the existing Compliance Search named... "
	Write-Host $script:SearchName -ForegroundColor Yellow
	Write-Host "...containing the query..."
	Write-Host $script:ContentMatchQuery -ForegroundColor Yellow
	Write-Host "==========================================================================="
	Get-ComplianceSearch -Identity "$script:SearchName"
	ComplianceSearch
}

#Function to Re-Run a previous eDiscovery Search. Commented out because this needs work.  When this gets handed off to the ShowEDiscoverysearchMenu function, that function and others further in the workflow require variables that assume you got there by first running a ComplianceSearch.
#Function RunPreviousEDiscoverySearch {
	##create an empty array for existing eDiscovery MailboxSearches
	#$script:EDiscoverySearches = @()
	##set up an Integer to use for tagging each existing eDiscovery Mailbox Search with a number
	#$I = 1
	#$script:EDiscoverySearches = Get-MailboxSearch
	##For every eDiscovery Mailbox Search found, add a NoteProperty named Search number, assign it with our integer, and then increase the Integer by 1 so it's ready for the next Compliance Search #in the array.
#	$Script:EDiscoverySearches | %{$_ | Add-Member -NotePropertyName SearchNumber -NotePropertyValue $I -Force; $I++}
	##set the Integer back to 1 so we can display a list of existing eDiscovery Mailbox Searches with the SearchNumber in a bracket so it's displayed similar to our other menus.
	#$I = 1
		#foreach ($script:EDiscoverySearch in $script:EDiscoverySearches){
			#Write-Host [$I] $Script:EDiscoverySearch.Name
			#$I++
		#}
	##after looking through all of the eDiscovery Mailbox Searches in the array, decrease the Integer by 1 so that we can display the last used value in the instruction below.
	#$I--
	#$Script:EDiscoverySearchNumberSelection = Read-Host -Prompt "Please select enter a selection from above (1 - $I), and Press Enter to continue"
	##set up variables so our ComplianceSearch Function will run
	#$Script:SelectedEDiscoverySearch = $Script:EDiscoverySearches | Where {$_.SearchNumber -eq $Script:EDiscoverySearchNumberSelection}
	#$script:EDiscoverySearchName = $Script:SelectedEDiscoverySearch.Name
	#$script:ThisEDiscoverySearch = Get-MailboxSearch $script:EDiscoverySearchName
	#Write-Host "==========================================================================="
	#Write-Host "Loading the eDiscovery Search named... "
	#Write-Host $Script:EDiscoverySearchName -ForegroundColor Yellow
	#Write-Host "==========================================================================="
	#Get-MailboxSearch -Identity "$Script:EDiscoverySearchName"
#	
	#Start-Sleep 1
	#ShowEDiscoverySearchMenu
#}


#Function for the Compliance Search Creation and Execution
Function ComplianceSearch {
	# If SelectedComplianceSearch (used to re-run an existing search) is Null, go through the full process of creating a Searchname, Checking name length, notifying about subjects being wildcard searches, and creating the search.  Otherwise, bypass all that and get right to running the existing search.
	If ($Script:SelectedComplianceSearch -eq $null){
		#If UserSetSearchNameChoice is 1 (meaning the user didn't choose to set their own Search Name), Set SearchName based on SearchType
		If ($Script:UserSetSearchNameChoice -eq '1'){
			switch ($script:SearchType){
					'1'{
						$script:SearchName = "Remove Subject [$script:Subject] Sender [$script:Sender] DateRange [$script:DateRange] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'2'{
						$script:SearchName = "Remove Subject [$script:Subject] DateRange [$script:DateRange] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'3'{
						$script:SearchName = "Remove Subject [$script:Subject] Sender [$script:Sender] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'4'{
						$script:SearchName = "Remove Subject [$script:Subject] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'5'{
						$script:SearchName = "Remove Sender [$script:Sender] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'6'{
						$script:SearchName = "Remove Sender [$script:Sender] DateRange [$script:DateRange] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
					'7'{
						$script:SearchName = "Remove Exchange Location [$script:ExchangeLocation] Phishing Message"
					}
					'8'{
						$script:SearchName = "SADPhishes Pre-Built Suspicious Attachment Types Search Exchange Location [$script:ExchangeLocation]"
					}
					'9'{
						$script:SearchName = "SADPhishes Headers Parsed Subject [$script:Subject] DateRange [$script:DateRange] Sender [$script:Sender] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
					}
			}
			#If an AttachmentName has been specified, Modify SearchName to include it.  
			if ($script:AttachmentName -ne $null){
				$script:SearchName = $script:SearchName + " with Attachment [" + $script:AttachmentName + "]"
				# If a ContentMatchQuery is already set, modify $script:ContentMatchQuery to include the attachment.
				If ($script:ContentMatchQuery -ne $null){
				$script:ContentMatchQuery = "(Attachment:'$script:AttachmentName') AND " + $script:ContentMatchQuery
				}
				# If an AttachmentName has been specified, and a ContentMatchQuery is NOT already set, set the ContentMatchQuery.
				If ($script:ContentMatchQuery -eq $null){
				$script:ContentMatchQuery = "(Attachment:'$script:AttachmentName')"
				}
			}	
			## Timestamp the SearchName (to make it unique), then Create and Execute a New Compliance Search based on the user set Variables
			#$script:TimeStamp = Get-Date -Format o | foreach {$_ -replace ":", "."} #Timestamp for search name not used anymore, but leaving the var here for now.
			#$script:SearchName = $script:SearchName + " " + $script:TimeStamp
		}	
		
		# Name the Compliance Search using the Name that has been built using the search criteria, followed by an integer. To handle repeat searches of matching criteria, increase the integer until you hit a Search name that doesn't already exit.
		$I = 1;
		$Script:ComplianceSearches = Get-ComplianceSearch;
			while ($true){
				$found = $false
				$script:ThisComplianceSearchRun = "$script:SearchName$I"
				foreach ($Script:ComplianceSearch in $Script:ComplianceSearches){
					if ($Script:ComplianceSearch.Name -eq $Script:ThisComplianceSearchRun){
						$found = $true;
						break;
					}
				}
				if (!$found){
					break;
				}
				$I++;
			}
		$Script:SearchName = "$Script:SearchName$I"
		# If the Compliance SearchName is >200 characters, prompt the user to supply a new SearchName, then append an Integer. To handle repeat searches with the same user-defined name, increase the integer until you hit a Search name that doesn't already exist.
		Do {
			If ($Script:SearchName.length -gt 200) {
				Write-Host "============WARNING - The Search Name is too long!============" -ForegroundColor Red
				Write-Host "============WARNING - The Search Name is too long!============" -ForegroundColor Red
				Write-Host "============WARNING - The Search Name is too long!============" -ForegroundColor Red
				Write-Host "This Search Name for your Compliance Search..."
				Write-Host $Script:SearchName -ForegroundColor Yellow
				Write-Host "...is this many Characters in length..."
				Write-Host $Script:SearchName.length -ForegroundColor Yellow
				Write-Host "...and that is greater than the 200 Characters that Microsoft allows."
				Write-Host "Please supply a new Search Name so that SADPhishes can proceed."
				$Script:SearchName = Read-Host -Prompt "After reading the information above, please enter a new Search Name that is less than 198 Characters."
				$I = 1;
				$Script:ComplianceSearches = Get-ComplianceSearch;
					while ($true){
						$found = $false
						$script:ThisComplianceSearchRun = "$script:SearchName$I"
						foreach ($Script:ComplianceSearch in $Script:ComplianceSearches){
							if ($Script:ComplianceSearch.Name -eq $Script:ThisComplianceSearchRun){
								$found = $true;
								break;
							}
						}
						if (!$found){
							break;
						}
						$I++;
					}
				$Script:SearchName = "$Script:SearchName$I"
			}
		}
		Until ($Script:SearchName.length -le 200)
		Write-Host "==========================================================================="
		Write-Host "Creating a new Compliance Search with the name..."
		Write-Host $script:SearchName -ForegroundColor Yellow
		if ($script:AddDescription -eq '2') {
			Write-Host "...with the description..."
			Write-Host $Script:SearchDescription -ForegroundColor Yellow
		}
		Write-Host "...using the query..."
		Write-Host $script:ContentMatchQuery -ForegroundColor Yellow
		Write-Host "==========================================================================="
		
		#If a Subject was specified, warn the user about Microsoft returning results with additional text before or after the subject that was defined.
		if ($script:Subject -ne $null){
			Write-Host "===========================================================================" -ForegroundColor Yellow
			Write-Host "Warning: Your Compliance Search contained a Subject [$script:Subject]."             -ForegroundColor Yellow
			Write-Host "When you use the Subject property in a query, the search returns all"        -ForegroundColor Yellow
			Write-Host "messages in which the subject line contains the text you are searching for." -ForegroundColor Yellow
			Write-Host "The query doesn't only return exact matches.  For example, if you search"    -ForegroundColor Yellow
			Write-Host "(Subject:SADPhishes), your results will include messages with the subject"   -ForegroundColor Yellow
			Write-Host "'SADPhishes', but also messages with the subjects 'SADPhishes is good!' and" -ForegroundColor Yellow
			Write-Host "'RE: Screw SADPhishes. it sucks!'"                                           -ForegroundColor Yellow
			Write-Host " "                                                                           -ForegroundColor Yellow
			Write-Host "This is just how the Microsoft Exchange Content Search works."               -ForegroundColor Yellow
			Write-Host " "                                                                           -ForegroundColor Yellow
			Write-Host "Please take this into consideration when using the Search Results."          -ForegroundColor Yellow
			Write-Host "===========================================================================" -ForegroundColor Yellow
			Read-Host -Prompt "Please press Enter after reading the warning above."
		
		}
		switch ($script:AddDescription) {
			'1' {
			New-ComplianceSearch -Name "$script:SearchName" -ExchangeLocation $script:ExchangeLocation -ContentMatchQuery $script:ContentMatchQuery
			}
			'2' {
			New-ComplianceSearch -Name "$script:SearchName" -ExchangeLocation $script:ExchangeLocation -ContentMatchQuery $script:ContentMatchQuery -Description "$Script:SearchDescription"
			}
		}
	}	
	Start-ComplianceSearch -Identity "$script:SearchName"
	Get-ComplianceSearch -Identity "$script:SearchName"
	#Display status, then results of Compliance Search
	do{
		$script:ThisSearch = Get-ComplianceSearch -Identity $script:SearchName
		Start-Sleep 2
		Write-Host $script:ThisSearch.Status
	}
	until ($script:ThisSearch.status -match "Completed")

	Write-Host "==========================================================================="
	Write-Host The search returned...
	Write-Host $script:ThisSearch.Items Items -ForegroundColor Yellow
	Write-Host That match the query...
	Write-Host $script:ContentMatchQuery -ForegroundColor Yellow
	ThisSearchMailboxCount
	Write-Host "==========================================================================="
	#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
	if ($script:SearchType -match "8"){
		Write-host "===================================================="  -ForegroundColor Red
		Write-Host "Take the Search Results above and Investigate." -ForegroundColor Red
		Write-host "===================================================="  -ForegroundColor Red
		ShowNoDeleteMenu
	}
	#If the search was any other type, show the regular Actions menu that allows Delete.
	ShowMenu
}

#Function to count and list Mailboxes with Search Hits.  Code mostly taken from a MS TechNet article.
Function ThisSearchMailboxCount {
	$script:ThisSearchResults = $script:ThisSearch.SuccessResults;
	if (($script:ThisSearch.Items -le 0) -or ([string]::IsNullOrWhiteSpace($script:ThisSearchResults))){
               Write-Host "!!!The Compliance Search didn't return any useful results!!!" -ForegroundColor Red
	}
	$script:mailboxes = @() #create an empty array for mailboxes
	$script:ThisSearchResultsLines = $script:ThisSearchResults -split '[\r\n]+'; #Split up the Search Results at carriage return and line feed
	foreach ($script:ThisSearchResultsLine in $script:ThisSearchResultsLines){
		# If the Search Results Line matches the regex, and $matches[2] (the value of "Item count: n") is greater than 0)
		if ($script:ThisSearchResultsLine -match 'Location: (\S+),.+Item count: (\d+)' -and $matches[2] -gt 0){ 
			# Add the Location: (email address) for that Search Results Line to the $mailboxes array
			$script:mailboxes += $matches[1]; 
		}
	}
	$script:MailboxesWithHitsCount = $script:mailboxes.count
	Write-Host "Number of mailboxes that have Search Hits..."
	Write-Host $script:mailboxes.Count -ForegroundColor Yellow
	Write-Host "List of mailboxes that have Search Hits..."
	write-Host $script:mailboxes -ForegroundColor Yellow
	if ($script:MailboxesWithHitsCount -gt 499) {
		Write-Host "============WARNING - There are 500 or more Mailboxes with results!============" -ForegroundColor Red
		Write-Host "Microsoft's Compliance Search can search everywhere, but only returns the top" -ForegroundColor Red
		Write-Host "500 Mailboxes with the most hits that match the search!" -ForegroundColor Red
		Write-Host " " 
		Write-Host "If you use this search to delete Email Items, you will need to run the same" -ForegroundColor Red
		Write-Host "query again to return more mailboxes if there are more than 500 with hits." -ForegroundColor Red
		Read-Host -Prompt "Please press Enter after reading the warning above."
	}
}

#Function to show the full action menu of options
Function MenuOptions{
	Write-host "===================================================="
	Write-Host "COMPLIANCE SEARCH ACTIONS MENU" -ForegroundColor Green
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the Compliance Search results."
	Write-Host "[2] Delete the Items (move them to Deleted Recoverable Items). WARNING: No automated way to restore them!"
	Write-Host "[3] Create an Exchange In-Place eDiscovery Search from the Compliance Search results."
	Write-Host "[4] Delete this search and Return to the Search Options Menu."
	Write-Host "[5] Return to the Search Options Menu."
	}
	
#Function for full action menu
Function ShowMenu{
	Do{
		MenuOptions
		$script:MenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1 - 5), and press Enter'
		switch ($script:MenuChoice){
			'1'{
			$script:ThisSearch | Format-List
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-Host "Please review the output above" -ForegroundColor Red
			Write-host "After reviewing, please make another selection below"  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			ShowMenu
			}
			
			'2'{
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				$script:DangerousPurge = Read-Host -Prompt 'After reading the warning above, would you like to proceed? [Y]es or [Q]uit'
				Do {
					switch ($script:DangerousPurge){
						'Y'{
							$script:PurgeSuffix = "_purge"
							$script:PurgeName = $script:SearchName + $script:PurgeSuffix
							Write-Host "==========================================================================="
							Write-Host "Creating/Running a Compliance Search Purge Action with the name..."
							Write-Host $script:PurgeName -ForegroundColor Yellow
							Write-Host "==========================================================================="
							New-ComplianceSearchAction -SearchName "$script:SearchName" -Purge #-PurgeType SoftDelete
								do{
									$script:ThisPurge = Get-ComplianceSearchAction -Identity $script:PurgeName
									Start-Sleep 2
									Write-Host $script:ThisPurge.Status
								}
								until ($script:ThisPurge.Status -match "Completed")
							$script:ThisPurge | Format-List
							$script:ThisPurgeResults = $script:ThisPurge.Results
							#commented out - problems with the matching when ThisPurge.Results contains details for multiple mailboxes (if more than 1 was included in Search Results)...it rolls to new lines so the matches don't work because the final } is not on the same line.  Will review this sometime in the future.
							#$Script:ThisPurgeResultsMatches = $script:ThisPurgeResults -match '^Purge Type: SoftDelete; Item count: (\d*); Total size (\d*); Details: {(.*)}$'
							$Script:ThisPurgeResultsMatches = $script:ThisPurgeResults -match 'Purge Type: SoftDelete; Item count: (\d*); Total size (\d*);.*'
							If ($script:ThisPurgeResultsMatches){
								$Script:ThisPurgeResultsItemCount = $matches[1]
								$Script:ThisPurgeResultsTotalSize = $matches[2]
							#commented out - see note above
							#	$Script:ThisPurgeResultsDetails = $matches[3]
								}
							Write-Host "==========================================================="
							Write-Host "SADPhishes Purged this many Items..."
							Write-Host $Script:ThisPurgeResultsItemCount -ForegroundColor Yellow
							Write-Host "...with a total size of..."
							Write-Host $Script:ThisPurgeResultsTotalSize -ForegroundColor Yellow
							#commented out - see note above
							#Write-Host "Potentially useful details below..."
							#Write-host $Script:ThisPurgeResultsDetails -ForegroundColor Yellow
							Write-Host "==========================================================="
							#
							# CONTINUE HERE.  IF $Script:ThisPurgeResultsItemCount is not 0, get this to loop through until it is 0.
							#
							#
							#
							#
							If ($script:ThisPurgeResultsItemCount -eq "0"){
									Write-Host "SADPhishes did not find any items to delete!" -ForegroundColor Red
									Write-Host "SADPhishes did not find any items to delete!" -ForegroundColor Red
									Write-Host "SADPhishes did not find any items to delete!" -ForegroundColor Red
									Write-Host "The initial Compliance Search returned this many items...  "
									Write-Host $script:ThisSearch.Items Items -ForegroundColor Yellow
									Write-Host "...but the Delete/Purge occurred on this many items..."
									Write-Host $Script:ThisPurgeResultsItemCount -ForegroundColor Yellow
									Write-Host "That should be an indication that all of the Items returned by the Compliance Search are already located in the Deleted Recoverable Items folder of each Mailbox!" -ForegroundColor Yellow
									Write-Host "==========================================================="
									Write-Host "You can use the In-Place eDiscovery Search (option 3 presented in the SADPhishes Compliance Search Actions Menu, after the initial search is run) to confirm if that is true."
									Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
									ClearSADPhishesVars
									SearchTypeMenu
								}					
							Write-host "==================================================================================="
							Write-Host "Note: Microsoft's Compliance Search Purge Actions will remove a maximum of 10" -ForegroundColor Yellow
							Write-Host "items per mailbox at one time.  They say it's designed that way because it's" -ForegroundColor Yellow
							Write-Host "an Incident Response Tool and the limit helps ensure that messages are quickly" -ForegroundColor Yellow
							Write-Host "removed." -ForegroundColor Yellow
							Write-Host "The SADPhishes author believes that, in some IR scenarios, that makes sense." -ForegroundColor Yellow
							Write-Host "In other scenarios, it's an unfortunate restriction that should have a bypass " -ForegroundColor Yellow
							Write-host "method.  We have tested looping purges until the Purged Item count hits 0, but" -ForegroundColor Yellow
							Write-host "haven't found a consistent way to get it to work from this same SADPhishes session." -ForegroundColor Yellow
							Write-host "==================================================================================="
							Write-host "If you think this Purge may have left items behind, you should run another Search" -ForegroundColor Yellow
							Write-host "and Delete/Purge until the Item count displayed above is 0." -ForegroundColor Yellow
							Write-Host "The current Purge is complete." -ForegroundColor Red
							Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
							ClearSADPhishesVars
							SearchTypeMenu
						}
						'q'{
							Read-Host -Prompt "Please press Enter to return to the Compliance Search Actions Menu"
							ShowMenu
						}
					}
				}
				Until ($script:DangerousPurge -eq 'q')
			}
			
			'3'{
			CreateEDiscoverySearch
			}
			'4'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
			'5'{
			Write-Host "The previous Compliance Search has not been deleted. Returning to the Search Options Menu" -ForegroundColor Red
			ClearSADPhishesVars
			SearchTypeMenu
			}
			
			'q'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
		}
	}
	Until ($script:MenuChoice -eq 'q')
}

#Function to show the No Delete action menu of options (for Suspicious Attachment Types Search)
Function NoDeleteMenuOptions{
	Write-Host "COMPLIANCE SEARCH ACTIONS MENU (No Delete)" -ForegroundColor Green
	Write-Host "Note: As a precaution, the delete option is not available for a Suspicious Attachment Types Search." -ForegroundColor Yellow
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the search results."
	Write-Host "[2] Create an Exchange In-Place eDiscovery Search from the results."
	Write-Host "[3] Delete this search and Return to the Search Options Menu."
	}
	
#Function for No Delete menu (for Suspicious Attachment Types Search)
Function ShowNoDeleteMenu{
	Do{
		NoDeleteMenuOptions
		$script:NoDeleteMenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2 or 3) and press Enter'
		switch ($script:NoDeleteMenuChoice){
			'1'{
			$script:ThisSearch | Format-List
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-Host "Please review the output above" -ForegroundColor Red
			Write-host "After reviewing, please make another selection below"  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			ShowNoDeleteMenu
			}
			
			'2'{
			CreateEDiscoverySearch
			}
			
			'3'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
			
			'q'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
		}
	}
	Until ($script:MenuChoice -eq 'q')
}

	
#Function to create an eDiscovery Search. Code mostly taken from a MS TechNet article.  
 Function CreateEDiscoverySearch{
	$script:ThisSearchResults = $script:ThisSearch.SuccessResults;
	if (($script:ThisSearch.Items -le 0) -or ([string]::IsNullOrWhiteSpace($script:ThisSearchResults))){
               Write-Host "!!!The Compliance Search didn't return any useful results!!!" -ForegroundColor Red
	}
	$script:mailboxes = @() #create an empty array for mailboxes
	$script:ThisSearchResultsLines = $script:ThisSearchResults -split '[\r\n]+'; #Split up the Search Results at carriage return and line feed
	foreach ($script:ThisSearchResultsLine in $script:ThisSearchResultsLines){
		# If the Search Results Line matches the regex, and $matches[2] (the value of "Item count: n") is greater than 0)
		if ($script:ThisSearchResultsLine -match 'Location: (\S+),.+Item count: (\d+)' -and $matches[2] -gt 0){ 
			# Add the Location: (email address) for that Search Results Line to the $mailboxes array
			$script:mailboxes += $matches[1]; 
		}
	}
	#Name the the EDiscoverySearch (MailboxSearch) using the Compliance Search's name, followed by _MBSearch, followed by an integer. increase the integer until you hit a name that doesn't already exist.
	$script:EDiscoverySearchName = $script:SearchName + "_MBSearch";
	$I = 1;
	$script:MailboxSearches = Get-MailboxSearch;
		while ($true){
			$found = $false
			$script:ThisEDiscoverySearchRun = "$script:EDiscoverySearchName$I"
			foreach ($script:MailboxSearch in $script:MailboxSearches){
				if ($script:MailboxSearch.Name -eq $script:ThisEDiscoverySearchRun){
					$found = $true;
					break;
				}
		}
		if (!$found){
			break;
		}
		$I++;
		}
	$script:ThisEDiscoverySearchName = "$script:EDiscoverySearchName$i"
	Write-Host "==========================================================================="
	Write-Host "Creating a new In-Place eDiscovery Search with the name..."
	Write-Host "$script:ThisEDiscoverySearchName" -ForegroundColor Yellow
	Write-Host "...that will search against these mailboxes..."
	Write-Host $script:mailboxes -ForegroundColor Yellow
	Write-Host "...using the Search Query..."
	Write-Host $script:ContentMatchQuery -ForegroundColor Yellow
	Write-Host "==========================================================================="
	New-MailboxSearch "$script:ThisEDiscoverySearchName" -SourceMailboxes $script:mailboxes -SearchQuery $script:ContentMatchQuery -EstimateOnly
	$script:ThisEDiscoverySearch = Get-MailboxSearch $script:ThisEDiscoverySearchName
	do{
		$script:ThisEDiscoverySearch = Get-MailboxSearch $script:ThisEDiscoverySearchName
		Start-Sleep 1
	}
	Until ($script:ThisEDiscoverySearch -ne $null)
	Write-Host "New In-Place eDiscovery Search Successfully Created!" -ForegroundColor Yellow
	ShowEDiscoverySearchMenu
}


#Function to show the eDiscovery Search Action menu of options
Function EDiscoverySearchMenuOptions{
	Write-host "===================================================="
	Write-Host "EDISCOVERY SEARCH ACTIONS MENU" -ForegroundColor Green
	Write-host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the In-Place eDiscovery Search."
	Write-Host "[2] Start the In-Place eDiscovery Search. (Experimental)"
	Write-Host "[3] Delete the In-Place eDiscovery Search and return to the Compliance Search Actions Menu."
	Write-Host "[4] Return to the top level Search Options Menu"
	}

#Function for the eDiscovery Search Action menu
Function ShowEDiscoverySearchMenu {
	EDiscoverySearchMenuOptions
	$script:EDiscoverySearchMenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or 3), and press Enter'
	Do {
		Switch ($script:EDiscoverySearchMenuChoice){
			'1'{
				$script:ThisEDiscoverySearch | Format-List
				Write-host "===================================================="  -ForegroundColor Red
				Write-host "===================================================="  -ForegroundColor Red
				Write-host "===================================================="  -ForegroundColor Red
				Write-Host "Please review the output above" -ForegroundColor Red
				Write-host "After reviewing, please make another selection below"  -ForegroundColor Red
				Write-host "===================================================="  -ForegroundColor Red
				Write-host "===================================================="  -ForegroundColor Red
				Write-host "===================================================="  -ForegroundColor Red
				ShowEDiscoverySearchMenu
			}
			
			'2'{
				Do {
					Write-Host "WARNING: Executing an In-Place eDiscovery Search created by SADPhishes is EXPERIMENTAL!" -ForegroundColor Red
					Write-Host "WARNING: Previous versions of SADPhishes were knocking the Mailbox Server offline due to" -ForegroundColor Red
					Write-Host "WARNING: a search property that worked for Compliance Searches but not eDiscovery Searches." -ForegroundColor Red
					Write-Host "WARNING: While the error has not been encountered in testing this version, you may not" -ForegroundColor Red
					Write-Host "WARNING: want to run it in Production." -ForegroundColor Red
					Write-Host "You have been warned." -ForegroundColor Red
					$script:DangerousEDiscoverySearch = Read-Host -Prompt 'After reading the warning above, would you like to proceed with executing the search? [Y]es or [Q]uit'
					switch ($script:DangerousEDiscoverySearch){
						'Y'{
							Write-Host "This might blow up.  You're on your own to clean up the mess." -ForegroundColor Red
							Write-Host "==========================================================================="
							Write-Host "Starting the new In-Place eDiscovery Search with the name..."
							Write-Host "$script:ThisEDiscoverySearchName" -ForegroundColor Yellow
							Write-Host "...that will search against these mailboxes..."
							Write-Host $script:mailboxes -ForegroundColor Yellow
							Write-Host "...using the Search Query..."
							Write-Host $script:ContentMatchQuery -ForegroundColor Yellow
							Write-Host "==========================================================================="
							Write-Host "Please wait for the In-Place eDiscovery Search to complete..." -ForegroundColor Yellow
							Write-Host "==========================================================================="
							Start-MailboxSearch -Identity $script:ThisEDiscoverySearchName
								do{
								$script:ThisEDiscoverySearch = Get-MailboxSearch $script:ThisEDiscoverySearchName
								Start-Sleep 2
								Write-Host $script:ThisEDiscoverySearch.Status
								}
								until ($script:ThisEDiscoverySearch.Status -match "EstimateSucceeded")
							$script:ThisEDiscoverySearchPreviewURL = $script:ThisEDiscoverySearch.PreviewResultsLink
							Write-Host "==========================================================================="
							Write-Host "The In-Place eDiscovery Search has completed."
							Write-Host "You can use this URL to Preview the Results..." 
							Write-Host $script:ThisEDiscoverySearchPreviewURL -ForegroundColor Yellow
							Write-host "---------------------------------------------------------------------------"
							Write-Host "If you need to Copy those results to a Discovery Mailbox, or Export them"
							Write-Host "to a PST file, please use Exchange Administrative Center's Compliance "
							Write-Host "Management In-Place eDiscovery workflow to proceed with those actions."
							Write-Host "You might want to do that to confirm if any of the items in this search are" -ForegroundColor Yellow
							Write-Host "located in the Deleted Recoverable Items folder.  The PST export generates a " -ForegroundColor Yellow
							Write-Host "CSV file with that information." -ForegroundColor Yellow
							Write-Host "==========================================================================="
							Do {
								$Script:LaunchEDiscoveryURL = Read-Host -Prompt "Would you like to open the Results URL using your default browser? [Y]es or [N]o"
								switch ($Script:LaunchEDiscoveryURL){
									'Y'{
										Write-Host "Launching the link in your browser..."
										Start-Sleep 1
										Write-Host "Please return to this SADPhishes session when you are done viewing the results."
										Start-Sleep 2
										Start-Process -FilePath $script:ThisEDiscoverySearchPreviewURL
										Read-Host -Prompt "After you finish viewing the search results, please press Enter to continue."
										ReturnToComplianceSearchActionsAfterEDiscoveryDone
									}
									'N'{
										ReturnToComplianceSearchActionsAfterEDiscoveryDone
									}
								}
								
							}
							Until ($Script:LaunchEDiscoveryURL -eq 'N')
						}
						'q'{
							Write-Host "Proceeding to return to the Compliance Search Actions Menu..."
							Do{
								Write-Host "==========================================================================="
								Write-Host "Do you want to Remove the new In-Place eDiscovery Search with the name..."
								Write-Host "$script:ThisEDiscoverySearchName" -ForegroundColor Yellow
								Write-Host "...or do you want to leave it in place?"
								Write-Host "[1] Delete the eDiscovery Search and return to the Compliance Search Actions Menu."
								Write-Host "[2] Return to the Compliance Search Actions Menu without deleting."
								$script:DangerousEDiscoverySearchQuitChoice = Read-Host -Prompt 'Please enter a selection from the menu (1 or 2) and press Enter.'
								switch ($script:DangerousEDiscoverySearchQuitChoice){
									'1'{
										Remove-MailboxSearch -Identity $script:ThisEDiscoverySearchName
										Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
										Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
										#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
										if ($script:SearchType -match "8"){
											ShowNoDeleteMenu
										}
										#If the search was any other type, show the regular Actions menu that allows Delete.
										ShowMenu
									}
									'2'{
										#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
										if ($script:SearchType -match "8"){
											ShowNoDeleteMenu
										}
										#If the search was any other type, show the regular Actions menu that allows Delete.
										ShowMenu
									}
								}
							}
							Until ($script:DangerousEDiscoverySearchQuitChoice -eq '1')
						}
					}
				}
				until ($script:DangerousEDiscoverySearch -eq 'q')
			}
			
			'3'{
				Remove-MailboxSearch -Identity $script:ThisEDiscoverySearchName
				Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
				Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
				#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
				if ($script:SearchType -match "8"){
					ShowNoDeleteMenu
				}
				#If the search was any other type, show the regular Actions menu that allows Delete.
				ShowMenu
			}
			'4'{
				Write-Host "The eDiscovery Search has not been deleted. Returning to the Search Options Menu." -ForegroundColor Red
				ClearSADPhishesVars
				SearchTypeMenu
			}
			
			'q'{
				Remove-MailboxSearch -Identity $script:ThisEDiscoverySearchName
				Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
				Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
				#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
				if ($script:SearchType -match "8"){
					ShowNoDeleteMenu
				}
				#If the search was any other type, show the regular Actions menu that allows Delete.
				ShowMenu	
			}
		}
	}
	Until ($script:EDiscoverySearchMenuChoice -eq 'q')
}

#Function to return to the Compliance Search Actions menu after an eDiscovery Search has been completed.
Function ReturnToComplianceSearchActionsAfterEDiscoveryDone {
	Write-Host "SADPhishes will now return to the Compliance Search Actions menu where you" -ForegroundColor Yellow
	Write-Host "will have the option to delete all of the emails with Search Hits." -ForegroundColor Yellow
	Write-Host "The eDiscovery Searches that were created during this session are not being" -ForegroundColor Yellow
	Write-Host "deleted." -ForegroundColor Yellow
	Read-Host "Please review all of the information above then press Enter to return to the Compliance Search Actions Menu."
	#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
	if ($script:SearchType -match "8"){
		ShowNoDeleteMenu
	}
	#If the search was any other type, show the regular Actions menu that allows Delete.
	ShowMenu
}

#Function to select a FileName to Open using a dialog box
Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}


#Function to Create all SADPhishes Vars and set to Null
Function CreateSADPhishesNullVars {
	$script:AddDescription = $null
	$script:AttachmentName = $null
	$script:AttachmentNameSelection = $null
	$Script:ComplianceSearch = $null
	$Script:ComplianceSearches = $null
	$Script:ComplianceSearchNumberSelection = $null
	$script:ContentMatchQuery = $null
	$script:DangerousEDiscoverySearch = $null
	$script:DangerousEDiscoverySearchQuitChoice = $null
	$script:DangerousSearch = $null
	$script:DateEnd = $null
	$Script:DateFromHeader = $null
	$Script:DateFromHeader = $null
	$Script:DateFromHeaderDayOfMonth = $null
	$Script:DateFromHeaderDayOfWeek = $null
	$Script:DateFromHeaderFormatted = $null
	$Script:DateFromHeaderMonth = $null
	$Script:DateFromHeaderMonthNum = $null
	$Script:DateFromHeaderYear = $null
	$Script:DateHeaderMatches = $null
	$script:DateRange = $null
	$script:DateRangeSeparator = $null
	$script:DateStart = $null
	$script:EDiscoverySearch = $null
	$script:EDiscoverySearches = $null
	$script:EDiscoverySearchMenuChoice = $null
	$script:EDiscoverySearchName = $null
	$script:EDiscoverySearchName = $null
	$Script:EDiscoverySearchNumberSelection = $null
	$script:EmailHeadersFile = $null
	$script:EmailHeadersLine = $null
	$script:EmailHeadersLines = $null
	$script:ExchangeLocation = $null
	$script:ExchangeSearchLocation = $null
	$Script:FromHeaderMatches = $null
	$script:LaunchEDiscoveryURL = $null
	$script:mailboxes = $null
	$script:MailboxesWithHitsCount = $null
	$script:MailboxSearch = $null
	$script:MailboxSearches = $null
	$script:MenuChoice = $null
	$script:NoDeleteMenuChoice = $null
	$script:PurgeName = $null
	$script:PurgeSuffix = $null
	$script:SearchDescription = $null
	$script:SearchName = $null
	$script:SearchType = $null
	$Script:SelectedComplianceSearch = $null
	$Script:SelectedEDiscoverySearch = $null
	$script:Sender = $null
	$script:Subject = $null
	$Script:SubjectHeaderMatches = $null
	$Script:ThisComplianceSearchRun = $null
	$script:ThisEDiscoverySearch = $null
	$script:ThisEDiscoverySearchName = $null
	$script:ThisEDiscoverySearchPreviewURL = $null
	$script:ThisEDiscoverySearchRun = $null
	$script:ThisPurge = $null
	$script:ThisSearch = $null
	$script:ThisSearchResults = $null
	$script:ThisSearchResultsLine = $null
	$script:ThisSearchResultsLines = $null
	$script:TimeStamp = $null
	$Script:UseDateFromHeaderFile = $null
	$Script:UserSetSearchNameChoice = $null
	$Script:UseSenderFromHeaderFile = $null
	$Script:UseSubjectFromHeaderFile = $null
}

#Function to clear all of the Vars set by SADPhishes
Function ClearSADPhishesVars {
	Clear-Variable -Name AddDescription -Scope Script
	Clear-Variable -Name AttachmentName -Scope Script
	Clear-Variable -Name AttachmentNameSelection -Scope Script
	Clear-Variable -Name ComplianceSearch -Scope Script
	Clear-Variable -Name ComplianceSearches -Scope Script
	Clear-Variable -Name ComplianceSearchNumberSelection -scope Script
	Clear-Variable -Name ContentMatchQuery -Scope Script
	Clear-Variable -Name DangerousEDiscoverySearch -Scope Script
	Clear-Variable -Name DangerousEDiscoverySearchQuitChoice -Scope Script
	Clear-Variable -Name DangerousSearch -Scope Script
	Clear-Variable -Name DateEnd -Scope Script
	Clear-Variable -Name DateFromHeader -Scope Script
	Clear-Variable -Name DateFromHeader -Scope Script
	Clear-Variable -Name DateFromHeaderDayOfMonth -Scope Script
	Clear-Variable -Name DateFromHeaderDayOfWeek -Scope Script
	Clear-Variable -Name DateFromHeaderFormatted -Scope Script
	Clear-Variable -Name DateFromHeaderMonth -Scope Script
	Clear-Variable -Name DateFromHeaderMonthNum -Scope Script
	Clear-Variable -Name DateFromHeaderYear -Scope Script
	Clear-Variable -Name DateHeaderMatches -Scope Script
	Clear-Variable -Name DateRange -Scope Script
	Clear-Variable -Name DateRangeSeparator -Scope Script
	Clear-Variable -Name DateStart -Scope Script
	Clear-Variable -Name EDiscoverySearch -Scope Script
	Clear-Variable -Name EDiscoverySearches -Scope Script
	Clear-Variable -Name EDiscoverySearchMenuChoice -Scope Script
	Clear-Variable -Name EDiscoverySearchName -Scope Script
	Clear-Variable -Name EDiscoverySearchName -Scope Script 
	Clear-Variable -Name EDiscoverySearchNumberSelection -Scope Script
	Clear-Variable -Name EmailHeadersFile -Scope Script
	Clear-Variable -Name EmailHeadersLine -Scope Script
	Clear-Variable -Name EmailHeadersLines -Scope Script
	Clear-Variable -Name ExchangeLocation -Scope Script
	Clear-Variable -Name ExchangeSearchLocation -Scope Script
	Clear-Variable -Name FromHeaderMatches -Scope Script
	Clear-Variable -Name LaunchEDiscoveryURL -Scope Script
	Clear-Variable -Name mailboxes -Scope Script
	Clear-Variable -Name MailboxesWithHitsCount -Scope Script
	Clear-Variable -Name MailboxSearch -Scope Script
	Clear-Variable -Name MailboxSearches -Scope Script
	Clear-Variable -Name MenuChoice -Scope Script
	Clear-Variable -Name NoDeleteMenuChoice -Scope Script
	Clear-Variable -Name PurgeName -Scope Script
	Clear-Variable -Name PurgeSuffix -Scope Script
	Clear-Variable -Name SearchDescription -Scope Script
	Clear-Variable -Name SearchName -Scope Script
	Clear-Variable -Name SearchType -Scope Script
	Clear-Variable -Name SelectedComplianceSearch -Scope Script
	Clear-Variable -Name SelectedEDiscoverySearch -Scope Script
	Clear-Variable -Name Sender -Scope Script
	Clear-Variable -Name Subject -Scope Script
	Clear-Variable -Name SubjectHeaderMatches -Scope Script
	Clear-Variable -Name ThisComplianceSearchRun -Scope Script
	Clear-Variable -Name ThisEDiscoverySearch -Scope Script
	Clear-Variable -Name ThisEDiscoverySearchName -Scope Script
	Clear-Variable -Name ThisEDiscoverySearchPreviewURL -Scope Script
	Clear-Variable -Name ThisEDiscoverySearchRun -Scope Script
	Clear-Variable -Name ThisPurge -Scope Script
	Clear-Variable -Name ThisSearch -Scope Script
	Clear-Variable -Name ThisSearchResults -Scope Script
	Clear-Variable -Name ThisSearchResultsLine -Scope Script
	Clear-Variable -Name ThisSearchResultsLines -Scope Script
	Clear-Variable -Name TimeStamp -Scope Script
	Clear-Variable -Name UseDateFromHeaderFile -Scope Script
	Clear-Variable -Name UserSetSearchNameChoice -Scope Script
	Clear-Variable -Name UseSenderFromHeaderFile -Scope Script
	Clear-Variable -Name UseSubjectFromHeaderFile -Scope Script
}

#Function to print all SADPhishes Vars
Function PrintSADPhishesVars {
	Write-Host AddDescription [$script:AddDescription]
	Write-Host AttachmentName [$script:AttachmentName]
	Write-Host AttachmentNameSelection [$script:AttachmentNameSelection]
	Write-Host ComplianceSearch [$script:ComplianceSearch]
	Write-Host ComplianceSearches [$script:ComplianceSearches]
	Write-Host ComplianceSearchNumberSelection [$script:ComplianceSearchNumberSelection]
	Write-Host ContentMatchQuery [$script:ContentMatchQuery]
	Write-Host DangerousEDiscoverySearch [$script:DangerousEDiscoverySearch]
	Write-Host DangerousEDiscoverySearchQuitChoice [$script:DangerousEDiscoverySearchQuitChoice]
	Write-Host DangerousSearch [$script:DangerousSearch]
	Write-Host DateEnd [$script:DateEnd]
	Write-Host DateFromHeader [$script:DateFromHeader]
	Write-Host DateFromHeader [$script:DateFromHeader]
	Write-Host DateFromHeaderDayOfMonth [$script:DateFromHeaderDayOfMonth]
	Write-Host DateFromHeaderDayOfWeek [$script:DateFromHeaderDayOfWeek]
	Write-Host DateFromHeaderFormatted [$Script:DateFromHeaderFormatted]
	Write-Host DateFromHeaderMonth [$script:DateFromHeaderMonth]
	Write-Host DateFromHeaderMonthNum [$Script:DateFromHeaderMonthNum]
	Write-Host DateFromHeaderYear [$script:DateFromHeaderYear]
	Write-Host DateHeaderMatches [$script:DateHeaderMatches]
	Write-Host DateRange [$script:DateRange]
	Write-Host DateRangeSeparator [$script:DateRangeSeparator]
	Write-Host DateStart [$script:DateStart]
	Write-Host EDiscoverySearch [$script:EDiscoverySearch]
	Write-Host EDiscoverySearches [$script:EDiscoverySearches]
	Write-Host EDiscoverySearchMenuChoice [$script:EDiscoverySearchMenuChoice]
	Write-Host EDiscoverySearchName [$script:EDiscoverySearchName]
	Write-Host EDiscoverySearchName [$script:EDiscoverySearchName]
	Write-Host EDiscoverySearchNumberSelection [$Script:EDiscoverySearchNumberSelection]
	Write-Host EmailHeadersFile [$script:EmailHeadersFile]
	Write-Host EmailHeadersLine [$script:EmailHeadersLine]
	Write-Host EmailHeadersLines [$script:EmailHeadersLines]
	Write-Host ExchangeLocation [$script:ExchangeLocation]
	Write-Host ExchangeSearchLocation [$script:ExchangeSearchLocation]
	Write-Host FromHeaderMatches [$script:FromHeaderMatches]
	Write-Host LaunchEDiscoveryURL [$script:LaunchEDiscoveryURL]
	Write-Host mailboxes [$script:mailboxes]
	Write-Host MailboxesWithHitsCount [$script:MailboxesWithHitsCount]
	Write-Host MailboxSearch [$script:MailboxSearch]
	Write-Host MailboxSearches [$script:MailboxSearches]
	Write-Host MenuChoice [$script:MenuChoice]
	Write-Host NoDeleteMenuChoice [$script:NoDeleteMenuChoice]
	Write-Host PurgeName [$script:PurgeName]
	Write-Host PurgeSuffix [$script:PurgeSuffix]
	Write-Host SearchDescription [$script:SearchDescription]
	Write-Host SearchName [$script:SearchName]
	Write-Host SearchType [$script:SearchType]
	Write-Host SelectedComplianceSearch [$script:SelectedComplianceSearch]
	Write-Host SelectedEDiscoverySearch [$Script:SelectedEDiscoverySearch]
	Write-Host Sender [$script:Sender]
	Write-Host Subject [$script:Subject]
	Write-Host SubjectHeaderMatches [$script:SubjectHeaderMatches]
	Write-Host ThisComplianceSearchRun [$Script:ThisComplianceSearchRun]
	Write-Host ThisEDiscoverySearch [$script:ThisEDiscoverySearch]
	Write-Host ThisEDiscoverySearchName [$script:ThisEDiscoverySearchName]
	Write-Host ThisEDiscoverySearchPreviewURL [$script:ThisEDiscoverySearchPreviewURL]
	Write-Host ThisEDiscoverySearchRun [$script:ThisEDiscoverySearchRun]
	Write-Host ThisPurge [$script:ThisPurge]
	Write-Host ThisSearch [$script:ThisSearch]
	Write-Host ThisSearchResults [$script:ThisSearchResults]
	Write-Host ThisSearchResultsLine [$script:ThisSearchResultsLine]
	Write-Host ThisSearchResultsLines [$script:ThisSearchResultsLines]
	Write-Host TimeStamp [$script:TimeStamp]
	Write-Host UseDateFromHeaderFile [$Script:UseDateFromHeaderFile]
	Write-Host UserSetSearchNameChoice [$Script:UserSetSearchNameChoice]
	Write-Host UseSenderFromHeaderFile [$script:UseSenderFromHeaderFile]
	Write-Host UseSubjectFromHeaderFile [$script:UseSubjectFromHeaderFile]
}

#Drop the user into the DisplayBanner function (and then Search Type Menu) to begin the process.
DisplayBanner