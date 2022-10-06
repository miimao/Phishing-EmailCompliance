function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        Write-Host("Module $m is already imported.") -F Green
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            Write-Host("Module $m Has been imported.") -F Green
            Import-Module $m #-Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                Write-Host("Module $m Was not found. Attempting Install") -F Red
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m #-Verbose
            }
            else {

                # If the module is not imported, not available and not in the online gallery then abort
                Write-Host("Module $m not imported, not available and not in an online gallery, exiting.")
                EXIT 1
            }
        }
    }
}


#Check Execution policy and changed to remote signed if not already
Write-Host("Checking Execution Policy...")
$execPolicy = Get-ExecutionPolicy
if ($execPolicy -ne "UnRestricted"){
    Write-Host("Execution Policy is Not UnRestricted, Attempting to set correct policy...") -F Red
    Set-ExecutionPolicy UnRestricted
    Write-Host("OK") -F Green
}
else{
    Write-Host("Execution Policy is 'UnRestricted'") -F Green
}

#Check for "ExchangeOnlineManagement PS Module"
Write-Host("Checking for 'ExchangeOnlineManagement' PowerShell Module...")
Load-Module("ExchangeOnlineManagement")

#Login the O365 Exchange module
$AdminUser = Read-Host "Enter The admin account of the O365 environment.`n"
Connect-IPPSSession -UserPrincipalName $AdminUser

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
	Write-Host " "
	Write-host "------- Debugging Options -------"
	Write-Host "    [X] gci variable:"
	Write-Host "    [Y] Print Vars"
	Write-Host "    [Z] Clear Vars"
	Write-Host " "
	Write-Host "------------- Quit --------------"
	Write-Host "    [Q] Quit"
	Write-Host "---------------------------------"
}	
	
#Function for Search Type Menu
Function SearchTypeMenu{
	Do {	
		SearchTypeOptions
		CreateNullVars
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
							ClearVars
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
				$script:AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. Sketchy.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'8'{
				Write-Host "You have chosen to conduct the Pre-Built Suspicious Attachment Types Search." -ForegroundColor Yellow
				Write-Host "This search will return a list of Mailboxes that contain Attachments with specific file extensions." -ForegroundColor Yellow
				Write-Host "This search is a Search-Only option, with no Delete built into the Workflow." -ForegroundColor Yellow
				Write-Host "Take these results and investigate." -ForegroundColor Yellow
				Read-Host -Prompt "After you have read the information about this Suspicious Attachment Search, Press Enter to continue."
				$script:ContentMatchQuery = "((Attachment:'.ade') OR (Attachment:'.adp') OR (Attachment:'.apk') OR (Attachment:'.bas') OR (Attachment:'.bat') OR (Attachment:'.chm') OR (Attachment:'.cmd') OR (Attachment:'.com') OR (Attachment:'.cpl') OR (Attachment:'.dll') OR (Attachment:'.exe') OR (Attachment:'.hta') OR (Attachment:'.inf') OR (Attachment:'.iqy') OR (Attachment:'.jar') OR (Attachment:'.js') OR (Attachment:'.jse') OR (Attachment:'.lnk') OR (Attachment:'.mht') OR (Attachment:'.msc') OR (Attachment:'.msi') OR (Attachment:'.msp') OR (Attachment:'.mst') OR (Attachment:'.ocx') OR (Attachment:'.pif') OR (Attachment:'.pl') OR (Attachment:'.ps1') OR (Attachment:'.reg') OR (Attachment:'.scr') OR (Attachment:'.sct') OR (Attachment:'.shs') OR (Attachment:'.slk') OR (Attachment:'.sys') OR (Attachment:'.vb') OR (Attachment:'.vbe') OR (Attachment:'.vbs') OR (Attachment:'.wsc') OR (Attachment:'.wsf') OR (Attachment:'.wsh'))"
				ExchangeSearchLocationMenu
			}
			'9'{
				Write-Host "You have chosen to have open a Text file containing the Headers" -ForegroundColor Yellow
				Write-Host "from a sample EMail.  Please select the text file to open in the dialog box" -ForegroundColor Yellow
				Write-Host "that will open when you proceed." -ForegroundColor Yellow
				Read-Host -Prompt "After you have read the information above, Press Enter to proceed."
				ParseEmailHeadersFile
			}
			'10'{
				RunPreviousComplianceSearch
			}
			'q'{
				Write-Host "All Done!" -ForegroundColor Yellow
				Exit
			}
			'x'{
			Get-ChildItem variable:
			}
			'y'{
			PrintVars
			}
			'z'{
			ClearVars
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
			Write-Host "Found this Sender Address..." -ForegroundColor Yellow
			Write-Host $script:Sender
		}
        
		If ($script:SubjectHeaderMatches){
			$Script:Subject = $matches[1]
			Write-Host "Found this Subject..." -ForegroundColor Yellow
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
				Write-Host "Decoded plain text Subject for Search..."
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
		Write-Host "Found this Date in the Headers..." -ForegroundColor Yellow
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
				ClearVars
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
				ClearVars
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
				ClearVars
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
				$script:AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. Sketchy.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'q'{
				ClearVars
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
				$script:ExchangeLocation = "All"
				UserSetSeachNameMenu
			}
			'2'{
				$script:ExchangeLocation = Read-Host -Prompt 'Please enter the EMail Address of the MailBox or Group you would like to search within'
				UserSetSeachNameMenu
			}
			'q'{
				ClearVars
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
	Write-Host "If you don't need to specify your own name, This script will automatically create a name based on the search criteria you have specified." -ForegroundColor Yellow
	Write-Host "If you aren't sure what to choose, pick No so you can see how the script builds Search Names." -ForegroundColor Yellow
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
				ClearVars
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
				ClearVars
				SearchTypeMenu			
			}
		}
	}
	until ($Script:AddDescription -eq 'q')
}



#Function to Re-Run a previous Compliance Search
Function RunPreviousComplianceSearch {
	Write-Host "Listing all of the existing Compliance Searches" -ForegroundColor Yellow
	Write-Host "They will be in the the format '[#] SearchNameHere', where # is the integer you will use to select which search to run." -ForegroundColor Yellow
	Read-Host -Prompt "After reading the information above, Please press Enter to continue."
	#create an empty array for existing ComplianceSearches
	$script:ComplianceSearches = @()
	#set up an Integer to use for tagging each existing Compliance Search with a number
	$I = 1
	$script:ComplianceSearches = Get-ComplianceSearch
	#For every Compliance Search found, add a NoteProperty named Search number, assign it with our integer, and then increase the Integer by 1 so it's ready for the next Compliance Search in the array.
	$Script:ComplianceSearches | ForEach-Object{$_ | Add-Member -NotePropertyName SearchNumber -NotePropertyValue $I -Force; $I++}
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
        $Script:ComplianceSearchNumberSelectionInt = [int]$Script:ComplianceSearchNumberSelection
	}
    
	Until ($Script:ComplianceSearchNumberSelectionInt -ge 1 -and $Script:ComplianceSearchNumberSelectionInt -le $I)
	#set up variables so our ComplianceSearch Function will run
	$Script:SelectedComplianceSearch = $Script:ComplianceSearches | Where-Object {$_.SearchNumber -eq $Script:ComplianceSearchNumberSelection}
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
						$script:SearchName = "Pre-Built Suspicious Attachment Types Search Exchange Location [$script:ExchangeLocation]"
					}
					'9'{
						$script:SearchName = "Headers Parsed Subject [$script:Subject] DateRange [$script:DateRange] Sender [$script:Sender] ExchangeLocation [$script:ExchangeLocation] Phishing Message"
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
				$script:ThisComplianceSearchRun = "$script:SearchName-$I"
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
		$Script:SearchName = "$Script:SearchName-$I"
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
				Write-Host "Please supply a new Search Name so that you can proceed."
				$Script:SearchName = Read-Host -Prompt "After reading the information above, please enter a new Search Name that is less than 198 Characters."
				$I = 1;
				$Script:ComplianceSearches = Get-ComplianceSearch;
					while ($true){
						$found = $false
						$script:ThisComplianceSearchRun = "$script:SearchName-$I"
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
				$Script:SearchName = "$Script:SearchName-$I"
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
			Write-Host "(Subject:Sketchy Email), your results will include messages with the subject"   -ForegroundColor Yellow
			Write-Host "'Sketchy Email', but also messages with the subjects 'Sketchy Emails is good!' and" -ForegroundColor Yellow
			Write-Host "'RE: Screw Sketchy Emails. it sucks!'"                                           -ForegroundColor Yellow
			Write-Host " "                                                                           -ForegroundColor Yellow
			Write-Host "This is just how the Microsoft Exchange Content Search works."               -ForegroundColor Yellow
			Write-Host " "                                                                           -ForegroundColor Yellow
			Write-Host "Please take this into consideration when using the Search Results."          -ForegroundColor Yellow
			Write-Host "===========================================================================" -ForegroundColor Yellow
			Read-Host -Prompt "Please press Enter after reading the warning above."
		
		}
		switch ($script:AddDescription) {
			'1' {
			New-ComplianceSearch -Name "$script:SearchName"  -ContentMatchQuery $script:ContentMatchQuery -ExchangeLocation $script:ExchangeLocation
			}
			'2' {
			New-ComplianceSearch -Name "$script:SearchName"  -ContentMatchQuery $script:ContentMatchQuery -Description "$Script:SearchDescription" -ExchangeLocation $script:ExchangeLocation
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
	Write-Host "[3] Delete this search and Return to the Search Options Menu."
	Write-Host "[4] Return to the Search Options Menu."
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
							New-ComplianceSearchAction -SearchName "$script:SearchName" -Purge -PurgeType SoftDelete
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
							Write-Host "Purged this many Items..."
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
									Write-Host "Did not find any items to delete!" -ForegroundColor Red
									Write-Host "Did not find any items to delete!" -ForegroundColor Red
									Write-Host "Did not find any items to delete!" -ForegroundColor Red
									Write-Host "The initial Compliance Search returned this many items...  "
									Write-Host $script:ThisSearch.Items Items -ForegroundColor Yellow
									Write-Host "...but the Delete/Purge occurred on this many items..."
									Write-Host $Script:ThisPurgeResultsItemCount -ForegroundColor Yellow
									Write-Host "That should be an indication that all of the Items returned by the Compliance Search are already located in the Deleted Recoverable Items folder of each Mailbox!" -ForegroundColor Yellow
									Write-Host "==========================================================="
									Write-Host "You can use the In-Place eDiscovery Search (option 3 presented in the Compliance Search Actions Menu, after the initial search is run) to confirm if that is true."
									Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
									ClearVars
									SearchTypeMenu
								}					
							Write-host "==================================================================================="
							Write-Host "Note: Microsoft's Compliance Search Purge Actions will remove a maximum of 10" -ForegroundColor Yellow
							Write-Host "items per mailbox at one time.  They say it's designed that way because it's" -ForegroundColor Yellow
							Write-Host "an Incident Response Tool and the limit helps ensure that messages are quickly" -ForegroundColor Yellow
							Write-Host "removed." -ForegroundColor Yellow
							Write-host "==================================================================================="
							Write-host "If you think this Purge may have left items behind, you should run another Search" -ForegroundColor Yellow
							Write-host "and Delete/Purge until the Item count displayed above is 0." -ForegroundColor Yellow
							Write-Host "The current Purge is complete." -ForegroundColor Red
							Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
							ClearVars
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
				Remove-ComplianceSearch -Identity $script:SearchName
				Write-Host "The search has been deleted." -ForegroundColor Red
				Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
				ClearVars
				SearchTypeMenu
			}
			'4'{
				Write-Host "The previous Compliance Search has not been deleted. Returning to the Search Options Menu" -ForegroundColor Red
				ClearVars
				SearchTypeMenu
			}
			
			'q'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearVars
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
	Write-Host "[2] Delete this search and Return to the Search Options Menu."
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
				Remove-ComplianceSearch -Identity $script:SearchName
				Write-Host "The search has been deleted." -ForegroundColor Red
				Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
				ClearVars
				SearchTypeMenu
			}

			# '3'{
			# CreateEDiscoverySearch
			# }
			
			'q'{
			Remove-ComplianceSearch -Identity $script:SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearVars
			SearchTypeMenu
			}
		}
	}
	Until ($script:MenuChoice -eq 'q')
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


#Function to Create all Vars and set to Null
Function CreateNullVars {
	$script:AddDescription = $null
	$script:AttachmentName = $null
	$script:AttachmentNameSelection = $null
	$Script:ComplianceSearch = $null
	$Script:ComplianceSearches = $null
	$Script:ComplianceSearchNumberSelection = $null
	$Script:ComplianceSearchNumberSelectionInt = $null
	$script:ContentMatchQuery = $null
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
	$script:EmailHeadersFile = $null
	$script:EmailHeadersLine = $null
	$script:EmailHeadersLines = $null
	$script:ExchangeLocation = $null
	$script:ExchangeSearchLocation = $null
	$Script:FromHeaderMatches = $null
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
	$script:Sender = $null
	$script:Subject = $null
	$Script:SubjectHeaderMatches = $null
	$Script:ThisComplianceSearchRun = $null
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

#Function to clear all of the Vars
Function ClearVars {
	Clear-Variable -Name AddDescription -Scope Script
	Clear-Variable -Name AttachmentName -Scope Script
	Clear-Variable -Name AttachmentNameSelection -Scope Script
	Clear-Variable -Name ComplianceSearch -Scope Script
	Clear-Variable -Name ComplianceSearches -Scope Script
	Clear-Variable -Name ComplianceSearchNumberSelection -scope Script
	Clear-Variable -Name ComplianceSearchNumberSelectionInt -scope Script
	Clear-Variable -Name ContentMatchQuery -Scope Script
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
	Clear-Variable -Name EmailHeadersFile -Scope Script
	Clear-Variable -Name EmailHeadersLine -Scope Script
	Clear-Variable -Name EmailHeadersLines -Scope Script
	Clear-Variable -Name ExchangeLocation -Scope Script
	Clear-Variable -Name ExchangeSearchLocation -Scope Script
	Clear-Variable -Name FromHeaderMatches -Scope Script
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
	Clear-Variable -Name Sender -Scope Script
	Clear-Variable -Name Subject -Scope Script
	Clear-Variable -Name SubjectHeaderMatches -Scope Script
	Clear-Variable -Name ThisComplianceSearchRun -Scope Script
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

#Function to print all Vars
Function PrintVars {
	Write-Host AddDescription [$script:AddDescription]
	Write-Host AttachmentName [$script:AttachmentName]
	Write-Host AttachmentNameSelection [$script:AttachmentNameSelection]
	Write-Host ComplianceSearch [$script:ComplianceSearch]
	Write-Host ComplianceSearches [$script:ComplianceSearches]
	Write-Host ComplianceSearchNumberSelection [$script:ComplianceSearchNumberSelection]
	Write-Host ComplianceSearchNumberSelectionInt [$Script:ComplianceSearchNumberSelectionInt]
	Write-Host ContentMatchQuery [$script:ContentMatchQuery]
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
	Write-Host EmailHeadersFile [$script:EmailHeadersFile]
	Write-Host EmailHeadersLine [$script:EmailHeadersLine]
	Write-Host EmailHeadersLines [$script:EmailHeadersLines]
	Write-Host ExchangeLocation [$script:ExchangeLocation]
	Write-Host ExchangeSearchLocation [$script:ExchangeSearchLocation]
	Write-Host FromHeaderMatches [$script:FromHeaderMatches]
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
	Write-Host Sender [$script:Sender]
	Write-Host Subject [$script:Subject]
	Write-Host SubjectHeaderMatches [$script:SubjectHeaderMatches]
	Write-Host ThisComplianceSearchRun [$Script:ThisComplianceSearchRun]
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

SearchTypeMenu