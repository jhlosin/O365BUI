######################################
#                                    #
# Connection to Office 365   # 
#                                    #
######################################

function Connect-Office365 {
    <#
	.SYNOPSIS
        Connect Office 365 MSOL service and Exchange Online.
    #>
    $LiveCred = get-credential
    $global:Session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $LiveCred -Authentication basic -AllowRedirection
    Import-PSSession $global:Session365
    connect-MsolService -credential $LiveCred
}

function Disconnect-Office365 {
    <#
	.SYNOPSIS
        Remove connected Office 365 MSOL service and Exchange Online sessions.
    #>
    Get-PSSession | Remove-PSSession -Confirm:$false
    Remove-Module MSOnline
    Import-Module MSOnline    
}

######################################
#                                    #
#       MSOL Services Util           # 
#                                    #
######################################

function BulkNewOrUpdate-MsolUser {
    <#
	.SYNOPSIS
        It will create Windows Azure MsolUsers in bulk and modify attritbues to the ones in the CSV if already exists.		
	.EXAMPLE
		PS> BulkNew-MsolUser -CsvLocation 'C:\AllUsers.csv'
        ===List of Properties(column) in CSV===
        FirstName(required)
        LastName(required)
        UserPrincipalName(PrimarySmtpAddress)(required)
        Fax
        Department
        MobilePhone
        PhoneNumber
        StreetAddress
        City
        Title
        State
        PostalCode
        Country
        UsageLocation
        Manager 
        Company
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)][string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $users = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
        $currentTime = Get-CurrentDateString
        
        # select option for assigning a license
        while(1) {
            $AssignLicense = Read-Host 'do you want to assign a license for users?(y/n)'
            if ($AssignLicense -eq 'y' -or $AssignLicense -eq 'n') {
                break
            } else {
                Write-Warning "Only y or n is allowed. try again."
            }
        }
        # if you chose y
        if ($AssignLicense -eq 'y') {
            #select a license
            $license = (Get-MsolAccountSku | Out-GridView -Title "Select the license to assign" -PassThru).AccountSkuId
        # else
        } else {
            $license = ''
        }
                
		# for each user in users
        foreach ($user in $users)
        {
            $UserParams = @{
                'FirstName' = $user.FirstName
                'LastName' = $user.LastName
                'UserPrincipalName' = $user.UserPrincipalName
                'fax' = $user.fax
                'Department' = $user.Department
                'MobilePhone' = $user.MobilePhone
                'PhoneNumber' = $user.PhoneNumber
                'DisplayName' =  $user.FirstName + " " + $user.LastName
                'StreetAddress' = $user.StreetAddress
                'Title' = $user.Title
                'city' = $user.city
                'state' = $user.state
                'postalCode' = $user.postalCode
                'country' = $user.country
                'UsageLocation' = $user.UsageLocation
            }

            #check if this user already exists in MS online
            $thisUser = Get-MsolUser -UserPrincipalName $UserParams.UserPrincipalName -ErrorAction SilentlyContinue
            
            # if thisUser does not exist
            if (!$thisUser) {
                # create a new msol user
                Write-Host ("creating a new msoluser " + $UserParams.DisplayName)
                if ($license) {
                    New-MsolUser @UserParams -ErrorAction SilentlyContinue -LicenseAssignment $license
                } else {
                    New-MsolUser @UserParams -ErrorAction SilentlyContinue
                }
	        } else { # if already exist
                try {
                    # update attributes.
                    $message = ("The user " + $UserParams.DisplayName + " already exists. trying to update attritbutes")
                    Write-Log -Message $message -Path $csvFolder"\Log_BulkNewOrUpdate-MsolUser-"$currentTime".log"
                    set-MsolUser @UserParams -ErrorAction Stop
        
                    # add the license if you selected prior
                    if ($license) {
                        try {
                            Set-MsolUserLicense -UserPrincipalName $thisUser.UserPrincipalName -AddLicenses $license -ErrorAction Stop
                            $successMessage = 'A license has been assigned to' + $thisUser.UserPrincipalName
                            Write-Log -Message $successMessage -Path $csvFolder"\Log_BulkNewOrUpdate-MsolUser-"$currentTime".log"
                        } catch {
                            $message = 'user-' + $thisUser.UserPrincipalName + ' ' + $_
                            Write-Log -Message $message -Path $csvFolder"\Log_BulkNewOrUpdate-MsolUser-"$currentTime".log" -Level Error
                        }
                    }
                } catch {
                    $message = 'user ' + $thisUser.UserPrincipalName + ' ' + $_
                    Write-Log -Message $message -Path $csvFolder"\Log_BulkNewOrUpdate-MsolUser-"$currentTime".log" -Level Error
                }
            }
            
            ## handle other attributes ##
            #wait until the user is populated
                $UPN = $UserParams.UserPrincipalName
                $emailUser = Get-User $UPN -ErrorAction SilentlyContinue
                while (1) {
                    if($emailUser) {
                        $message = "The user $UPN has been created"
                        Write-Log -Message $message -Path $csvFolder"\Log_BulkNewOrUpdate-MsolUser-"$currentTime".log" -Level Info
                        break
                    } else {
                        $emailUser = ''
                        Write-Warning "waiting for the user to be created"
                        sleep 10
                        $emailUser = Get-User $UPN -ErrorAction SilentlyContinue
                    }
                }

            # max company attribute is 64, cut it if it is longer than 64
            if ($user.company.length -gt 64) {
                $user.company = $user.company.Substring(0,63)
            }

            # set up other attritbutes
            if ($user.Manager) {
                Set-User -Identity $UPN -Manager $user.Manager -Company $user.company
            } else {
                Set-User -Identity $UPN -Company $company
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkNewOrUpdate-MsolUser-"$currentTime".log" -ForegroundColor Yellow
    }
}

function BulkRemove-MsolUser {
    <#
	.SYNOPSIS
        It will forcefully remove Windows Azure MsolUsers for the user given in the csv.
	.EXAMPLE
		PS> BulkNew-MsolUser -CsvLocation 'C:\msoluserstoremove.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to remove
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)][string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        $users = import-csv $CsvLocation
        $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
        $currentTime = Get-CurrentDateString
                
		# for each user in users
        foreach ($user in $users)
        {
            $userName = $user.userPrincipalName
            Try {
                Remove-MsolUser -UserPrincipalName $user.userPrincipalName -Force:$true -ErrorAction stop
                $successMessage = "the user" + $userName + "has been removed."
                Write-Log -Message $successMessage -Path $csvFolder"\Log_BulkRemove-MsolUser-"$currentTime".log"
            } Catch {
                #Write-Warning "Error occured: $_"
                #$_ | Out-File $csvFolder"\Log_BulkRemove-MsolUser-"$nowString".txt" -Append
                Write-Log -Message "The user, $userName, does not exist" -Path $csvFolder"\Log_BulkRemove-MsolUser-"$currentTime".log" -Level error
            }
        }

        Write-host "The process has been finished. Log has been saved in the csv folder as Log_BulkRemove-MsolUser.txt." -ForegroundColor Yellow
	}
}

function BulkSet-MsolLicense {
    <#
	.SYNOPSIS
        It will assign license you select for the user given in the csv or all user mailboxes.
	.EXAMPLE
		PS> BulkSet-MsolLicense -CsvLocation 'C:\users.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
        This below will assign a license selected for all user mailboxes in the tenant.
        PS> BulkSet-MsolLicense -AllUserMailboxes
       		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to assign a license with
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$false)][string]$CsvLocation,

        [Parameter(Mandatory=$false)][bool]$AllUserMailboxes=$false
	)
	process {
        # import the user list csv to a variable users
        $license = (Get-MsolAccountSku | Out-GridView -Title "Select the license to assign" -PassThru).AccountSkuId
        $currentTime = Get-CurrentDateString

        if ($AllUserMailboxes) {
            $users = get-mailbox -ResultSize unlimited | Where-Object {$_.recipientTypeDetails -like "UserMailbox"}

        } else {
            $users = import-csv $CsvLocation
            $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
        }
		# for each user in users
        $path = "Log_BulkSet-MsolUserLicense-" + $currentTime + ".log"
        foreach ($user in $users)
        {
            $userName = $user.userPrincipalName
            $CsvLocation
            Try {
                # add the license you selected prior
                #if (!$user.UsageLocation) {
                #    set-MsolUser -UserPrincipalName $userName -UsageLocation US
                #}
                Set-MsolUserLicense -UserPrincipalName $userName -AddLicenses $license -ErrorAction SilentlyContinue

                $successMessage = "the user" + $userName + "has been assigned with a license " + $license
                Write-Log -Message $successMessage -Path $path
            } Catch {
                # if failed, log the error
                #$path = "Log_BulkRemove-MsolUser-" + $currentTime + ".log"
                Write-Log -Message $_ -Path $path -Level "error"
            }
        }

        Write-host "The process has been finished. Log has been saved in Log_BulkSet-MsolLicense.txt." -ForegroundColor Yellow
	}
}

function BulkReset-MsolPassword {
    <#
	.SYNOPSIS
        It will reset passwords for the user(s) given in the csv.
	.EXAMPLE
		PS> BulkReset-MsolPassword -CsvLocation 'C:\msolusers.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
        FirstName
        LastName
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to remove
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)][string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $users = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
        $currentTime = Get-CurrentDateString
        
        # select option for assigning a license
        while(1) {
            $setRandomPassword = Read-Host 'Do you want to set random passwords or a generic one for all users?(r for Random/g for generic)'
            $forceChangePassword = Read-Host 'Do you want users to change their password at the next login?(y/n)'
            if ( ($setRandomPassword -eq 'r' -or $setRandomPassword -eq 'g') -and ($forceChangePassword -eq 'y' -or $forceChangePassword -eq 'n') ) {
                break
            } else {
                Write-Warning "not allowed input. try again."
            }
        }
                
		# for each user in users
        foreach ($user in $users)
        {
            # if you chose r(random)
            if ($setRandomPassword -eq 'r') {
                if ($forceChangePassword -eq 'y') {
                    try {
                        $password = Set-MsolUserPassword -UserPrincipalName $user.userPrincipalName -ForceChangePassword:$true -ErrorAction Stop
                        $password | select-object `
                                        @{label='UserPrincipalName'; Expression={$user.userPrincipalName}} `
                                        , @{label='FirstName'; Expression={$user.Firstname}} `
                                        , @{label='LastName'; Expression={$user.LastName}} `
                                        , @{label='password'; Expression={$_}} `
                            | export-csv MSOLUserNewPasswords.csv -NoTypeInformation -Append
                        $message = 'password for ' + $user.UserPrincipalName +  ' has been reset.'
                        Write-Log -Message $message -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log"
                    } catch {
                        Write-Log -Message $_ -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log" -Level Error
                    }
                } else {
                    try {
                        $password = Set-MsolUserPassword -UserPrincipalName $user.userPrincipalName
                        $password | select-object `
                                        @{label='UserPrincipalName'; Expression={$user.userPrincipalName}} `
                                        , @{label='FirstName'; Expression={$user.Firstname}} `
                                        , @{label='LastName'; Expression={$user.LastName}} `
                                        , @{label='password'; Expression={$_}} `
                            | export-csv MSOLUserNewPasswords.csv -NoTypeInformation -Append
                        $message = 'password for ' + $user.UserPrincipalName +  ' has been reset.'
                        Write-Log -Message $message -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log"
                    } catch {
                        Write-Log -Message $_ -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log" -Level Error
                    }
                }
            # for g(generic)
            } else {
                $newPassword = Read-Host 'put the new password for all users. it should be combination of number and alphabet with at least one Capital and at least 8 characters length.(e.g People12)'
                if ($forceChangePassword -eq 'y') {
                    try {
                        $password = Set-MsolUserPassword -UserPrincipalName $user.userPrincipalName -ForceChangePassword:$true -NewPassword $newPassword
                        $password | select-object `
                                        @{label='UserPrincipalName'; Expression={$user.userPrincipalName}} `
                                        , @{label='FirstName'; Expression={$user.Firstname}} `
                                        , @{label='LastName'; Expression={$user.LastName}} `
                                        , @{label='password'; Expression={$_}} `
                            | export-csv MSOLUserNewPasswords.csv -NoTypeInformation -Append
                        $message = 'password for ' + $user.UserPrincipalName +  ' has been reset.'
                        Write-Log -Message $message -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log"
                    } catch {
                        Write-Log -Message $_ -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log" -Level Error
                    }
                } else {
                    try {
                        $password = Set-MsolUserPassword -UserPrincipalName $user.userPrincipalName -NewPassword $newPassword
                        $password | select-object `
                                        @{label='UserPrincipalName'; Expression={$user.userPrincipalName}} `
                                        , @{label='FirstName'; Expression={$user.Firstname}} `
                                        , @{label='LastName'; Expression={$user.LastName}} `
                                        , @{label='password'; Expression={$_}} `
                            | export-csv MSOLUserNewPasswords.csv -NoTypeInformation -Append
                        $message = 'password for ' + $user.UserPrincipalName +  ' has been reset.'
                        Write-Log -Message $message -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log"
                    } catch {
                        Write-Log -Message $_ -Path $csvFolder"\Log_BulkSet-MsolLicense-"$currentTime".log" -Level Error
                    }
                }
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkSet-MsolLicense-"$currentTime".log" -ForegroundColor Yellow
        Write-host "The password info is saved on MSOLUserNewPasswords.csv" -ForegroundColor Yellow
    }
}


function BulkEmail-UserPassword {
    <#
	.SYNOPSIS
        It will distirbute emails wit user password given in a csv file.
	.EXAMPLE
		PS> BulkEmail-UserPassword -CsvLocation 'C:\newpassword.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required) = recipient Email Address
        password(required)
        FirstName(required)
        LastName(required)
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to remove
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)][string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $users = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
        $currentTime = Get-CurrentDateString
        $getCred = Get-Credential

		# for each user in users
        foreach ($user in $users)
        {
            #send an email to the user
            $mycreds = New-Object System.Management.Automation.PSCredential ($getCred.UserName, $getCred.Password)
            $date = Get-Date
            $smtp = "smtp.office365.com" 
            $to = $user.userPrincipalName
            $from = "DaleCarnegieEmailService@metrocsg.microsoftonline.com" 
            $subject = "IMPORTANT - Your new Dale Carnegie Office 365 account"
            $firstName = $user.firstName
            $tempPassword = $user.password
            $body = @" 

**This message has been sent to all dalecarnegie.com users on behalf of your IT team!**<br><br>

Dear $firstName<br><br>

You are now able to log on to your new Dale Carnegie Office 365 email account.  We are in the process of migrating all of your mail, contacts and calendar entries, so don't expect to see all of your history just yet; but you can now log on and test your user credentials to this new platform.<br><br>

To sign on, open any web browser and go to https://login.microsoftonline.com/<br>
Your user name is <b>$to</b><br>
Your temporary password is :  <b>$tempPassword</b><br><br>

Upon first log-on you will be prompted to enter a unique password.<br><br>

If you run into any issues or have any problems please refer to our education newsletter (sent to you earlier in the week) or reach out to us for assistance.  We are here help, it's in the name  - ithelp@dalecarnegie.com <br><br>

Thank you and welcome to Dale Carnegie's Office 365!<br>

"@
            try {
                Send-MailMessage -SmtpServer $smtp -To $to -Subject $subject -Credential $mycreds -UseSsl -Port "587" -Body $body -From $from -BodyAsHtml -ErrorAction Stop
                $message = 'An email has been sent to '+ $to
                Write-Log -Message $message -Path $csvFolder"\Log_BulkEmail-UserPassword-"$currentTime".log"
            } catch {
                Write-Log -Message $_ -Path $csvFolder"\Log_BulkEmail-UserPassword-"$currentTime".log"
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkEmail-UserPassword-"$currentTime".log" -ForegroundColor Yellow
    }
}

######################################
#                                    #
#       Exchange Online Utils        #
#                                    #
######################################

function BulkNew-MailContact {
    <#
	.SYNOPSIS
        It will create mailcontact in bulk for all items in a given CSV.
	.EXAMPLE
		PS> BulkNew-MailContact -CsvLocation 'C:\conatcts.csv'
        ===List of Properties(column) in CSV===
        name(required)
        externalEmailAddress(required)
        FirstName
        Lastname
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $users = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $currentTime = Get-CurrentDateString
		
		# for each user in users
        foreach ($item in $users)
        {
            $thisUser = Get-MailContact $item.name -ErrorAction SilentlyContinue
            if (!$thisUser) {
                try {
                    New-MailContact -Name $item.name -ExternalEmailAddress $item.ExternalEmailAddress -DisplayName $item.name -FirstName $item.firstName -LastName $item.lastName -ErrorAction stop
                    $message = $item.name + " has been created."
                    Write-Host $message
                    Write-Log -Message $message -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log"
                } catch {
                    Write-Log -Message $_ -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log" -Level Error
                }
                
            } else {
                $message = "The contact " + $item.name + " already exists. skipping..."
                Write-Host $message
                Write-Log -Message $message -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log" -Level Warn
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkNew-MailContact-"$currentTime".log" -ForegroundColor Yellow
    }
}


function BulkUpdate-ProxyAddresses {
    <#
	.SYNOPSIS
        It will update proxy address for Dale Carnegie for mailboxes in a given csv.
	.EXAMPLE
		PS> BulkChange-ProxyAddresses -CsvLocation 'C:\mailboxes.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
        Alias(required)
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $mailboxes = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $currentTime = Get-CurrentDateString
		
		# for each user in users
        foreach ($mailbox in $mailboxes)
        {
            $thisMailbox = Get-Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
            if ($thisMailbox) {
                Write-Host ("updating proxy addresses for " + $thisMailbox.displayname)
                $thisUser = Get-User -Identity $thisMailbox.PrimarySmtpAddress
                
                $proxyAddresses = $thisMailbox.EmailAddresses
                $simplifiedProxyAddresses = $proxyAddresses | %{$_.toLower()}

                # add ones in csv file
                $csvAliases = $mailbox.alias.split(",")
                for ($i=0; $i -lt $csvAliases.Length; $i++) {
                    $item = "smtp:" + $csvAliases[$i].toLower()
                    if ($csvAliases[$i] -and !$simplifiedProxyAddresses.Contains($item)) {
                        # only add when not exists
                        $proxyAddresses.add("smtp:"+$csvAliases[$i])
                    }
                }

                $simplifiedProxyAddresses = $proxyAddresses | %{$_.toLower()}
                
                #[string[]]$tempProxy = $proxyAddresses | out-string -stream

                # add additional aliases for dalecarnegie domains
                $firstName = ($thisUser.FirstName).replace(" ","")
                $lastName = ($thisUser.LastName).replace(" ","")
                if($firstName) {
                    $alias1 = ("smtp:$($firstName +"_" + $lastName)@dalecarnegie.com")
                    $alias2 = ("smtp:$($firstName +"_" + $lastName)@dale-carnegie.com")
                    $alias3 = ("smtp:$($firstName +"." + $lastName)@dalecarnegie.edu")
                    $alias4 = ("smtp:$($firstName +"." + $lastName)@dale-carnegie.com")
                } else {
                    $alias1 = "smtp:$lastName@dalecarnegie.com"
                    $alias2 = "smtp:$lastName@dale-carnegie.com"
                    $alias3 = "smtp:$lastName@dalecarnegie.edu"
                    $alias4 = "smtp:$lastName@dale-carnegie.com"
                }
                #Write-Host "simplifiedProxyAddresses"
                #$simplifiedProxyAddresses
                
                if (!$simplifiedProxyAddresses.Contains($alias1.ToLower())) {
                    $proxyAddresses.Add($alias1)
                }
                if (!$simplifiedProxyAddresses.Contains($alias2.ToLower())) {
                    $proxyAddresses.Add($alias2)
                }
                if (!$simplifiedProxyAddresses.Contains($alias3.ToLower())) {
                    $proxyAddresses.Add($alias3)
                }
                if (!$simplifiedProxyAddresses.Contains($alias4.ToLower())) {
                    $proxyAddresses.Add($alias4)
                }
                # remove duplicates if exists
                $proxyAddresses = $proxyAddresses | Sort-Object | Get-Unique
                #Write-Host "proxyAddresses"
                #$proxyAddresses
                # apply it to the mailbox
                try {
                    Set-Mailbox -Identity $thisMailbox.identity -EmailAddresses $proxyAddresses -ErrorAction stop
                } catch {
                    Write-Log -Message $_ -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log" -Level Error
                }
                
            } else {
                Write-Host ("The mailbox " + $mailbox.UserPrincipalName + " does not exist. skipping...") -ForegroundColor red
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkUpdate-ProxyAddresses-"$currentTime".log" -ForegroundColor Yellow
    }
}


function BulkSet-Mailboxes {
    <#
	.SYNOPSIS
        It will update mailbox attritbues in bulk for mailboxes given in a csv.
	.EXAMPLE
		PS> BulkSet-Mailboxes -CsvLocation 'C:\mailboxes.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
        ForwardingSMTPAddress(required)
        TODO : will need to add more attributes available
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        try {
            $mailboxes = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $currentTime = Get-CurrentDateString
		
		# for each user in users
        foreach ($mailbox in $mailboxes)
        {
            $thisMailbox = Get-Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
            if ($thisMailbox) {
                if ($mailbox.ForwardingSMTPAddress) {
                    try {
                        Set-Mailbox -Identity $thisMailbox.identity -ForwardingSmtpAddress $mailbox.forwardingSMTPAddress -ErrorAction stop
                    } catch {
                        Write-Log -Message $_ -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log" -Level Error
                    }
                } 
            } else {
                $message = "The contact " + $mailbox.UserPrincipalName + " already exists. skipping..."
                Write-Host $message
                Write-Log -Message $message -Path $csvFolder"\Log_BulkNew-MailContact-"$currentTime".log" -Level Warn
            }
        }
	}
    end {
        Write-host "The process has been finished. Log has been saved in Log_BulkSet-Mailboxes-"$currentTime".log" -ForegroundColor Yellow
    }
}


function Check-MailboxExistence {
    <#
	.SYNOPSIS
        It will check if a mailbox exists in Exchange Online for mailboxes listed in the given csv and create a csv of not existing mailboxes.
	.EXAMPLE
		PS> check-MailboxExistence -CsvLocation 'C:\mailboxes.csv'
        ===List of Properties(column) in CSV===
        UserPrincipalName(required)
        displayName
        FirstName
        LastName
        or any other attributes(it will be good to have them all when you want to import them using the exported csv)
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        $currentTime = Get-CurrentDateString

        # import the user list csv to a variable users
        try {
            $mailboxes = import-csv $CsvLocation
        } catch {
           Write-Warning "The csv file does not exist. Please try again."
           return 
        }
        $mailboxes = import-csv $CsvLocation
        $csvFolder = $csvLocation.Substring(0, $CsvLocation.LastIndexOf("\"))
		$count = 0
		# for each user in users
        Write-Host ("Checking mailboxes.... please wait") -ForegroundColor DarkGray
        foreach ($mailbox in $mailboxes)
        {
            $thisMailbox = Get-Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
            Write-Host ("Checking " + $mailbox.userprincipalname) -ForegroundColor DarkGray
            $nonExistUsers = @{}
            if (!$thisMailbox) {
                Write-Host ("The mailbox " + $mailbox.userprincipalname + " does not exist.") -ForegroundColor Yellow
                $count = $count + 1
                #append one to a csv
                $mailbox | export-csv $csvFolder\missingMailboxes.csv -NoTypeInformation -Append
            }
        }

        if (!$count) {
            Write-Host ("Done. All the mailboxes exist in Exchange Online") -BackgroundColor Cyan
        } else {
            Write-Host ("Done. Total $count mailbox(es) is(are) not in Exchange Online. The list has been saved in $csvFolder\missingMailboxes.csv") -BackgroundColor Cyan
        }
	}
    end {
        #Write-host "The process has been finished. Log has been saved in Log_Check-MailboxExistence-"$currentTime".log" -ForegroundColor Yellow
    }
}

function BulkNew-DistributionGroups {}

######################################
#                                    #
#          Utility Functions         #
#                                    #
######################################

function Get-CurrentDateString {
    $now = Get-Date
    $currentTimeString = $now.Year.ToString() + "-" + $now.Month.ToString() +"-"+ $now.Day.ToString() +" "+ $now.Hour.ToString() +"h"+ $now.Minute.ToString() + 'm'
    return $currentTimeString
}

function Write-Log { 
    <# 
    .Synopsis 
       Write-Log writes a message to a specified log file with the current time stamp. 
    .DESCRIPTION 
       The Write-Log function is designed to add logging capability to other scripts. 
       In addition to writing output and/or verbose you can write to a log file for 
       later debugging. 
    .NOTES 
       Created by: Jason Wasser @wasserja 
       Modified: 11/24/2015 09:30:19 AM   
 
       Changelog: 
        * Code simplification and clarification - thanks to @juneb_get_help 
        * Added documentation. 
        * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks 
        * Revised the Force switch to work as it should - thanks to @JeffHicks 
 
       To Do: 
        * Add error handling if trying to create a log file in a inaccessible location. 
        * Add ability to write $Message to $Verbose or $Error pipelines to eliminate 
          duplicates. 
    .PARAMETER Message 
       Message is the content that you wish to add to the log file.  
    .PARAMETER Path 
       The path to the log file to which you would like to write. By default the function will  
       create the path and file if it does not exist.  
    .PARAMETER Level 
       Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
    .PARAMETER NoClobber 
       Use NoClobber if you do not wish to overwrite an existing file. 
    .EXAMPLE 
       Write-Log -Message 'Log message'  
       Writes the message to c:\Logs\PowerShellLog.log. 
    .EXAMPLE 
       Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
       Writes the content to the specified log file and creates the path and file specified.  
    .EXAMPLE 
       Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
       Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
    .LINK 
       https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
    #> 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path='C:\Logs\PowerShellLog.log', 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}