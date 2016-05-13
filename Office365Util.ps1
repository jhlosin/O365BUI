######################################
#                                    #
# Initial Connection to Office 365   # 
#                                    #
######################################

function Connect-Office365
{
    <#
	.SYNOPSIS
        Connect Office 365 MSOL service and Exchange Online.
    #>
    $LiveCred = get-credential
    $global:Session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $LiveCred -Authentication basic -AllowRedirection
    Import-PSSession $global:Session365

    connect-MsolService -credential $LiveCred
}


######################################
#                                    #
#       MSOL Services Util           # 
#                                    #
######################################

function BulkNew-MsolUser {
    <#
	.SYNOPSIS
        It will create Windows Azure MsolUsers in bulk and modify attritbues to the ones in the CSV if already exists.		
	.EXAMPLE
		PS> BulkNew-MsolUser -CsvLocation 'C:\AllUsers.csv'
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
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
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)][string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        $users = import-csv $CsvLocation
        # select a license
        $license = (Get-MsolAccountSku | Out-GridView -Title "Select the license to assign" -PassThru).AccountSkuId
                
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

            $thisUser = Get-MsolUser -UserPrincipalName $UserParams.UserPrincipalName -ErrorAction SilentlyContinue
            # if thisUser does not exist
            if (!$thisUser) {
	            # create a new msol user
                Write-Host ("creating a new msoluser " + $UserParams.DisplayName)
                New-MsolUser @UserParams
	        }

            # if already exist
            if ($thisUser) {
                # update attributes.
                Write-Host ("The user " + $UserParams.DisplayName + " already exists. trying to update attritbutes")
                set-MsolUser @UserParams
        
                # add the license you selected prior
                Set-MsolUserLicense -UserPrincipalName $thisUser.UserPrincipalName -AddLicenses $license -ErrorAction SilentlyContinue
            }

            # max company attribute is 64, cut it if it is longer than 64
            if ($user.company.length -gt 64) {
                ## HERE
                $company = $user.manager.Substring(0,63)
            }
            # set up other attritbutes
            Set-User -Identity $thisUser.UserPrincipalName -Manager $user.Manager -Company $company

        }
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
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
        ===List of Properties(column) in CSV===
        name(required)
        externalEmailAddress(required)
        FirstName
        Lastname
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        $users = import-csv $CsvLocation
		
		# for each user in users
        foreach ($item in $items)
        {
            $thisUser = Get-MailContact $item.name -ErrorAction SilentlyContinue
            if (!$thisUser) {
                Write-Host ("creating " + $item.name)
                New-MailContact -Name $item.name -ExternalEmailAddress $item.ExternalEmailAddress -DisplayName $item.name -FirstName $item.firstName -LastName $item.lastName
            } else {
                Write-Host ("The contact " + $item.name + " already exists. skipping...")
            }
        }
	}
}


function BulkSet-Mailbox {
    <#
	.SYNOPSIS
        It will create mailcontact in bulk for all items in a given CSV.
	.EXAMPLE
		PS> BulkNew-MailContact -CsvLocation 'C:\conatcts.csv'
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
        ===List of Properties(column) in CSV===
        name(required)
        externalEmailAddress(required)
        FirstName
        Lastname
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        $users = import-csv $CsvLocation
		
		# for each user in users
        foreach ($item in $items)
        {
            $thisUser = Get-MailContact $item.name -ErrorAction SilentlyContinue
            if (!$thisUser) {
                Write-Host ("creating " + $item.name)
                New-MailContact -Name $item.name -ExternalEmailAddress $item.ExternalEmailAddress -DisplayName $item.name -FirstName $item.firstName -LastName $item.lastName
            } else {
                Write-Host ("The contact " + $item.name + " already exists. skipping...")
            }
        }
	}
}

#### TODO
# forwarding
# proxyaddresses

#enable in-place hold.
#create contacts
# set up forwarding and aliases

#Set-Mailbox -Identity $thisUser.userPrincipalName -ForwardingSmtpAddress $ForwardingTo
