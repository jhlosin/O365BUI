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

function BulkNew-MsolUser {
    <#
	.SYNOPSIS
        It will create Windows Azure MsolUsers in bulk and modify attritbues if already exists.		
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

        $NewUserParams = @{
            'UserPrincipalName' = $Username
            'Name' = $Username
            'GivenName' = $FirstName
            'Surname' = $LastName
            'Title' = $Title
            'SamAccountName' = $Username
            'AccountPassword' = (ConvertTo-SecureString $DefaultPassword -AsPlainText -Force)
            'Enabled' = $true
            'Initials' = $MiddleInitial
            'Path' = "$Location,$DomainDn"
            'ChangePasswordAtLogon' = $true
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
                'Manager' = $user.manager
                'Company' = $user.Company

            }

            $thisUser = Get-MsolUser -UserPrincipalName $user.primarySMTP -ErrorAction SilentlyContinue
            # if thisUser does not exist
            if (!$thisUser) {
	            # create a new msol user
                Write-Host ("creating a new msoluser " + $UserParams.DisplayName)
                New-MsolUser @UserParams

                # set up other attritbutes
                Set-User -Identity $UserParams.UserPrincipalName -Manager $UserParams.Manager -Company $UserParams.Company
	        }
        }
	}
}

function BulkSet-MsolUser {
    <#
	.SYNOPSIS
        It will create Windows Azure MsolUsers in bulk and modify attritbues if already exists.		
	.EXAMPLE
		PS> BulkSet-MsolUser -CsvLocation 'C:\AllUsers.csv'
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
        ===List of Properties(column) in CSV===
        FirstName(required)
        LastName(required)
        UserPrincipalName(PrimarySmtpAddress)(required)
        Fax
        Department
        MobilePhone
        OfficePhone
        OfficeAddress
        City
        Title
        State
        PostalCode
        Country
        UsageLocation
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
            $FirstName = $user.FirstName
            $LastName = $user.LastName
            $fax = $user.fax
            $PhoneNumber = $user.OfficePhone
            $DisplayName =  $FirstName + " " + $LastName
            $UserName = $user.PrimarySMTP
            $StreetAddress = $user.OfficeAddress
            $city = $user.city
            $state = $user.state
            $postalCode = $user.zip
            $country = $user.country
            $UsageLocation = $user.UsageLocation
            $ForwardingTo = $user.ForwardingTo

            $Manager = $user.Manager
            $Company = $user.Company

            $thisUser = Get-MsolUser -UserPrincipalName $user.primarySMTP -ErrorAction SilentlyContinue
            # if thisUser does not exist
            if (!$thisUser) {
	            # create a new msol user
                Write-Host "creating a new msoluser $DisplayName"
                New-MsolUser -DisplayName $DisplayName -UserPrinciPalName $UserName `
                    -FirstName $FirstName -LastName $LastName -Country $Country -City $city `
                    -StreetAddress $StreetAddress -fax $fax -PhoneNumber $PhoneNumber -postalCode $postalCode `
                    -UsageLocation $UsageLocation -LicenseAssignment $license -State $state

                # set up other attritbutes
                Set-User -Identity $thisUser.UserPrincipalName -Manager $Manager -Company $Company
	        }
    
            # if alrady exist
            if ($thisUser) {
                # try to update attributes.
                Write-Host "The user $DisplayName already exists. trying to update attritbutes"
                set-MsolUser -UserPrinciPalName $thisUser.userPrincipalName `
                    -FirstName $FirstName -LastName $LastName -Country $Country -City $city `
                    -StreetAddress $StreetAddress -fax $fax -PhoneNumber $PhoneNumber -postalCode $postalCode `
                    -UsageLocation $UsageLocation -State $state
        
                # add the license you selected prior
                Set-MsolUserLicense -UserPrincipalName $thisUser.UserPrincipalName -AddLicenses $license -ErrorAction SilentlyContinue

                # set up forwarding and aliases
                # TODO: need to create contacts first.
                #Set-Mailbox -Identity $thisUser.userPrincipalName -ForwardingSmtpAddress $ForwardingTo

                # set up other attritbutes
                Set-User -Identity $thisUser.UserPrincipalName -Manager $Manager -Company $Company

                
            }
        }
	}
}

######################################
#                                    #
#       Exchange Online Utils        # 
#                                    #
######################################

function BulkEnable-inPlaceHold {
    <#
	.SYNOPSIS
        It will enable in-placehold for all users in a given CSV.
	.EXAMPLE
		PS> BulkEnabl-inPlaceHold -CsvLocation 'C:\AllUsers.csv'
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
        ===List of Properties(column) in CSV===
        PrimarySMTP(or alias)
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
        foreach ($user in $users)
        {
        }
	}
}


#### TODO
# forwarding
# proxyaddresses
# company - set-User
# manager - set-User

#enable in-place hold.