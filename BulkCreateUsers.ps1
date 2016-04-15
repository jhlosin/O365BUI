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

function New-BulkMsolUser {
    <#
	.SYNOPSIS
        It will create Windows Azure MsolUsers in bulk and modify attritbues if already exists.		
	.EXAMPLE
		PS> New-BulkMsolUser -CsvLocation '.\DaleCarnegie_LotusNotes_Jan2016\AllUsers.csv'
		
	.PARAMETER CsvLocation
	 	Location of the CSV file of users you want to import
        List of Properties in CSV
        FirstName
        LastName
        Fax
        OfficePhone
        PrimarySMTP
        OfficeAddress
        City
        State
        Zip
        Country
        UsageLocation
        ForwardingTo
        Manager
        Company
	#>
    [CmdletBinding()]
	param (
		[parameter(Mandatory=$true]
        [string]$CsvLocation
	)
	process {
        # import the user list csv to a variable users
        $users = import-csv $CsvLocation
		
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
            $license = (Get-MsolAccountSku | Out-GridView -Title "Select the license to assign" -PassThru).AccountSkuId
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
        
                # add exchange plan2 license
                Set-MsolUserLicense -UserPrincipalName $thisUser.UserPrincipalName -AddLicenses $license -ErrorAction SilentlyContinue

                # set up forwarding and aliases
                # TODO: need to create contacts first.
                #Set-Mailbox -Identity $thisUser.userPrincipalName -ForwardingSmtpAddress $ForwardingTo

                # set up other attritbutes
                Set-User -Identity $thisUser.UserPrincipalName -Manager $Manager -Company $Company

                # TODO: enable in-place hold
            }
        }
	}
}

function Set-
#### TODO
# forwarding
# proxyaddresses
# company - set-User
# manager - set-User

#enable in-place hold.