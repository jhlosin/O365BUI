# Office 365 & Exchange Powershell Util
utility functions from one of my recent project

## Table of Contents

1. [Usage](#Usage)
1. [Requirements](#requirements)
1. [Utility Description](#description)
1. [Contributing](#contributing)

## Usage

1. clone or download it and run Office365Util.ps1 from powershell.([How to run a powershell script](http://stackoverflow.com/questions/2035193/how-to-run-a-powershell-script))
2. Always try get-help office365utilcmdname to see full description of a cmdlet to see how to use one. e.g) Get-Help BulkNewOrUpdate-MsolUser -Full

## Requirements

- powershell v3.0 or above.
- Execution policy may need to be changed to "unretricted". Have a look [this](https://technet.microsoft.com/en-us/library/ee176961.aspx)

## Description

### MSOL Services Utils
- BulkNewOrUpdate-MsolUser
  - create Windows Azure MsolUsers in bulk and modify attritbues to the ones in the CSV if already exists.

- BulkRemove-MsolUser
  - forcefully remove Windows Azure MsolUsers for the user given in the csv.

- BulkSet-MsolLicense
  - assign license you select for the user given in the csv or all user mailboxes.

- BulkReset-MsolPassword
  - reset passwords for the user(s) given in the csv.

- BulkEmail-UserPassword
  - distirbute emails wit user password given in a csv file.

- Get-Office365UserInfo
  - get office 365 user info from the connected tenant.

### Exchange Online Utils
- BulkNew-MailContact
  - create mailcontact in bulk for all items in a given CSV.

- BulkUpdate-ProxyAddresses
  - update proxy address for Dale Carnegie for mailboxes in a given csv.

- BulkSet-Mailboxes
  - update mailbox attritbues in bulk for mailboxes given in a csv.

- Check-MailboxExistence
  - check if a mailbox exists in Exchange Online for mailboxes listed in the given csv and create a csv of not existing mailboxes.

- Get-EOMailboxLogonStatistics
  - Get mailbox login statistics from the connected Exchange Online Session

- BulkNew-DistributionGroups
  - Creates distribution groups in Exchange Online for ones given in the csv

### Onprem Exchange Utils
- Get-AllMailboxStatistics
  - Get all mailbox statistics from an onprem Exchange server.

### etc
- Connect-Office365
- Disconnect-Office365

## Contributing

- I would love anything that can be useful to save our time from tedious repeatitive tasks.
- Fork it and let me know when you have anything useful for office 365 or exchange tasks or that can improve the code.
- For any suggestions, create an issue from [here](https://github.com/jhlosin/Office-365-Exchange-Util/issues)
