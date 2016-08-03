# Office 365 & Exchange Powershell Util
utility functions from one of my recent project

## Table of Contents

1. [Usage](#Usage)
1. [Requirements](#requirements)
1. [Utility Description](#description)
1. [Contributing](#contributing)

## Usage

1. clone or download it and run Office365Util.ps1 from powershell. 

## Requirements

- powershell v3.0 or above.
- Execution policy may need to be changed to "unretricted". Have a look [this](https://technet.microsoft.com/en-us/library/ee176961.aspx)

## Description

### MSOL Services Utils
- BulkNewOrUpdate-MsolUser
- BulkRemove-MsolUser
- BulkSet-MsolLicense
- BulkReset-MsolPassword
- BulkEmail-UserPassword
- Get-Office365UserInfo

### Exchange Online Utils
- BulkNew-MailContact
- BulkUpdate-ProxyAddresses
- BulkSet-Mailboxes
- Check-MailboxExistence
- Get-EOMailboxLogonStatistics
- BulkNew-DistributionGroups

### Onprem Exchange Utils
- Get-AllMailboxStatistics

### etc
- Connect-Office365
- Disconnect-Office365

## Contributing

- I would love to get anything that can be useful to save our time from tedious repeatitive tasks. 
- Fork it and let me know when you have anything useful for office 365 or exchange tasks.
