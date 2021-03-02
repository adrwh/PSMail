# PSMail

Console based PowerShell Office365 email client.

```
                               _ _ 
     _ __  ___ _ __ ___   __ _(_) |
    | '_ \/ __| '_ ` _ \ / _` | | |
    | |_) \__ \ | | | | | (_| | | |
    | .__/|___/_| |_| |_|\__,_|_|_|
    |_| 

    Get-PSMail to read, reply or forward email

    Send-PSmail to send new email

    Example: Send-PSMail -to "email@domain.com" -subject "hey psmail" -body "test from psmail"

```

PSMail was an experiment, a weekend hack job to learn and develop PowerShell.  I don't necessarily have intentions to take this much further, but it is a proof-of-concept (POC), and funnily enough, it works!

## Use case
* PSMail lets your read, reply, forward and create email, atttached to your Office365 mailbox.  It uses the Microsoft Graph API for all operations and is designed for those that enjoy the console/terminal.
* PSMail is also an option for headless systems, and those that do not have a web browser installed.

## Setup
PSMail requires an Azure Active Directory app setup with `User.Read, Mail.ReadWrite and Mail.Send` delegated Api permissions assigned.  This allows the user to read and write in their own mailbox, and requires the user to consent to these permissions when they first connect/login.

The application uses Device Code/Authorization flow and follows the MS documented process https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code

## Install
1. git clone git@github.com:adrwh/PSMail.git
1. Create your config file (see below)
1. dot source the PowerShell script
1. Start reading email

## Config
Once you have cloned the repository to your computer, create a file named `Config.psd1` in the same directory.
Add the following content and replace your ClientID and TenantID with values from your own Azure app.
```
@{
    ClientID = 'nnnnn'
    TenantID = 'nnnnn'
}
```
