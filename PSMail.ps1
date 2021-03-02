$VerbosePreference = "Continue"
$Script:Authentication = ""
$Script:Token = ""
$Config = Import-PowerShellDataFile -Path Config.psd1

#region


$DeviceCodeParams = @{
    Method = 'POST'
    Uri    = "https://login.microsoftonline.com/$($Config.TenantID)/oauth2/v2.0/devicecode"
    Body   = @{
        client_id = $Config.ClientId
        scope     = "User.Read Mail.ReadWrite Mail.Send" 
    }
}

$DeviceCode = Invoke-RestMethod @DeviceCodeParams -Verbose:$false
Write-Host $DeviceCode.message -ForegroundColor Yellow
Set-Clipboard -Value $DeviceCode.user_code
Start-Process $DeviceCode.verification_uri

$TokenParams = @{
    Method = 'POST'
    Uri    = "https://login.microsoftonline.com/$($Config.TenantID)/oauth2/v2.0/token"
    Body   = @{
        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $DeviceCode.device_code
        client_id  = $Config.ClientId
    }
}
$Authentication = $null

$TimeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()
while ([string]::IsNullOrEmpty($Authentication.access_token)) {
    if ($TimeoutTimer.Elapsed.TotalSeconds -gt 60) {
        
        throw 'Login timed out, please try again.'
    }
    try {
        $Authentication = Invoke-RestMethod @TokenParams -ErrorAction Stop -Verbose:$false
    }
    catch {
        $Message = $_.ErrorDetails.Message | ConvertFrom-Json
        if ($Message.error -ne "authorization_pending") {
            throw
        }
    }
    Start-Sleep -Seconds 5 -Verbose
}

$Token = $Authentication.access_token | ConvertTo-SecureString -AsPlainText


#endregion

function Get-RefreshToken {
    param (
        $refresh_token
    )

    $TokenParams = @{
        Method = 'POST'
        Uri    = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body   = @{
            grant_type    = "refresh_token"
            refresh_token = $Authentication.refresh_token
            client_id     = $ClientId
        }
    }
    
    try {
        $Authentication = Invoke-RestMethod @TokenParams -ErrorAction Stop -Verbose:$false
        return $Authentication
    }
    catch {
        $_
    }
    
}

function Get-PSMail {
    # Write-Verbose "Getting email.."
    $Messages = Invoke-RestMethod -Authentication OAuth -Token $Token -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages' -Verbose:$false
    $Index = 0
    $EmailMessages = $Messages.value.foreach{
        [pscustomobject]@{
            Id          = $_.id
            Index       = $Index
            Received    = $_.receivedDateTime
            Sender      = $_.from.emailAddress.address
            Subject     = $_.subject
            BodyPreview = $_.bodyPreview
            WebLink     = $_.webLink
        } 
        $Index++
    }
    
    [Console]::ForegroundColor = "Green"
    $EmailMessages | Select-Object Index, Received, Sender, Subject | Format-Table -AutoSize
    [Console]::ResetColor()

    @'
    Type r0 to reply to email index 0
    Type f3 to forward email index 3
    Type p6 to preview email index 6 in the terminal & browser
    Type d8 to delete email index 8

'@
    $EmailAction = Read-Host -Prompt "Select your email"
    $EmailItem = $EmailMessages["$($EmailAction[1])"]
    switch -regex ($EmailAction) {
        'p\d' {
            [Console]::BackgroundColor = "Black"
            [Console]::ForegroundColor = "Yellow"
            $EmailItem | Select-Object Received, Sender, Subject | Format-Table -AutoSize
            $EmailItem.bodyPreview
            "`n"
            [Console]::ResetColor() 
            $Answer = Read-Host -Prompt "Read it online y/n?"
            if ($Answer -eq "y") {
                Start-Process $EmailItem.webLink;
            }
            Get-PSMail; Break 
        }
        'r\d' { Send-PSMail -Id $EmailItem.id -Reply; Break }
        'f\d' { Send-PSMail -Id $EmailItem.id -Forward; Break }
        'd\d' { Send-PSMail -Id $EmailItem.id -Delete; Break }
        Default {}
    }
}

function Send-PSMail {
    param (
        # Parameter help description
        [Parameter(ParameterSetName = 'New')]
        [String]$To,
        [Parameter(ParameterSetName = 'New')]
        [String]$Subject,
        [Parameter(ParameterSetName = 'New')]
        [String]$Body,
        [Switch]$Reply,
        [Switch]$Forward,
        [Switch]$Delete,
        [String]$Id
    )

    switch ($PSBoundParameters.Keys) {
        'Reply' { $Action = 'reply'; Break }
        'Forward' { $Action = 'forward'; Break }
        'Delete' { $Action = 'forward'; Break }
        Default {}
    }

    if ($PsCmdlet.ParameterSetName -eq "New") {
        $Uri = 'https://graph.microsoft.com/v1.0/me/messages'
        $Method = "Post"
        $DraftMessagePost = @{
            "subject"      = $Subject
            "body"         = @{
                "content" = $Body
            }
            "toRecipients" = @(
                @{
                    "emailAddress" = @{
                        "address" = $To
                    }
                }
            )
        }
    }
    elseif ($PSBoundParameters.Delete) {
        $Uri = "https://graph.microsoft.com/v1.0/me/messages/$Id/"
        $Method = "Delete"
    }
    else {
        $Uri = "https://graph.microsoft.com/v1.0/me/messages/$Id/$Action"
        $Method = "Post"
        $Comment = Read-Host -Prompt "Type your comments.."
        $DraftMessagePost = @{
            "comment" = $Comment
        }
    }


    $Params = @{
        Authentication = "OAuth"
        ContentType    = "application/json"
        Method         = $Method
        Token          = $Token
        Body           = $DraftMessagePost | ConvertTo-Json -Depth 3
        Uri            = $Uri
        Verbose        = $false
    }

    try {
        Write-Verbose "Processing email.."
        Invoke-RestMethod @Params
        Get-PSMail
    }
    catch {
        
    }    
}

$Me = Invoke-RestMethod -Authentication OAuth -Token $Token -Uri 'https://graph.microsoft.com/v1.0/me' -Verbose:$false
"`n"
Write-Host ("Hello {0}" -f $Me.displayName) -ForegroundColor Yellow

@'
                                _ _ 
     _ __  ___ _ __ ___   __ _(_) |
    | '_ \/ __| '_ ` _ \ / _` | | |
    | |_) \__ \ | | | | | (_| | | |
    | .__/|___/_| |_| |_|\__,_|_|_|
    |_|                            

    Get-PSMail to retrieve ,display, reply or forward email

    Send-PSmail to send new email

    Example: Send-PSMail -to "email@domain.com" -subject "hey psmail" -body "test from psmail"


'@