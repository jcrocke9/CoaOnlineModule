#Require -Version 5.0

using namespace System;
using namespace System.Text;
using namespace System.Diagnostics;
using namespace System.Linq;
using namespace System.Collections.Generic;
Import-Module ActiveDirectory;
Import-Module CoaLoggingModule -Function Add-CoaWriteToLog
<#
    .Synopsis
    Post-creation Exchange Online mailbox configuration for new accounts.

    .Description
    These configuration modifications put the mailbox in-policy for COA business needs.

    .Parameter NumberOfDays
    Default -120. The number of days backward (negative!) in time to gather new accounts and ensure they are in-policy.

    .Example
    # Runs the configuration for the new mailboxes created in the past 120 days.
    Set-CoaMailboxConfiguration

    .Example
    # Runs the configuration for the new mailboxes created in the past 30 days.
    Set-CoaMailboxConfiguration -NumberOfDays 30
#>
function Set-CoaMailboxConfiguration {
    Param (
        # Default -120. The number of days backward (negative) in time to gather new accounts and ensure they are in-policy.
        [Parameter()]
        [int]
        $NumberOfDays = -120,
        [string]$FileName = "MailboxConfiguration",
        [string]$RoleAssignmentPolicy = "COA Default Role Assignment Policy",
        [string]$ClientAccessPolicyName = "COAOWAMailboxPolicy",
        [int]$LitigationHoldDuration = 1,
        [string]$ExchangeOnlineAdminAccount = "COA Administrator",
        [string]$RetentionPolicyE3 = "COA Policy",
        [string]$RetentionPolicyK1 = "COA F1 Policy",
        [string]$CoaSkuInformationWorkers = "ALEXANDRIAVA1:ENTERPRISEPACK",
        [string]$CoaSkuFirstlineWorkers = "ALEXANDRIAVA1:DESKLESSPACK"
    )
    $logCode = "Start"
    $writeTo = "Starting Mailbox Configuration Script"
    Add-CoaAdd-CoaWriteToLog -writeTo $writeTo -logCode $logCode -FilePath $FileName

    $UserList = @()
    Clear-Variable UserList

    $ErrorActionPreference = "Stop"
    $UserList = Get-Mailbox -ResultSize unlimited -Filter {(ArchiveStatus -eq $None)} | Where-Object {$_.WhenCreated -gt (get-date).AddDays($NumberOfDays)} | Select-Object -ExpandProperty userPrincipalName
    Write-Output "User count found: " + $UserList.Count
    $logCode = "Get"
    $writeTo = "User count found: " + $UserList.Count
    foreach ($upn in $UserList) {

        try {
            Enable-Mailbox -identity $upn –Archive
            $writeTo = "Enable-Mailbox: identity $upn Archive"
            $logCode = "Success"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch {
                    
            $writeTo = "Enable-Mailbox: identity $upn Archive"
            $logCode = "Error"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
        }
        try {
    
            Set-Mailbox -identity $upn –LitigationHoldEnabled $True –LitigationHoldDuration $LitigationHoldDuration –RoleAssignmentPolicy $RoleAssignmentPolicy
            $writeTo = "Set-Mailbox: Successfully set mailbox $upn litigation hold and assignment policy"
            $logCode = "Success"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch {
    
            $writeTo = "Set-Mailbox: FAILED set mailbox $upn litigation hold and assignment policy"
            $logCode = "Error"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
        }
        try {
            Add-MailboxPermission -identity $upn -User $ExchangeOnlineAdminAccount -AccessRights fullaccess -InheritanceType all -AutoMapping $false
            $writeTo = "Add-MailboxPermission: Successfully added $upn mailbox permission for $ExchangeOnlineAdminAccount"
            $logCode = "Success"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch {
    
            $writeTo = "Add-MailboxPermission: FAILED added $upn mailbox permission for COA Admin"
            $logCode = "Error"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
        }
    
        try {
            Set-Clutter -Identity $upn -Enable $false
            $writeTo = "Set-Clutter: Successfully set mailbox $upn clutter to false"
            $logCode = "Success"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch {
    
            $writeTo = "Set-Clutter: FAILED set mailbox $upn clutter to false"
            $logCode = "Error"
            
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
        }
    
        $RecipientTypeDetails = (Get-Mailbox -Identity $upn).RecipientTypeDetails
    
        if ($RecipientTypeDetails -eq "UserMailbox") {
    
            try {
                Set-CASMailbox -Identity $upn –OWAMailboxPolicy $ClientAccessPolicyName 
                Set-CASMailbox -Identity $upn -PopEnabled $false
                Set-CASMailbox -Identity $upn -ImapEnabled $false
                $writeTo = "Set-CASMailbox: Successfully set mailbox $upn client access permissions"
                $logCode = "Success"
                
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            catch {
    
                $writeTo = "Set-CASMailbox: FAILED set mailbox $upn client access permissions"
                $logCode = "Error"
                
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
            }
    
            try {
                $upn = (Get-MsolUser -SearchString $upn).UserPrincipalName
                $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
            }
            catch {
                $writeTo = "Get-MSOLUser -UserPrincipalName $upn FAILED"
                $logCode = "Error"
                
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                Continue
            }
            
            # E3
            if ($LicenseLineItem -contains $CoaSkuInformationWorkers) { 
                try {
                    Set-Mailbox -Identity $upn -RetentionPolicy $RetentionPolicyE3
                    $writeTo = "Set-Mailbox: Successfully set mailbox $upn policy to $RetentionPolicyE3"
                    $logCode = "Success"
                    
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                }
                catch {
    
                    $writeTo = "Set-Mailbox: FAILED set mailbox $upn policy to $RetentionPolicyE3"
                    $logCode = "Error"
                    
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
                }
            }
    
            # K1
            if ($LicenseLineItem -contains $CoaSkuFirstlineWorkers) {
                try {
                    Set-Mailbox -Identity $upn -RetentionPolicy $RetentionPolicyK1
                    $writeTo = "Set-Mailbox: Successfully set mailbox $upn policy to $RetentionPolicyK1"
                    $logCode = "Success"
                    
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                }
                catch {
    
                    $writeTo = "Set-Mailbox: FAILED set mailbox $upn policy to $RetentionPolicyK1"
                    $logCode = "Error"
                    
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
    
                }
            }
    
            Clear-Variable LicenseLineItem
    
        }
        Clear-Variable upn
    }
    
    Clear-Variable UserList
    $ErrorActionPreference = "Continue"
}

function WriteToLog {
    param([string]$logLineTime, [string]$writeTo, [string]$logCode)
    $logFileDate = Get-Date -UFormat "%Y%m%d"
    $logLineInfo = "`t$([Environment]::UserName)`t$([Environment]::MachineName)`t"
    $logLine = $logLineTime
    $logLine += $logLineInfo
    $logLine += $logCode; $logLine += "`t"
    $logLine += $writeTo
    $logLine | Out-File -FilePath "C:\Logs\NewUserScript_$logFileDate.log" -Append -NoClobber
    Return;
}

function OpenLog {
    $logLineTime = (Get-Date).ToString()
    $logCode = "Start"
    $writeTo = "Starting New User script"
    WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    Return;
}
function SendAnEmail {
    param ([string]$subject, [string]$body)
    $emailAddress = "$([Environment]::UserName)@alexandriava.gov"
    Send-MailMessage -To $emailAddress -From "COA New User Module <noreply@alexandriava.gov>" -Subject $subject -Body $body -SmtpServer "smtp.alexgov.net" -Port 25
    $writeTo = "Send-MailMessage`t$subject`t$body"
    $logCode = "Email"
    $logLineTime = (Get-Date).ToString()
    WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    Return;
}

function QueryAdToValidateUsers {
    param ([string]$samAccountName)
    $ErrorActionPreference = "stop"
    try {
        $mail = Get-ADUser $samAccountName -Properties mail | Select-Object mail -ExpandProperty mail
        if (!$mail) {$mail = "None set"}
        $department = Get-ADUser $samAccountName -Properties department | Select-Object department -ExpandProperty department
        $userAccountControl = Get-ADUser $samAccountName -Properties userAccountControl | Select-Object userAccountControl -ExpandProperty userAccountControl
        if ($userAccountControl -eq 514) {
            $userAccountControl = "User Disabled"
        }
        elseif ($userAccountControl -eq 512) {
            $userAccountControl = "Active"
        }
        elseif ($userAccountControl -eq 66048) {
            $userAccountControl = "Has a non-expiring password"
        }
        else {
            $problemUsers += $samAccountName
            $writeTo = "Get-ADUser: $samAccountName`t$userAccountControl"
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
            SendAnEmail -subject "New-MSOLUser Error" -body "Get-ADUser: samAccountName cannot be found for $samAccountName"
        }
        $writeTo = "$samAccountName | $userAccountControl | $mail | $department"
        $logCode = "Start"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }
    catch {
        if ($user.Length -gt 20) {
            $writeTo = "The user name $samAccountName is more than 20 characters"
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
            SendAnEmail -subject "New-MSOLUser Error" -body "The user name $samAccountName is more than 20 characters; no account was created."
        }
        $problemUsers.Add($samAccountName);
        $writeTo = "Get-ADUser: samAccountName cannot be found for $samAccountName"
        $logCode = "Error"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        SendAnEmail -subject "New-MSOLUser Error" -body "Get-ADUser: samAccountName cannot be found for $samAccountName"
    }
    $ErrorActionPreference = "continue"
    Return;
}
#endregion

#region: Sets the mail and SMTP attributes, if needed
function SetMailAndSmtpAttributes {
    param([string]$user)
    $mail = Get-ADUser $user -Properties mail | Select-Object mail -ExpandProperty mail
    if (!$mail) {
        try {
            Set-ADUser -Identity $user -EmailAddress "$user@alexandriava.gov" -ErrorAction Stop
            $writeTo = "Set-ADUser: Successfully added email address to $user@alexandriava.gov"
            $logCode = "Success"
        }
        catch {
            $logCode = "Error"
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
        }    
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }
    Start-Sleep -Seconds 5
    $mail = Get-ADUser $user -Properties mail | Select-Object mail -ExpandProperty mail
    $SMTP = Get-ADUser -Identity $user -Properties proxyaddresses | Select-Object proxyaddresses -ExpandProperty proxyaddresses
    if (!$SMTP) {
        try {
            Set-ADUser -Identity $user -Add @{Proxyaddresses = "SMTP:" + $mail } -ErrorAction Stop
            Set-ADUser -Identity $user -Add @{targetAddress = "SMTP:" + $mail } -ErrorAction Stop
            Set-ADUser -Identity $user -Replace @{mailNickname = $user}
            $writeTo = "Set-ADUser: Successfully set SMTP address to SMTP:$mail"
            $logCode = "Success"
        }
        catch {
            $logCode = "Error"
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
        }    
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }

    Return;
}

function SetLicenseAttributeE3 {
    param([string]$user)
    $extensionAttribute13 = Get-ADUser -Identity $user -Properties extensionAttribute13 | Select-Object -ExpandProperty extensionAttribute13
    if (!$extensionAttribute13) {
        try {
            Set-ADUser -Identity $user -Add @{extensionAttribute13 = "E3"}
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = E3"
            $logCode = "Success"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        
        }
        catch {
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode 
        }
    }
    else {
        try {
            Set-ADUser -Identity $user -Replace @{extensionAttribute13 = "E3"}
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = E3"
            $logCode = "Success"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        
        }
        catch {
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode 
        }
    }
}

function SetLicenseAttributeK1 {
    param([string]$user)
    $extensionAttribute13 = Get-ADUser -Identity $user -Properties extensionAttribute13 | Select-Object -ExpandProperty extensionAttribute13
    if (!$extensionAttribute13) {
        try {
            Set-ADUser -Identity $user -Add @{extensionAttribute13 = "K1"}
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = K1"
            $logCode = "Success"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        
        }
        catch {
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode 
        }
    }
    else {
        try {
            Set-ADUser -Identity $user -Replace @{extensionAttribute13 = "K1"}
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = K1"
            $logCode = "Success"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        
        }
        catch {
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode 
        }
    }
}
class UserObject {
    [string]$samAccountName
    [string]$License
}
$global:UsersToWorkThrough = [System.Collections.Generic.List[UserObject]]::new();
<#
    .Synopsis
    Sets new mailbox accounts up with the standard policies of COA

    .Description
    Sets new mailbox accounts up with: email address, customAttribute13, smtp, targetAddress

    .Example
    # Sets Exchange attributes of given single samAccountName with E3
    New-CoaUser joe.crockett | Set-CoaExchangeAttributes

    .Example
    # Sets Exchange attributes of given single samAccountName with K1
    New-CoaUser joe.crockett -Firstline | Set-CoaExchangeAttributes
    
    .Example
    # Sets Exchange attributes of prepopulated New-CoaUser
    New-CoaUser joe.crockett
    New-CoaUser heladio.martinez -Firstline
    Set-CoaExchangeAttributes

#>
function Set-CoaExchangeAttributes {
    Param (
        [parameter(
            Position = 0,
            ValueFromPipeline = $false)]
        [System.Collections.Generic.List[UserObject]]
        $UserList = $global:UsersToWorkThrough,
        [parameter(
            Position = 1,
            ValueFromPipeline = $true)]    
        [UserObject]
        $SingleUser
    )
    $problemUsers = [System.Collections.Generic.List[System.Object]]::new();    
    # $standardUsers = [System.Collections.Generic.List[System.Object]]::new();
    # $basicUsers = [System.Collections.Generic.List[System.Object]]::new();
    # $standardLicenseName = "emailStandard_createAlexID"
    $basicLicenseName = "emailBasic_createAlexID"

    OpenLog
    if ($SingleUser) {
        QueryAdToValidateUsers -samAccountName $SingleUser.samAccountName
        # Remove problem users
        SetMailAndSmtpAttributes -user $SingleUser.samAccountName
        if ($SingleUser.License -eq $basicLicenseName) {
            SetLicenseAttributeK1 -user $SingleUser.samAccountName
        }
        else {
            SetLicenseAttributeE3 -user $SingleUser.samAccountName
        }
        $global:UsersToWorkThrough.Remove($SingleUser);
    }
    else {
        foreach ($samAccountName in $UserList ) {
            QueryAdToValidateUsers -samAccountName $samAccountName.samAccountName
        }

        foreach ($problemUser in $problemUsers) {
            $UserList.Remove("$problemUser");
        }

        foreach ($user in $UserList) {
            SetMailAndSmtpAttributes -user $user.samAccountName
        }

        foreach ($user in $UserList) {
            if ($user.License -eq $basicLicenseName) {
                SetLicenseAttributeK1 -user $user.samAccountName
            }
            else {
                SetLicenseAttributeE3 -user $user.samAccountName
            }
        }
        $global:UsersToWorkThrough.Clear()
    }
}


function New-CoaUser {
    Param (
        [parameter(Mandatory = $true,
            Position = 0)] 
        [string]$SamAccountName,
        [switch]$Firstline
    )
    $user = $null
    $user = [UserObject]::new()
    if ($Firstline) {
        $user.License = "emailBasic_createAlexID"
    }
    else {
        $user.License = "emailStandard_createAlexID"
    }
    $user.samAccountName = $samAccountName
    $global:UsersToWorkThrough.Add($user)
    Write-Output $global:UsersToWorkThrough
    # Write-Output "Use New-CoaAssignments $UsersToWorkThrough to complete."
}
Export-ModuleMember -Function Set-CoaMailboxConfiguration, Set-CoaExchangeAttributes