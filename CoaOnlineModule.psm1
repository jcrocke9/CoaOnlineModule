#Require -Version 5.0

using namespace System;
using namespace System.Text;
using namespace System.Diagnostics;
using namespace System.Linq;
using namespace System.Collections.Generic;
Import-Module ActiveDirectory;
Import-Module -Name C:\alex\CoaOnlineModule\CoaLoggingModule.psm1 -Function Add-CoaWriteToLog
<#
    .Synopsis
    Use this cmdlet to set new variables over the default ones.

    .Description
    Sets the script variables that are in common use across the functions
#>
function Set-CoaVariables {
    Param (
        # Default -120. The number of days backward (negative) in time to gather new accounts and ensure they are in-policy.
        [Parameter()]
        [int]
        $NumberOfDays = -120,
        [string]$FileName = "MailboxConfiguration",
        [string]$RoleAssignmentPolicy = "COA Default Role Assignment Policy",
        [string]$ClientAccessPolicyName = "COAOWAMailboxPolicy",
        [int]$LitigationHoldDuration = 730,
        [string]$ExchangeOnlineAdminAccount = "COA Administrator",
        [string]$RetentionPolicyE3 = "COA Policy",
        [string]$RetentionPolicyK1 = "COA F1 Policy",
        [string]$CoaSkuInformationWorkers = "ALEXANDRIAVA1:ENTERPRISEPACK_GOV",
        [string]$CoaSkuFirstlineWorkers = "ALEXANDRIAVA1:DESKLESSPACK_GOV",
        [string]$CoaSkuExoArchive = "ALEXANDRIAVA1:EXCHANGEARCHIVE_ADDON",
        [string]$CoaSkuExoAtp = "ALEXANDRIAVA1:ATP_ENTERPRISE_GOV",
        [string]$standardLicenseName = "emailStandard_createAlexID",
        [string]$basicLicenseName = "emailBasic_createAlexID",
        [string]$Domain = "alexandriava.gov"
    )
    $Script:NumberOfDays = $NumberOfDays;
    $Script:FileName = $FileName
    $Script:RoleAssignmentPolicy = $RoleAssignmentPolicy
    $Script:ClientAccessPolicyName = $ClientAccessPolicyName
    $Script:LitigationHoldDuration = $LitigationHoldDuration
    $Script:ExchangeOnlineAdminAccount = $ExchangeOnlineAdminAccount
    $Script:RetentionPolicyE3 = $RetentionPolicyE3
    $Script:RetentionPolicyK1 = $RetentionPolicyK1
    $Script:CoaSkuInformationWorkers = $CoaSkuInformationWorkers
    $Script:CoaSkuFirstlineWorkers = $CoaSkuFirstlineWorkers
    $Script:CoaSkuExoArchive = $CoaSkuExoArchive
    $Script:CoaSkuExoAtp = $CoaSkuExoAtp
    $Script:StandardLicenseName = $standardLicenseName
    $Script:BasicLicenseName = $basicLicenseName
    $Script:Domain = $Domain
}
<#
    .Synopsis
    Use this cmdlet to view new variables over the default ones.

    .Description
    View the script variables that are in common use across the functions
#>
function Get-CoaVariables {
    $Private:CoaVariables = [ordered]@{
        NumberOfDays               = $Script:NumberOfDays;
        FileName                   = $Script:FileName;
        RoleAssignmentPolicy       = $Script:RoleAssignmentPolicy;
        ClientAccessPolicyName     = $Script:ClientAccessPolicyName;
        LitigationHoldDuration     = $Script:LitigationHoldDuration;
        ExchangeOnlineAdminAccount = $Script:ExchangeOnlineAdminAccount;
        RetentionPolicyE3          = $Script:RetentionPolicyE3;
        RetentionPolicyK1          = $Script:RetentionPolicyK1;
        CoaSkuInformationWorkers   = $Script:CoaSkuInformationWorkers;
        CoaSkuFirstlineWorkers     = $Script:CoaSkuFirstlineWorkers;
        CoaSkuExoArchive           = $Script:CoaSkuExoArchive;
        CoaSkuExoAtp               = $Script:CoaSkuExoAtp;
        standardLicenseName        = $Script:StandardLicenseName;
        basicLicenseName           = $Script:BasicLicenseName;
        Domain                     = $Script:Domain;
    }
    Write-Output $Private:CoaVariables
}
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
        $NumberOfDays = $Script:NumberOfDays,
        [string]$FileName = $Script:FileName,
        [string]$RoleAssignmentPolicy = $Script:RoleAssignmentPolicy,
        [string]$ClientAccessPolicyName = $Script:ClientAccessPolicyName,
        [int]$LitigationHoldDuration = $Script:LitigationHoldDUration,
        [string]$ExchangeOnlineAdminAccount = $Script:ExchangeOnlineAdminAccount,
        [string]$RetentionPolicyE3 = $Script:RetentionPolicyE3,
        [string]$RetentionPolicyK1 = $Script:RetentionPolicyK1,
        [string]$CoaSkuInformationWorkers = $Script:CoaSkuInformationWorkers,
        [string]$CoaSkuFirstlineWorkers = $Script:CoaSkuFirstlineWorkers
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
            Enable-Mailbox -identity $upn -Archive
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
    
            Set-Mailbox -identity $upn -LitigationHoldEnabled $True -LitigationHoldDuration $LitigationHoldDuration -RoleAssignmentPolicy $RoleAssignmentPolicy
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
                Set-CASMailbox -Identity $upn -OWAMailboxPolicy $ClientAccessPolicyName 
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
#region: Sets Active Directory attributes for Exchange Online
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
            Show-CoaCustomError -subject "New-MSOLUser Error" -body "Get-ADUser: samAccountName cannot be found for $samAccountName"
        }
        $writeTo = "$samAccountName`t$userAccountControl`t$mail`t$department";
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
            Show-CoaCustomError -subject "New-MSOLUser Error" -body "The user name $samAccountName is more than 20 characters; no account was created."
        }
        $problemUsers.Add($samAccountName);
        $writeTo = "Get-ADUser: samAccountName cannot be found for $samAccountName"
        $logCode = "Error"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        Show-CoaCustomError -subject "New-MSOLUser Error" -body "Get-ADUser: samAccountName cannot be found for $samAccountName"
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
            Set-ADUser -Identity $user -EmailAddress "$user@$Script:Domain" -ErrorAction Stop
            $writeTo = "Set-ADUser: Successfully added email address to $user@$Script:Domain"
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
$global:CoaUsersToWorkThrough = [System.Collections.Generic.List[UserObject]]::new();
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
        $UserList = $global:CoaUsersToWorkThrough,
        [parameter(
            Position = 1,
            ValueFromPipeline = $true)]    
        [UserObject]
        $SingleUser
    )
    $problemUsers = [System.Collections.Generic.List[System.Object]]::new();    
    # $standardUsers = [System.Collections.Generic.List[System.Object]]::new();
    # $basicUsers = [System.Collections.Generic.List[System.Object]]::new();

    OpenLog
    if ($SingleUser) {
        QueryAdToValidateUsers -samAccountName $SingleUser.samAccountName
        # Remove problem users
        SetMailAndSmtpAttributes -user $SingleUser.samAccountName
        if ($SingleUser.License -eq $Script:BasicLicenseName) {
            SetLicenseAttributeK1 -user $SingleUser.samAccountName
        }
        else {
            SetLicenseAttributeE3 -user $SingleUser.samAccountName
        }
        $global:CoaUsersToWorkThrough.Remove($SingleUser);
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
            if ($user.License -eq $Script:BasicLicenseName) {
                SetLicenseAttributeK1 -user $user.samAccountName
            }
            else {
                SetLicenseAttributeE3 -user $user.samAccountName
            }
        }
        # $global:CoaUsersToWorkThrough.Clear()
    }
}
#endregion
#region: Sets ExO Attributes

<# class EmailUser {
    [string]$samAccountName
    [string]$license
    [string]$emailRequired
    [string]$sys_created_by
    [string]$sys_created_on
} #>

<# function WriteToLog {
    param([string]$logLineTime, [string]$writeTo, [string]$logCode)
    $logFileDate = Get-Date -UFormat "%Y%m%d"
    $logLineInfo = "`t$([Environment]::UserName)`t$([Environment]::MachineName)`t"
    $logLine = $logLineTime
    $logLine += $logLineInfo
    $logLine += $logCode; $logLine += "`t"
    $logLine += $writeTo
    $logLine | Out-File -FilePath "C:\Logs\NewUserScript_$logFileDate.log" -Append -NoClobber
    Return;
} #>

function OpenLog {
    $logLineTime = (Get-Date).ToString()
    $logCode = "Start"
    $writeTo = "Starting New User script"
    WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    Return;
}

<# function SortTheUserLicenses {
    Import-Csv $pathToExcel | ForEach-Object {
        $impUser = New-Object EmailUser;
        $impUser.samAccountName = $_.firstName + "." + $_.lastName;
        $impUser.license = $_.emailType;
        $impUser.emailRequired = $_.emailRequired;
        $impUser.sys_created_by = $_.sys_created_by;
        $impUser.sys_created_on = $_.sys_created_on;
        if ($impUser.emailRequired -eq "Yes") {
            $usersFromExcel.Add($impUser.samAccountName.ToString());
        }
        
    }
} #>

function Show-CoaCustomError {
    param ([string]$subject, [string]$body)
    Write-Error "`n$subject`n$body"
    Return;
}

function basicLicensePack {
    $disabledPlans = @()
    $O365License = New-MsolLicenseOptions -AccountSkuId $Script:CoaSkuFirstlineWorkers -DisabledPlans $disabledPlans
    Return $O365License;
}
function standardLicensePack {
    $disabledPlans = @()
    $disabledPlans += "YAMMER_ENTERPRISE"
    $O365License = New-MsolLicenseOptions -AccountSkuId $Script:CoaSkuInformationWorkers -DisabledPlans $disabledPlans
    Return $O365License;
}

function Set-ValidateUsersUpn {
    param([UserObject]$SingleUser); 
    $user = $SingleUser.samAccountName;
    $arrayFromGet = @()
    $arrayFromGet += Get-MsolUser -SearchString $user | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
    if ($arrayFromGet.Count -eq 1) {
        $upn = $arrayFromGet[0]
        $writeTo = "Get-MsolUser`t$user`tSearchString returned: $upn"
        $logCode = "Get"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }
    else {
        if ($arrayFromGet.Count -gt 1) {
            $errMsg = "Either the samAccountName was empty, or the search returned more than one value."
        }
        else {
            $errMsg = "The user $user cannot be found in MSOL, and has been removed from processing."
        }
        $writeTo = $errMsg
        $logCode = "Else"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        Show-CoaCustomError -subject "New-MSOLUser Error" -body "The user $user cannot be found in MSOL, and has been removed from processing."
        # Need to skip to the next iteration
    }

    if ($upn -like "*onmicrosoft*") {        
        try {
            Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName "$user@$Script:Domain" -ErrorAction Stop
            $writeTo = "Set-MsolUserPrincipalName: Successfully set upn to $user@$Script:Domain"
            $logCode = "Success"
        }
        catch {
            $logCode = "Error"
            "Set-MsolUserPrincipalName: $upn Error: $_" | Tee-Object -Variable writeTo
        }    
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode        

        $upn = Get-MsolUser -UserPrincipalName "$user@$Script:Domain" | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
        $script:upnArray.Add($SingleUser)
        Return;
    }
    elseif ($upn -like "*$Script:Domain") {
        $script:upnArray.Add($SingleUser)
        $writeTo = "Get-MsolUserPrincipalName: UPN need not be set"
        $logCode = "Get"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        Return;
    }
    else {
        $writeTo = "The user $user cannot be found in MSOL, and has been removed from processing."
        $logCode = "Error"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        Show-CoaCustomError -subject "New-MSOLUser Error" -body "The user $user cannot be found in MSOL, and has been removed from processing."
        Return;
    }
}

function SetLicense {
    param([string]$upn, [string]$Licenses, [System.Object]$O365License, [string]$sys_created_by, [string]$licenseDisplayName);

    $Location = (Get-MSOLUser -UserPrincipalName $upn).UsageLocation
    if (!$Location) {
        try {
            Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
            $writeTo = "Set-MsolUser: Successfully set location for $upn"
            $logCode = "Success"
        }
        catch {
            $logCode = "Error"
            "Set-MsolUser: $upn Error: $_" | Tee-Object -Variable writeTo
        }    
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }

    $LicenseLineItem = @()
    $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId

    if ($LicenseLineItem -contains $Script:CoaSkuInformationWorkers -or $LicenseLineItem -contains $Script:CoaSkuFirstlineWorkers) {
        $writeTo = "Get-MsolUserLicense: $upn already contains: $LicenseLineItem"
        $logCode = "Get" 
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }
    else {
        try {
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $Licenses -LicenseOptions $O365License -ErrorAction Stop
            $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
            Show-CoaCustomError -subject "Account created for $upn" -body "An Office 365 account has been created for $upn. The account has been assigned a $licenseDisplayName license. This was requested by $sys_created_by on $sys_created_on. Please reach out to IT Services if you find an issue."
            $writeTo = "Set-MsolUserLicense: Successfully added $LicenseLineItem to $upn"
            $logCode = "Success"
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        }
        catch {
            $logCode = "Error"
            "Set-MsolUserLicense: $upn Error: $_" | Tee-Object -Variable writeTo
            $logLineTime = (Get-Date).ToString()
            WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode

        }
        if ($LicenseLineItem -contains $Script:CoaSkuFirstlineWorkers ) {
            try {
                Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $Script:CoaSkuExoArchive
                $writeTo = "Set-MsolUserLicense: identity $upn Archive"
                $logCode = "Success"
                $logLineTime = (Get-Date).ToString()
                WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
            }
            catch {
                $logCode = "Error"
                "Set-Mailbox: FAILED adding Archive license TO $UPN" | Tee-Object -Variable writeTo
                $logLineTime = (Get-Date).ToString()
                WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
            }    
        }
    }
}

function Set-CoaExoAttributes {
    Param (
        [parameter(
            Position = 0,
            ValueFromPipeline = $false)]
        [System.Collections.Generic.List[UserObject]]
        $UserList = $global:CoaUsersToWorkThrough,
        [parameter(
            Position = 1,
            ValueFromPipeline = $true)]    
        [UserObject]
        $SingleUser
    )
    $script:upnArray = [System.Collections.Generic.List[UserObject]]::new();
    $script:standardUsers = [System.Collections.Generic.List[System.Object]]::new();
    $script:basicUsers = [System.Collections.Generic.List[System.Object]]::new();
    [System.Object]$O365License;
    OpenLog
    
    foreach ($user in $UserList) {
        Set-ValidateUsersUpn -SingleUser $user
        if ($user.license -eq $Script:StandardLicenseName) {
            $script:standardUsers.Add($user);
        }
        if ($user.license -eq $Script:BasicLicenseName) {
            $script:basicUsers.Add($user);
        }
    }
    
    foreach ($upnFO in $script:upnArray) {
        $upn = $upnFO.samAccountName.ToString()
        $upn += "@"
        $upn += $Script:Domain
        $local:baseUpn = $upnFO.samAccountName.ToString()
        Add-CoaWriteToLog -writeTo "$local:baseUpn`t$upn" -logCode "Info" -FileName "NewUserScript"
        :outer
        foreach ($user in $script:standardUsers) {
            $samAccountName = $user.samAccountName.ToString()             
            $sys_created_by = $env:USERNAME.ToString()
            if ($samAccountName -eq $local:baseUpn) {
                $licenseDisplayName = "Standard"
                $pack = standardLicensePack
                $license = $Script:CoaSkuInformationWorkers
                SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
                $global:CoaUsersToWorkThrough.Remove($upnFO);
                break :outer
            }
        }
        foreach ($user in $script:basicUsers) {
            $samAccountName = $user.samAccountName.ToString() 
            $sys_created_by = $env:USERNAME.ToString();
            if ($samAccountName -eq $local:baseUpn) {
                $licenseDisplayName = "Basic"
                $pack = basicLicensePack
                $license = $Script:CoaSkuFirstlineWorkers
                SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
                $global:CoaUsersToWorkThrough.Remove($upnFO);
                break :outer
            }
        }
    } 
}
#endregion
#region: New-CoaUser
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
        $user.License = $Script:BasicLicenseName
    }
    else {
        $user.License = $Script:StandardLicenseName
    }
    $user.samAccountName = $samAccountName
    $global:CoaUsersToWorkThrough.Add($user)
    Write-Output $global:CoaUsersToWorkThrough
}
function Remove-CoaUser {
    param(
        [parameter(Mandatory = $true,
            Position = 0)]
        [string]$SamAccountName
    )    
    #region
    [string]$upn = ""
    $arrayFromGet = @()
    $arrayFromGet += Get-MsolUser -SearchString $SamAccountName | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
    if ($arrayFromGet.Count -eq 1) {
        $upn = $arrayFromGet[0]
        $writeTo = "Get-MsolUser`t$SamAccountName`tSearchString returned: $upn"
        $logCode = "Get"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
    }
    else {
        if ($arrayFromGet.Count -gt 1) {
            $errMsg = "Either the samAccountName was empty, or the search returned more than one value."
        }
        else {
            $errMsg = "The user $SamAccountName cannot be found in MSOL, and has been removed from processing."
        }
        $writeTo = $errMsg
        $logCode = "Else"
        $logLineTime = (Get-Date).ToString()
        WriteToLog -logLineTime $logLineTime -writeTo $writeTo -logCode $logCode
        Show-CoaCustomError -subject "New-MSOLUser Error" -body "The user $user cannot be found in MSOL, and has been removed from processing."
        # Need to skip to the next iteration
        break
    }
    #endregion
    $LicenseLineItem
    $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
    Add-CoaWriteToLog -writeTo "Get-MsolUser`t$upn`t$LicenseLineItem" -logCode "Success" -FileName "RemoveUserScript"
    try {
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $LicenseLineItem -ErrorAction Stop -ErrorVariable err
        Add-CoaWriteToLog -writeTo "Set-MsolUserLicense`t$upn`t$LicenseLineItem" -logCode "Success" -FileName "RemoveUserScript"
    }
    catch {
        Add-CoaWriteToLog -writeTo "Set-MsolUserLicense`t$upn`t$licenses`t$err" -logCode "Error" -FileName "RemoveUserScript"
    }
}
function Clear-CoaUser {
    $global:CoaUsersToWorkThrough.Clear()
}
#endregion
Set-CoaVariables
Export-ModuleMember -Function Set-CoaMailboxConfiguration, Set-CoaExchangeAttributes, Set-CoaExoAttributes, New-CoaUser, Get-CoaVariables, Set-CoaVariables, Clear-CoaUser, Remove-CoaUser