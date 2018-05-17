#Require -Version 5.1

using namespace System;
using namespace System.Text;
using namespace System.Diagnostics;
using namespace System.Linq;
using namespace System.Collections.Generic;
Import-Module ActiveDirectory;
Import-Module CoaLoggingModule;
#region: Variables
<#
    .Synopsis
    Use this cmdlet to set new variables over the default ones.

    .Description
    Sets the script variables that are in common use across the functions
#>
function Set-CoaVariables
{
    [CmdletBinding()]
    Param (
        # Default -120. The number of days backward (negative) in time to gather new accounts and ensure they are in-policy.
        [Parameter()]
        [int]
        $NumberOfDays = -120,
        [string]$RoleAssignmentPolicy = "COA Default Role Assignment Policy",
        [string]$ClientAccessPolicyName = "COAOWAMailboxPolicy",
        [int]$LitigationHoldDuration = 730,
        [string]$ExchangeOnlineAdminAccount = "COA Administrator",
        [string]$authOrig = "CN=O365 Administrator,OU=Admin,OU=Generics,OU=O365,DC=alexgov,DC=net",
        [string]$RetentionPolicyE3 = "COA Policy",
        [string]$RetentionPolicyK1 = "COA F1 Policy",
        [string]$RetentionPolicyTermOfficial = "Termination Retention Policy",
        [string]$RetentionPolicyDeptHead = "COA Department Head Policy",
        [string]$CoaSkuInformationWorkers = "ALEXANDRIAVA1:ENTERPRISEPACK_GOV",
        [string]$CoaSkuFirstlineWorkers = "ALEXANDRIAVA1:DESKLESSPACK_GOV",
        [string[]]$CoaSkuFirstlineWorkersDisabledPlans,
        [string[]]$CoaSkuInformationWorkersDisabledPlans = "RMS_S_ENTERPRISE_GOV",
        [string]$CoaSkuExoArchive = "ALEXANDRIAVA1:EXCHANGEARCHIVE_ADDON_GOV",
        [string]$CoaSkuExoAtp = "ALEXANDRIAVA1:ATP_ENTERPRISE_GOV",
        [string]$CoaSkuEms = "ALEXANDRIAVA1:EMS",
        [string]$standardLicenseName = "emailStandard_createAlexID",
        [string]$basicLicenseName = "emailBasic_createAlexID",
        [string]$Domain = "alexandriava.gov"
    )
    $Script:NumberOfDays = $NumberOfDays;
    $Script:RoleAssignmentPolicy = $RoleAssignmentPolicy
    $Script:ClientAccessPolicyName = $ClientAccessPolicyName
    $Script:LitigationHoldDuration = $LitigationHoldDuration
    $Script:ExchangeOnlineAdminAccount = $ExchangeOnlineAdminAccount
    $Script:authOrig = $authOrig
    $Script:RetentionPolicyE3 = $RetentionPolicyE3
    $Script:RetentionPolicyK1 = $RetentionPolicyK1
    $Script:RetentionPolicyTermOfficial = $RetentionPolicyTermOfficial
    $Script:RetentionPolicyDeptHead = $RetentionPolicyDeptHead
    $Script:CoaSkuInformationWorkers = $CoaSkuInformationWorkers
    $Script:CoaSkuFirstlineWorkers = $CoaSkuFirstlineWorkers
    $Script:CoaSkuFirstlineWorkersDisabledPlans = $CoaSkuFirstlineWorkersDisabledPlans
    $Script:CoaSkuInformationWorkersDisabledPlans = $CoaSkuInformationWorkersDisabledPlans
    $Script:CoaSkuExoArchive = $CoaSkuExoArchive
    $Script:CoaSkuExoAtp = $CoaSkuExoAtp
    $Script:CoaSkuEms = $CoaSkuEms
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
function Get-CoaVariables
{
    [CmdletBinding()]
    $Private:CoaVariables = [ordered]@{
        NumberOfDays                          = $Script:NumberOfDays;
        RoleAssignmentPolicy                  = $Script:RoleAssignmentPolicy;
        ClientAccessPolicyName                = $Script:ClientAccessPolicyName;
        LitigationHoldDuration                = $Script:LitigationHoldDuration;
        ExchangeOnlineAdminAccount            = $Script:ExchangeOnlineAdminAccount;
        authOrig                              = $Script:authOrig
        RetentionPolicyE3                     = $Script:RetentionPolicyE3;
        RetentionPolicyK1                     = $Script:RetentionPolicyK1;
        RetentionPolicyTermOfficial           = $Script:RetentionPolicyTermOfficial;
        RetentionPolicyDeptHead               = $Script:RetentionPolicyDeptHead;
        CoaSkuInformationWorkers              = $Script:CoaSkuInformationWorkers;
        CoaSkuFirstlineWorkers                = $Script:CoaSkuFirstlineWorkers;
        CoaSkuInformationWorkersDisabledPlans = $Script:CoaSkuInformationWorkersDisabledPlans
        CoaSkuFirstlineWorkersDisabledPlans   = $Script:CoaSkuFirstlineWorkersDisabledPlans
        CoaSkuExoArchive                      = $Script:CoaSkuExoArchive;
        CoaSkuExoAtp                          = $Script:CoaSkuExoAtp;
        CoaSkuEms                             = $Script:CoaSkuEms
        standardLicenseName                   = $Script:StandardLicenseName;
        basicLicenseName                      = $Script:BasicLicenseName;
        Domain                                = $Script:Domain;
    }
    Write-Output $Private:CoaVariables
}
#endregion
#region: Mailbox Configuration
function Set-CoaMailboxConfiguration
{
    [CmdletBinding()]
    Param (
        # Default -120. The number of days backward (negative) in time to gather new accounts and ensure they are in-policy.
        [Parameter()]
        [int]
        $NumberOfDays = $Script:NumberOfDays,
        [string]$FileName = "MailboxConfiguration",
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
    Add-CoaWriteToLog -writeTo $writeTo -logCode $logCode -FilePath $FileName

    $UserList = @()
    Clear-Variable UserList

    $Global:ErrorActionPreference = "Stop"
    $UserList = Get-Mailbox -ResultSize unlimited -Filter {(ArchiveStatus -eq $None)} | Where-Object {$_.WhenCreated -gt (get-date).AddDays($NumberOfDays)} | Select-Object -ExpandProperty userPrincipalName
    Write-Output "User count found: " + $UserList.Count
    $logCode = "Get"
    $writeTo = "User count found: " + $UserList.Count
    foreach ($upn in $UserList)
    {
        try
        {
            Add-MailboxPermission -identity $upn -User $ExchangeOnlineAdminAccount -AccessRights fullaccess -InheritanceType all -AutoMapping $false
            $writeTo = "Add-MailboxPermission: Successfully added $upn mailbox permission for $ExchangeOnlineAdminAccount"
            $logCode = "Success"
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch
        {
            Write-Output $_.Exception.Message
            $writeTo = "Add-MailboxPermission: FAILED added $upn mailbox permission for COA Admin"
            $logCode = "Error"
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        try
        {
            Set-Clutter -Identity $upn -Enable $false
            $writeTo = "Set-Clutter: Successfully set mailbox $upn clutter to false"
            $logCode = "Success"
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        catch
        {
            Write-Output $_.Exception.Message
            $writeTo = "Set-Clutter: FAILED set mailbox $upn clutter to false"
            $logCode = "Error"
            Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
        }
        $RecipientTypeDetails = (Get-Mailbox -Identity $upn).RecipientTypeDetails
        if ($RecipientTypeDetails -eq "UserMailbox")
        {
            try
            {
                Enable-Mailbox -identity $upn -Archive
                $writeTo = "Enable-Mailbox: identity $upn Archive"
                $logCode = "Success"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            catch
            {
                Write-Output $_.Exception.Message
                $writeTo = "Enable-Mailbox: identity $upn Archive"
                $logCode = "Error"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            try
            {
                Set-Mailbox -identity $upn -LitigationHoldEnabled $True -LitigationHoldDuration $LitigationHoldDuration -RoleAssignmentPolicy $RoleAssignmentPolicy
                $writeTo = "Set-Mailbox: Successfully set mailbox $upn litigation hold and assignment policy"
                $logCode = "Success"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            catch
            {
                Write-Output $_.Exception.Message
                $writeTo = "Set-Mailbox: FAILED set mailbox $upn litigation hold and assignment policy"
                $logCode = "Error"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            try
            {
                Set-CASMailbox -Identity $upn -OWAMailboxPolicy $ClientAccessPolicyName
                Set-CASMailbox -Identity $upn -PopEnabled $false
                Set-CASMailbox -Identity $upn -ImapEnabled $false
                $writeTo = "Set-CASMailbox: Successfully set mailbox $upn client access permissions"
                $logCode = "Success"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            catch
            {
                Write-Output $_.Exception.Message
                $writeTo = "Set-CASMailbox: FAILED set mailbox $upn client access permissions"
                $logCode = "Error"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
            }
            try
            {
                $upn = (Get-MsolUser -SearchString $upn).UserPrincipalName
                $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
            }
            catch
            {
                Write-Output $_.Exception.Message
                $writeTo = "Get-MSOLUser -UserPrincipalName $upn FAILED"
                $logCode = "Error"
                Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                Continue
            }
            # E3
            if ($LicenseLineItem -contains $CoaSkuInformationWorkers)
            {
                try
                {
                    Set-Mailbox -Identity $upn -RetentionPolicy $RetentionPolicyE3
                    $writeTo = "Set-Mailbox: Successfully set mailbox $upn policy to $RetentionPolicyE3"
                    $logCode = "Success"
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                }
                catch
                {
                    Write-Output $_.Exception.Message
                    $writeTo = "Set-Mailbox: FAILED set mailbox $upn policy to $RetentionPolicyE3"
                    $logCode = "Error"
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                }
            }
            # K1
            if ($LicenseLineItem -contains $CoaSkuFirstlineWorkers)
            {
                try
                {
                    Set-Mailbox -Identity $upn -RetentionPolicy $RetentionPolicyK1
                    $writeTo = "Set-Mailbox: Successfully set mailbox $upn policy to $RetentionPolicyK1"
                    $logCode = "Success"
                    Add-CoaWriteToLog -FileName $FileName -writeTo $writeTo -logCode $logCode
                }
                catch
                {
                    Write-Output $_.Exception.Message
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
    $Global:ErrorActionPreference = "Continue"
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
}
#endregion
#region: Sets Active Directory attributes for Exchange Online
function OpenLog
{
    $logCode = "Start"
    $writeTo = "Starting New User script"
    $CurrentFileName = "NewUser"
    Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    Return;
}
function QueryAdToValidateUsers
{
    param ([string]$samAccountName)
    $Global:ErrorActionPreference = "Stop"
    try
    {
        $mail = Get-ADUser $samAccountName -Properties mail | Select-Object mail -ExpandProperty mail
        if (!$mail)
        {$mail = "None set"
        }
        $department = Get-ADUser $samAccountName -Properties department | Select-Object department -ExpandProperty department
        $userAccountControl = Get-ADUser $samAccountName -Properties userAccountControl | Select-Object userAccountControl -ExpandProperty userAccountControl
        if ($userAccountControl -eq 514)
        {
            $userAccountControl = "User Disabled"
        }
        elseif ($userAccountControl -eq 512)
        {
            $userAccountControl = "Active"
        }
        elseif ($userAccountControl -eq 66048)
        {
            $userAccountControl = "Has a non-expiring password"
        }
        else
        {
            $problemUsers += $samAccountName
            $writeTo = "Get-ADUser: $samAccountName`t$userAccountControl"
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
            Write-Output "Get-ADUser: samAccountName cannot be found for $samAccountName"
        }
        $writeTo = "$samAccountName`t$userAccountControl`t$mail`t$department";
        $logCode = "Start"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    catch
    {
        $problemUsers.Add($samAccountName);
        if ($samAccountName.Length -gt 20)
        {
            Write-Output $_.Exception.Message
            Write-Output "The user name $samAccountName is more than 20 characters"
            $writeTo = "The user name $samAccountName is more than 20 characters"
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
        else
        {
            Write-Output $_.Exception.Message
            $writeTo = "Get-ADUser: samAccountName cannot be found for $samAccountName"
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
    $Global:ErrorActionPreference = "continue"
    Return;
}
#endregion
#region: Sets the mail and SMTP attributes, if needed
#region: Private functions for AD
function SetMailAndSmtpAttributes
{
    param([string]$user)
    $mail = Get-ADUser $user -Properties mail | Select-Object mail -ExpandProperty mail
    if (!$mail)
    {
        try
        {
            Set-ADUser -Identity $user -EmailAddress "$user@$Script:Domain" -ErrorAction Stop
            $writeTo = "Set-ADUser: Successfully added email address to $user@$Script:Domain"
            $logCode = "Success"
        }
        catch
        {
            Write-Output $_.Exception.Message
            $logCode = "Error"
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
        }
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    Start-Sleep -Seconds 5
    $mail = Get-ADUser $user -Properties mail | Select-Object mail -ExpandProperty mail
    $SMTP = Get-ADUser -Identity $user -Properties proxyaddresses | Select-Object proxyaddresses -ExpandProperty proxyaddresses
    if (!$SMTP)
    {
        try
        {
            Set-ADUser -Identity $user -Add @{Proxyaddresses = "SMTP:" + $mail } -ErrorAction Stop
            Set-ADUser -Identity $user -Add @{targetAddress = "SMTP:" + $mail } -ErrorAction Stop
            Set-ADUser -Identity $user -Replace @{mailNickname = $user}
            $writeTo = "Set-ADUser: Successfully set SMTP address to SMTP:$mail"
            $logCode = "Success"
        }
        catch
        {
            Write-Output $_.Exception.Message
            $logCode = "Error"
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
        }
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }

    Return;
}
function SetLicenseAttributeE3
{
    param([string]$user)
    $extensionAttribute13 = Get-ADUser -Identity $user -Properties extensionAttribute13 | Select-Object -ExpandProperty extensionAttribute13
    if (!$extensionAttribute13)
    {
        try
        {
            Set-ADUser -Identity $user -Add @{extensionAttribute13 = "E3"} -ErrorAction Stop
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = E3"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        }
        catch
        {
            Write-Output $_.Exception.Message
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
    else
    {
        try
        {
            Set-ADUser -Identity $user -Replace @{extensionAttribute13 = "E3"} -ErrorAction Stop
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = E3"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        }
        catch
        {
            Write-Output $_.Exception.Message
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
}
function SetLicenseAttributeK1
{
    param([string]$user)
    $extensionAttribute13 = Get-ADUser -Identity $user -Properties extensionAttribute13 | Select-Object -ExpandProperty extensionAttribute13
    if (!$extensionAttribute13)
    {
        try
        {
            Set-ADUser -Identity $user -Add @{extensionAttribute13 = "K1"} -ErrorAction Stop
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = K1"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        }
        catch
        {
            Write-Output $_.Exception.Message
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
    else
    {
        try
        {
            Set-ADUser -Identity $user -Replace @{extensionAttribute13 = "K1"} -ErrorAction Stop
            $writeTo = "Set-ADUser`t$user`tSet extensionAttribute13 = K1"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        }
        catch
        {
            Write-Output $_.Exception.Message
            "Set-ADUser: $user Error: $_" | Tee-Object -Variable writeTo
            $logCode = "Error"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
}
#endregion

function Set-CoaExchangeAttributes
{
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
    [CmdletBinding()]
    Param (
        [parameter(
            Position = 0,
            ValueFromPipeline = $false)]
        [System.Collections.Generic.List[UserObject]]
        $UserList = $Global:CoaUsersToWorkThrough,
        [parameter(
            Position = 1,
            ValueFromPipeline = $true)]
        [UserObject]
        $SingleUser
    )
    $problemUsers = [System.Collections.Generic.List[System.Object]]::new();
    OpenLog
    if ($SingleUser)
    {
        QueryAdToValidateUsers -samAccountName $SingleUser.samAccountName
        # Remove problem users
        SetMailAndSmtpAttributes -user $SingleUser.samAccountName
        if ($SingleUser.License -eq $Script:BasicLicenseName)
        {
            SetLicenseAttributeK1 -user $SingleUser.samAccountName
        }
        else
        {
            SetLicenseAttributeE3 -user $SingleUser.samAccountName
        }
        return $SingleUser
    }
    else
    {
        foreach ($samAccountName in $UserList )
        {
            QueryAdToValidateUsers -samAccountName $samAccountName.samAccountName
        }

        foreach ($problemUser in $problemUsers)
        {
            $foundProblemUser = $UserList.Find( {param($u) $u.samAccountName -match $problemUser})
            $UserList.Remove($foundProblemUser);
        }

        foreach ($user in $UserList)
        {
            SetMailAndSmtpAttributes -user $user.samAccountName
        }

        foreach ($user in $UserList)
        {
            if ($user.License -eq $Script:BasicLicenseName)
            {
                SetLicenseAttributeK1 -user $user.samAccountName
            }
            else
            {
                SetLicenseAttributeE3 -user $user.samAccountName
            }
        }
    }
}
#endregion
#region: Sets ExO Attributes
#region: Private functions for ExO
function basicLicensePack
{
    $O365License = New-MsolLicenseOptions -AccountSkuId $Script:CoaSkuFirstlineWorkers -DisabledPlans $Script:CoaSkuFirstlineWorkersDisabledPlans
    Return $O365License;
}
function standardLicensePack
{
    $O365License = New-MsolLicenseOptions -AccountSkuId $Script:CoaSkuInformationWorkers -DisabledPlans $Script:CoaSkuInformationWorkersDisabledPlans
    Return $O365License;
}
function Set-ValidateUsersUpn
{
    param([UserObject]$SingleUser);
    $user = $SingleUser.samAccountName;
    $arrayFromGet = @()
    $arrayFromGet += Get-MsolUser -SearchString $user | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
    if ($arrayFromGet.Count -eq 1)
    {
        $upn = $arrayFromGet[0]
        $writeTo = "Get-MsolUser`t$user`tSearchString returned: $upn"
        $logCode = "Get"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    else
    {
        if ($arrayFromGet.Count -gt 1)
        {
            $errMsg = "Either the samAccountName was empty, or the search returned more than one value."
        }
        else
        {
            $errMsg = "The user $user cannot be found in MSOL, and has been removed from processing."
        }
        $writeTo = $errMsg
        $logCode = "Else"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        Write-Output $errMsg
    }

    if ($upn -like "*onmicrosoft*")
    {
        try
        {
            Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName "$user@$Script:Domain" -ErrorAction Stop
            $writeTo = "Set-MsolUserPrincipalName: Successfully set upn to $user@$Script:Domain"
            $logCode = "Success"
        }
        catch
        {
            Write-Output $_.Exception.Message
            $logCode = "Error"
            "Set-MsolUserPrincipalName: $upn Error: $_" | Tee-Object -Variable writeTo
        }
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        $upn = Get-MsolUser -UserPrincipalName "$user@$Script:Domain" | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
        $script:upnArray.Add($SingleUser)
        Return;
    }
    elseif ($upn -like "*$Script:Domain")
    {
        $script:upnArray.Add($SingleUser)
        $writeTo = "Get-MsolUserPrincipalName: UPN need not be set"
        $logCode = "Get"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        Return;
    }
    else
    {
        $writeTo = "The user $user cannot be found in MSOL, and has been removed from processing."
        $logCode = "Error"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        Write-Output $writeTo
        Exit;
    }
}
function SetLicense
{
    param([string]$upn, [string]$Licenses, [System.Object]$O365License, [string]$sys_created_by, [string]$licenseDisplayName);

    $Location = (Get-MSOLUser -UserPrincipalName $upn).UsageLocation
    if (!$Location)
    {
        try
        {
            Set-MsolUser -UserPrincipalName $upn -UsageLocation "US" -ErrorAction Stop
            $writeTo = "Set-MsolUser: Successfully set location for $upn"
            $logCode = "Success"
        }
        catch
        {
            Write-Output $_.Exception.Message
            $logCode = "Error"
            "Set-MsolUser: $upn Error: $_" | Tee-Object -Variable writeTo
        }
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }

    $LicenseLineItem = @()
    $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId

    if ($LicenseLineItem -contains $Script:CoaSkuInformationWorkers -or $LicenseLineItem -contains $Script:CoaSkuFirstlineWorkers)
    {
        $writeTo = "Get-MsolUserLicense: $upn already contains: $LicenseLineItem"
        $logCode = "Get"
        $CurrentFileName = "NewUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    else
    {
        try
        {
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $Licenses -LicenseOptions $O365License -ErrorAction Stop
            $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
            Write-Output "An Office 365 account has been created for $upn. The account has been assigned a $licenseDisplayName license."
            $writeTo = "Set-MsolUserLicense: Successfully added $LicenseLineItem to $upn"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
        catch
        {
            Write-Output $_.Exception.Message
            $logCode = "Error"
            "Set-MsolUserLicense: $upn Error: $_" | Tee-Object -Variable writeTo
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

        }
        if ($LicenseLineItem -contains $Script:CoaSkuFirstlineWorkers )
        {
            try
            {
                Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $Script:CoaSkuExoArchive -ErrorAction Stop
                $writeTo = "Set-MsolUserLicense: identity $upn Archive"
                $logCode = "Success"
                $CurrentFileName = "NewUser"
                Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
            }
            catch
            {
                Write-Output $_.Exception.Message
                $logCode = "Error"
                "Set-Mailbox: FAILED adding Archive license TO $UPN" | Tee-Object -Variable writeTo
                $CurrentFileName = "NewUser"
                Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
            }
        }
        try
        {
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $Script:CoaSkuEms -ErrorAction Stop
            $writeTo = "Set-MsolUserLicense: Successfully added EMS to $upn"
            $logCode = "Success"
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
        catch
        {
            $logCode = "Error"
            "Set-Mailbox: FAILED adding EMS license TO $UPN" | Tee-Object -Variable writeTo
            $CurrentFileName = "NewUser"
            Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        }
    }
}
#endregion
<#
    .SYNOPSIS
    Sets the Exchange Online Attributes after a sync
    .DESCRIPTION
    Once a new user object has been created, and a sync has been performed, this cmdlet will add the attributes needed in Exchange Online
    .PARAMETER UserList
    Uses the NewCoaUser objects that were created prior to running the cmdlet;
    .PARAMETER SingleUser
    Uses the NewCoaUser objects that were created with the pipeline cmdlet;
    .EXAMPLE
    Set-CoaExoAttributes -UserList $CoaUsersToWorkThrough
    .EXAMPLE
    New-CoaUser test.user | Set-CoaExchangeAttributes | Set-CoaExoAttributes
#>
function Set-CoaExoAttributes
{
    [CmdletBinding()]
    Param (
        [parameter(
            Position = 0,
            ValueFromPipeline = $false)]
        [System.Collections.Generic.List[UserObject]]
        $UserList = $Global:CoaUsersToWorkThrough,
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
    if ($SingleUser)
    {
        Set-ValidateUsersUpn -SingleUser $SingleUser
        $upn = $SingleUser.samAccountName.ToString()
        $upn += "@"
        $upn += $Script:Domain
        $local:baseUpn = $SingleUser.samAccountName.ToString()
        Add-CoaWriteToLog -writeTo "$local:baseUpn`t$upn" -logCode "Info" -FileName "NewUser"
        if ($SingleUser.License -eq $Script:BasicLicenseName)
        {
            $licenseDisplayName = "Basic"
            $pack = basicLicensePack
            $license = $Script:CoaSkuFirstlineWorkers
            SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
            $Global:CoaUsersToWorkThrough.Remove($SingleUser);
        }
        else
        {
            $licenseDisplayName = "Standard"
            $pack = standardLicensePack
            $license = $Script:CoaSkuInformationWorkers
            SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
            $Global:CoaUsersToWorkThrough.Remove($SingleUser);
        }
    }
    else
    {
        foreach ($user in $UserList)
        {
            Set-ValidateUsersUpn -SingleUser $user
            if ($user.license -eq $Script:StandardLicenseName)
            {
                $script:standardUsers.Add($user);
            }
            if ($user.license -eq $Script:BasicLicenseName)
            {
                $script:basicUsers.Add($user);
            }
        }
        foreach ($upnFO in $script:upnArray)
        {
            $upn = $upnFO.samAccountName.ToString()
            $upn += "@"
            $upn += $Script:Domain
            $local:baseUpn = $upnFO.samAccountName.ToString()
            Add-CoaWriteToLog -writeTo "$local:baseUpn`t$upn" -logCode "Info" -FileName "NewUser"
            :outer
            foreach ($user in $script:standardUsers)
            {
                $samAccountName = $user.samAccountName.ToString()
                $sys_created_by = $env:USERNAME.ToString()
                if ($samAccountName -eq $local:baseUpn)
                {
                    $licenseDisplayName = "Standard"
                    $pack = standardLicensePack
                    $license = $Script:CoaSkuInformationWorkers
                    SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
                    $Global:CoaUsersToWorkThrough.Remove($upnFO);
                    break :outer
                }
            }
            foreach ($user in $script:basicUsers)
            {
                $samAccountName = $user.samAccountName.ToString()
                $sys_created_by = $env:USERNAME.ToString();
                if ($samAccountName -eq $local:baseUpn)
                {
                    $licenseDisplayName = "Basic"
                    $pack = basicLicensePack
                    $license = $Script:CoaSkuFirstlineWorkers
                    SetLicense -upn $upn -Licenses $license -O365License $pack -sys_created_by $sys_created_by -licenseDisplayName $licenseDisplayName
                    $Global:CoaUsersToWorkThrough.Remove($upnFO);
                    break :outer
                }
            }
        }
        $Global:CoaUsersToWorkThrough.Clear()
    }
}
#endregion
#region: Private Remove-CoaUser functions
function Remove-CoaUserActiveDirectory
{
    param (
        # User to update attributes in active directory
        [Parameter(Mandatory = $true)]
        [string]
        $SamAccountName
    )
    try
    {
        Set-ADUser $SamAccountName -Replace @{authOrig = $Script:authOrig} -ErrorAction Stop
        $writeTo = "Set-ADUser`t$SamAccountName`tReplace authOrig"
        $logCode = "Success"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    catch
    {
        Write-Output $_.Exception.Message
        $writeTo = "Set-ADUser`t$SamAccountName`tCould not replace authOrig`t${$_.Exception.Message}"
        $logCode = "Error"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    try
    {
        Set-ADUser $SamAccountName -Replace @{msExchHideFromAddressLists = "TRUE"} -ErrorAction Stop
        $writeTo = "Set-ADUser`t$SamAccountName`tReplace msExchHideFromAddressLists"
        $logCode = "Success"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    catch
    {
        Write-Output $_.Exception.Message
        $writeTo = "Set-ADUser`t$SamAccountName`tCould not replace msExchHideFromAddressLists`t${$_.Exception.Message}"
        $logCode = "Error"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    try
    {
        Get-ADPrincipalGroupMembership -Identity $SamAccountName | Select-Object samAccountName -ExpandProperty samAccountName | ForEach-Object {
            $groupName = $_
            if ($groupName -ne "Domain Users")
            {
                try
                {
                    Remove-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $groupName -Confirm:$false -ErrorAction Stop
                    $writeTo = "Remove-ADPrincipalGroupMembership:`t$SamAccountName`t$groupName"
                    $logCode = "Success"
                    $CurrentFileName = "RemoveUser"
                    Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
                }
                catch
                {
                    Write-Output $_.Exception.Message
                    $writeTo = "Remove-ADPrincipalGroupMembership:`t$SamAccountName`t$groupName"
                    $logCode = "Error"
                    $CurrentFileName = "RemoveUser"
                    Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
                }
            }
        }
    }
    catch
    {
        Write-Output $_.Exception.Message
        $writeTo = "Set-ADUser`t$SamAccountName`tCould not find for Group Membership`t${$_.Exception.Message}"
        $logCode = "Error"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode

    }
}
function Remove-CoaUserMsol
{
    param (
        # User to update attributes in Microsoft Online
        [Parameter(Mandatory = $true)]
        [string]
        $SamAccountName
    )
    [string]$upn = ""
    $arrayFromGet = @()
    $arrayFromGet += Get-MsolUser -SearchString $SamAccountName | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName
    if ($arrayFromGet.Count -eq 1)
    {
        $upn = $arrayFromGet[0]
        $writeTo = "Get-MsolUser`t$SamAccountName`tSearchString returned: $upn"
        $logCode = "Get"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
    }
    else
    {
        if ($arrayFromGet.Count -gt 1)
        {
            $errMsg = "Either the samAccountName was empty, or the search returned more than one value."
        }
        else
        {
            $errMsg = "The user $SamAccountName cannot be found in MSOL, and has been removed from processing."
        }
        $writeTo = $errMsg
        $logCode = "Else"
        $CurrentFileName = "RemoveUser"
        Add-CoaWriteToLog -FileName $CurrentFileName -writeTo $writeTo -logCode $logCode
        Write-Output $errMsg
        # Need to skip to the next iteration
        break
    }
    $LicenseLineItem
    $LicenseLineItem = (Get-MSOLUser -UserPrincipalName $upn).Licenses.AccountSkuId
    Add-CoaWriteToLog -writeTo "Get-MsolUser`t$upn`t$LicenseLineItem" -logCode "Success" -FileName "RemoveUser"
    try
    {
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $LicenseLineItem -ErrorAction Stop -ErrorVariable err
        Add-CoaWriteToLog -writeTo "Set-MsolUserLicense`tRemove`t$upn`t$LicenseLineItem" -logCode "Success" -FileName "RemoveUser"
    }
    catch
    {
        Write-Output $_.Exception.Message
        Add-CoaWriteToLog -writeTo "Set-MsolUserLicense`tRemove`t$upn`t$licenses`t${$_.Exception.Message}" -logCode "Error" -FileName "RemoveUser"
    }
}
function Set-CoaUserRetentionPolicy
{
    param (
        # User to update attributes in Exchange Online
        [Parameter(Mandatory = $true)]
        [string]
        $SamAccountName
    )
    try
    {
        # Get Mailbox and turn the retention, type, ProhibitSendReceiveQuota, ProhibitSendQuota, IssueWarningQuota
        Set-Mailbox -identity $SamAccountName -RetentionPolicy $Script:RetentionPolicyTermOfficial -ErrorAction Stop
        Add-CoaWriteToLog -writeTo "Set-Mailbox`t$SamAccountName`t$Script:RetentionPolicyTermOfficial" -logCode "Success" -FileName "RemoveUser"
    }
    catch
    {
        Write-Output $_.Exception.Message
        Add-CoaWriteToLog -writeTo "Set-Mailbox`t$SamAccountName`t$Script:RetentionPolicyTermOfficial`t${$_.Exception.Message}" -logCode "Error" -FileName "RemoveUser"
    }
}
#endregion
#region: New-CoaUser
class UserObject
{
    [string]$samAccountName
    [string]$License
}
$Global:CoaUsersToWorkThrough = [System.Collections.Generic.List[UserObject]]::new();
<#
    .SYNOPSIS
    Creates the new user object that can be used by other cmdlets to complete the attributes needed for new user creation.
    .DESCRIPTION
    A new user consists of the SamAccountName and which license you will assign. This cmdlet preps those two things, and can hand them off to Set-CoaExchangeAttributes, Set-CoaExoAttributes to complete the attributes needed for new user creation
    .PARAMETER SamAccountName
    Specifies the samAccountName for the user
    .PARAMETER Firstline
    Switch used to make the user a Firstline worker; default is Enterprise worker
    .EXAMPLE
    New-CoaUser joe.crockett
    .EXAMPLE
    New-CoaUser joe.crockett -Firstline
#>
function New-CoaUser
{
    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true,
            Position = 0)]
        [string]$SamAccountName,
        [switch]$Firstline
    )
    $user = $null
    $user = [UserObject]::new()
    if ($Firstline)
    {
        $user.License = $Script:BasicLicenseName
    }
    else
    {
        $user.License = $Script:StandardLicenseName
    }
    $user.samAccountName = $samAccountName
    $Global:CoaUsersToWorkThrough.Add($user)
    if ($Global:CoaUsersToWorkThrough.Count -eq 1)
    {
        return $user
    }
    else
    {
        Write-Output $Global:CoaUsersToWorkThrough
    }
}
<#
    .SYNOPSIS
    Clears that variable CoaUsersToWorkThrough
    .DESCRIPTION
    When you use the New-CoaUser, it stores the users in a global variable called CoaUsersToWorkThrough.
    .EXAMPLE
    Clear-CoaUser
#>
function Clear-CoaUser
{
    $Global:CoaUsersToWorkThrough.Clear()
}

#endregion
#region: Remove-CoaUser
<#
    .SYNOPSIS
    Removes the user's license & user groups, and updates the attributes on the AD object.
    .DESCRIPTION
    Removes the user's license and updates the following attributes: authOrig, msExchHideFromAddressLists; also removes the groups the user is a member of.
    .PARAMETER SamAccountName
    Specifies the samAccountName for the user
    .PARAMETER CommaSeparatedSamAccountNames
    A comma separated list of samAccountNames to work through
    .EXAMPLE
    # Removes the user license, groups, and updates the attributes for one account
    Remove-CoaUser -SamAccountName test.user
    .EXAMPLE
    # Removes the user license, groups, and updates the attributes for a comma separated list of users
    Remove-CoaUser -CommaSeparatedSamAccountNames "test.user1","test.user2","test.user3"
#>
function Remove-CoaUser
{
    [CmdletBinding(,
        DefaultParameterSetName = "SingleUser")]
    param(
        [parameter(Position = 0,
            Mandatory = $true,
            ParameterSetName = "SingleUser")]
        [string]$SamAccountName,
        # Parameter for taking comma separated samAccountNames
        [Parameter(Position = 0,
            Mandatory = $true,
            ParameterSetName = "MultipleUsers")]
        [string[]]
        $CommaSeparatedSamAccountNames
    )
    if ($PSCmdlet.ParameterSetName -eq "SingleUser")
    {
        Remove-CoaUserActiveDirectory -SamAccountName $SamAccountName
        Remove-CoaUserMsol -SamAccountName $SamAccountName
        Set-CoaUserRetentionPolicy -SamAccountName $SamAccountName
    }
    else
    {
        foreach ($IndividualUser in $CommaSeparatedSamAccountNames)
        {
            Write-Output $IndividualUser
            Remove-CoaUserActiveDirectory -SamAccountName $IndividualUser
            Remove-CoaUserMsol -SamAccountName $IndividualUser
            Set-CoaUserRetentionPolicy -SamAccountName $IndividualUser
        }
    }
}
#endregion
Set-CoaVariables
Export-ModuleMember -Function Set-CoaMailboxConfiguration, Set-CoaExchangeAttributes, Set-CoaExoAttributes, New-CoaUser, Get-CoaVariables, Set-CoaVariables, Clear-CoaUser, Remove-CoaUser
