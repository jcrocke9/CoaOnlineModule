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
<#
    .Synopsis
    Sets new mailbox accounts up with the standard policies of COA

    .Description
    Sets new mailbox accounts up with: 
#>
Export-ModuleMember -Function Set-CoaMailboxConfiguration 