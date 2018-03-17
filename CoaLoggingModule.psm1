<#
    .Synopsis
    Logs the events of New and Set verbs to a file.

    .Description
    Changes happen with new or set verbs. This module logs those changes to a file.

    .Parameter writeTo
    Logs the event.

    .Parameter logCode
    Logs the COA custom status code: Success, Error

    .Parameter FilePath
    The location of the text log file. Include the last backslash when writing.

    .Parameter FileName
    The descriptive name of the log file. Do not include the file extenstion. It will be a .log.

    .Example
    # Write data to the log file
    Add-CoaWriteToLog -writeTo "Logged an example line." -logCode "Success" -FilePath "\\Its01\deptfiles\Its\EmailTeam\Logs\" -FileName "exampleLog"
#>
function Add-CoaWriteToLog {
    param(
        [string]$writeTo, 
        [string]$logCode,
        [string]$FilePath = "\\Its01\deptfiles\Its\EmailTeam\Logs\",
        [parameter(Mandatory = $true)][string]$FileName = "noNameLog"
    )
    $logLineTime = (Get-Date).ToString() 
    $logFileDate = Get-Date -UFormat "%Y%m%d"
    $logLineInfo = "`t$([Environment]::UserName)`t$([Environment]::MachineName)`t"
    $logLine = $null
    $logLine = $logLineTime
    $logLine += $logLineInfo
    $logLine += $logCode; $logLine += "`t"
    $logLine += $writeTo
    $logLine | Out-File -FilePath "$FilePath$FileName`_$logFileDate.log" -Append -NoClobber
    Clear-Variable logLine -Scope global
    Clear-Variable writeTo -Scope global
    Clear-Variable logLineTime -Scope global
    Clear-Variable logCode -Scope global
}
Export-ModuleMember -Function Add-CoaWriteToLog