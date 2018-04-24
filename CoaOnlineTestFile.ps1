if ($Global:CoaUsersToWorkThrough)
{
    $Global:CoaUsersToWorkThrough.Clear()
}
$Global:Error.Clear();
# Remove-Module -Name CoaOnlineModule -Verbose -Force -ErrorAction SilentlyContinue
# Import-Module -Name C:\alex\CoaOnlineModule -Verbose -Force
# Set-CoaVariables -Domain "joecrockett.io" -CoaSkuInformationWorkers "joecrockett:DEVELOPERPACK" -CoaSkuFirstlineWorkers "joecrockett:DEVELOPERPACK" -CoaSkuExoArchive "" -CoaSkuExoAtp "" -authOrig "CN=O365 Administrator,OU=Admin,OU=Generics,OU=O365,DC=alexfstest,DC=net"

# New-CoaUser -SamAccountName Module.Test5 -Firstline | Set-CoaExchangeAttributes | Set-CoaExoAttributes
# New-CoaUser -SamAccountName Module.Test6 | Set-CoaExchangeAttributes | Set-CoaExoAttributes

New-CoaUser -SamAccountName module.test5
New-CoaUser -SamAccountName module.test6 -Firstline
Set-CoaExchangeAttributes -UserList $CoaUsersToWorkThrough
Set-CoaExoAttributes -UserList $CoaUsersToWorkThrough

# Remove-CoaUser -SamAccountName Module.Test5
# Remove-CoaUser -CommaSeparatedSamAccountNames "module.test5","module.test6"
