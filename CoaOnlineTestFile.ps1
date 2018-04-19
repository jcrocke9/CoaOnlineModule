if ($Global:CoaUsersToWorkThrough)
{
    $Global:CoaUsersToWorkThrough.Clear()
}
$Global:Error.Clear();
# Remove-Module -Name CoaOnlineModule -Verbose -Force -ErrorAction SilentlyContinue
# Import-Module -Name C:\alex\CoaOnlineModule -Verbose -Force
# Set-CoaVariables -Domain "joecrockett.io" -CoaSkuInformationWorkers "joecrockett:DEVELOPERPACK" -CoaSkuFirstlineWorkers "joecrockett:DEVELOPERPACK" -CoaSkuExoArchive "" -CoaSkuExoAtp "" -authOrig "CN=O365 Administrator,OU=Admin,OU=Generics,OU=O365,DC=alexfstest,DC=net"

New-CoaUser -SamAccountName Module.Test -Firstline | Set-CoaExchangeAttributes | Set-CoaExoAttributes

# New-CoaUser -SamAccountName test.user3 -Firstline
# Set-CoaExchangeAttributes -UserList $CoaUsersToWorkThrough
# Set-CoaExoAttributes -UserList $CoaUsersToWorkThrough

# Remove-CoaUser -SamAccountName Module.Test1
# Remove-CoaUser -CommaSeparatedSamAccountNames "test.user3","test.test05"
