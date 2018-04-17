
New-CoaUser -SamAccountName test.test05 -Firstline 
New-CoaUser -SamAccountName test.user3 -Firstline
Set-CoaExchangeAttributes -UserList $CoaUsersToWorkThrough
Set-CoaExoAttributes -UserList $CoaUsersToWorkThrough

# Remove-CoaUser -CommaSeparatedSamAccountNames "test.user3","test.test05"