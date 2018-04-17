Import-Module MSOnline, C:\Alex\CoaOnlineModule\CoaOnlineModule.psm1
#Connect-MsolService
#Connect-EXOPSSession
$credential = Get-Credential
Connect-MsolService -Credential $credential
$session = New-PSSession -ConfigurationName Microsoft.Exchange -Credential $credential -ConnectionUri 'https://ps.outlook.com/powershell/' -Authentication Basic -AllowRedirection 
Import-PSSession $session -AllowClobber

Set-CoaVariables -Domain "qa01alexandriava.net" -CoaSkuInformationWorkers "qa01alexandriava:ENTERPRISEPACK" -CoaSkuFirstlineWorkers "qa01alexandriava:DESKLESSPACK" -CoaSkuExoArchive "" -CoaSkuExoAtp "" -authOrig "CN=O365 Administrator,OU=Admin,OU=Generics,OU=O365,DC=alexfstest,DC=net"