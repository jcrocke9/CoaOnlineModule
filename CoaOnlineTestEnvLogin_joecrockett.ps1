Import-Module MSOnline, C:\Alex\CoaOnlineModule
#Connect-MsolService
#Connect-EXOPSSession
$credential = Get-Credential
Connect-MsolService -Credential $credential
$session = New-PSSession -ConfigurationName Microsoft.Exchange -Credential $credential -ConnectionUri 'https://ps.outlook.com/powershell/' -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber

Set-CoaVariables -Domain "joecrockett.io" -CoaSkuInformationWorkers "joecrockett:DEVELOPERPACK" -CoaSkuFirstlineWorkers "joecrockett:DEVELOPERPACK" -CoaSkuExoArchive "" -CoaSkuExoAtp "" -authOrig "CN=O365 Administrator,OU=Admin,OU=Generics,OU=O365,DC=alexfstest,DC=net" -CoaSkuInformationWorkersDisabledPlans "FORMS_PLAN_E5" -CoaSkuEms ""
