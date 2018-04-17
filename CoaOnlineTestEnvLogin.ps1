Import-Module MSOnline, C:\Alex\CoaOnlineModule\CoaOnlineModule.psm1
Connect-MsolService
Connect-EXOPSSession
Set-CoaVariables -Domain "qa01alexandriava.net" -CoaSkuInformationWorkers "qa01alexandriava:ENTERPRISEPACK"`
    -CoaSkuFirstlineWorkers "qa01alexandriava:DESKLESSPACK" -CoaSkuExoArchive "" -CoaSkuExoAtp ""
Get-Command -Module C:\Alex\CoaOnlineModule\CoaOnlineModule.psm1