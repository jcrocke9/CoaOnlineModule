# What it does
* License Microsoft Online users & mailboxes
* Remove users licenses
* Set common Exchange Online policies to mailboxes
# Why
Using multi-factor authentication with Exchange Online requires launching an interactive session in the [Exchange Online Remote PowerShell Module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps). For ease of use, I've packaged the series of scripts into a module which can be co-loaded and used within the session seamlessly.
# How to start
1. Launch PowerShell as administrator
2. Run the install cmdlet with the name of the module
```PowerShell
    Install-Module -Name "CoaOnlineModule"
```
3. Close the shell
4. Open Exchange Online Remote PowerShell Module
5. Follow steps in the [wiki](https://github.com/jcrocke9/CoaOnlineModule/wiki), under **How to start everyday**
