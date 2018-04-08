Import-Module C:\alex\CoaModule\CoaOnlineModule.psm1

<# Describe "New-CoaUser" {
    It "creates an object to seed a user account" {
        $user = New-CoaUser joe.crockett
        $user | Should -BeOfType [Object[]]
    }
} #>
$user = New-CoaUser joe.crockett
$user