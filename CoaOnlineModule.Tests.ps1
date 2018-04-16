Import-Module C:\alex\CoaOnlineModule\CoaOnlineModule.psm1 
class UserObject {
    [string]$samAccountName
    [string]$License
}

Describe "New-CoaUser" {
    
    Context 'Test of object type' {
        It "Does not throw" {
            [UserObject]::new()
        }
        It 'Users should not have an empty samAccountName' {
            (New-CoaUser joe.c) | Should -BeOfType [System.Collections.Generic.List[UserObject]]
        }
    }
    <# Context 'Test of pipeline' {
        It "Should Not BeNullOrEmpty" {
            New-CoaUser test.user | Set-CoaExchangeAttributes | Should -Not -BeNullOrEmpty
        }
    } #>
}
