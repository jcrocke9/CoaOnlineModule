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
    }
    Context 'creates an object to seed a user account' {
        Mock New-Object {UserObject} -Verifiable
        It 'Returns a user' {
            New-CoaUser joe.c
            Assert-MockCalled New-Object 1 {
                
                $TypeName -eq 'UserObject'
            }
        }
    }
    <# Context 'Test of pipeline' {
        It "Should Not BeNullOrEmpty" {
            New-CoaUser test.user | Set-CoaExchangeAttributes | Should -Not -BeNullOrEmpty
        }
    } #>
}
