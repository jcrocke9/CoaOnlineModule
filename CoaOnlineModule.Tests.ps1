Import-Module C:\alex\CoaOnlineModule\CoaOnlineModule.psm1 
<# class UserObject {
    [string]$samAccountName
    [string]$License
} #>

Describe "New-CoaUser" {
    
    Context 'Test of object type' {
        $userTestName = "joe.c"
        <# It "Does not throw" {
            [UserObject]::new()
        } #>
        [int]$numUsers = $CoaUsersToWorkThrough.Count
        $userObjectArr = [System.Collections.Generic.List[PSObject]]::new()
        Do {
            $userObject = "UserObject"
            $userObjectArr.Add($userObject)
        } while ($userObjectArr.Count -le $numUsers) 
        It "Create a user" {
            $userTest = New-CoaUser $userTestName
            $userTest | Should -Be $userObjectArr
        }
        foreach ($userTest in $CoaUsersToWorkThrough) {
            It "User should have a samAccountName" {
                [bool]($userTest.PSObject.Properties.Name -match "samAccountName") | Should -Be $true
            }
        }
    }
    <# Context 'Test of pipeline' {
        It "Should Not BeNullOrEmpty" {
            New-CoaUser test.user | Set-CoaExchangeAttributes | Should -Not -BeNullOrEmpty
        }
    } #>

    InModuleScope CoaOnlineModule {
        Context 'Test of object type' {
            $userTestName = "joe.c"
            <# It "Does not throw" {
                [UserObject]::new()
            } #>
            [int]$numUsers = $CoaUsersToWorkThrough.Count
            $userObjectArr = [System.Collections.Generic.List[PSObject]]::new()
            Do {
                $userObject = "UserObject"
                $userObjectArr.Add($userObject)
            } while ($userObjectArr.Count -le $numUsers) 
            It "Create a user" {
                $userTest = New-CoaUser $userTestName
                $userTest | Should -Be $userObjectArr
            }
            foreach ($userTest in $CoaUsersToWorkThrough) {
                It "User should have a samAccountName" {
                    [bool]($userTest.PSObject.Properties.Name -match "samAccountName") | Should -Be $true
                }
            }
        }
    }
    Context 'Clear the variable' {
        It "Should clear the variable without issue" {
            Clear-CoaUser
        }
    }
}
