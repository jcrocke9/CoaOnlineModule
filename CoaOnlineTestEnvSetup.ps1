# RUN THIS SCRIPT ONCE
$ListOfPolicies = @(
    "Termination Retention Policy",
    "COA Department Head Policy",
    "COA F1 Policy",
    "COA Policy",
    "COA PZ Policy"
    )

foreach ($policy in $ListOfPolicies) {
    New-RetentionPolicy $policy
}

$ListOfRoleAssignmentPolicies = @("COA Default Role Assignment Policy")

foreach ($item in $ListOfRoleAssignmentPolicies) {
    New-RoleAssignmentPolicy -Name $item
}

$ListOfClientAccessRules = @("COAOWAMailboxPolicy")

foreach ($item in $ListOfClientAccessRules) {
    New-ClientAccessRule -Name $item
}

