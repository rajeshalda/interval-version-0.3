



---------------------------------------------------------
#for Group 

# Connect to Azure AD (if not already connected)
Connect-AzureAD

# Retrieve the group object details
Get-AzureADGroup -SearchString "Infra-Group"

#   2aa5a15b-9ed8-478d-9029-ec52469fdb36

------------------------------------------------------

#first create a policy


# Create a Teams Application Access Policy
New-CsApplicationAccessPolicy -Identity "AppAccessPolicyForGroup" -AppIds "64083eeb-dfe6-4134-a720-f33fa67b237b"

----------------------------------------------------------------------------------------------------------------------------


#Not Working 🛠️

# Assign the application access policy to the group
Grant-CsApplicationAccessPolicy -PolicyName "AppAccessPolicyForGroup" -Identity "2aa5a15b-9ed8-478d-9029-ec52469fdb36"

------------------------------------------------------------------------------------------------------------------------------
verify the policy by viewing the policy list:

# List all application access policies
Get-CsApplicationAccessPolicy

--------------------------------------------------------------------------------------------------------------------------

#its gives the list of all members in the group

# Configuration
$GroupName = "Infra-Group"  # Name of the Azure AD group

# Connect to Azure AD
Write-Host "Connecting to Azure AD..."
Connect-AzureAD -ErrorAction Stop

# Step 1: Retrieve the Group
Write-Host "Retrieving group: $GroupName..."
$Group = Get-AzureADGroup -SearchString $GroupName

if (-not $Group) {
    Write-Host "Error: Group '$GroupName' not found. Exiting..." -ForegroundColor Red
    exit
}

Write-Host "Group found: $($Group.DisplayName)" -ForegroundColor Green
Write-Host "Group Object ID: $($Group.ObjectId)" -ForegroundColor Cyan

# Step 2: Get All Members of the Group
Write-Host "Fetching all members of the group..."
$GroupMembers = Get-AzureADGroupMember -ObjectId $Group.ObjectId

if (-not $GroupMembers) {
    Write-Host "Error: No members found in the group." -ForegroundColor Yellow
} else {
    Write-Host "Found $($GroupMembers.Count) members in the group:" -ForegroundColor Green
    $GroupMembers | ForEach-Object {
        Write-Host "DisplayName: $($_.DisplayName), UPN: $($_.UserPrincipalName), ObjectId: $($_.ObjectId)" -ForegroundColor Cyan
    }
}

Write-Host "Group member retrieval completed." -ForegroundColor Cyan

------------------------------------------------------------------------------------------------------------------------------------

# its will apply the policy to all member one by one.
#other domain cannot grant to this domain like nathcorp as we have test on on-microsoft trial Tenet.

# Configuration
$GroupName = "Infra-Group"                # Name of the Azure AD group
$PolicyName = "AppAccessPolicyForGroup"  # Name of the Teams Application Access Policy

# Connect to Azure AD
Write-Host "Connecting to Azure AD..."
Connect-AzureAD -ErrorAction Stop

# Step 1: Retrieve the Group
Write-Host "Retrieving group: $GroupName..."
$Group = Get-AzureADGroup -SearchString $GroupName

if (-not $Group) {
    Write-Host "Error: Group '$GroupName' not found. Exiting..." -ForegroundColor Red
    exit
}

Write-Host "Group found: $($Group.DisplayName)" -ForegroundColor Green
Write-Host "Group Object ID: $($Group.ObjectId)" -ForegroundColor Cyan

# Step 2: Get All Members of the Group
Write-Host "Fetching all members of the group..."
$GroupMembers = Get-AzureADGroupMember -ObjectId $Group.ObjectId

if (-not $GroupMembers) {
    Write-Host "Error: No members found in the group. Exiting..." -ForegroundColor Red
    exit
}

Write-Host "Found $($GroupMembers.Count) members in the group." -ForegroundColor Green

# Step 3: Assign the Policy to Each Member
foreach ($User in $GroupMembers) {
    Try {
        Write-Host "Assigning policy '$PolicyName' to user: $($User.UserPrincipalName)..."
        Grant-CsApplicationAccessPolicy -PolicyName $PolicyName -Identity $User.UserPrincipalName
        Write-Host "Policy assigned successfully to $($User.UserPrincipalName)." -ForegroundColor Green
    } Catch {
        Write-Host "Error assigning policy to $($User.UserPrincipalName): $_" -ForegroundColor Red
    }
}

Write-Host "Policy assignment process completed." -ForegroundColor Cyan

---------------------------------------------------------------------------------------------------------

