# Microsoft Graph API - Attendance and Meeting Insights Script

This PowerShell script uses the Microsoft Graph API and Teams Application Access Policies to retrieve Teams meetings and attendance data for a specific user over the past 7 days. It supports **group-based policy management** to assign access to users in bulk.

---

## Features
- Retrieves Teams meetings for the specified user.
- Displays meeting details, including:
  - **Subject**
  - **Start/End Time**
  - **Online Meeting Details**
  - **Attendance Reports** with participant names, emails, and attendance durations.
- Leverages **Teams Application Access Policies** for access control.
- Supports group-based policy management:
  - Retrieve group members.
  - Apply application access policies to all members in the group.
- Handles both meeting descriptions (if available) and attendance insights.

---

## Prerequisites

### **1. Azure AD App Registration**
- Register an app in the Azure AD Portal.
- Generate a **Client ID** and **Client Secret**.
- Assign the following **Application Permissions** under Microsoft Graph:
  - `Calendars.Read`
  - `OnlineMeetingArtifact.Read.All`
  - `User.Read.All`
- Grant **admin consent** for these permissions.

---

### **2. Teams Application Access Policies**
**Application Access Policies** are required to allow your app to access Teams meetings and attendance data.

1. **Connect to Microsoft Teams**:
   Make sure the Teams PowerShell module is installed and connected:
   ```powershell
   Install-Module -Name PowerShellGet -Force -SkipPublisherCheck
   Install-Module -Name MicrosoftTeams
   Connect-MicrosoftTeams
   ```

2. **Retrieve the Group**:
   Find the Azure AD group you want to target:
   ```powershell
   # Connect to Azure AD
   Connect-AzureAD

   # Retrieve the group object
   Get-AzureADGroup -SearchString "Infra-Group"
   ```

   Example Output:
   ```
   DisplayName: Infra-Group
   ObjectId: 2aa5a15b-9ed8-478d-9029-ec52469fdb36
   ```

3. **Create the Policy**:
   Create a Teams Application Access Policy for your app:
   ```powershell
   New-CsApplicationAccessPolicy -Identity "AppAccessPolicyForGroup" -AppIds "64083eeb-dfe6-4134-a720-f33fa67b237b"
   ```

4. **Assign the Policy to the Group**:
   Bulk-apply the policy to all members of the group:
   ```powershell
   # Configuration
   $GroupName = "Infra-Group"
   $PolicyName = "AppAccessPolicyForGroup"

   # Retrieve the group
   $Group = Get-AzureADGroup -SearchString $GroupName

   # Get group members
   $GroupMembers = Get-AzureADGroupMember -ObjectId $Group.ObjectId

   # Assign policy to all members
   foreach ($User in $GroupMembers) {
       Try {
           Write-Host "Assigning policy '$PolicyName' to user: $($User.UserPrincipalName)..."
           Grant-CsApplicationAccessPolicy -PolicyName $PolicyName -Identity $User.UserPrincipalName
           Write-Host "Policy assigned successfully to $($User.UserPrincipalName)." -ForegroundColor Green
       } Catch {
           Write-Host "Error assigning policy to $($User.UserPrincipalName): $_" -ForegroundColor Red
       }
   }
   ```

5. **Verify the Policy**:
   Confirm the policy was created:
   ```powershell
   Get-CsApplicationAccessPolicy
   ```

---

## Setup Instructions

1. **Clone or Download**:
   Clone or download this repository to your local machine.

2. **Update Configuration**:
   Open the script file and update the following configuration details:
   ```powershell
   $TenantId = "YOUR_TENANT_ID"
   $ClientId = "YOUR_CLIENT_ID"
   $ClientSecret = "YOUR_CLIENT_SECRET"
   $TargetUserId = "USER_UPN_OR_OBJECT_ID"  # Example: "user@contoso.com"
   ```

3. **Run the Script**:
   Execute the script in PowerShell:
   ```powershell
   .\YourScriptName.ps1
   ```

---

## Script Workflow

### **Attendance and Meeting Insights**
1. **Authentication**:
   The script retrieves an access token using the **Client Credentials Flow**.
   
2. **Fetch User Info**:
   It retrieves details of the specified target user.

3. **Retrieve Meetings**:
   Fetches all meetings within the last 7 days using the `/events` endpoint.

4. **Attendance Reports**:
   - Extracts online meeting details.
   - Decodes the meeting ID to fetch attendance records.
   - Displays participant names, emails, and total attendance durations.

5. **Error Handling**:
   Catches and displays API errors or missing data gracefully.

---

## Example Outputs

### **Meeting Insights**
```
Connected using application permissions.
Retrieving information for user: user@contoso.com...
Target User ID: abcdef12-3456-7890-abcd-ef1234567890
Retrieving meetings from 2024-11-11T00:00:00Z to 2024-11-18T00:00:00Z...

Processing meeting: Weekly Sync 🛠️
Meeting Description:
This is a weekly sync meeting to discuss team updates and progress.

Attendance Report ID: 12345abc-6789-def0-1234-56789abcdef0
User: John Doe, Email: john.doe@contoso.com, Total Time: 240 seconds
User: Jane Smith, Email: jane.smith@contoso.com, Total Time: 360 seconds
```

### **Group-Based Policy Management**
```
Connecting to Azure AD...
Retrieving group: Infra-Group...
Group found: Infra-Group
Group Object ID: 2aa5a15b-9ed8-478d-9029-ec52469fdb36

Fetching all members of the group...
Found 3 members in the group:
DisplayName: John Doe, UPN: john.doe@contoso.com
DisplayName: Jane Smith, UPN: jane.smith@contoso.com
DisplayName: Bob Lee, UPN: bob.lee@contoso.com

Assigning policy 'AppAccessPolicyForGroup' to user: john.doe@contoso.com...
Policy assigned successfully to john.doe@contoso.com.

Assigning policy 'AppAccessPolicyForGroup' to user: jane.smith@contoso.com...
Policy assigned successfully to jane.smith@contoso.com.
```

---

## API Permissions and Justification

| **Permission Name**              | **Type**      | **Purpose**                                                                   |
|----------------------------------|---------------|-------------------------------------------------------------------------------|
| `User.Read.All`                  | Application   | Retrieve user information (`/users/{userId}`).                                |
| `Calendars.Read`                 | Application   | Read calendar events (`/users/{userId}/events`).                              |
| `OnlineMeetingArtifact.Read.All` | Application   | Fetch meeting attendance reports (`/onlineMeetings/.../attendanceReports`).   |

---

## Troubleshooting

- **403 Forbidden**:
  - Ensure the app has the correct API permissions.
  - Verify that the **Teams Application Access Policy** is correctly assigned to the user or group.

- **Group Policy Assignment Issues**:
  - Ensure the group exists in Azure AD.
  - Confirm group members and their UPNs are correct.

---

## License
This project is licensed under the MIT License - see the LICENSE file for details.
