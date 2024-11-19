# Microsoft Graph API - Attendance and Meeting Insights Script

This PowerShell script uses the Microsoft Graph API and Teams Application Access Policies to retrieve Teams meetings and attendance data for a specific user over the past 7 days. It authenticates using the **Client Credentials Flow** and ensures secure access via access policies.

---

## Features
- Retrieves Teams meetings for the specified user.
- Displays meeting details, including:
  - **Subject**
  - **Start/End Time**
  - **Online Meeting Details**
  - **Attendance Reports** with participant names, emails, and attendance durations.
- Leverages **Teams Application Access Policies** for access control.
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
**Application Access Policies** are required to allow your app to access Teams meetings and attendance data. Follow these steps to configure:

1. **Create an Application Access Policy**:
   Use the following PowerShell command to create the policy:
   ```powershell
   New-CsApplicationAccessPolicy -PolicyName "AppAccessPolicyForGraph" -Description "Policy for Graph API access" -AppIds "64083eeb-dfe6-4134-a720-f33fa67b237b"
   ```
   Replace the **AppIds** value with your Azure AD app's **Client ID**.

2. **Assign the Policy to a Group**:
   To assign the policy to a group, first ensure the target users are in the group:
   ```powershell
   Grant-CsApplicationAccessPolicy -PolicyName "AppAccessPolicyForGraph" -Identity "Infra-Group"
   ```
   Replace `Infra-Group` with the name of the group containing users whose Teams data the app can access.

3. **Verify Policy Assignment**:
   Check if the policy is assigned correctly using:
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

## Example Output

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

---

## API Permissions and Justification

| **Permission Name**              | **Type**      | **Purpose**                                                                   |
|----------------------------------|---------------|-------------------------------------------------------------------------------|
| `User.Read.All`                  | Application   | Retrieve user information (`/users/{userId}`).                                |
| `Calendars.Read`                 | Application   | Read calendar events (`/users/{userId}/events`).                              |
| `OnlineMeetingArtifact.Read.All` | Application   | Fetch meeting attendance reports (`/onlineMeetings/.../attendanceReports`).   |

---

## Teams Application Access Policies

To restrict app access to specific users or groups, configure **Teams Application Access Policies**:
- **Policy Creation**: Use `New-CsApplicationAccessPolicy`.
- **Policy Assignment**: Use `Grant-CsApplicationAccessPolicy`.
- **Verification**: Use `Get-CsApplicationAccessPolicy`.

---

## License
This project is licensed under the MIT License - see the LICENSE file for details.

---

## Contributing
Feel free to fork this repository, submit pull requests, or report issues.

---

## Troubleshooting
- **403 Forbidden**:
  - Ensure the app has the correct API permissions.
  - Verify that the **Teams Application Access Policy** is correctly assigned to the user or group.
- **No Data Found**:
  - Check if the target user has valid Teams meetings in the last 7 days.
  - Verify the target user’s UPN or object ID.

For any issues, contact the repository owner or create a GitHub issue.
