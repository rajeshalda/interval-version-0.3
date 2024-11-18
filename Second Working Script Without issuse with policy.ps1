# Configuration details
$TenantId = "22ee08ca-64e4-4a31-808d-9dc6b9e35882"
$ClientId = "64083eeb-dfe6-4134-a720-f33fa67b237b"
$ClientSecret = "e5b8Q~R3eIJSenx~b3~EOXyHkbHyTFxjPTCowbq7"
$TargetUserId = "rajesh-sir-tenet@M365x39618365.onmicrosoft.com"  # Replace with target user's UPN or object ID

# Helper function to get access token for application authentication
function Get-ApplicationAccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    $TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $Body
    return $TokenResponse.access_token
}

# Get Access Token
$AccessToken = Get-ApplicationAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
$Headers = @{ Authorization = "Bearer $AccessToken" }
Write-Host "Connected using application permissions."

# Step 1: Get user information
Write-Host "Retrieving information for user: $TargetUserId..."
$TargetUser = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$TargetUserId" -Headers $Headers
if (-not $TargetUser) {
    Write-Host "Failed to retrieve target user information. Exiting..." -ForegroundColor Red
    exit
}
Write-Host "Target User ID: $($TargetUser.id)" -ForegroundColor Cyan

# Step 2: Get meetings from the last 7 days
$StartDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Retrieving meetings from $StartDate to $EndDate..."

$Uri = "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/events?"
$Filter = [System.Web.HttpUtility]::UrlEncode("start/dateTime ge '$StartDate' and end/dateTime le '$EndDate'")
$MeetingsUri = "$Uri`$filter=$Filter&`$select=subject,start,end,onlineMeeting,onlineMeetingUrl"

$Meetings = Invoke-RestMethod -Method Get -Uri $MeetingsUri -Headers $Headers

if ($Meetings.value) {
    foreach ($Meeting in $Meetings.value) {
        Write-Host "`nProcessing meeting: $($Meeting.subject)" -ForegroundColor Green

        if ($Meeting.onlineMeeting) {
            # Decode Join URL
            $DecodedJoinUrl = [System.Web.HttpUtility]::UrlDecode($Meeting.onlineMeeting.joinUrl)
            if ($DecodedJoinUrl -match "19:meeting_([^@]+)@thread.v2") {
                $MeetingId = "19:meeting_" + $matches[1] + "@thread.v2"
            }

            if ($DecodedJoinUrl -match '"Oid":"([^"]+)"') {
                $OrganizerOid = $matches[1]
            }

            $FormattedString = "1*${OrganizerOid}*0**${MeetingId}"
            $Base64MeetingId = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($FormattedString))

            Try {
                Write-Host "Retrieving attendance report for Base64 Meeting ID: $Base64MeetingId..."
                $AttendanceReport = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/onlineMeetings/$Base64MeetingId/attendanceReports" -Headers $Headers

                if ($AttendanceReport.value) {
                    $ReportId = $AttendanceReport.value[0].id
                    Write-Host "Attendance Report ID: $ReportId" -ForegroundColor Green

                    # Retrieve Attendance Records
                    $AttendanceRecords = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/onlineMeetings/$Base64MeetingId/attendanceReports/$ReportId/attendanceRecords" -Headers $Headers

                    if ($AttendanceRecords.value) {
                        foreach ($Record in $AttendanceRecords.value) {
                            Write-Host "User: $($Record.identity.user.displayName), Total Time: $($Record.totalAttendanceInSeconds) seconds"
                        }
                    } else {
                        Write-Host "No attendance records found for this meeting." -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No attendance report found for this meeting." -ForegroundColor Yellow
                }
            } Catch {
                Write-Host "Error retrieving attendance report: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Not a Teams meeting. Skipping..." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "No meetings found in the last 7 days." -ForegroundColor Yellow
}
