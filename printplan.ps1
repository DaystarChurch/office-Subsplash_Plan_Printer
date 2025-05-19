
function Get-FluroAuthToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
        [Parameter(Mandatory = $true)]
        [SecureString]$Password
    )

    $body = "grant_type=password&username=$([uri]::EscapeDataString($Username))&password=$([uri]::EscapeDataString($Password))"

    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "User-Agent"   = "PowerShell"
        "Accept"       = "*/*"
    }

    $response = Invoke-RestMethod -Uri "https://api.fluro.io/token/login" -Method Post -Headers $headers -Body $body

    return @{
        Token  = $response.token
        Expiry = $response.expiry
    }
}

function New-FluroServiceFilter {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate,
        [Parameter(Mandatory = $true)]
        [string]$Timezone
    )

    return @{
        sort = @{
            sortKey = "startDate"
            sortDirection = "asc"
            sortType = "date"
        }
        filter = @{
            operator = "and"
            filters = @(
                @{
                    operator = "and"
                    filters = @(
                        @{
                            key = "status"
                            comparator = "in"
                            values = @("active", "draft", "archived")
                        }
                    )
                }
            )
        }
        search = ""
        includeArchived = $false
        allDefinitions = $true
        searchInheritable = $false
        includeUnmatched = $true
        startDate = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        endDate = $EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        timezone = $Timezone
    }
}
function Get-FluroServices {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AuthToken,
        [Parameter(Mandatory = $true)]
        [object]$FilterBody
    )

    $headers = @{
        "Authorization" = "Bearer $AuthToken"
        "Content-Type"  = "application/json"
        "Accept"        = "*/*"
    }

    $bodyJson = $FilterBody | ConvertTo-Json -Depth 10

    $response = Invoke-RestMethod -Uri "https://api.fluro.io/content/service/filter" -Method Post -Headers $headers -Body $bodyJson

    return $response
}

function New-PlanHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object]$JsonBody,
        [string[]]$Teams
    )

    # Load JSON
    $json = $JsonBody

    # Get teams from JSON if not specified
    if (-not $Teams -or $Teams.Count -eq 0) {
        $Teams = $json.plans[0].teams
    }

    # Prepare table headers
    $headers = @('Time', 'Detail') + $Teams

    # Get plan start time as DateTime (assume UTC in JSON)
    $planStartUtc = [datetime]::Parse($json.plans[0].startDate)
    $localTZ = [System.TimeZoneInfo]::Local

    # Extract service title, date/time, and versioning info
    $plan = $json.plans[0]
    $serviceTitle = $plan.title
    $serviceDateTimeUtc = [datetime]::Parse($plan.startDate)
    $serviceDateTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($serviceDateTimeUtc, [System.TimeZoneInfo]::FindSystemTimeZoneById("Mountain Standard Time"))
    $serviceDateTimeStr = $serviceDateTimeLocal.ToString("dddd, MMMM d, yyyy 'at' h:mm tt")

    # Versioning data
    $lastUpdatedUT = [datetime]::Parse($json.updated)
    $lastUpdatedLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($lastUpdatedUT, $localTZ).ToString("yyyy-MM-dd hh:mm:ss")
    $lastUpdatedBy = $json.updatedBy
    $printTime = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")

    # Build HTML
    $html = @"
<html>
<head>
<style>
    body { font-family: Segoe UI, Arial, sans-serif; }
    .header-row { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 1.5em; }
    .service-info { }
    .service-title { font-size: 2em; font-weight: bold; margin-bottom: 0.2em; }
    .service-subtitle { font-size: 1.1em; color: #555; }
    .version-info { text-align: right; font-size: 0.95em; color: #444; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 4px; vertical-align: top; text-align: left; }
    tr.song { background: #e6f7ff; }
    tr.breaker { background: #f9f9f9; font-weight: bold; }
    tr.start { background: #d9f7be; }
    .duration { font-size: 0.9em; color: #888; display: block; }
    /* Add more CSS as needed */
</style>
</head>
<body>
<div class="header-row">
    <div class="service-info">
        <div class="service-title">$serviceTitle</div>
        <div class="service-subtitle">$serviceDateTimeStr</div>
    </div>
    <div class="version-info">
        <div><strong>Last updated:</strong> $lastUpdatedLocal</div>
        <div><strong>Updated by:</strong> $lastUpdatedBy</div>
        <div><strong>Printed:</strong> $printTime</div>
    </div>
</div>
<table>
    <thead>
        <tr>
"@

    foreach ($header in $headers) {
        $html += "            <th>$header</th>`n"
    }
    $html += "        </tr>
    </thead>
    <tbody>
"

    # Calculate running time
    $runningTime = 0

    # Loop through each schedule entry
    foreach ($row in $json.plans[0].schedules) {
        $type = $row.type
        if (-not $type) { $type = "normal" }
        $duration = $row.duration

        # Calculate actual time for this row, convert to local timezone
        $rowTimeUtc = $planStartUtc.AddSeconds($runningTime)
        $rowTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($rowTimeUtc, $localTZ)
        $timeStr = $rowTimeLocal.ToString("HH:mm")

        # Duration in minutes, rounded
        $durationMin = [math]::Round($duration / 60)
        $durationLine = ""
        if ($durationMin -gt 0) {
            $durationLine = "<br/><span class='duration'>$durationMin min</span>"
        }

        # Add the title to the detail column, keeping HTML formatting
        $detailText = ""
        if ($row.title) {
            $detailText = "<strong>$($row.title)</strong>"
        }
        if ($row.detail) {
            if ($detailText) {
                $detailText += "<br/>$($row.detail)"
            } else {
                $detailText = $row.detail
            }
        }

        $html += "        <tr class='$type'>`n"
        $html += "            <td class='col-time'>$timeStr$durationLine</td>`n"
        $html += "            <td class='col-detail details'>$detailText</td>`n"
        foreach ($team in $Teams) {
            $cell = ""
            if ($row.notes -and $row.notes.$team) {
                $cell = $row.notes.$team
            }
            $teamClass = "col-" + ($team -replace '[^a-zA-Z0-9\-]', '-').ToLower()
            $html += "            <td class='$teamClass'>$cell</td>`n"
        }
        $html += "        </tr>`n"
        $runningTime += $duration
    }

    $html += @"
    </tbody>
</table>
</body>
</html>
"@

    return $html
}

function Convert-PlanHtmlToPdf {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlanHtml,
        [Parameter(Mandatory = $true)]
        [string]$OutPath
    )

    # Create a temporary HTML file
    $tempHtml = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'
    Set-Content -Path $tempHtml -Value $PlanHtml -Encoding UTF8

    # Build msedge command
    $msedgePath = "msedge"
    $arguments = @(
        "--headless"
        "--disable-gpu"
        "--run-all-compositor-stages-before-draw"
        "--print-to-pdf=""$OutPath"""
        "`"$tempHtml`""
    ) -join ' '

    # Run msedge to print to PDF
    $process = Start-Process -FilePath $msedgePath -ArgumentList $arguments -NoNewWindow -Wait -PassThru

    # Clean up temp file
    Remove-Item $tempHtml -ErrorAction SilentlyContinue
}


### Date Filtering Setup
$now = Get-Date
$localTimezone = "America/Edmonton" # Set your local timezone here reference: https://www.timeanddate.com/time/zones/
# Find the next Sunday (including today if today is Sunday)
$daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
$nextSunday = $now.Date.AddDays($daysUntilSunday)
$endDate = $nextSunday.AddDays(1).AddMilliseconds(-1) # End of Sunday
# Debug information
Write-Debug "Local Timezone: $localTimezone"
Write-Debug "Current Date: $now"
Write-Debug "Days until next Sunday: $daysUntilSunday"
Write-Debug "Next Sunday: $nextSunday"
Write-Debug "End Date: $endDate"
###

#$token = (Get-FluroAuthToken -Username "user" -Password "pass").Token
