[CmdletBinding()]
param (
    [Parameter()]
    [switch]$LoginSubsplash,
    [Parameter()]
    [switch]$ListTeams,
    [Parameter()]
    [switch]$ListServices, 
    [Parameter()]
    [switch]$PrintSongs,
    [Parameter()]
    [switch]$PrintPlan,
    [Parameter()]
    [switch]$Headless,
    [Parameter()]
    [string]$serviceid,
    [Parameter()]
    [string[]]$Teams
)
#region Functions
function Get-FluroAuthToken {
    param(
        [Parameter(Mandatory = $true)]
        [pscredential]$FluroCreds
    )
    $UserName = $FluroCreds.UserName
    if (-not $UserName) {
        Write-Error "Username is empty. Please provide a valid username."
        return
    }
    $Password = $FluroCreds.Password
    if (-not $Password) {
        Write-Error "Password is empty. Please provide a valid password."
        return
    }
    $Password = $($Password | ConvertFrom-SecureString -AsPlainText)
    $body = "grant_type=password&username=$([uri]::EscapeDataString($UserName))&password=$([uri]::EscapeDataString($Password))"

    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "User-Agent"   = "PowerShell"
        "Accept"       = "*/*"
    }

    Try {
        $response = Invoke-RestMethod -Uri "https://api.fluro.io/token/login" -Method Post -Headers $headers -Body $body -StatusCodeVariable statusCode
    }
    Catch {
        Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
        return 
    }

    return @{
        Response   = $response
        StatusCode = $statusCode
        Token      = $response.token
        Expiry     = $response.expiry
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
        sort              = @{
            sortKey       = "startDate"
            sortDirection = "asc"
            sortType      = "date"
        }
        filter            = @{
            operator = "and"
            filters  = @(
                @{
                    operator = "and"
                    filters  = @(
                        @{
                            key        = "status"
                            comparator = "in"
                            values     = @("active", "draft", "archived")
                        }
                    )
                }
            )
        }
        search            = ""
        includeArchived   = $false
        allDefinitions    = $true
        searchInheritable = $false
        includeUnmatched  = $true
        startDate         = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        endDate           = $EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        timezone          = $Timezone
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

function Get-FluroServiceById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AuthToken,
        [Parameter(Mandatory = $true)]
        [string]$ServiceId
    )

    $headers = @{
        "Authorization" = "Bearer $AuthToken"
        "Content-Type"  = "application/json"
        "Accept"        = "*/*"
    }

    $url = "https://api.fluro.io/content/get/$ServiceId"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Error "Failed to retrieve service with ID $ServiceId from Fluro API."
        return $null
    }
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
            }
            else {
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

function Set-FluroCreds {

    $flurocredinput = Get-Credential -Message "Enter your Subsplash credentials"
    $flurocreds = New-Object System.Management.Automation.PSCredential ($flurocredinput.UserName, $flurocredinput.Password)
    try {
        Export-Clixml -Path "fluro.xml" -InputObject $flurocreds -ErrorAction Stop
        Write-Host "Fluro credentials saved successfully."
        return $flurocreds
    }
    catch {
        Write-Error "Failed to save Fluro credentials. Please check file permissions."
        Write-Host "Credentials not saved. Script will not run without saved credentials."
        Write-Host "Please run the script with -LoginSubsplash to set up credentials after confirming you can write to the directory."
        return $null
    }
}
#endregion
#### Main Script Execution

#region Handle "LoginSubsplash" parameter
if ($LoginSubsplash) {
    Write-Host "Setting up Fluro credentials..."
    $flurocreds = Set-FluroCreds
    if ($null -eq $flurocreds) {
        Write-Error "Failed to set Fluro credentials. Please check file permissions."
        Write-Host "Credentials not set. Script will not run without saved credentials."
        exit 1
    }
    try {
        $fluroauth = Get-FluroAuthToken -FluroCreds $flurocreds
    }
    catch {
        Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
        exit 1
    }
    if ($fluroauth.StatusCode -ne 200) {
        Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
        exit 1
    }
    else {
        Write-Host "FLogin test successful. You can now run the script without specifying -LoginSubsplash."
    }
    exit 0
}
#endregion

#region Saved Credential Management
# Credentials are saved in fluro.xml in the same directory as this script; password is encrypted. 
# If the file doesn't exist, the script will prompt for credentials and create the file. 
# If -headless is specified, then the script will exit with an error instead of prompting.
###
# Check if the credentials file exists
if (Get-Item -Path "fluro.xml" -ErrorAction SilentlyContinue) {
    Write-Host "Fluro credentials found. Loading..."
    try { 
        $flurocreds = Import-Clixml -Path "fluro.xml"
    }
    catch {
        Write-Error "Failed to load Fluro credentials. Please check file permissions."
        Write-Host "Credentials not loaded. Script will not run without saved credentials."
        Write-Host "Please run the script with -LoginSubsplash to set up credentials after confirming you can write to the directory."
        exit 1
    }    
}
elseif ($headless) {
    # If -headless is specified and the credentials file doesn't exist, exit with an error
    Write-Error "Credentials not found. Please run the script with -LoginSubsplash to set up credentials."
    Write-Host "Credentials not found. Script will not run without saved credentials."
    exit 1 
}
else {
    # If the credentials file doesn't exist and -headless is not specified, prompt for credentials
    Write-Host "Subsplash credentials not found. Please enter your credentials."
    $flurocreds = Set-FluroCreds
    if ($null -eq $flurocreds) {
        Write-Error "Failed to set Fluro credentials. Please check file permissions."
        Write-Host "Credentials not set. Script will not run without saved credentials."
        exit 1
    }
}

Write-Host "Fluro credentials loaded successfully. Username is $($flurocreds.UserName). Accessing Fluro API..."
# Test the credentials by getting an auth token
try {
    $fluroauth = Get-FluroAuthToken -FluroCreds $flurocreds
}
catch {
    Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
    exit 1
}
if ($fluroauth.StatusCode -ne 200) {
    Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
    exit 1
}
else {
    Write-Host "Fluro API authentication successful."
    $token = $fluroauth.Token
}
#endregion

#region List Services
if ($ListServices) {
    # Set up the date range for the search (next Sunday to end of that Sunday)
    $now = Get-Date
    $localTimezone = "America/Edmonton"
    $daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
    $nextSunday = $now.Date.AddDays($daysUntilSunday)
    $endDate = $nextSunday.AddDays(1).AddMilliseconds(-1)

    # Create the filter body for the API request
    $filterBody = New-FluroServiceFilter -StartDate $nextSunday -EndDate $endDate -Timezone $localTimezone

    # Get the services from the API
    $services = Get-FluroServices -AuthToken $token -FilterBody $filterBody

    if (-not $services -or $services.Count -eq 0) {
        Write-Host "No services found."
        exit 0
    }

    $localTZ = [System.TimeZoneInfo]::Local
    $services | Select-Object `
        @{Name="Title";Expression={$_.title}},
        @{Name="Start (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.startDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
        @{Name="End (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.endDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
        @{Name="_id";Expression={$_. _id}} |
        Format-Table -AutoSize

    exit 0
}
#endregion

#region Search for Services 
if (-not $serviceid) {
    # Set up the date range for the search (next Sunday to end of that Sunday)
    $now = Get-Date
    $localTimezone = "America/Edmonton"
    $daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
    $nextSunday = $now.Date.AddDays($daysUntilSunday)
    $endDate = $nextSunday.AddDays(1).AddMilliseconds(-1)

    # Create the filter body for the API request
    $filterBody = New-FluroServiceFilter -StartDate $nextSunday -EndDate $endDate -Timezone $localTimezone

    # Get the services from the API
    $services = Get-FluroServices -AuthToken $token -FilterBody $filterBody

    # Check if services were retrieved
    if (-not $services -or $services.Count -eq 0) {
        Write-Error "No services found."
        exit 1
    }

    if ($services.Count -eq 1) {
        # Only one service found, use its ID
        $serviceid = $services[0]._id
        Write-Host "Service ID: $serviceid"
    }
    elseif ($services.Count -gt 1) {
        if ($Headless) {
            Write-Error "Multiple services found. Please run without -headless to select a service."
            exit 1
        }
        Write-Host "Multiple services found. Please select one:"
        $localTZ = [System.TimeZoneInfo]::Local
        $servicesDisplay = $services | Select-Object `
            @{Name="Title";Expression={$_.title}},
            @{Name="Start (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.startDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
            @{Name="End (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.endDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
            @{Name="_id";Expression={$_. _id}}

        $selectedService = $servicesDisplay | Out-GridView -Title "Select a Service" -PassThru
        if (-not $selectedService) {
            Write-Error "No service selected. Exiting."
            exit 1
        }
        $serviceid = $selectedService._id
        Write-Host "Selected Service ID: $serviceid"
    }
}
else {
    Write-Host "Service ID provided: $serviceid. Skipping service search."
}
#endregion