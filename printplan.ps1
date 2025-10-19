[CmdletBinding()]
param()  # No parameters; all config comes from environment

#region Functions
# ---------------------------
# Helpers / configuration load
# ---------------------------
function Get-EnvOrDefault {
    param([string]$Name, [object]$Default = $null)
    $value = (Get-Item -Path "Env:$Name" -ErrorAction SilentlyContinue).Value
    if ($value -and $value.Trim().Length -gt 0) { return $value }
    return $Default
}

function Get-JsonFromEnv {
    param([string]$Name)
    $envVar = (Get-Item -Path "Env:$Name" -ErrorAction SilentlyContinue).Value
    if (-not $envVar) { return $null }
    try { return ($envVar | ConvertFrom-Json -Depth 50) }
    catch {
        Write-Error "Environment variable '$Name' is not valid JSON. $_"
        exit 2
    }
}

function Get-JsonFile {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path -Path $Path)) {
        Write-Error "PLAN_PROFILES_FILE points to '$Path' but the file does not exist."
        exit 2
    }
    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        return ($raw | ConvertFrom-Json -Depth 50)
    } catch {
        Write-Error "Failed to parse JSON from '$Path'. $_"
        exit 2
    }
}
# ---------------------------
# Subsplash API functions
# ---------------------------
function Get-SubsplashAuthToken {
    param(
        [Parameter(Mandatory = $true)]
        [pscredential]$SubsplashCreds
    )
    Write-Debug "Get-SubsplashAuthToken function called."
    $UserName = $SubsplashCreds.UserName
    Write-Debug "Username: $UserName"
    if (-not $UserName) {
        Write-Error "Username is empty. Please provide a valid username."
        return
    }
    $Password = $SubsplashCreds.Password
    if (-not $Password) {
        Write-Error "Password is empty. Please provide a valid password."
        return
    }
    Write-Debug "Converting password to plain text for auth string."
    $Password = $($Password | ConvertFrom-SecureString -AsPlainText)
    Write-Debug "Constructing body for authentication request."
    $body = "grant_type=password&username=$([uri]::EscapeDataString($UserName))&password=$([uri]::EscapeDataString($Password))"
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "User-Agent"   = "PowerShell"
        "Accept"       = "*/*"
    }

    Try {
        Write-Debug "Sending authentication request to Subsplash API."
        $response = Invoke-RestMethod -Uri "https://api.fluro.io/token/login" -Method Post -Headers $headers -Body $body -StatusCodeVariable statusCode
    }
    Catch {
        Write-Error "Failed to authenticate with Subsplash API. $_"
        return $_
    }

    return @{
        Response   = $response
        StatusCode = $statusCode
        Token      = $response.token
        Expiry     = $response.expiry
    }
}

function New-SubsplashServiceFilter {
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
function Get-SubsplashServices {
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

function Get-SubsplashServiceById {
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
        Write-Error "Failed to retrieve service with ID $ServiceId from Subsplash API."
        return $null
    }
}

function Get-SubsplashPlanDetails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AuthToken,
        [Parameter(Mandatory = $true)]
        [string]$PlanId
    )

    # Construct the API URL
    $url = "https://api.fluro.io/content/get/$PlanId"

    # Set up the headers for the API request
    $headers = @{
        "Authorization" = "Bearer $AuthToken"
        "Content-Type"  = "application/json"
        "Accept"        = "*/*"
    }

    try {
        # Make the API request
        Write-Debug "Sending request to Subsplash API for plan details. URL: $url"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        Write-Debug "Plan details retrieved successfully."
        return $response
    }
    catch {
        Write-Error "Failed to retrieve plan details for Plan ID $PlanId. $_"
        return $null
    }
}
# ---------------------------
# Plansheet rendering functions
# ---------------------------
function New-PlanHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object]$plandetails,
        [object]$Teams,
        [string]$PlanName,
        [string]$orientation = "landscape",
        [string]$CssPath = "print.css"   # Optional: path to CSS file
    )
    Write-Debug "New-PlanHtml function called."
    Write-Debug "PlanName: $PlanName"
    Write-Debug "Orientation: $orientation"
    Write-Debug "Teams: $($Teams | ConvertTo-Json -Depth 10)"
    # Load JSON

    # Get teams from JSON if not specified
    if (-not $Teams -or $Teams.Count -eq 0) {
        $Teams = $plandetails.teams
    }

    # Prepare table headers
    $headers = @('Time', 'Detail') + $Teams

    # Get plan start time as DateTime (assume UTC in JSON)
    $planStartUtc = [datetime]::Parse($plandetails.startDate)
    $localTZ = [System.TimeZoneInfo]::Local

    # Extract service title, date/time, and versioning info
    $plan = $plandetails
    $serviceTitle = $plan.event.title
    $serviceDateTimeUtc = [datetime]::Parse($plan.startDate)
    $serviceTz = [System.TimeZoneInfo]::FindSystemTimeZoneById($timezone)  # e.g., "America/Edmonton"
    $serviceDateTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($serviceDateTimeUtc, $serviceTz)
    $serviceDateTimeStr = $serviceDateTimeLocal.ToString("dddd, MMMM d, yyyy 'at' h:mm tt")

    # Versioning data
    $lastUpdatedUT = [datetime]::Parse($plan.updated)
    $lastUpdatedLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($lastUpdatedUT, $localTZ).ToString("yyyy-MM-dd hh:mm:ss")
    $lastUpdatedBy = $plan.updatedBy
    $printTime = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")

    # Load CSS from file
    Write-Debug "Trying to load CSS from file. Path: $CssPath"
    $cssContent = ""
    if (Test-Path $CssPath) {
        Write-Debug "CSS file found. Loading content."
        $cssContent = Get-Content $CssPath -Raw
    } else {
        Write-Debug "CSS file not found. Using default CSS."
        Write-Warning "CSS file '$CssPath' not found. Using default CSS."
        $cssContent = @"
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
"@
    }
    # Ensure @page is always present
    $cssContent += "`n@page {margin: 0.25in; padding: 0; size: letter $orientation;}`n"
    # Build HTML
    $html = @"
<html>
<head>
<style>
$cssContent
</style>
</head>
<body>
<div class="document">
<div class="header-row">
    <div class="service-info">
        <div class="service-title">$serviceTitle</div>
        <div class="service-subtitle">$serviceDateTimeStr</div>
    </div>
    <div class="plansheet-info">
        <div class="plan-name">$PlanName</div>
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
    foreach ($row in $plan.schedules) {
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
            $detailText = "<span class='detail-title'>$($row.title)</span>"
        }
        if ($row.detail) {
            if ($detailText) {
                $detailText += "<br/><span class='detail-text'>$($row.detail)</span>"
            }
            else {
                $detailText = $row.detail
            }
        }

        $html += "        <tr class='$type'>`n"
        $html += "            <td class='col-time'>$timeStr$durationLine</td>`n"
        $html += "            <td class='col-detail'>$detailText</td>`n"
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
</div>
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

    # Create a temp HTML file
    $tempHtml = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'
    Set-Content -Path $tempHtml -Value $PlanHtml -Encoding UTF8

    try {
        # Call the Python-based CLI that's on PATH inside the container
        # Equivalent CLI exists as "python -m weasyprint" as well.
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "weasyprint"
        $psi.ArgumentList.Add($tempHtml)
        $psi.ArgumentList.Add($OutPath)
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute = $false

        $p = [System.Diagnostics.Process]::Start($psi)
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()

        if ($p.ExitCode -ne 0) {
            Write-Error "WeasyPrint failed (exit $($p.ExitCode)). STDERR:`n$stderr"
            throw "WeasyPrint conversion failed."
        }
    }
    finally {
        Remove-Item $tempHtml -ErrorAction SilentlyContinue
    }
}
#endregion

# ---------------------------
#region Load Environment Variables
# ---------------------------
$TIMEZONE    = (Get-EnvOrDefault 'TIMEZONE' 'America/Edmonton')   # IANA
$OUTPUT_DIR  = (Get-EnvOrDefault 'OUTPUT_DIR' '/data')
$EMPTY_OUTPUT_DIR = (Get-EnvOrDefault 'EMPTY_OUTPUT_DIR' 'false')
$SERVICE_ID  = (Get-EnvOrDefault 'SERVICE_ID')
#$SEARCH_MODE = (Get-EnvOrDefault 'SEARCH_MODE')                   # e.g., 'next-sunday'
#$TITLE_CONTAINS = (Get-EnvOrDefault 'TITLE_CONTAINS')
#$FILE_STEM   = (Get-EnvOrDefault 'FILE_STEM')                       # optional CSV
$CSSPATH    = (Get-EnvOrDefault 'PLAN_CSS_PATH' '/app/print.css')
$PROFILES    = Get-JsonFromEnv 'PLAN_PROFILES'
$PROFILES_ENV  = if ($env:PLAN_PROFILES) { $env:PLAN_PROFILES } else { $null }
$PROFILES_FILE = if ($env:PLAN_PROFILES_FILE) { $env:PLAN_PROFILES_FILE } else { $null }
$KEEP_HTML   = (Get-EnvOrDefault 'KEEP_HTML' 'false')

$SUBSPLASH_USER  = Get-EnvOrDefault 'SUBSPLASH_USERNAME'
$SUBSPLASH_PASS  = Get-EnvOrDefault 'SUBSPLASH_PASSWORD'

# Validate required env vars
if (-not $SUBSPLASH_USER -or -not $SUBSPLASH_PASS) {
    Write-Error "SUBSPLASH_USERNAME/SUBSPLASH_PASSWORD must be set."
    exit 2
}
if (-not (Test-Path -Path $OUTPUT_DIR)) {
    Write-Debug "OUTPUT_DIR '$OUTPUT_DIR' does not exist. Creating it."
    try {
        New-Item -Path $OUTPUT_DIR -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Error "Failed to create OUTPUT_DIR '$OUTPUT_DIR'. $_"
        exit 2
    }
}
if ($EMPTY_OUTPUT_DIR -eq 'true') {
    Write-Debug "EMPTY_OUTPUT_DIR is true. Clearing contents of $OUTPUT_DIR"
    try {
        Get-ChildItem -Path $OUTPUT_DIR -Recurse | Remove-Item -Force -Recurse
    }
    catch {
        Write-Error "Failed to clear OUTPUT_DIR '$OUTPUT_DIR'. $_"
        exit 2
    }
}
# ---------------------------
# Build PSCredential from env
# ---------------------------
$secure = ConvertTo-SecureString $SUBSPLASH_PASS -AsPlainText -Force
$subsplashcreds = [pscredential]::new($SUBSPLASH_USER, $secure)
#endregion
# ---------------------------

#### Main Script Execution
Write-Debug "Starting printplan.ps1 script execution."
#region Verify Subsplash credentials
if (-not $subsplashcreds) {
    Write-Debug "Subsplash credentials not provided."
    Write-Error "Subsplash credentials not provided. Exiting."
    exit 1
}
# Test the credentials by getting an auth token
Write-Debug "Testing Subsplash credentials..."
try {
    $subsplashauth = Get-SubsplashAuthToken -SubsplashCreds $subsplashcreds
}
catch {
    Write-Debug "Function to authenticate with Subsplash API failed."
    Write-Error "Failed to authenticate with Subsplash API. $_"
    exit 1
}
if ($subsplashauth.StatusCode -ne 200) {
    Write-Debug "Subsplash API authentication failed. Status code: $($subsplashauth.StatusCode)"
    Write-Debug "Subsplash API authentication failed. Response: $($subsplashauth.Response | ConvertTo-Json -Depth 10)"
    Write-Error "Failed to authenticate with Subsplash API. Please check your credentials."
    exit 1
}
else {
    Write-Host "Subsplash API authentication successful." -ForegroundColor Green
    $token = $subsplashauth.Token
}
#endregion

#region Get or Search for Service ID
Write-Debug "Get or Search for Service ID"

# Use $SERVICE_ID if provided
if ($SERVICE_ID) {
    Write-Debug "SERVICE_ID environment variable provided: $SERVICE_ID"
    $serviceid = $SERVICE_ID
    Write-Host "Using SERVICE_ID from environment: $serviceid" -ForegroundColor Green
}
else {
    # Search for services
    Write-Debug "No SERVICE_ID provided. Searching for services."
    $now = Get-Date
    $localTimezone = $timezone
    $daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
    $nextSunday = $now.Date.AddDays($daysUntilSunday)
    $endDate = $nextSunday.AddDays(1).AddMilliseconds(-1)
    Write-Debug "Next Sunday: $nextSunday, End Date: $endDate"
    Write-Host "Searching for services from $nextSunday to $endDate in timezone $localTimezone." -ForegroundColor Green

    $filterBody = New-SubsplashServiceFilter -StartDate $nextSunday -EndDate $endDate -Timezone $localTimezone
    Write-Host "Getting list of services from Subsplash API..." -ForegroundColor Green
    $services = Get-SubsplashServices -AuthToken $token -FilterBody $filterBody

    if (-not $services -or $services.Count -eq 0) {
        Write-Error "No services found."
        exit 1
    }

    # Filter only services with plans
    Write-Debug "Filtering services to only those with plans."
    $servicesWithPlans = @()
    foreach ($svc in $services) {
        $svcDetails = Get-SubsplashServiceById -AuthToken $token -ServiceId $svc._id
        if ($svcDetails -and $svcDetails.plans -and $svcDetails.plans.Count -gt 0) {
            $servicesWithPlans += $svc
        }
    }
    $services = $servicesWithPlans
    Write-Debug "Filtered services with plans: $($services.Count)"

    if (-not $services -or $services.Count -eq 0) {
        Write-Error "No services with plans found."
        exit 1
    }

    # If only one service, use its ID
    if ($services.Count -eq 1) {
        $serviceid = $services[0]._id
        Write-Host "Only one service with a plan found. Using Service ID: $serviceid" -ForegroundColor Green
        Write-Host "Service Title: $($services[0].title)" -ForegroundColor Green
    }
    else {
        # More than one service: print list and exit with instructions
        Write-Host "Multiple services with plans found. Please specify a service ID using the SERVICE_ID environment variable or script parameter." -ForegroundColor Yellow
        $services | Select-Object @{Name="Title";Expression={$_.title}}, @{Name="ID";Expression={$_. _id}} | Format-Table -AutoSize
        Write-Host "`nRerun the script with the desired service ID, e.g.:" -ForegroundColor Yellow
        Write-Host "    `$env:SERVICE_ID = <service_id>; .\printplan.ps1" -ForegroundColor Cyan
        exit 0
    }
}
#endregion

#region Get Service Details
Write-Debug "Get Service Details"
    Write-Host "Getting service details for ID: $serviceid" -ForegroundColor Green
    $serviceDetails = Get-SubsplashServiceById -AuthToken $token -ServiceId $serviceid
    if (-not $serviceDetails) {
        Write-Debug "Function to get service details failed. $_"
        Write-Error "Failed to retrieve service details. Exiting."
        exit 1
    }

    # Handle multiple plans
    if ($serviceDetails.plans.Count -gt 1) {
        Write-Debug "Multiple plans found for this service: $($serviceDetails.plans.Count)"
        Write-Error "More than one plan found for this service. Cannot select a plan in non-interactive mode. Exiting."
        exit 1
    } elseif ($serviceDetails.plans.Count -eq 0) {
        Write-Error "No plans found for this service. Exiting."
        exit 1
    }

#endregion

#region Build Profile List
Write-Debug "Plansheet Rendering"
Write-Debug "Building list of plansheet profiles..."
Write-Host "Building list of plansheet profiles..." -ForegroundColor Green
# Determine profiles from env vars or file
# Initialize profiles array
$profiles = @()

if ($PROFILES_ENV) {
    # Highest precedence: inline JSON in env var
    try { $profiles = $PROFILES_ENV | ConvertFrom-Json -Depth 50 }
    catch {
        Write-Error "PLAN_PROFILES env var is not valid JSON. $_"
        exit 2
    }
}
elseif ($PROFILES_FILE) {
    # Second: JSON file path
    try { $profiles = Get-JsonFile -Path $PROFILES_FILE }
    catch {
        Write-Error "Failed to load PLAN_PROFILES_FILE. $_"
        exit 2
    }
}
else {
    # Fallback: one profile with all plan teams
    $profiles = @(@{ Name = "All Teams"; Teams = $serviceDetails.plans[0].teams; orientation = "landscape" })
}

# basic validation
if (-not $profiles -or $profiles.Count -eq 0) {
    Write-Error "No plan profiles resolved. Provide PLAN_PROFILES, PLAN_PROFILES_FILE, TEAMS, or rely on plan teams."
    exit 2
}

Write-Debug "Resolved $($profiles.Count) plansheet profiles."
#endregion
#region Render Plansheet(s)
Write-Debug "Rendering plansheet."
if (-not $serviceDetails -or -not $serviceDetails.plans -or $serviceDetails.plans.Count -eq 0) {
    Write-Error "No service details or plans found. Cannot render plansheet." 
    exit 1
}
#Get complete plan details for the first plan
Write-Host "Getting complete plan details for Plan ID: $($serviceDetails.plans[0]._id)" -ForegroundColor Green
$planDetails = Get-SubsplashPlanDetails -AuthToken $token -PlanId $serviceDetails.plans[0]._id
if (-not $planDetails) {
    Write-Error "Failed to retrieve plan details for Plan ID $($serviceDetails.plans[0]._id). Exiting."
    exit 1
}
Write-Debug "Rendering plansheet for service ID: $serviceid, looping though profiles."
foreach ($planprofile in $profiles) {
    if (-not $planprofile.Teams -or $planprofile.Teams.Count -eq 0) {
        Write-Host "No teams found in profile '$($planprofile.Name)'. Skipping..." -ForegroundColor Red
        continue
    }
    $Teams = $planprofile.Teams
    $safeProfileName = ($planprofile.Name -replace '[^a-zA-Z0-9\-]', '-').Trim('-') -replace '-+', '-'
    $safePlanTitle = ($planDetails.title -replace '[^a-zA-Z0-9\-]', '-').Trim('-') -replace '-+', '-'
    Write-Host "Rendering plansheet for profile '$($planprofile.Name)' with teams: $($Teams -join ', ')" -ForegroundColor Green
    Write-Debug "Teams: $($Teams | ConvertTo-Json -Depth 10)"
    try {
        Write-Debug "Generating HTML for profile '$($planprofile.Name)'..."
        $html = New-PlanHtml -plandetails $planDetails -Teams $Teams -PlanName $planprofile.Name -orientation $planprofile.orientation
    }
    catch {
        Write-Error "Failed to generate HTML for profile '$($planprofile.Name)'. $_"
        continue
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $outputBaseName = "$($safeProfileName)_$($safePlanTitle)_$timestamp"
    $outputPdfPath = Join-Path -Path $OUTPUT_DIR -ChildPath ($outputBaseName + ".pdf")
    $outputHtmlPath = Join-Path -Path $OUTPUT_DIR -ChildPath ($outputBaseName + ".html")

    # Always save PDF
    try {
        Write-Debug "Converting HTML to PDF..."
        Convert-PlanHtmlToPdf -PlanHtml $html -OutPath $outputPdfPath
        Write-Host "Plansheet PDF saved to: $outputPdfPath" -ForegroundColor Magenta
    }
    catch {
        Write-Error "Failed to convert HTML to PDF or save to $outputPdfPath. $_"
    }
    # Save HTML if requested
    if ($KEEP_HTML -eq 'true') {
            try {
                Write-Debug "Writing HTML to file..."
                Set-Content -Path $outputHtmlPath -Value $html -Encoding UTF8
                Write-Host "Plansheet HTML saved to: $outputHtmlPath" -ForegroundColor Magenta        
            }
            catch {
                Write-Error "Failed to write HTML to $outputHtmlPath. $_"
            }
    }
}
#endregion
Write-Debug "Script completed."
Write-Host "Script completed." -ForegroundColor Green