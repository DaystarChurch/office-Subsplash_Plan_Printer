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
    Write-Log "Failed to parse JSON from environment variable '$Name'. Error details: $_" -Level "ERROR"
    Write-Error "Environment variable '$Name' ('$envVar') is not valid JSON. Please check the value and format. Error details: $_"
    exit 2
    }
}

function Get-JsonFile {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path -Path $Path)) {
    Write-Log "PLAN_PROFILES_FILE points to '$Path' but the file does not exist." -Level "ERROR"
    Write-Error "The PLAN_PROFILES_FILE environment variable points to '$Path', but the file does not exist. Please check the path and try again."
        exit 2
    }
    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        return ($raw | ConvertFrom-Json -Depth 50)
    } catch {
    Write-Log "Failed to parse JSON from file '$Path'. Error details: $_" -Level "ERROR"
    Write-Error "Could not parse JSON from file '$Path'. Please check the file contents and format. Error details: $_"
        exit 2
    }
}
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    # Define log level order
    $levels = @("DEBUG", "INFO", "WARNING", "ERROR", "VERBOSE")
    $levelOrder = @{
        "VERBOSE" = 0
        "DEBUG"   = 1
        "INFO"    = 2
        "WARNING" = 3
        "ERROR"   = 4
    }
    # Default to INFO if $LOGLEVEL is not set or invalid
    $currentLevel = $levelOrder[$LOGLEVEL]
    if ($null -eq $currentLevel) { $currentLevel = $levelOrder["INFO"] }
    $msgLevel = $levelOrder[$Level]
    if ($null -eq $msgLevel) { $msgLevel = $levelOrder["INFO"] }

    # Only log if message level is >= configured level
    if ($msgLevel -ge $currentLevel) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] - [$Level] - $Message"
        try {
            Add-Content -Path $LOGPATH -Value $logEntry
        }
        catch {
            Write-Host "Failed to write to log file '$LOGPATH'. $_" -ForegroundColor Red
            Write-Error "Write-Log: Could not write to log file '$LOGPATH'. $_"
        }
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
    Write-Log "Get-SubsplashAuthToken function called." -Level "DEBUG"
    $UserName = $SubsplashCreds.UserName
    Write-Log "Username: $UserName" -Level "DEBUG"
    if (-not $UserName) {
        Write-Log "Username is empty. Please provide a valid username." -Level "ERROR"
        return
    }
    $Password = $SubsplashCreds.Password
    if (-not $Password) {
        Write-Log "Password is empty. Please provide a valid password." -Level "ERROR"
        return
    }
    Write-Log "Converting password to plain text for auth string." -Level "DEBUG"
    $Password = $($Password | ConvertFrom-SecureString -AsPlainText)
    Write-Log "Constructing body for authentication request." -Level "DEBUG"
    $body = "grant_type=password&username=$([uri]::EscapeDataString($UserName))&password=$([uri]::EscapeDataString($Password))"
    Write-Log "Auth Body: $body" -Level "VERBOSE"
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "User-Agent"   = "PowerShell"
        "Accept"       = "*/*"
    }
    Write-Log "Auth Headers: $($headers | ConvertTo-Json -Depth 10)" -Level "VERBOSE"

    Try {
        Write-Log "Sending authentication request to Subsplash API." -Level "DEBUG"
        $response = Invoke-RestMethod -Uri "https://api.fluro.io/token/login" -Method Post -Headers $headers -Body $body -StatusCodeVariable statusCode
        Write-Log "Authentication successful. Status Code: $statusCode" -Level "DEBUG"
        Write-Log "Auth Response: $($response | ConvertTo-Json -Depth 10)" -Level "VERBOSE"
    }
    Catch {
        Write-Log "Failed to authenticate with Subsplash API. $_" -Level "ERROR"
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
    Write-Log "New-SubsplashServiceFilter function called. StartDate: $StartDate, EndDate: $EndDate, Timezone: $Timezone" -Level "DEBUG"
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
    Write-Log "Get-SubsplashServices function called. Response count: $($response.items.Count)" -Level "DEBUG"
    Write-Log "Get-SubsplashServices function called. Response: $($response | ConvertTo-Json -Depth 10)" -Level "VERBOSE"

    return $response
}

function Get-SubsplashServiceById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AuthToken,
        [Parameter(Mandatory = $true)]
        [string]$ServiceId
    )
    Write-Log "Get-SubsplashServiceById function called. ServiceId: $ServiceId" -Level "DEBUG"
    $headers = @{
        "Authorization" = "Bearer $AuthToken"
        "Content-Type"  = "application/json"
        "Accept"        = "*/*"
    }

    $url = "https://api.fluro.io/content/get/$ServiceId"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        Write-Log "Service with ID $ServiceId retrieved successfully." -Level "DEBUG"
        Write-Log "Service details: $($response | ConvertTo-Json -Depth 10)" -Level "VERBOSE"
        return $response
    }
    catch {
        Write-Log "Failed to retrieve service with ID $ServiceId from Subsplash API." -Level "ERROR"
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
        Write-Log "Sending request to Subsplash API for plan details. URL: $url" -Level "DEBUG"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        Write-Log "Plan details retrieved successfully." -Level "DEBUG"
        return $response
    }
    catch {
        Write-Log "Failed to retrieve plan details for Plan ID $PlanId. $_" -Level "ERROR"
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
    Write-Log "New-PlanHtml function called." -Level "DEBUG"
    Write-Log "PlanName: $PlanName" -Level "DEBUG"
    Write-Log "Orientation: $orientation" -Level "DEBUG"
    Write-Log "Teams: $($Teams | ConvertTo-Json -Depth 10)" -Level "DEBUG"
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
    Write-Log "Trying to load CSS from file. Path: $CssPath" -Level "DEBUG"
    $cssContent = ""
    if (Test-Path $CssPath) {
        Write-Log "CSS file found. Loading content." -Level "DEBUG"
        $cssContent = Get-Content $CssPath -Raw
    } else {
        Write-Log "CSS file not found. Using default CSS." -Level "INFO"
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
    Write-Log "Convert-PlanHtmlToPdf function called. OutPath: $OutPath" -Level "DEBUG"
    # Create a temp HTML file
    $tempHtml = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'
    Set-Content -Path $tempHtml -Value $PlanHtml -Encoding UTF8
    Write-Log "Temporary HTML file created: $tempHtml" -Level "DEBUG"

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
            Write-Log "WeasyPrint failed (exit $($p.ExitCode)). STDERR:`n$stderr" -Level "ERROR"
            throw "WeasyPrint conversion failed."
        }
    }
    finally {
        Remove-Item $tempHtml -ErrorAction SilentlyContinue
        Write-Log "Temporary HTML file removed: $tempHtml" -Level "DEBUG"
    }
}
#endregion

# ---------------------------
#region Load Environment Variables
# ---------------------------
$LOGPATH   = (Get-EnvOrDefault 'LOGPATH' '/data/printplan.log')
$LOGLEVEL  = (Get-EnvOrDefault 'LOGLEVEL' 'INFO')  # DEBUG, INFO, WARNING, ERROR
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
    Write-Log "SUBSPLASH_USERNAME/SUBSPLASH_PASSWORD must be set. Cannot continue without credentials." -Level "ERROR"
    Write-Error "Environment variables SUBSPLASH_USERNAME and SUBSPLASH_PASSWORD must be set. Please provide valid credentials and try again."
    exit 2
}
if (-not (Test-Path -Path $OUTPUT_DIR)) {
    Write-Host "OUTPUT_DIR '$OUTPUT_DIR' does not exist. Attempting to create it..." -ForegroundColor Yellow
    Write-Log "OUTPUT_DIR '$OUTPUT_DIR' does not exist. Attempting to create directory." -Level "INFO"
    try {
        New-Item -Path $OUTPUT_DIR -ItemType Directory -Force | Out-Null
        Write-Log "OUTPUT_DIR '$OUTPUT_DIR' created successfully." -Level "DEBUG"
    }
    catch {
    Write-Log "Failed to create OUTPUT_DIR '$OUTPUT_DIR'. Error details: $_" -Level "ERROR"
    Write-Error "Could not create output directory '$OUTPUT_DIR'. Please check permissions and available disk space. Error details: $_"
        exit 2
    }
}
if ($EMPTY_OUTPUT_DIR -eq 'true') {
    Write-Host "EMPTY_OUTPUT_DIR is true. Clearing contents of $OUTPUT_DIR" -ForegroundColor Yellow
    Write-Log "EMPTY_OUTPUT_DIR is true. Clearing contents of $OUTPUT_DIR" -Level "INFO"
    try {
        Get-ChildItem -Path $OUTPUT_DIR -Recurse | Remove-Item -Force -Recurse
        Write-Log "OUTPUT_DIR '$OUTPUT_DIR' cleared successfully." -Level "DEBUG"
    }
    catch {
    Write-Log "Failed to clear OUTPUT_DIR '$OUTPUT_DIR'. Error details: $_" -Level "ERROR"
    Write-Error "Could not clear contents of output directory '$OUTPUT_DIR'. Please check permissions. Error details: $_"
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
Write-Log "Starting printplan.ps1 script execution."
#region Verify Subsplash credentials
if (-not $subsplashcreds) {
    Write-Log "Subsplash credentials object not created. Check SUBSPLASH_USERNAME and SUBSPLASH_PASSWORD." -Level "ERROR"
    Write-Error "Subsplash credentials not provided or invalid. Please check SUBSPLASH_USERNAME and SUBSPLASH_PASSWORD environment variables."
    exit 1
}
# Test the credentials by getting an auth token
Write-Log "Testing Subsplash credentials..." -Level "DEBUG"
Write-Host "Authenticating with Subsplash API..." -ForegroundColor Green
try {
    $subsplashauth = Get-SubsplashAuthToken -SubsplashCreds $subsplashcreds
}
catch {
    Write-Log "Function to authenticate with Subsplash API failed. Error details: $_" -Level "ERROR"
    Write-Error "Subsplash API authentication failed for user '$SUBSPLASH_USER'. Please verify credentials and network connectivity. Error details: $_"
    exit 1
}
if ($subsplashauth.StatusCode -ne 200) {
    Write-Log "Subsplash API authentication failed. Status code: $($subsplashauth.StatusCode). Response: $($subsplashauth.Response | ConvertTo-Json -Depth 10)" -Level "ERROR"
    Write-Error "Subsplash API authentication failed for user '$SUBSPLASH_USER'. Status code: $($subsplashauth.StatusCode). Please check credentials and try again."
    exit 1
}
else {
    Write-Log "Subsplash API authentication successful."
    Write-Host "Subsplash API authentication successful." -ForegroundColor Green
    $token = $subsplashauth.Token
}
#endregion

#region Get or Search for Service ID
Write-Log "Get or Search for Service ID"

# Use $SERVICE_ID if provided
if ($SERVICE_ID) {
    Write-Log "Using SERVICE_ID from environment: $SERVICE_ID" -Level "INFO"
    $serviceid = $SERVICE_ID
    Write-Host "Using SERVICE_ID from environment: $serviceid" -ForegroundColor Green
}
else {
    # Search for services
    Write-Log "No SERVICE_ID provided. Searching for services."
    $now = Get-Date
    Write-Log "Current date/time: $now" -Level "VERBOSE"
    $localTimezone = $timezone
    $daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
    $nextSunday = $now.Date.AddDays($daysUntilSunday)
    $endDate = $nextSunday.AddDays(1).AddMilliseconds(-1)
    Write-Log "Next Sunday: $nextSunday, End Date: $endDate" -Level "VERBOSE"
    Write-Host "Searching for services from $nextSunday to $endDate in timezone $localTimezone." -ForegroundColor Green

    $filterBody = New-SubsplashServiceFilter -StartDate $nextSunday -EndDate $endDate -Timezone $localTimezone
    Write-Log "Service Filter Body: $($filterBody | ConvertTo-Json -Depth 10)" -Level "DEBUG"
    Write-Log "Getting list of services from Subsplash API..." -Level "INFO"
    Write-Host "Getting list of services from Subsplash API..." -ForegroundColor Green
    $services = Get-SubsplashServices -AuthToken $token -FilterBody $filterBody

    if (-not $services -or $services.Count -eq 0) {
    Write-Log "No services found in the specified date range ($nextSunday to $endDate, timezone: $localTimezone)." -Level "ERROR"
    Write-Error "No services found in the specified date range ($nextSunday to $endDate, timezone: $localTimezone). Please check the Subsplash API or adjust your search criteria."
        exit 1
    }

    # Filter only services with plans
    Write-Log "Filtering services to only those with plans." -Level "DEBUG"
    $servicesWithPlans = @()
    foreach ($svc in $services) {
        $svcDetails = Get-SubsplashServiceById -AuthToken $token -ServiceId $svc._id
        if ($svcDetails -and $svcDetails.plans -and $svcDetails.plans.Count -gt 0) {
            $servicesWithPlans += $svc
        }
    }
    $services = $servicesWithPlans
    Write-Log "Filtered services with plans: $($services.Count)" -Level "DEBUG"

    if (-not $services -or $services.Count -eq 0) {
    Write-Log "No services with plans found in the specified date range ($nextSunday to $endDate, timezone: $localTimezone)." -Level "ERROR"
    Write-Error "No services with plans found in the specified date range ($nextSunday to $endDate, timezone: $localTimezone). Please check the Subsplash API or adjust your search criteria."
        exit 1
    }

    # If only one service, use its ID
    if ($services.Count -eq 1) {
        $serviceid = $services[0]._id
    Write-Log "Only one service with a plan found. Using Service ID: $serviceid. Service Title: $($services[0].title)" -Level "INFO"
    Write-Host "Only one service with a plan found. Using Service ID: $serviceid (Title: $($services[0].title))" -ForegroundColor Green
    }
    else {
        # More than one service: print list and exit with instructions
        Write-Log "Multiple services with plans found. Please specify a service ID using the SERVICE_ID environment variable or script parameter." -Level "INFO"
        Write-Host "Multiple services with plans found. Please specify a service ID using the SERVICE_ID environment variable or script parameter." -ForegroundColor Yellow
        $services | Select-Object @{Name="Title";Expression={$_.title}}, @{Name="ID";Expression={$_. _id}} | Format-Table -AutoSize
        Write-Log "Service List:`n$($services | Select-Object @{Name="Title";Expression={$_.title}}, @{Name="ID";Expression={$_. _id}} | Out-String)" -Level "INFO"
        Write-Host "`nRerun the script with the desired service ID, e.g.:" -ForegroundColor Yellow
        Write-Host "    `$env:SERVICE_ID = <service_id>; .\printplan.ps1" -ForegroundColor Cyan
        Write-Log "Exiting script due to multiple services found." -Level "INFO"
        exit 0
    }
}
#endregion

#region Get Service Details
Write-Log "Get Service Details" -Level "INFO"
    Write-Log "Getting service details for ID: $serviceid" -Level "DEBUG"
    Write-Host "Getting service details for ID: $serviceid" -ForegroundColor Green
    $serviceDetails = Get-SubsplashServiceById -AuthToken $token -ServiceId $serviceid
    Write-Log "Service Details: $($serviceDetails | ConvertTo-Json -Depth 10)" -Level "VERBOSE"
    if (-not $serviceDetails) {
    Write-Log "Function to get service details for ID '$serviceid' failed. Error details: $_" -Level "ERROR"
    Write-Error "Failed to retrieve service details for ID '$serviceid'. Please check the Subsplash API and network connectivity. Error details: $_"
        exit 1
    }

    # Handle multiple plans
    if ($serviceDetails.plans.Count -gt 1) {
    Write-Log "Multiple plans ($($serviceDetails.plans.Count)) found for service ID '$serviceid'. Cannot process multiple plans in non-interactive mode." -Level "ERROR"
    Write-Error "More than one plan found for service ID '$serviceid'. Please specify a plan or run in interactive mode."
        exit 1
    } elseif ($serviceDetails.plans.Count -eq 0) {
    Write-Error "No plans found for service ID '$serviceid'. Please check the Subsplash API or service configuration."
    Write-Log "No plans found for service ID '$serviceid'. Exiting." -Level "ERROR"
        exit 1
    }

#endregion

#region Build Profile List
Write-Log "Building list of plansheet profiles..." -Level "INFO"
Write-Host "Building list of plansheet profiles..." -ForegroundColor Green
# Determine profiles from env vars or file
# Initialize profiles array
$profiles = @()

if ($PROFILES_ENV) {
    Write-Log "Using PLAN_PROFILES env var." -Level "INFO"
    # Highest precedence: inline JSON in env var
    try { $profiles = $PROFILES_ENV | ConvertFrom-Json -Depth 50 }
    catch {
        Write-Log "Failed to parse PLAN_PROFILES env var as JSON. $_" -Level "ERROR"
        Write-Error "PLAN_PROFILES env var is not valid JSON. $_"
        exit 1
    }
        exit 1
}
elseif ($PROFILES_FILE) {
    # Second: JSON file path
    Write-Log "Using PLAN_PROFILES_FILE at path: $PROFILES_FILE" -Level "INFO"
    try { $profiles = Get-JsonFile -Path $PROFILES_FILE }
    catch {
        Write-Log "Failed to load PLAN_PROFILES_FILE. $_" -Level "ERROR"
        Write-Error "Failed to load PLAN_PROFILES_FILE. $_"
        exit 2
    }
    Write-Log "Parsed PLAN_PROFILES_FILE: $($profiles | ConvertTo-Json -Depth 10)" -Level "DEBUG"
}
else {
    # Fallback: one profile with all plan teams
    Write-Log "No PLAN_PROFILES or PLAN_PROFILES_FILE provided. Using all plan teams as single profile." -Level "INFO"
    $profiles = @(@{ Name = "All Teams"; Teams = $serviceDetails.plans[0].teams; orientation = "landscape" })
}

# basic validation
if (-not $profiles -or $profiles.Count -eq 0) {
    Write-Log "No plan profiles could be resolved. Please set PLAN_PROFILES, PLAN_PROFILES_FILE, or ensure the plan contains team information." -Level "ERROR"
    Write-Error "No plan profiles could be resolved. Please set PLAN_PROFILES, PLAN_PROFILES_FILE, or ensure the plan contains team information."
    exit 2
}

Write-Log "Resolved $($profiles.Count) plansheet profiles." -Level "INFO"
#endregion
#region Render Plansheet(s)
Write-Log "Rendering plansheet." -Level "INFO"
if (-not $serviceDetails -or -not $serviceDetails.plans -or $serviceDetails.plans.Count -eq 0) {
    Write-Log "No service details or plans found. Cannot render plansheet. Check previous errors for details." -Level "ERROR"
    Write-Error "No service details or plans found. Cannot render plansheet. Please check previous errors for details."
    exit 1
}
#Get complete plan details for the first plan
Write-Log "Getting complete plan details for Plan ID: $($serviceDetails.plans[0]._id)" -Level "INFO"
Write-Host "Getting complete plan details for Plan ID: $($serviceDetails.plans[0]._id)" -ForegroundColor Green
$planDetails = Get-SubsplashPlanDetails -AuthToken $token -PlanId $serviceDetails.plans[0]._id
if (-not $planDetails) {
    Write-Log "Failed to retrieve plan details for Plan ID $($serviceDetails.plans[0]._id). Error details: $_" -Level "ERROR"
    Write-Error "Failed to retrieve plan details for Plan ID $($serviceDetails.plans[0]._id). Please check the Subsplash API and network connectivity. Error details: $_"
    exit 1
}
Write-Log "Rendering plansheet for service ID: $serviceid, looping though profiles." -Level "DEBUG"
foreach ($planprofile in $profiles) {
    if (-not $planprofile.Teams -or $planprofile.Teams.Count -eq 0) {
        Write-Log "No teams found in profile '$($planprofile.Name)'. Skipping profile. Teams property is empty or missing." -Level "WARNING"
        Write-Host "No teams found in profile '$($planprofile.Name)'. Skipping this profile. Please check the profile configuration." -ForegroundColor Yellow
        continue
    }
    $Teams = $planprofile.Teams
    $safeProfileName = ($planprofile.Name -replace '[^a-zA-Z0-9\-]', '-').Trim('-') -replace '-+', '-'
    $safePlanTitle = ($planDetails.title -replace '[^a-zA-Z0-9\-]', '-').Trim('-') -replace '-+', '-'
    Write-Log "Rendering plansheet for profile '$($planprofile.Name)' with teams: $($Teams -join ', ')" -Level "INFO"
    Write-Log "Safe Profile Name: $safeProfileName, Safe Plan Title: $safePlanTitle" -Level "DEBUG"
    Write-Host "Rendering plansheet for profile '$($planprofile.Name)' with teams: $($Teams -join ', ')" -ForegroundColor Green
    Write-Log "Teams: $($Teams | ConvertTo-Json -Depth 10)" -Level "DEBUG"
    try {
        Write-Log "Generating HTML for profile '$($planprofile.Name)'..." -Level "DEBUG"
        $html = New-PlanHtml -plandetails $planDetails -Teams $Teams -PlanName $planprofile.Name -orientation $planprofile.orientation
    }
    catch {
        Write-Log "Failed to generate HTML for profile '$($planprofile.Name)'. Teams: $($Teams -join ', '). Error details: $_" -Level "WARNING"
        Write-Host "Failed to generate HTML for profile '$($planprofile.Name)'. Teams: $($Teams -join ', '). Please check the plan and profile configuration. Error details: $_" -ForegroundColor Yellow
        continue
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $outputBaseName = "$($safeProfileName)_$($safePlanTitle)_$timestamp"
    $outputPdfPath = Join-Path -Path $OUTPUT_DIR -ChildPath ($outputBaseName + ".pdf")
    $outputHtmlPath = Join-Path -Path $OUTPUT_DIR -ChildPath ($outputBaseName + ".html")
Write-Log "Output PDF Path: $outputPdfPath" -Level "DEBUG"
Write-Log "Output HTML Path: $outputHtmlPath" -Level "DEBUG"
    # Always save PDF
    try {
        Write-Log "Converting HTML to PDF..." -Level "DEBUG"
        Convert-PlanHtmlToPdf -PlanHtml $html -OutPath $outputPdfPath
        Write-Host "Plansheet PDF saved to: $outputPdfPath" -ForegroundColor Magenta
        Write-Log "Plansheet PDF saved to: $outputPdfPath" -Level "INFO"
    }
    catch {
        Write-Log "Failed to convert HTML to PDF or save to '$outputPdfPath'. Profile: '$($planprofile.Name)'. Error details: $_" -Level "WARNING"
        Write-Host "Failed to convert HTML to PDF or save to '$outputPdfPath'. Profile: '$($planprofile.Name)'. Please check the HTML content and output directory. Error details: $_" -ForegroundColor Yellow
    }
    # Save HTML if requested
    if ($KEEP_HTML -eq 'true') {
        Write-Log "KEEP_HTML is true. Saving HTML file." -Level "DEBUG"
        Write-Log "Saving HTML to: $outputHtmlPath" -Level "INFO"
        try {
            Write-Log "Writing HTML to file..." -Level "DEBUG"
            Set-Content -Path $outputHtmlPath -Value $html -Encoding UTF8
            Write-Host "Plansheet HTML saved to: $outputHtmlPath" -ForegroundColor Magenta        
            Write-Log "Plansheet HTML saved to: $outputHtmlPath" -Level "INFO"
        }
        catch {
            Write-Log "Failed to write HTML to '$outputHtmlPath'. Profile: '$($planprofile.Name)'. Error details: $_" -Level "WARNING"
            Write-Host "Failed to write HTML to '$outputHtmlPath'. Profile: '$($planprofile.Name)'. Please check the output directory and file permissions. Error details: $_" -ForegroundColor Yellow
        }
    }
}
#endregion
Write-Log "Script completed." -Level "INFO"
Write-Host "Script completed." -ForegroundColor Green