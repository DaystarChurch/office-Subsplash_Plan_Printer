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
# Fluro API functions
# ---------------------------
function Get-FluroAuthToken {
    param(
        [Parameter(Mandatory = $true)]
        [pscredential]$FluroCreds
    )
    Write-Debug "Get-FluroAuthToken function called."
    $UserName = $FluroCreds.UserName
    Write-Debug "Username: $UserName"
    if (-not $UserName) {
        Write-Error "Username is empty. Please provide a valid username."
        return
    }
    $Password = $FluroCreds.Password
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
        Write-Debug "Sending authentication request to Fluro API."
        $response = Invoke-RestMethod -Uri "https://api.fluro.io/token/login" -Method Post -Headers $headers -Body $body -StatusCodeVariable statusCode
    }
    Catch {
        Write-Error "Failed to authenticate with Fluro API. $_"
        return $_
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

function Get-FluroPlanDetails {
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
        Write-Debug "Sending request to Fluro API for plan details. URL: $url"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        Write-Debug "Plan details retrieved successfully."
        return $response
    }
    catch {
        Write-Error "Failed to retrieve plan details for Plan ID $PlanId. $_"
        return $null
    }
}

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
    $serviceDateTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($serviceDateTimeUtc, [System.TimeZoneInfo]::FindSystemTimeZoneById("Mountain Standard Time"))
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
    # Check for WeasyPrint exe in script directory
    if (-not (Get-Item -path "weasyprint.exe" -ErrorAction SilentlyContinue)) {
        Write-Error "WeasyPrint executable not found in script directory. Please download WeasyPrint and place weasyprint.exe in the script directory."
        throw "WeasyPrint executable not found."
    }

    # Create a temporary HTML file
    $tempHtml = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'
    Set-Content -Path $tempHtml -Value $PlanHtml -Encoding UTF8
    Write-Debug "Temporary HTML file created at: $tempHtml"

    try {
        # Run WeasyPrint to convert to PDF
        $weasyPrintPath = Join-Path $scriptDir "weasyprint.exe"
        $process = Start-Process -FilePath $weasyPrintPath -ArgumentList "`"$tempHtml`"", "`"$OutPath`"" -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Error "WeasyPrint exited with code $($process.ExitCode). PDF may not have been created."
        }
    }
    catch {
        Write-Error "Failed to run WeasyPrint for PDF conversion: $_"
        throw
    }
    finally {
        # Clean up temp file
        try {
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

$FLURO_USER  = Get-EnvOrDefault 'FLURO_USERNAME'
$FLURO_PASS  = Get-EnvOrDefault 'FLURO_PASSWORD'

# Validate required env vars
if (-not $FLURO_USER -or -not $FLURO_PASS) {
    Write-Error "FLURO_USERNAME/FLURO_PASSWORD must be set."
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
$secure = ConvertTo-SecureString $FLURO_PASS -AsPlainText -Force
$flurocreds = [pscredential]::new($FLURO_USER, $secure)
#endregion
# ---------------------------

#### Main Script Execution
Write-Debug "Starting printplan.ps1 script execution."
#region Verify Fluro credentials
if (-not $flurocreds) {
    Write-Debug "Fluro credentials not provided."
    Write-Error "Fluro credentials not provided. Exiting."
    exit 1 
}
# Test the credentials by getting an auth token
Write-Debug "Testing Fluro credentials..."
try {
    $fluroauth = Get-FluroAuthToken -FluroCreds $flurocreds
}
catch {
    Write-Debug "Function to authenticate with Fluro API failed."
    Write-Error "Failed to authenticate with Fluro API. $_"
    exit 1
}
if ($fluroauth.StatusCode -ne 200) {
    Write-Debug "Fluro API authentication failed. Status code: $($fluroauth.StatusCode)"
    Write-Debug "Fluro API authentication failed. Response: $($fluroauth.Response | ConvertTo-Json -Depth 10)"
    Write-Error "Failed to authenticate with Fluro API. Please check your credentials."
    exit 1
}
else {
    Write-Host "Fluro API authentication successful." -ForegroundColor Green
    $token = $fluroauth.Token
}
#endregion

#region List or Search Services
# If -ListServices is specified, list all services for the next Sunday
# If -serviceid is not specified, search for services for the next Sunday
Write-Debug "List or Search Services"
if ($ListServices -or -not $serviceid) {
    Write-Debug "ListServices or service search triggered."
    # Set up the date range for the search (next Sunday to end of that Sunday)
    $now = Get-Date
    $localTimezone = $timezone
    $daysUntilSunday = (7 - [int]$now.DayOfWeek) % 7
    $nextSunday = $now.Date.AddDays($daysUntilSunday)
    $endDate = $nextSunday.AddDays(1).AddMilliseconds(-1)
    Write-Debug "Next Sunday: $nextSunday, End Date: $endDate"
    Write-Host "Searching for services from $nextSunday to $endDate in timezone $localTimezone." -ForegroundColor Green
    # Create the filter body for the API request
    Write-Debug "Creating filter body for API request."
    $filterBody = New-FluroServiceFilter -StartDate $nextSunday -EndDate $endDate -Timezone $localTimezone
    Write-Host "Getting list of services from Fluro API..." -ForegroundColor Green
    # Get the services from the API
    Write-Debug "Getting services from Fluro API..."
    $services = Get-FluroServices -AuthToken $token -FilterBody $filterBody

    if (-not $services -or $services.Count -eq 0) {
        if ($ListServices) {
            Write-Host "No services found." -ForegroundColor Yellow
            exit 0
        } else {
            Write-Error "No services found."
            exit 1
        }
    }

    if ($ListServices) {
        Write-Debug "Services found: $($services.Count)"
        Write-Debug "Services: $($services | ConvertTo-Json -Depth 10)"
        Write-Debug "Displaying services in grid view."
        $localTZ = [System.TimeZoneInfo]::Local
        $services | Select-Object `
            @{Name="Title";Expression={$_.title}},
            @{Name="Start (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.startDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
            @{Name="End (Local)";Expression={ [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$_.endDate, $localTZ).ToString("yyyy-MM-dd HH:mm") }},
            @{Name="_id";Expression={$_. _id}} |
            Format-Table -AutoSize
        Write-Debug "Exiting after listing services."
        exit 0
    }
    Write-Debug "Services found: $($services.Count)"
    # Filter out services that do not have plans by checking each service with Get-FluroServiceById
    Write-Debug "Filtering services to only those with plans."
    $servicesWithPlans = @()
    foreach ($svc in $services) {
        $svcDetails = Get-FluroServiceById -AuthToken $token -ServiceId $svc._id
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
    # If no service ID is specified, use the ID from the found service, or prompt the user to select a service if multiple are found
    if (-not $serviceid) {
        if ($services.Count -eq 1) {
            # Only one service found, use its ID
            $serviceid = $services[0]._id
            Write-Host "Service ID: $serviceid" -ForegroundColor Green
            Write-Host "Service Title: $($services[0].title)" -ForegroundColor Green
        }
        elseif ($services.Count -gt 1) {
            if ($Headless) {
                Write-Error "Multiple services found. Please run without -headless to select a service."
                exit 1
            }
            Write-Debug "Multiple services found. Displaying in grid view."
            Write-Host "Multiple services found. Please select one:" -ForegroundColor Yellow
            # Display the services in a grid view for selection
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
            Write-Host "Selected Service ID: $serviceid" -ForegroundColor Green
            Write-Host "Selected Service Title: $($selectedService.title)" -ForegroundColor Green
        }
    }
} else {
    Write-Debug "Service ID provided: $serviceid"
    Write-Host "Service ID provided: $serviceid. Skipping service search." -ForegroundColor Green
}
#endregion

#region Get Service Details
Write-Debug "Get Service Details"
if ($serviceid) {
    Write-Debug "Service ID provided: $serviceid"
    Write-Host "Getting service details for ID: $serviceid" -ForegroundColor Green
    $serviceDetails = Get-FluroServiceById -AuthToken $token -ServiceId $serviceid
    if (-not $serviceDetails) {
        Write-Debug "Function to get service details failed. $_"
        Write-Error "Failed to retrieve service details. Exiting."
        exit 1
    }

    # Handle multiple plans
    if ($serviceDetails.plans.Count -gt 1) {
        if ($Headless) {
            Write-Debug "Headless mode specified. More than one plan found for this service."
            Write-Error "More than one plan found for this service. Cannot select a plan in headless mode. Exiting."
            exit 1
        } else {
            Write-Warning "More than one plan found for this service. Please select one:"
            write-Debug "Displaying plans in grid view."
            $plansDisplay = $serviceDetails.plans | Select-Object `
                @{Name="Index";Expression={ [array]::IndexOf($serviceDetails.plans, $_) }},
                title,
                startDate
            $selectedPlan = $plansDisplay | Out-GridView -Title "Select a Plan" -PassThru
            if (-not $selectedPlan) {
                Write-Error "No plan selected. Exiting."
                exit 1
            }
            # Set $serviceDetails.plans to only the selected plan
            Write-Debug "Selected plan: $($selectedPlan.title)"
            $serviceDetails.plans = @($serviceDetails.plans[$selectedPlan.Index])
        }
    } elseif ($serviceDetails.plans.Count -eq 0) {
        Write-Error "No plans found for this service. Exiting."
        exit 1
    }
}
else {
    Write-Error "No service ID provided. Exiting."
    exit 1
}
#endregion

#region List Teams
if ($ListTeams) {
    Write-Debug "ListTeams parameter specified. Listing teams."
    if (-not $serviceDetails -or -not $serviceDetails.plans -or $serviceDetails.plans.Count -eq 0) {
        Write-Error "No service details or plans found. Cannot list teams."
        exit 1
    }
    $teams = $serviceDetails.plans[0].teams
    if (-not $teams -or $teams.Count -eq 0) {
        Write-Host "No teams found in the plan." -ForegroundColor Yellow
    } else {
        Write-Host "Teams in this plan titled '$($serviceDetails.plans[0].title)':"
        $teams | ForEach-Object { Write-Host "- $_" -ForegroundColor White}
    }
    exit 0
}
#endregion

#region Render Plansheet HTML
# Get list of plansheet profiles
Write-Debug "Plansheet Rendering"
Write-Debug "Getting list of plansheet profiles..."
Write-Host "Getting list of plansheet profiles..." -ForegroundColor Green
if ($config -and $config.planprofiles) {
    $planprofilelist = $config.planprofiles
    Write-Host "Loaded plan profiles from config." -ForegroundColor Green
    Write-Debug "Plan profiles: $($planprofilelist | ConvertTo-Json -Depth 10)"
} else {
    # Default: single profile with all teams from the plan
    Write-Debug "No plan profiles found in config. Using default profile with all teams."
    $allTeams = $serviceDetails.plans[0].teams
    $planprofilelist = @(@{ Name = "All Teams"; Teams = $allTeams })
    Write-Host "No plan profiles found in config. Using default profile with all teams." -ForegroundColor Yellow
}
if ($PrintPlan) {
    Write-Debug "PrintPlan parameter specified. Rendering plansheet."
    if (-not $serviceDetails -or -not $serviceDetails.plans -or $serviceDetails.plans.Count -eq 0) {
        Write-Error "No service details or plans found. Cannot render plansheet." 
        exit 1
    }
    #Get complete plan details for the first plan
    Write-Host "Getting complete plan details for Plan ID: $($serviceDetails.plans[0]._id)" -ForegroundColor Green
    $planDetails = Get-FluroPlanDetails -AuthToken $token -PlanId $serviceDetails.plans[0]._id
    if (-not $planDetails) {
        Write-Error "Failed to retrieve plan details for Plan ID $($serviceDetails.plans[0]._id). Exiting."
        exit 1
    }
    Write-Debug "Rendering plansheet for service ID: $serviceid, looping though profiles."
    foreach ($planprofile in $planprofilelist) {
        if ($planprofile.Teams.Count -eq 0) {
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
        #write-Debug "$html"
        if ($Headless) {
            Write-Debug "Headless mode specified. Saving to PDF."
            Write-Host "Rendering plansheet in headless mode. Saving to PDF..." -ForegroundColor Green
            $outputFileName = "$($safeProfileName)_$($safePlanTitle)_$(Get-Date -Format 'yyyyMMdd_HHmm').pdf"
            $outputPath = Join-Path -Path $outputdir -ChildPath $outputFileName
            Write-Debug "Output path: $outputPath"
            try {
                Write-Debug "Converting HTML to PDF..."
                Convert-PlanHtmlToPdf -PlanHtml $html -OutPath $outputPath
                Write-Host "Plansheet saved to: $outputPath" -ForegroundColor Magenta
            }
            catch {
                Write-Debug "Failed to convert HTML to PDF. $_"
                Write-Error "Failed to convert HTML to PDF or save to $outputPath. $_"
                continue
            }
        } else {
            Write-Debug "GUI mode specified. Saving to HTML."
            Write-Host "Rendering plansheet in GUI mode. Opening in browser..." -ForegroundColor Green
            $outputFileName = "$($safeProfileName)_$($safePlanTitle)_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
            $outputPath = Join-Path -Path $outputdir -ChildPath $outputFileName
            Write-Debug "Output path: $outputPath"
            try {
                Write-Debug "Writing HTML to file..."
                Set-Content -Path $outputPath -Value $html -Encoding UTF8
            }
            catch {
                Write-Debug "Failed to write HTML to file. $_"
                Write-Error "Failed to write HTML to $outputPath. $_"
                continue
            }
            Write-Host "Plansheet saved to: $outputPath" -ForegroundColor Magenta
            # Open the HTML file in MSEdge
            try {
                Write-Debug "Opening HTML file in browser..."
                Start-Process "msedge" -ArgumentList "--new-tab", "`"$outputPath`""
            }
            catch {
                Write-Error "Failed to open $outputPath in browser. $_"
            }
        }
    }
}
#endregion
Write-Debug "Script completed."
Write-Host "Script completed." -ForegroundColor Green