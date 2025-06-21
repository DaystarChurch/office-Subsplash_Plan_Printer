# Subsplash Plan Printing Script Documentation

## Overview

This PowerShell script (printplan.ps1) automates the retrieval, formatting, and printing of service plans from the Subsplash Management API. It supports interactive and headless (automated) operation, credential management, configuration via JSON, and output as HTML or PDF. The script is designed for teams who need to generate printable plansheets for services, including team assignments and schedules.

---

## Features

- **Authenticate with Subsplash Management API** using securely stored credentials.
- **Filters services** based on the next Sunday date.
- **Select and retrieve service details** (including multiple plans per service).
- **List teams** involved in a plan.
- **Render plansheets** as HTML or PDF, with customizable profiles and team columns.
- **Configurable output directory, timezone, and profiles** via a JSON config file.
- **Supports both interactive and headless (automation) modes.**

---

## Prerequisites

- **PowerShell 7+**: Ensure you have PowerShell 7 or later installed.
- **Microsoft Edge**: Required for PDF generation and HTML preview.
- **Internet Access**: The script needs to access the Subsplash Management API.
- **Credentials**: You must have valid Subsplash Management credentials stored securely.
- **JSON Config File**: Optional, for advanced configuration (see below).

---

## Usage

### Parameters

| Parameter         | Type      | Description                                                                 |
|-------------------|-----------|-----------------------------------------------------------------------------|
| `-LoginSubsplash` | Switch    | Prompts for Subsplash Management credentials and saves them securely.                      |
| `-ListTeams`      | Switch    | Lists all teams in the selected plan and exits.                             |
| `-ListServices`   | Switch    | Lists all services for the next Sunday and exits.                           |
| `-PrintSongs`     | Switch    | (Reserved for future use.)                                                  |
| `-PrintPlan`      | Switch    | Generates and outputs the plansheet (HTML or PDF).                          |
| `-Headless`       | Switch    | Runs in headless mode (no GUI prompts, outputs PDF).                        |
| `-serviceid`      | String    | Specify a Subsplash Management service ID directly (skips service selection).              |
| `-Teams`          | String[]  | Specify which teams to include in the plansheet (overrides config/profile).  |
| `-configpath`     | String    | Path to a JSON config file for advanced options.                            |

---

### Example Commands

*Service ID can be ommitted if there is only one service available, or you can specify it directly.*

- **Set up credentials:**

  ```powershell
  .\printplan.ps1 -LoginSubsplash
  ```

- **List services for next Sunday:**

  ```powershell
  .\printplan.ps1 -ListServices
  ```

- **List teams for a specific service:**

  ```powershell
  .\printplan.ps1 -serviceid "<SERVICE_ID>" -ListTeams
  ```

- **Print plansheet for a service (interactive HTML):**

  ```powershell
  .\printplan.ps1 -serviceid "<SERVICE_ID>" -PrintPlan
  ```

- **Print plansheet for a service (headless PDF):**

  ```powershell
  .\printplan.ps1 -serviceid "<SERVICE_ID>" -PrintPlan -Headless
  ```

- **Use a custom config file:**

  ```powershell
  .\printplan.ps1 -configpath "C:\path\to\config.json" -PrintPlan
  ```

- **Schedule a task to run weekly: (may require admin rights on computer)**

  ```powershell
  $action = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-File C:\path\to\printplan.ps1 -Headless -PrintPlan -configpath C:\path\to\config.json"
  $trigger = New-ScheduledTaskTrigger -Weekly -DaysofWeek Sunday -At 8am
  Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "WeeklyPlanPrint" -User "SYSTEM"
  ```

---

## Configuration File (`config.json`)

You can provide a JSON config file to customize script behavior. Example:

```json
{
  "timezone": "America/Edmonton",
  "destinationpath": "C:\\Plansheets",
  "planprofiles": [
    { "Name": "All Teams", "orientation": "landscape", "Teams": ["Band", "Tech", "Host"] },
    { "Name": "Band Only", "orientation": "landscape", "Teams": ["Band"] }
  ]
}
```

- **timezone**: IANA or Windows timezone ID (default: `America/Edmonton`)
- **destinationpath**: Output directory for generated files
- **planprofiles**: Array of profiles, each with a name, print orientation (landscape or portrait) and list of teams to include as columns

---

## Credential Management

- Credentials are stored securely in cred.xml (encrypted for the current user).
- Use `-LoginSubsplash` to set or update credentials.
- If credentials are missing and not in headless mode, you will be prompted to enter them.
- `-headless` mode will exit with an error if credentials are not found.

---

## Output

- **HTML**: Opens in your default browser (Edge) for review/printing.
- **PDF**: Saved directly to the output directory (requires Microsoft Edge installed).

---

## Error Handling

- The script provides clear error messages for missing credentials, failed authentication, missing config, or API errors.
- In headless mode, the script will exit with an error if user interaction is required.

---

## Customization

- You can modify the default CSS by providing a print.css file in the script directory or via the config.
- Plan profiles allow you to generate multiple plansheets with different team columns in one run.

---

## Troubleshooting

- **Authentication errors**: Re-run with `-LoginSubsplash` to reset credentials.
- **No services found**: Check your date range, credentials, or API access.
- **PDF not generated**: Ensure Edge is installed and accessible via `msedge` command.

---
