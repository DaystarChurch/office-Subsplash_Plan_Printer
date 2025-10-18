# Subsplash Plan Printing Script (Container Edition)

## Overview

This PowerShell script (`printplan.ps1`) automates the retrieval, formatting, and printing of service plans from the Subsplash Management API. It is designed to run **headless** in a Docker container, with all configuration and credentials supplied via environment variables or a `.env` file. Output is generated as PDF (and optionally HTML), with customizable profiles and team columns.

---

## Features

- **Authenticate with Subsplash Management API** using credentials passed as environment variables.
- **Filter services** based on the next Sunday date or by direct service ID.
- **Select and retrieve service details** (including multiple plans per service).
- **List teams** involved in a plan.
- **Render plansheets** as PDF (and optionally HTML), with customizable profiles and team columns.
- **Configurable output directory, timezone, CSS, and profiles** via environment variables or external JSON files.
- **Supports fully automated, headless operation**—no prompts or interactive steps.

---

## Prerequisites

- **Docker**: All dependencies are packaged in the container image.
- **Internet Access**: The script needs to access the Subsplash Management API.
- **Subsplash Management Credentials**: Supplied via environment variables or `.env` file.
- **Optional**: `profiles.json` for plan profiles, custom CSS file.

---

## Usage

### Docker Compose File

A sample `docker-compose.yml` is provided below. Adjust environment variables and volume mounts as needed.

```yaml
version: "3.9"
services:
  planprinter:
    image: timothiasthegreat/subsplash_plan_printer
    env_file: [.env]
    volumes:
      - ./data:/data
```

### .env File

Create a `.env` file in the same directory as your `docker-compose.yml` to supply configuration variables.

#### Environment Variables

All configuration is supplied via environment variables or a `.env` file. The most common options are:

| Variable            | Description                                                                                   | Example Value                                 |
|---------------------|----------------------------------------------------------------------------------------------|-----------------------------------------------|
| SUBSPLASH_USERNAME      | Subsplash Management API username                                                            | you@example.com                               |
| SUBSPLASH_PASSWORD      | Subsplash Management API password                                                            | supersecret                                   |
| SERVICE_ID          | (Optional) Directly specify a service ID to print                                            | 6532abc123def4567890                          |
| TIMEZONE            | IANA timezone name (default: `America/Edmonton`)                                             | America/Edmonton                              |
| OUTPUT_DIR          | Output directory inside the container (default: `/data`)                                     | /data                                         |
| PLAN_PROFILES       | (Optional) JSON array of plan profiles (see below)                                           | [{"Name":"Media",...}]                        |
| PLAN_PROFILES_FILE  | (Optional) Path to a JSON file with plan profiles (e.g., `/data/profiles.json`)              | /data/profiles.json                           |
| EMPTY_OUTPUT_DIR    | (Optional) If `true`, empties the output directory before generating new files (default: false)| true                                          |
| KEEP_HTML           | (Optional) If `true`, keeps intermediate HTML files alongside PDFs (default: false)           | true                                          |

**Precedence for profiles:**  
`PLAN_PROFILES` (inline JSON) > `PLAN_PROFILES_FILE` (external file) > fallback to “All Teams” from plan.

---

#### Example `.env` file

```dotenv
SUBSPLASH_USERNAME=you@example.com
SUBSPLASH_PASSWORD=supersecret
TIMEZONE=America/Edmonton
OUTPUT_DIR=/data
PLAN_PROFILES_FILE=/data/profiles.json
CSSPATH=/app/print.css
EMPTY_OUTPUT_DIR=false
KEEP_HTML=false
```

---

#### Example `profiles.json`

```json
[
  { "Name": "Media", "Teams": ["Band","NOTES","MEDIA - PROPRESENTER"], "orientation": "portrait" },
  { "Name": "Sound", "Teams": ["Band","NOTES","SOUND"], "orientation": "portrait" },
  { "Name": "All Teams", "Teams": ["Band","NOTES","MEDIA - PROPRESENTER","SOUND","LIVE STREAM SOUND","LIVE STREAM","Lighting"], "orientation": "landscape" }
]
```

---

### Running the Container

```bash
docker compose up --abort-on-container-exit
```
Or directly with Docker:

```bash
docker run --rm --env-file .env -v "$PWD/data:/data" -w /app timothiasthegreat/subsplash_plan_printer:latest
```

---

## Output

- **PDF files** are saved to the output directory (`OUTPUT_DIR`, default `/data`).
- **HTML files** are optionally kept if `KEEP_HTML=true`.
- If `EMPTY_OUTPUT_DIR=true`, the output directory is cleared before each run.

---

## Customization

- **CSS**: Provide a custom CSS file and set `CSSPATH` to its path inside the container.
- **Plan Profiles**: Use `PLAN_PROFILES` (inline JSON) or `PLAN_PROFILES_FILE` (external file) to define multiple output profiles.

---

## Troubleshooting

- **Authentication errors**: Check `SUBSPLASH_USERNAME` and `SUBSPLASH_PASSWORD` in your `.env`.
- **No services found**: Check your date range, credentials, or API access.
- **Custom CSS not applied**: Confirm `CSSPATH` points to a valid file inside the container.
- **Output directory issues**: Make sure your host directory is mounted to `/data` and permissions are correct.

---

## Error Handling

- The script provides clear error messages for missing credentials, failed authentication, missing config, or API errors.
- In headless mode, the script will exit with an error if user interaction is required.

---

## Notes

- All configuration is via environment variables or files mounted into the container.
- No interactive prompts; all runs are headless and suitable for automation.
- Credentials should be managed securely—prefer Docker secrets for production.

---
