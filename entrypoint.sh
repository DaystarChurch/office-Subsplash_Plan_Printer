#!/bin/bash
set -e

# Set timezone if TIMEZONE env variable is provided
if [ -n "$TIMEZONE" ]; then
  echo "Setting timezone to $TIMEZONE"
  ln -snf "/usr/share/zoneinfo/$TIMEZONE" /etc/localtime
  echo "$TIMEZONE" > /etc/timezone
else
  echo "TIMEZONE env variable not set; using default container timezone."
fi

# Launch PowerShell script
exec pwsh -File /app/printplan.ps1
