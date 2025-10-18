# Use a small Python base with Debian Bookworm (stable)
FROM python:3.12-slim-bookworm

# Avoid tz/time warnings
ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1

# --- System dependencies ---
# WeasyPrint (current) needs Pango + HarfBuzz subset + fontconfig, plus fonts; tzdata and ca-certs are good hygiene.
# Ref: WeasyPrint "First Steps" (Debian ≥11) dependency guidance.
RUN apt-get update && apt-get install -y --no-install-recommends \
      libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0 \
      fontconfig fonts-dejavu tzdata ca-certificates curl gnupg \
    && rm -rf /var/lib/apt/lists/*
# (Optional) add more fonts if your plans use them: fonts-noto, fonts-noto-cjk, etc.

# --- Install WeasyPrint (Python) ---
# Use the maintained package from PyPI; CLI "weasyprint" becomes available on PATH.
RUN pip install --upgrade pip \
 && pip install "weasyprint"

# --- Install PowerShell 7 on Debian (official repo) ---
# Ref: Microsoft Learn "Installing PowerShell on Debian"
RUN . /etc/os-release \
 && curl -fsSL https://packages.microsoft.com/config/debian/${VERSION_ID}/packages-microsoft-prod.deb -o /tmp/packages-microsoft-prod.deb \
 && dpkg -i /tmp/packages-microsoft-prod.deb \
 && rm /tmp/packages-microsoft-prod.deb \
 && apt-get update \
 && apt-get install -y --no-install-recommends powershell \
 && rm -rf /var/lib/apt/lists/*

# --- App layout ---
WORKDIR /app
# Copy your script and optional CSS (rename the file to .ps1 for clarity)
COPY printplan.ps1 ./printplan.ps1
COPY print.css ./print.css

# The container will read/write configs and PDFs in /data (mapped from the host)
VOLUME ["/data"]

# Default to PowerShell as the entry. We'll pass script parameters at 'docker run' time.
ENTRYPOINT ["pwsh", "-File", "/app/printplan.ps1"]