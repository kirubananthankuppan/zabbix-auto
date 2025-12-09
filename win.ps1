
# ============================================================
# ZABBIX AGENT 2 - CONDITIONAL CLEAN + FULLY UNATTENDED INSTALL
# ============================================================

$ErrorActionPreference = "Stop"

# Force TLS 1.2 for reliable downloads (older Windows)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# --- CONFIG ---
$ZabbixServer   = "172.29.117.104"              # Your Zabbix Server IP
$HostMetadata   = "windows"                     # For auto-registration rules
$DeployRoot     = "C:\ZabbixDeploy"
$LogFile        = Join-Path $DeployRoot "install_log.txt"
$MsiFile        = Join-Path $DeployRoot "zabbix_agent2.msi"
$MsiUrl         = "https://cdn.zabbix.com/zabbix/binaries/stable/7.4/7.4.5/zabbix_agent2-7.4.5-windows-amd64-openssl.msi"
$AgentExePath   = "C:\Program Files\Zabbix Agent 2\zabbix_agent2.exe"
$ConfPath       = "C:\Program Files\Zabbix Agent 2\zabbix_agent2.conf"
$EventLogRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Application\Zabbix Agent 2"

# --- LOGGING ---
function Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$ts $msg"
    $line | Tee-Object -FilePath $LogFile -Append
    Write-Host $msg
}

# Ensure deploy/log folder exists BEFORE any logging
if (!(Test-Path $DeployRoot)) { New-Item -ItemType Directory -Path $DeployRoot | Out-Null }

Log "=== ZABBIX AGENT 2 CONDITIONAL CLEAN + INSTALL START ==="

# ------------------------------------------------------------
# STEP A: If Agent 2 exists -> FULL CLEANUP; else skip to install
# ------------------------------------------------------------
$existingSvc = Get-Service -Name "Zabbix Agent 2" -ErrorAction SilentlyContinue
if ($existingSvc) {
    Log "Service 'Zabbix Agent 2' exists -> performing full cleanup"

    # Stop & delete service
    try {
        Stop-Service $existingSvc.Name -Force -ErrorAction SilentlyContinue
        Log "Stopped service: $($existingSvc.Name)"
    } catch { Log "WARN: Failed to stop service: $($existingSvc.Name): $($_.Exception.Message)" }
    sc.exe delete "$($existingSvc.Name)" | Out-Null
    Log "Deleted service: $($existingSvc.Name)"

    # Also remove any other Zabbix-related services
    foreach ($n in @("Zabbix Agent","Zabbix Agent2","zabbix","zabbix_agent2")) {
        $svc = Get-Service -Name $n -ErrorAction SilentlyContinue
        if ($svc) {
            try { Stop-Service $svc.Name -Force -ErrorAction SilentlyContinue } catch {}
            sc.exe delete "$($svc.Name)" | Out-Null
            Log "Deleted service: $($svc.Name)"
        }
    }

    Start-Sleep -Seconds 2

    # Uninstall any MSI products named Zabbix*
    Log "Uninstalling MSI products named 'Zabbix*' (this can take a while)..."
    $products = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Zabbix*" }
    foreach ($p in $products) {
        try {
            Log "Uninstalling: $($p.Name)"
            $p.Uninstall() | Out-Null
        } catch {
            Log "WARN: Failed to uninstall $($p.Name): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2

    # Remove leftover folders (do NOT delete $DeployRoot)
    foreach ($p in @(
        "C:\Program Files\Zabbix Agent",
        "C:\Program Files\Zabbix Agent 2",
        "C:\Program Files (x86)\Zabbix",
        "C:\Zabbix"
    )) {
        if (Test-Path $p) {
            try {
                Remove-Item -Recurse -Force $p
                Log "Removed folder: $($p)"
            } catch {
                Log "WARN: Cannot remove $($p): $($_.Exception.Message)"
            }
        }
    }

    # Remove EventLog registry key (it can block re-install)
    if (Test-Path $EventLogRegKey) {
        try {
            Remove-Item -Path $EventLogRegKey -Recurse -Force
            Log "Removed EventLog registry key: $EventLogRegKey"
        } catch {
            Log "WARN: Failed to remove EventLog reg key: $($_.Exception.Message)"
        }
    } else {
        Log "No EventLog reg key found to remove."
    }

    Log "Cleanup complete. Proceeding to install."
} else {
    Log "No existing 'Zabbix Agent 2' service found. Proceeding to install."
}

# ------------------------------------------------------------
# STEP B: Download MSI
# ------------------------------------------------------------
if (Test-Path $MsiFile) { Remove-Item $MsiFile -Force }
Log "Downloading Zabbix Agent 2 MSI from $MsiUrl ..."
Invoke-WebRequest -Uri $MsiUrl -OutFile $MsiFile

# ------------------------------------------------------------
# STEP C: Install MSI silently, correct path
# ------------------------------------------------------------
Log "Installing Zabbix Agent 2..."
$msiArgs = @(
    "/i", "`"$MsiFile`"",
    "/qn", "/norestart",
    "SERVER=$ZabbixServer",
    "SERVERACTIVE=$ZabbixServer",
    "HOSTNAMEITEM=system.hostname",
    "HOSTMETADATA=$HostMetadata",
    "LOGTYPE=file",
    "LOGFILE=`"C:\Program Files\Zabbix Agent 2\zabbix_agent2.log`""
)
$proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru
Log "msiexec exit code: $($proc.ExitCode)"
if ($proc.ExitCode -ne 0) { throw "MSI installation failed with exit code $($proc.ExitCode)" }

Start-Sleep -Seconds 2

# ------------------------------------------------------------
# STEP D: Ensure service exists, start, set Automatic
# ------------------------------------------------------------
$svc2 = Get-Service -Name "Zabbix Agent 2" -ErrorAction SilentlyContinue
if (-not $svc2) {
    if (Test-Path $AgentExePath) {
        Log "Service not found; creating manually..."
        sc.exe create "Zabbix Agent 2" binPath= "`"$AgentExePath`" --config `"$ConfPath`"" start= auto | Out-Null
        Start-Sleep -Seconds 2
        $svc2 = Get-Service -Name "Zabbix Agent 2" -ErrorAction SilentlyContinue
    } else {
        throw "Agent executable not found at $AgentExePath"
    }
}

try { Start-Service $svc2.Name -ErrorAction SilentlyContinue } catch { Log "WARN: Start-Service failed: $($_.Exception.Message)" }
Set-Service -Name $svc2.Name -StartupType Automatic
Log "Service '$($svc2.Name)' is running and set to Automatic."

# ------------------------------------------------------------
# STEP E: Open Windows Firewall TCP 10050
# ------------------------------------------------------------
Log "Ensuring firewall rule for TCP 10050..."
Get-NetFirewallRule -DisplayName "Zabbix Agent 2" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
New-NetFirewallRule -DisplayName "Zabbix Agent 2" -Direction Inbound -Protocol TCP -LocalPort 10050 -Action Allow -Profile Any | Out-Null

# ------------------------------------------------------------
# STEP F: Final status
# ------------------------------------------------------------
Log "=== INSTALLATION COMPLETE ==="
Get-Service | Where-Object { $_.DisplayName -ilike "*Zabbix*" } | ForEach-Object {
    Log ("Service: {0} - {1}" -f $_.DisplayName, $_.Status)
}
Log "Configured: Server=$ZabbixServer (passive) & ServerActive=$ZabbixServer (active)"
Log "HostMetadata: '$HostMetadata' (used by autoregistration rules)"

# Exit cleanly (no prompts)
exit 0
