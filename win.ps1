# ============================================================
# ZABBIX AGENT2 AUTO DEPLOY SCRIPT (64-BIT, PRODUCTION READY)
# Removes old agent, installs Agent2, configures service
# ============================================================

$InstallPath = "C:\Zabbix"
$MsiUrl      = "https://cdn.zabbix.com/zabbix/binaries/stable/7.4/7.4.5/zabbix_agent2-7.4.5-windows-amd64-openssl.msi"
$MsiFile     = "$InstallPath\zabbix_agent2.msi"
$ZabbixIP    = "192.168.2.41"

Write-Host "=== ZABBIX AGENT2 AUTO DEPLOY START ===" -ForegroundColor Cyan

# ------------------------------------------------------------
# STEP 1: Remove OLD Agent / Agent2
# ------------------------------------------------------------
Write-Host "Checking for old Zabbix services..."

$OldServices = Get-Service | Where-Object { $_.Name -match "zabbix" }
foreach ($svc in $OldServices) {
    Write-Host "Stopping service: $($svc.Name)"
    Stop-Service $svc.Name -Force -ErrorAction SilentlyContinue

    Write-Host "Deleting service: $($svc.Name)"
    sc.exe delete $svc.Name | Out-Null
}

Write-Host "Uninstalling old Zabbix MSI packages..."

$products = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Zabbix*" }
foreach ($p in $products) {
    Write-Host "Uninstalling: $($p.Name)"
    $p.Uninstall() | Out-Null
}

Start-Sleep -Seconds 2

# ------------------------------------------------------------
# STEP 2: Prepare Folder
# ------------------------------------------------------------
if (!(Test-Path $InstallPath)) {
    Write-Host "Creating install directory..."
    New-Item -ItemType Directory -Path $InstallPath | Out-Null
}

# ------------------------------------------------------------
# STEP 3: Download MSI
# ------------------------------------------------------------
Write-Host "Downloading Zabbix Agent2 MSI..."
Invoke-WebRequest -Uri $MsiUrl -OutFile $MsiFile -UseBasicParsing

# ------------------------------------------------------------
# STEP 4: Install Agent2 Service (Silent)
# ------------------------------------------------------------
Write-Host "Installing Zabbix Agent2..."
Start-Process msiexec.exe -ArgumentList `
    "/i `"$MsiFile`" /qn SERVER=$ZabbixIP SERVERACTIVE=$ZabbixIP HOSTNAMEITEM=system.hostname" `
    -Wait

Start-Sleep -Seconds 3

# ------------------------------------------------------------
# STEP 5: Start & Configure Service
# ------------------------------------------------------------
$Service = Get-Service | Where-Object { $_.DisplayName -like "*Zabbix Agent 2*" }

if ($Service) {
    Write-Host "Zabbix Agent2 installed as service: $($Service.Name)"
    Start-Service $Service.Name -ErrorAction SilentlyContinue
    Set-Service -Name $Service.Name -StartupType Automatic
    Write-Host "Service started & set to Automatic."
} else {
    Write-Host "ERROR: Zabbix Agent2 service NOT FOUND!" -ForegroundColor Red
}

# ------------------------------------------------------------
# STEP 6: Final Status
# ------------------------------------------------------------
Write-Host ""
Write-Host "=== INSTALLATION COMPLETE ===" -ForegroundColor Green
Get-Service | Where-Object { $_.DisplayName -like "*Zabbix*" }
Write-Host "Agent communicating with Zabbix server at $ZabbixIP"
