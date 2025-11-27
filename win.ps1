Write-Host "Starting Zabbix Agent auto deployment..."

$ServerIP = "103.127.29.5"
$WorkDir = "C:\ZabbixAuto"
$AgentPath = "C:\Program Files\Zabbix Agent 2"
$ConfigFile = "$AgentPath\zabbix_agent2.conf"

# Correct Zabbix 7.0.5 LTS URL
$downloadUrl = "https://cdn.zabbix.com/zabbix/binaries/stable/7.0/7.0.5/zabbix_agent2-7.0.5-windows-amd64-openssl.zip"
$zipFile = "$WorkDir\zabbix.zip"

# Create working folder
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

# Detect hostname & IP
$Hostname = (hostname)
$IPAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.IPAddress -notlike "169.*" }).IPAddress | Select-Object -First 1

Write-Host "Detected Hostname: $Hostname"
Write-Host "Detected IP Address: $IPAddress"

# Download Zabbix Agent 2
Write-Host "Downloading Zabbix Agent 2 ZIP package..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile

# Extract ZIP
Write-Host "Extracting Zabbix Agent 2..."
$UnzipPath = "$WorkDir\unzipped"
Expand-Archive -Path $zipFile -DestinationPath $UnzipPath -Force

# Detect extracted folder name
$ExtractedFolder = Get-ChildItem -Path $UnzipPath | Where-Object { $_.PSIsContainer } | Select-Object -First 1
$ExtractedPath = $ExtractedFolder.FullName

# Create target folder
New-Item -ItemType Directory -Force -Path $AgentPath | Out-Null

# Copy binaries
Copy-Item "$ExtractedPath\bin\zabbix_agent2.exe" $AgentPath -Force
Copy-Item "$ExtractedPath\conf\zabbix_agent2.conf" $AgentPath -Force

# Configure agent
(Get-Content $ConfigFile) `
    -replace "^Server=.*", "Server=$ServerIP" `
    -replace "^ServerActive=.*", "ServerActive=$ServerIP" `
    -replace "^Hostname=.*", "Hostname=$Hostname" `
    | Set-Content $ConfigFile

# Install service
Write-Host "Installing Zabbix Agent 2 as service..."
& "$AgentPath\zabbix_agent2.exe" --install

# Start & enable service
Start-Service "Zabbix Agent 2"
Set-Service -Name "Zabbix Agent 2" -StartupType Automatic

Write-Host ""
Write-Host "Zabbix Agent 2 successfully installed & running!"
Write-Host "Machine IP: $IPAddress"
Write-Host "Hostname: $Hostname"
Write-Host "Server: $ServerIP"
