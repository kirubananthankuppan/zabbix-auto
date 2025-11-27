# ------------------------
# Fully AUTO Zabbix Install
# Supports any DHCP / isolated network
# ------------------------

$server = "10.2.0.8"
$msi    = "C:\zabbix_agent2.msi"
$url    = "https://cdn.zabbix.com/zabbix/binaries/stable/latest/zabbix_agent2-latest-windows-amd64-openssl.msi"

# Wait for network to come up
Write-Output "Waiting for network..."
while (-not (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet)) {
    Start-Sleep -Seconds 5
}

# Download Zabbix Agent
Write-Output "Downloading Zabbix Agent..."
Invoke-WebRequest -Uri $url -OutFile $msi

# Install
Write-Output "Installing Zabbix Agent..."
Start-Process msiexec.exe -Wait -ArgumentList `
    "/i $msi /qn SERVER=$server SERVERACTIVE=$server HOSTNAME=auto"

# Enable + Start Service
Set-Service -Name "Zabbix Agent 2" -StartupType Automatic
Start-Service "Zabbix Agent 2"

Write-Output "Zabbix Agent Installed Successfully!"
