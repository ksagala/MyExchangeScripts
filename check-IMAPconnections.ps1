# Enable IMAP protocol logging if not already enabled
# Set-ImapSettings -ProtocolLogEnabled $true

# Define log path for IMAP
$logPath = "$env:ExchangeInstallPath\Logging\Imap4"

# Get all IMAP log files
$logFiles = Get-ChildItem -Path $logPath -Filter "*.log"

# Parse logs for successful connections
$connections = foreach ($file in $logFiles) {
    Select-String -Path $file.FullName -Pattern "Connected" | ForEach-Object {
        $line = $_.Line
        # Extract username and client IP
        if ($line -match "User=(?<User>[^ ]+).*ClientIp=(?<IP>[^ ]+)") {
            [PSCustomObject]@{
                UserName = $matches['User']
                ClientIP = $matches['IP']
                LogFile  = $file.Name
            }
        }
    }
}

# Display unique users and IPs
$connections | Sort-Object UserName, ClientIP | Format-Table -AutoSize
