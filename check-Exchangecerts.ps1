#
# Script to check all expired and almost expired certificates on all Exchange servers in organization
#

# 1. Retrieving all Exchange servers in the organization
$ExchangeServers = Get-ExchangeServer | Where-Object { $_.AdminDisplayVersion -match "Version 15" } # Exchange 2013, 2016, 2019

$Report = @()
$ThresholdDays = 30 # Number of days to consider a certificate as "about to expire"

Write-Host "Starting scan of certificates on Exchange servers..." -ForegroundColor Cyan

foreach ($Server in $ExchangeServers) {
    Write-Host "Checking server: $($Server.Name)..." -ForegroundColor Gray
    
    try {
        # Pobranie certyfikatów z konkretnego serwera
        $Certs = Get-ExchangeCertificate -Server $Server.Name
        
        foreach ($Cert in $Certs) {
            $DaysRemaining = ($Cert.NotAfter - (Get-Date)).Days
            $Status = "OK"
            
            if ($DaysRemaining -lt 0) { $Status = "EXPIRED" }
            elseif ($DaysRemaining -le $ThresholdDays) { $Status = "WARNING: Expiring soon" }

            # Creating object with data
            $CertInfo = [PSCustomObject]@{
                Server        = $Server.Name
                Subject       = $Cert.Subject
                Services      = $Cert.Services
                Expiration    = $Cert.NotAfter
                DaysLeft      = $DaysRemaining
                Status        = $Status
                Thumbprint    = $Cert.Thumbprint
            }
            $Report += $CertInfo
        }
    }
    catch {
        Write-Warning "Cannot connect to server $($Server.Name)."
    }
}

# 2. Display results in console
$Report | Sort-Object DaysLeft | Format-Table Server, Subject, Services, Expiration, DaysLeft, Status -AutoSize

# 3. Optional 1: Export results to CSV file
# $Report | Export-Csv -Path "C:\Reports\Exchange_certs.csv" -NoTypeInformation -Encoding UTF8
# Write-Host "Report was saved in C:\Reports\Exchange_certs.csv" -ForegroundColor Green

# 3. Optional 2: Sent report via email
<# $ExpiringCerts = $Report | Where-Object { $_.DaysLeft -le $ThresholdDays }
 if ($ExpiringCerts) {
    Send-MailMessage -To "admin@yourdomain.com" -From "exchange@yourdomain.com" `
    -Subject "ALARM: Wygasające certyfikaty Exchange" `
    -Body ($ExpiringCerts | Out-String) -SmtpServer "your.smtp.server"
}#>