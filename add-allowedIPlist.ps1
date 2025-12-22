# 1. Get the current list to avoid overwriting it
$currentPolicy = Get-HostedConnectionFilterPolicy -Identity Default

# 2. Define the new IPs to add
# Assuming the CSV file has a header "IP"
$newIPs = Import-Csv -Path "C:/Scripts/Exchange/allowlist.csv" | Select-Object -ExpandProperty IP
# 3. Combine old and new lists
$combinedList = $currentPolicy.IPAllowList + $newIPs
# 4. Update the policy
Set-HostedConnectionFilterPolicy -Identity Default -IPAllowList $combinedList