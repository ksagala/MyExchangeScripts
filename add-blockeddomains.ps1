# 1. Get the current list to avoid overwriting it
$Policy = Get-HostedContentFilterPolicy -Identity Default

# 2. Define the new addresses or domains to add
# Assuming the CSV file has a header "Domain"
$newsenders = Import-Csv -Path "C:/Scripts/blockeddomains.csv" | Select-Object -ExpandProperty Domain

# 3. Combine old and new lists
$FinalList = $Policy.BlockedDomains + $newsenders

# 4. Update the policy
Set-HostedContentFilterPolicy -Identity Default -BlockedSenderDomains $FinalList

# 5. Confirm the update
Get-HostedContentFilterPolicy -Identity Default | Select-Object -ExpandProperty BlockedSenderDomains