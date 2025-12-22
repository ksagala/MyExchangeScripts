# 1. Get the current list to avoid overwriting it
$Policy = Get-HostedContentFilterPolicy -Identity Default

# 2. Define the new addresses or domains to add
# Assuming the CSV file has a header "Email"
$newsenders = Import-Csv -Path "C:/Scripts/blockedsenders.csv" | Select-Object -ExpandProperty Email

# 3. Combine old and new lists
$FinalList = $Policy.BlockedSenders + $newsenders

# 4. Update the policy
Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $FinalList
