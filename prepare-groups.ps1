#
# PowerShell script to create mail-enabled security groups in Exchange 2019
# based on a list of two-character strings from a text file.
#
# Path to the input text file containing two-character strings (one per line)
$inputFile = "C:\Scripts\Purview\group-codes.txt"

# Import the list of codes
$codes = Get-Content $inputFile | Where-Object { $_.Length -eq 2 }

foreach ($code in $codes) {
    $groupName = "${code}_label"
    # Create mail-enabled security group using Exchange Online PowerShell
    if (Get-DistributionGroup -Filter "Name -eq '$groupName'") {
        Write-Host "Group already exists: $groupName"
    }
    else {
        New-DistributionGroup -Name $groupName -DisplayName $groupName -Type Security -OrganizationalUnit "Purview" # -PrimarySmtpAddress "$
        Write-Host "Created group: $groupName"
    }
}