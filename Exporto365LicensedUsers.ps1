# Install the MSOnline Module if not already installed
# Install-Module -Name MSOnline

# Import MSOnline module
Import-Module MSOnline

# Function to collect licensed user details from a tenant
function Get-LicensedUserDetails {
    # Connect to MS Online Service with Modern Authentication
    Connect-MsolService

    # Get all users with a license
    Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true } | Select-Object @{Name="TenantDomain";Expression={(Get-MsolCompanyInformation).DisplayName}}, DisplayName, UserPrincipalName, @{Name="Licenses";Expression={($_.Licenses | ForEach-Object { $_.AccountSkuId }) -join ', '}}
}

# Define the path for the exported CSV
$csvPath = ".\ExportedLicences.csv"

# Array to hold all licensed user details across tenants
$allLicensedUserDetails = @()

if (Test-Path -Path $csvPath) {
    # Import existing data if the CSV already exists
    $allLicensedUserDetails = Import-Csv -Path $csvPath
}

# Ask for the number of tenants
$tenantCount = Read-Host "Enter the number of tenants"

for ($i = 1; $i -le $tenantCount; $i++) {
    # Inform the user to sign in for each tenant
    Write-Host "Please sign in for tenant $i"

    # Retrieve the licensed user details
    $licensedUserDetails = Get-LicensedUserDetails

    foreach ($user in $licensedUserDetails) {
        # Check if the user already exists in the array
        $existingUser = $allLicensedUserDetails | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
        if ($existingUser) {
            # Update the existing user's details
            $existingUser.Licenses = $user.Licenses
        } else {
            # Add new user details to the array
            $allLicensedUserDetails += $user
        }
    }
}

# Export all licensed user details to the CSV file
$allLicensedUserDetails | Export-Csv -Path $csvPath -NoTypeInformation -Force

# Output the path to the CSV file
Write-Host "All licensed users report saved to: $csvPath"
