<##
    Umer Mahmood
    Student ID: 001224010
    AUN1 Task 2: Restored AD and SQL Automation Scripts
    NOTE: Run each section in an elevated session under a Domain Admin or SQL sysadmin account.
         Use `Unblock-File -Path Restore-AD.ps1` if you see execution warnings.
##>

# ==================================================
# SECTION 1: Restore-AD.ps1  (Recreate Finance OU & import users)
# ==================================================

# Make sure the AD module is available before proceeding
Import-Module ActiveDirectory -ErrorAction Stop

try {
    # Sanity check to avoid permission errors—needs Domain Admin rights
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "Insufficient rights: Run this script as a Domain Admin."
        throw "Privilege check failed."
    }

    $ouDistinguishedName = "ou=Finance,dc=consultingfirm,dc=com"

    # Clean up existing OU if it exists
    Write-Host "-- Checking for existing Finance OU..."
    $existingOu = Get-ADOrganizationalUnit -Filter 'Name -eq "Finance"' -ErrorAction SilentlyContinue
    if ($existingOu) {
        Write-Host "Finance OU exists. Deleting..." -ForegroundColor Yellow
        Remove-ADOrganizationalUnit -Identity $existingOu -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "Finance OU deleted." -ForegroundColor Green
    } else {
        Write-Host "Finance OU not found. Proceeding to creation."
    }

    # Create the OU from scratch
    Write-Host "-- Creating Finance OU..."
    New-ADOrganizationalUnit -Name "Finance" -Path "dc=consultingfirm,dc=com" -ErrorAction Stop
    Write-Host "Finance OU created." -ForegroundColor Green

    # Import the CSV with user data
    $csvFile = Join-Path $PSScriptRoot "financePersonnel.csv"
    Write-Host "-- Importing users from: $csvFile"
    $users = Import-Csv -Path $csvFile -ErrorAction Stop

    # Loop through each entry and create a user
    foreach ($u in $users) {
        $fullName = "$($u.'First Name') $($u.'Last Name')"
        $userParams = @{ 
            GivenName       = $u.'First Name'
            Surname         = $u.'Last Name'
            Name            = $fullName
            DisplayName     = $fullName
            PostalCode      = $u.'Postal Code'
            OfficePhone     = $u.'Office Phone'
            MobilePhone     = $u.'Mobile Phone'
            Path            = $ouDistinguishedName
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }
        New-ADUser @userParams -ErrorAction Stop
        Write-Host "Created AD user: $fullName"
    }
    Write-Host "All finance personnel have been imported." -ForegroundColor Green

    # Export created user info to a text file
    $outputFile = Join-Path $PSScriptRoot "AdResults.txt"
    Write-Host "-- Exporting AD report to: $outputFile"
    Get-ADUser -Filter * -SearchBase $ouDistinguishedName -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "AD export complete." -ForegroundColor Green
}
catch {
    Write-Error "[AD Script Error] $($_.Exception.Message)"
    if ($_.Exception.Message -match 'Access is denied') {
        Write-Host "Hint: Verify Domain Admin rights and AD module availability." -ForegroundColor Red
    }
}
