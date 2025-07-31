<##
    Umer Mahmood
    Student ID: 001224010
    Task 2: Restore Active Directory

    NOTE: Run in an elevated session under a Domain Admin account.
    Use `Unblock-File -Path Restore-AD.ps1` if execution policy blocks the script.
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
        Set-ADOrganizationalUnit -Identity $existingOu -ProtectedFromAccidentalDeletion:$false -ErrorAction SilentlyContinue
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

    $created = 0
    foreach ($u in $users) {
        # Support multiple header styles in the CSV
        $first  = $u.'First Name'  ?? $u.FirstName  ?? $u.First_Name
        $last   = $u.'Last Name'   ?? $u.LastName   ?? $u.Last_Name
        $postal = $u.'Postal Code' ?? $u.PostalCode ?? $u.Postal_Code
        $office = $u.'Office Phone' ?? $u.OfficePhone ?? $u.Office_Phone
        $mobile = $u.'Mobile Phone' ?? $u.MobilePhone ?? $u.Mobile_Phone
        $sam    = $u.samAccount

        # Build a SamAccountName if one is not supplied
        if (-not $sam) {
            $base = ($first.Substring(0,1) + $last).ToLower()
            if ($base.Length -gt 19) { $base = $base.Substring(0,19) }
            $sam = $base
            $i = 1
            while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
                $trim = 19 - $i.ToString().Length
                $sam = $base.Substring(0,[Math]::Max(0,$trim)) + $i
                $i++
            }
        }

        $fullName = "$first $last"
        $userParams = @{
            SamAccountName  = $sam
            GivenName       = $first
            Surname         = $last
            Name            = $fullName
            DisplayName     = $fullName
            PostalCode      = $postal
            OfficePhone     = $office
            MobilePhone     = $mobile
            Path            = $ouDistinguishedName
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }
        New-ADUser @userParams -ErrorAction Stop
        Write-Host "Created AD user: $fullName ($sam)"
        $created++
    }
    Write-Host "$created users imported." -ForegroundColor Green

}
catch {
    Write-Error "[AD Script Error] $($_.Exception.Message)"
    if ($_.Exception.Message -match 'Access is denied') {
        Write-Host "Hint: Verify Domain Admin rights and AD module availability." -ForegroundColor Red
    }
}

# Export the Finance OU users for submission
Get-ADUser -Filter * -SearchBase $ouDistinguishedName -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > (Join-Path $PSScriptRoot 'AdResults.txt')
