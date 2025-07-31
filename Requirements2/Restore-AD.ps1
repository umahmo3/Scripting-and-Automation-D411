<##
    Umer Mahmood
    Student ID: 001224010
    Task 2: Restore Active Directory

    NOTE: Run this script in an elevated session as a Domain Admin.
    If the script is blocked, use: Unblock-File -Path Restore-AD.ps1
##>

# Load the AD module
Import-Module ActiveDirectory -ErrorAction Stop

try {
    # Check for Domain Admin rights
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "Insufficient rights: Run this script as a Domain Admin."
        throw "Privilege check failed."
    }

    $ouName   = "Finance"
    $domainDn = (Get-ADDomain).DistinguishedName

    Write-Host "-- Checking for existing Finance OU..."
    $existingOu = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction SilentlyContinue

    if ($existingOu) {
        Write-Host "Finance OU exists. Removing protection flag if set..." -ForegroundColor Yellow
        # Disable accidental deletion protection if enabled
        Set-ADObject -Identity $existingOu.DistinguishedName -ProtectedFromAccidentalDeletion:$false

        Write-Host "Deleting existing Finance OU..."
        Remove-ADOrganizationalUnit -Identity $existingOu.DistinguishedName -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "Finance OU deleted." -ForegroundColor Green
    } else {
        Write-Host "Finance OU not found. Proceeding to creation."
    }

    # Create new Finance OU
    Write-Host "-- Creating Finance OU..."
    New-ADOrganizationalUnit -Name $ouName -Path $domainDn -ErrorAction Stop
    Write-Host "Finance OU created." -ForegroundColor Green

    # Refresh OU DN
    $ouObject = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction Stop
    $ouDn     = $ouObject.DistinguishedName

    # Import users from CSV
    $csvPath = Join-Path $PSScriptRoot "financePersonnel.csv"
    Write-Host "-- Importing users from: $csvPath"
    $users = Import-Csv -Path $csvPath -ErrorAction Stop

    foreach ($u in $users) {
        $fullName = "$($u.First_Name) $($u.Last_Name)"
        $userParams = @{
            GivenName       = $u.First_Name
            Surname         = $u.Last_Name
            Name            = $fullName
            DisplayName     = $fullName
            PostalCode      = $u.PostalCode
            OfficePhone     = $u.OfficePhone
            MobilePhone     = $u.MobilePhone
            Path            = $ouDn
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }
        New-ADUser @userParams -ErrorAction Stop
        Write-Host "Created AD user: $fullName"
    }

    Write-Host "All finance users imported successfully." -ForegroundColor Green

    # Export results
    $exportPath = Join-Path $PSScriptRoot "AdResults.txt"
    Write-Host "-- Exporting AD user list to: $exportPath"
    Get-ADUser -Filter * -SearchBase $ouDn -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "AD export complete." -ForegroundColor Green
}
catch {
    Write-Error "[AD Script Error] $($_.Exception.Message)"
    if ($_.Exception.Message -match 'Access is denied') {
        Write-Host "Hint: Ensure you're running as a Domain Admin and that the OU is not protected." -ForegroundColor Red
    }
}
