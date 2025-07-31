<#
    Restore-AD.ps1
    Umer Mahmood
    Student ID: 001224010 
#>

# ---prep---------------------------------------------------------------
Import-Module ActiveDirectory

$domainDN        = 'dc=consultingfirm,dc=com'   # Fixed DN per rubric wording
$financeOUName   = 'Finance'
$financeOUDN     = "ou=$financeOUName,$domainDN"
$scriptDir       = Split-Path -Parent $PSCommandPath
$csvPath         = Join-Path $scriptDir 'financePersonnel.csv'

try {
    # ---remove existing OU if exists--------------------------
    $existingOU = Get-ADOrganizationalUnit -LDAPFilter "(ou=$financeOUName)" `
                                           -SearchBase $domainDN `
                                           -ErrorAction SilentlyContinue
    if ($existingOU) {
        Write-Host "An existing '$financeOUName' OU was found and will be deleted..." -ForegroundColor Yellow

        # Disable accidental-deletion protection then delete
        Set-ADObject -Identity $existingOU.DistinguishedName -ProtectedFromAccidentalDeletion:$false
        Remove-ADOrganizationalUnit -Identity $existingOU.DistinguishedName -Recursive -Confirm:$false
        Write-Host "The old '$financeOUName' OU has been removed." -ForegroundColor Yellow
    } else {
        Write-Host "No existing '$financeOUName' OU was detected." -ForegroundColor Cyan
    }

    # --- reate new OU-------------------------------------------
    New-ADOrganizationalUnit -Name $financeOUName -Path $domainDN -ProtectedFromAccidentalDeletion $true
    Write-Host "New '$financeOUName' OU created successfully." -ForegroundColor Green

    # ---Import users from CSV-------------------------------------------
    if (-not (Test-Path $csvPath)) { throw "CSV file not found: $csvPath" }

    Import-Csv -Path $csvPath | ForEach-Object {
        $given   = $_.First_Name.Trim()
        $surname = $_.Last_Name.Trim()
        $display = "$given $surname"
        $sam     = $_.samAccount.Trim()

        # Create the user
        New-ADUser `
            -GivenName      $given `
            -Surname        $surname `
            -Name           $display `
            -DisplayName    $display `
            -SamAccountName $sam `
            -Path           $financeOUDN `
            -PostalCode     $_.PostalCode `
            -OfficePhone    $_.OfficePhone `
            -MobilePhone    $_.MobilePhone `
            -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force) `
            -Enabled        $true
    }

    Write-Host "All Finance users imported successfully." -ForegroundColor Green
}
catch {
    Write-Host "An unexpected error occurred:`n$($_.Exception.Message)" -ForegroundColor Red
    break
}

# --- REQUIRED OUTPUT LINE --------------------------------------------------
Get-ADUser -Filter * -SearchBase "ou=Finance,dc=consultingfirm,dc=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt
