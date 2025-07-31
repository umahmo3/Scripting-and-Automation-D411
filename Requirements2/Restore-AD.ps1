<##
    Umer Mahmood
    Student ID: 001224010
    Restore-AD.ps1 - recreate Finance OU and import users
##>

try {
    $ouPath = "ou=Finance,dc=consultingfirm,dc=com"

    # Check for and delete any existing Finance OU without prompting
    Write-Host "Checking for existing Finance OU..."
    $existingOu = Get-ADOrganizationalUnit -LDAPFilter "(ou=Finance)" -ErrorAction SilentlyContinue
    if ($existingOu) {
        Write-Host "Finance OU already exists. Deleting it..." -ForegroundColor Yellow
        Remove-ADOrganizationalUnit -Identity $existingOu -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "Finance OU deleted." -ForegroundColor Green
    } else {
        Write-Host "Finance OU does not exist. Continuing..."
    }

    #Create Finance OU
    Write-Host "Creating Finance OU..."
    New-ADOrganizationalUnit -Name "Finance" -Path "dc=consultingfirm,dc=com" -ErrorAction Stop
    Write-Host "Finance OU created." -ForegroundColor Green

    #Import users from CSV
    $csvPath = Join-Path $PSScriptRoot "financePersonnel.csv"
    Write-Host "Importing users from $csvPath..."
    $users = Import-Csv -Path $csvPath -ErrorAction Stop
    foreach ($u in $users) {
        $displayName = "$($u.'First Name') $($u.'Last Name')"
        $params = @{
            GivenName       = $u.'First Name'
            Surname         = $u.'Last Name'
            Name            = $displayName
            DisplayName     = $displayName
            PostalCode      = $u.'Postal Code'
            OfficePhone     = $u.'Office Phone'
            MobilePhone     = $u.'Mobile Phone'
            Path            = $ouPath
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }
        New-ADUser @params -ErrorAction Stop
        Write-Host "Created user: $displayName"
    }
    Write-Host "All finance personnel imported." -ForegroundColor Green

    # Export results
    $exportPath = Join-Path $PSScriptRoot "AdResults.txt"
    Write-Host "Exporting AD results to $exportPath..."
    Get-ADUser -Filter * -SearchBase $ouPath -Properties DisplayName,PostalCode,OfficePhone,MobilePhone `
        | Select-Object DisplayName,PostalCode,OfficePhone,MobilePhone `
        | Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "Export completed."
}
catch {
    Write-Error "Error occurred in AD script: $($_.Exception.Message)"
}
