<##
    Umer Mahmood
    Student ID: 001224010
    Restore-AD.ps1 - recreate Finance OU and import users
##>

try {
    # Privilege check: must run under Domain Admin context
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "You may lack the required admin rights. Please run PowerShell as an elevated Domain Admin."
        throw "Insufficient privileges for AD changes."
    }

    $ouPath = "ou=Finance,dc=consultingfirm,dc=com"

    Write-Host "Checking for existing Finance OU..."
    $existingOu = Get-ADOrganizationalUnit -Filter 'Name -eq "Finance"' -ErrorAction SilentlyContinue
    if ($existingOu) {
        Write-Host "Finance OU already exists. Deleting it..." -ForegroundColor Yellow
        # Use -Recursive and suppress prompts
        Remove-ADOrganizationalUnit -Identity $existingOu -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "Finance OU deleted." -ForegroundColor Green
    } else {
        Write-Host "Finance OU does not exist. Continuing..."
    }

    Write-Host "Creating Finance OU..."
    New-ADOrganizationalUnit -Name "Finance" -Path "dc=consultingfirm,dc=com" -ErrorAction Stop
    Write-Host "Finance OU created." -ForegroundColor Green

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

    $exportPath = Join-Path $PSScriptRoot "AdResults.txt"
    Write-Host "Exporting AD results to $exportPath..."
    Get-ADUser -Filter * -SearchBase $ouPath -Properties DisplayName,PostalCode,OfficePhone,MobilePhone `
        | Select DisplayName,PostalCode,OfficePhone,MobilePhone `
        | Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "Export completed." -ForegroundColor Green
}
catch {
    Write-Error "Error in AD script: $($_.Exception.Message)"
    # If Access Denied, remind to check rights
    if ($_.Exception.Message -match 'Access is denied') {
        Write-Host "Hint: Ensure you have Domain Admin privileges and that the AD module is imported." -ForegroundColor Red
    }
}
