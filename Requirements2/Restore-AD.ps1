<##
Umer Mahmood
Student ID: 001224010
Restore AD script for Task 2
-- must be run as Domain Admin. Use Unblock-File if needed.
##>

Import-Module ActiveDirectory -ErrorAction Stop

try {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "Run as Domain Admin or this won’t work."
        throw "Permission error."
    }

    $ouName   = "Finance"
    $domainDn = (Get-ADDomain).DistinguishedName

    Write-Host "Looking for Finance OU..."
    $existingOu = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction SilentlyContinue

    if ($existingOu) {
        Write-Host "Finance OU found. Removing protection..."
        Set-ADObject -Identity $existingOu.DistinguishedName -ProtectedFromAccidentalDeletion:$false
        Write-Host "Deleting OU..."
        Remove-ADOrganizationalUnit -Identity $existingOu.DistinguishedName -Recursive -Confirm:$false -ErrorAction Stop
    } else {
        Write-Host "OU not found. Making a new one."
    }

    New-ADOrganizationalUnit -Name $ouName -Path $domainDn -ErrorAction Stop
    Write-Host "OU created"

    $ouObject = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction Stop
    $ouDn = $ouObject.DistinguishedName

    $csvPath = Join-Path $PSScriptRoot "financePersonnel.csv"
    $users = Import-Csv -Path $csvPath -ErrorAction Stop

    foreach ($u in $users) {
        $firstName = $u.'First Name'
        $lastName = $u.'Last Name'
        $displayName = "$firstName $lastName"

        $userParams = @{
            GivenName       = $firstName
            Surname         = $lastName
            Name            = $displayName
            DisplayName     = $displayName
            PostalCode      = $u.'Postal Code'
            OfficePhone     = $u.'Office Phone'
            MobilePhone     = $u.'Mobile Phone'
            Path            = $ouDn
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }

        New-ADUser @userParams -ErrorAction Stop
        Write-Host "Added: $displayName"
    }

    $exportPath = Join-Path $PSScriptRoot "AdResults.txt"
    Get-ADUser -Filter * -SearchBase $ouDn -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select-object DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "Finished. Users exported to AdResults.txt"
}
catch {
    Write-Error "Something went wrong: $($_.Exception.Message)"
    if ($_.Exception.Message -match 'Access is denied') {
        Write-Host "You probably need to run this as a Domain Admin." -ForegroundColor Red
    }
}
