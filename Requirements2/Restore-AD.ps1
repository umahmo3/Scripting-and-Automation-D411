<##
    Umer Mahmood
    Student ID: 001224010
    Restore-AD.ps1 - recreate Finance OU and import users
##>

try {
    $ouPath = "ou=Finance,dc=consultingfirm,dc=com"
    $existingOu = Get-ADOrganizationalUnit -LDAPFilter "(ou=Finance)" -ErrorAction SilentlyContinue
    if ($existingOu) {
        Write-Host "Finance OU already exists. Deleting it..."
        Remove-ADOrganizationalUnit -Identity $existingOu -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "Finance OU deleted."
    } else {
        Write-Host "Finance OU does not exist."
    }

    New-ADOrganizationalUnit -Name "Finance" -Path "dc=consultingfirm,dc=com" -ErrorAction Stop
    Write-Host "Finance OU created."

    $csvPath = Join-Path $PSScriptRoot 'financePersonnel.csv'
    $users = Import-Csv -Path $csvPath
    foreach ($u in $users) {
        $displayName = "$($u.First_Name) $($u.Last_Name)"
        $params = @{
            GivenName = $u.First_Name
            Surname = $u.Last_Name
            Name = $displayName
            DisplayName = $displayName
            SamAccountName = $u.samAccount
            Path = $ouPath
            PostalCode = $u.PostalCode
            OfficePhone = $u.OfficePhone
            MobilePhone = $u.MobilePhone
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled = $true
        }
        New-ADUser @params -ErrorAction Stop
    }
    Write-Host "Finance personnel imported."

    Get-ADUser -Filter * -SearchBase $ouPath -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > "$PSScriptRoot\AdResults.txt"
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}
