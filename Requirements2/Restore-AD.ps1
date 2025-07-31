# Umer Mahmood
# Student ID: 001224010
# Restore Active Directory OU and import users

Import-Module ActiveDirectory

try {
    # Check if OU exists
    $ou = Get-ADOrganizationalUnit -LDAPFilter "(ou=Finance)" -ErrorAction SilentlyContinue
    if ($ou) {
        Write-Host "Finance OU exists. Deleting OU..."
        Remove-ADOrganizationalUnit -Identity "OU=Finance,DC=consultingfirm,DC=com" -Recursive -Confirm:$false
        Write-Host "Finance OU deleted."
    } else {
        Write-Host "Finance OU does not exist."
    }

    # Create Finance OU
    New-ADOrganizationalUnit -Name "Finance" -Path "DC=consultingfirm,DC=com"
    Write-Host "Finance OU created."

    # Import users from CSV
    $users = Import-Csv ".\financePersonnel.csv"
    foreach ($user in $users) {
        $displayName = "$($user.'First Name') $($user.'Last Name')"
        New-ADUser `
            -Name $displayName `
            -GivenName $user.'First Name' `
            -Surname $user.'Last Name' `
            -DisplayName $displayName `
            -PostalCode $user.'Postal Code' `
            -OfficePhone $user.'Office Phone' `
            -MobilePhone $user.'Mobile Phone' `
            -Path "OU=Finance,DC=consultingfirm,DC=com" `
            -AccountPassword (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) `
            -Enabled $true
    }

    # Export users to AdResults.txt
    Get-ADUser -Filter * -SearchBase "OU=Finance,DC=consultingfirm,DC=com" `
        -Properties DisplayName,PostalCode,OfficePhone,MobilePhone | `
        Select-Object DisplayName,PostalCode,OfficePhone,MobilePhone | `
        Out-File -FilePath ".\AdResults.txt"

    Write-Host "AD user import and export complete."
}
catch {
    Write-Error "Error: $_"
}
