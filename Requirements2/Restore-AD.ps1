<##
Umer Mahmood  |  Student ID 001224010
Restore-AD.ps1 – final version
##>

Import-Module ActiveDirectory -ErrorAction Stop

try {
    $me = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw "Need Domain Admin" }

    $ouName   = 'Finance'
    $domainDn = (Get-ADDomain).DistinguishedName

    # remove old OU if exists
    if ($ou = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction SilentlyContinue) {
        Set-ADObject -Identity $ou -ProtectedFromAccidentalDeletion:$false
        Remove-ADOrganizationalUnit -Identity $ou -Recursive -Confirm:$false
    }

    # create fresh OU
    New-ADOrganizationalUnit -Name $ouName -Path $domainDn -ErrorAction Stop
    $ouDn = (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'").DistinguishedName
    Write-Host "OU located at: $ouDn"

    # import users
    $csvPath = Join-Path $PSScriptRoot 'financePersonnel.csv'
    $rows    = Import-Csv -Path $csvPath -ErrorAction Stop
    $count   = 0

    # detect headers dynamically
    $headers   = (Get-Content $csvPath -First 1) -split ','
    $fnHeader  = ($headers | Where-Object { $_ -match '(?i)first' })[0]
    $lnHeader  = ($headers | Where-Object { $_ -match '(?i)last' })[0]
    $pcHeader  = ($headers | Where-Object { $_ -match '(?i)postal' })[0]
    $opHeader  = ($headers | Where-Object { $_ -match '(?i)office' })[0]
    $mpHeader  = ($headers | Where-Object { $_ -match '(?i)mobile' })[0]

    foreach ($r in $rows) {
        # normalize header names
        $fn = ($r.'First Name' ?? $r.FirstName ?? $r.'First_Name') -replace '[^A-Za-z0-9]'
        $ln = ($r.'Last Name'  ?? $r.LastName  ?? $r.'Last_Name')  -replace '[^A-Za-z0-9]'
        if (-not ($fn && $ln)) { continue }

        # build unique SamAccountName
        $base = ($fn.Substring(0,1) + $ln).ToLower()
        $max  = 19
        $sam  = if ($base.Length -gt $max) { $base.Substring(0,$max) } else { $base }
        $i    = 1
        while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
            $suffix  = $i++
            $trimLen = $max - $suffix.ToString().Length
            $sam      = $base.Substring(0,[Math]::Max(0,$trimLen)) + $suffix
        }

        New-ADUser -SamAccountName $sam \
                   -GivenName $fn -Surname $ln -DisplayName "$fn $ln" \
                   -PostalCode ($r.'Postal Code' ?? $r.PostalCode) \
                   -OfficePhone ($r.'Office Phone' ?? $r.OfficePhone) \
                   -MobilePhone ($r.'Mobile Phone' ?? $r.MobilePhone) \
                   -Path $ouDn -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force) \
                   -Enabled $true -ErrorAction Stop
        Write-Host "Created: $sam"
        $count++
    }

    Write-Host "$count users created."
    if ($count -eq 0) { Write-Warning 'No users created. Verify CSV path and headers.' }

    Get-ADUser -Filter * -SearchBase $ouDn -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File (Join-Path $PSScriptRoot 'AdResults.txt') -Encoding UTF8
    Write-Host "Done. Results in AdResults.txt"
}
catch {
    Write-Error $_
}
