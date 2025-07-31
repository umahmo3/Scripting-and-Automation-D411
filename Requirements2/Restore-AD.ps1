<##
Umer Mahmood  |  Student ID 001224010
Restore-AD.ps1 – quick & dirty (rubric-ready)
##>

Import-Module ActiveDirectory -ErrorAction Stop

try {
    $me = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw "Need Domain Admin" }

    $ouName   = 'Finance'
    $domainDn = (Get-ADDomain).DistinguishedName

    if ($fou = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction SilentlyContinue) {
        Set-ADObject $fou -ProtectedFromAccidentalDeletion:$false
        Remove-ADOrganizationalUnit $fou -Recursive -Confirm:$false
    }
    New-ADOrganizationalUnit -Name $ouName -Path $domainDn -ErrorAction Stop
        $ouDn = (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'").DistinguishedName
    Write-Host "Using OU DN: $ouDn"

    $csv = Join-Path $PSScriptRoot 'financePersonnel.csv'
    $usedSam = @{}
    foreach ($row in Import-Csv $csv) {
        $fnRaw = $row.'First Name'; if (-not $fnRaw) { $fnRaw = $row.FirstName }
        $lnRaw = $row.'Last Name';  if (-not $lnRaw) { $lnRaw = $row.LastName }
        $postal = $row.'Postal Code'; if (-not $postal) { $postal = $row.PostalCode }
        $office = $row.'Office Phone'; if (-not $office) { $office = $row.OfficePhone }
        $mobile = $row.'Mobile Phone'; if (-not $mobile) { $mobile = $row.MobilePhone }

        Write-Host "Processing row: FirstName='$fnRaw', LastName='$lnRaw'"
        $fn = ($fnRaw -replace '[^A-Za-z0-9]').Trim(); $ln = ($lnRaw -replace '[^A-Za-z0-9]').Trim()
        if (-not ($fn -and $ln)) {
            Write-Host "  Skipped: missing names"
            continue
        }

        $base = ($fn.Substring(0,1)+$ln).ToLower()
        $sam = if ($base.Length -gt 19) { $base.Substring(0,19) } else { $base }
        $i=1
        while ($usedSam[$sam] -or (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue)) {
            $suffix = $i++
            $max = 19 - $suffix.ToString().Length
            $sam = ($base.Substring(0,[Math]::Max(0,$max))) + $suffix
        }
        $usedSam[$sam] = $true

        Write-Host "  Creating user: $sam ($fn $ln)"
        New-ADUser -SamAccountName $sam -GivenName $fn -Surname $ln -DisplayName "$fn $ln" `
            -PostalCode $postal -OfficePhone $office -MobilePhone $mobile `
            -Path $ouDn -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force) -Enabled $true -ErrorAction Stop
    }

    if ($usedSam.Count -eq 0) {
        Write-Warning "No users created. Check CSV headers and file path."
    } else {
        Write-Host "$($usedSam.Count) users created."
    }

    Get-ADUser -Filter * -SearchBase $ouDn -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File (Join-Path $PSScriptRoot 'AdResults.txt') -Encoding UTF8
    Write-Host "AdResults.txt generated with $($usedSam.Count) entries"

    Write-Host "AdResults.txt generated"
}
catch {
    Write-Error $_
}
