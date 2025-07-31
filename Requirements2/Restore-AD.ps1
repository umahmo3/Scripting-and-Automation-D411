<##
Umer Mahmood  |  Student ID 001224010
Restore-AD.ps1 – quick & dirty (rubric-ready)
##>

Import-Module ActiveDirectory -ErrorAction Stop

try {
    $me = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw "Need Domain Admin" }

    $ouName  = 'Finance'
    $domainDn = (Get-ADDomain).DistinguishedName

    if ($fou = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -ErrorAction SilentlyContinue) {
        Set-ADObject $fou -ProtectedFromAccidentalDeletion:$false
        Remove-ADOrganizationalUnit $fou -Recursive -Confirm:$false
    }
    New-ADOrganizationalUnit -Name $ouName -Path $domainDn
    $ouDn = (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'").DistinguishedName

    $csv = Join-Path $PSScriptRoot 'financePersonnel.csv'
    $usedSam = @{}
    foreach ($row in Import-Csv $csv) {
        $fnRaw = $row.'First Name'; if (-not $fnRaw) { $fnRaw = $row.FirstName }; if (-not $fnRaw) { $fnRaw = $row.'First_Name' }
        $lnRaw = $row.'Last Name';  if (-not $lnRaw) { $lnRaw = $row.LastName };  if (-not $lnRaw) { $lnRaw = $row.'Last_Name' }

        $postal = $row.'Postal Code';   if (-not $postal) { $postal = $row.PostalCode }
        $office = $row.'Office Phone';  if (-not $office) { $office = $row.OfficePhone }
        $mobile = $row.'Mobile Phone';  if (-not $mobile) { $mobile = $row.MobilePhone }

        $fn = ($fnRaw -replace '[^A-Za-z0-9]').Trim()
        $ln = ($lnRaw -replace '[^A-Za-z0-9]').Trim()
        if (-not ($fn -and $ln)) { Write-Host "Bad row, skipping"; continue }

        $sam = (($fn.Substring(0,1) + $ln).ToLower()).Substring(0,[Math]::Min(19, $fn.Length + $ln.Length))
        $i = 1
        while ($usedSam[$sam] -or (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue)) {
            $sam = "{0}{1}" -f $sam.Substring(0,18), $i; $i++
        }
        $usedSam[$sam] = $true

        $params = @{
            SamAccountName  = $sam
            GivenName       = $fn
            Surname         = $ln
            DisplayName     = "$fn $ln"
            PostalCode      = $postal
            OfficePhone     = $office
            MobilePhone     = $mobile
            Path            = $ouDn
            AccountPassword = (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
            Enabled         = $true
        }
        New-ADUser @params -ErrorAction Stop
    }

    Get-ADUser -Filter * -SearchBase $ouDn -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
        Select DisplayName,PostalCode,OfficePhone,MobilePhone |
        Out-File (Join-Path $PSScriptRoot 'AdResults.txt') -Encoding UTF8
    Write-Host "AdResults.txt generated"
}
catch {
    Write-Error $_
}
