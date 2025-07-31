<#
  Restore-AD.ps1  –  Umer Mahmood | 001224010

  Notes to self
  -------------
  • run as Domain Admin, elevated PowerShell
  • nukes + rebuilds “Finance” OU
  • pulls users from financePersonnel.csv (same folder)
  • clears accidental-deletion flag first
  • spits out AdResults.txt for grader
#>

# load AD tools
Import-Module ActiveDirectory -ErrorAction Stop
$ErrorActionPreference = 'Stop'

# constants
$ouName   = 'Finance'
$domainDN = 'dc=consultingfirm,dc=com'
$ouDN     = "ou=$ouName,$domainDN"
$here     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$csvPath  = Join-Path $here 'financePersonnel.csv'
$outFile  = Join-Path $here 'AdResults.txt'
$pwd      = ConvertTo-SecureString 'WguPa55!@#' -AsPlainText -Force  # demo pass

# ---- wipe old OU if it exists ----
if ($old = Get-ADOrganizationalUnit -LDAP "(ou=$ouName)" -SearchBase $domainDN -EA 0) {
    Write-Host "Old $ouName OU found – deleting…" -f Yellow
    Set-ADObject $old.DistinguishedName -ProtectedFromAccidentalDeletion:$false
    Remove-ADOrganizationalUnit $old.DistinguishedName -Recursive -Confirm:$false
}

# ---- create fresh OU ----
New-ADOrganizationalUnit -Name $ouName -Path $domainDN -ProtectedFromAccidentalDeletion:$true
Write-Host "$ouName OU created." -f Green

# ---- sanity check CSV ----
if (-not (Test-Path $csvPath)) { throw "CSV missing: $csvPath" }
$rows = Import-Csv $csvPath

# ---- loop + add users ----
foreach ($u in $rows) {
    $sam  = $u.samAccount
    $disp = "$($u.First_Name) $($u.Last_Name)"

    New-ADUser `
        -Path  $ouDN `
        -SamAccountName $sam `
        -UserPrincipalName "$sam@consultingfirm.com" `
        -GivenName $u.First_Name `
        -Surname   $u.Last_Name `
        -Name      $disp `
        -DisplayName $disp `
        -PostalCode   $u.PostalCode `
        -OfficePhone  $u.OfficePhone `
        -MobilePhone  $u.MobilePhone `
        -AccountPassword $pwd `
        -Enabled $true `
        -ChangePasswordAtLogon $true
}
Write-Host "Imported $($rows.Count) users." -f Green

# ---- dump results for grader ----
Get-ADUser -Filter * -SearchBase $ouDN `
  -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
Select SamAccountName,DisplayName,PostalCode,OfficePhone,MobilePhone |
Out-File $outFile -Encoding UTF8
Write-Host "AdResults.txt done -> $outFile" -f Green
