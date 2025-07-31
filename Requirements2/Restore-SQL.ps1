<#
  Restore-SQL.ps1
  Umer Mahmood • 001224010
#>

Import-Module SqlServer -ErrorAction Stop   

try {
    $dbName   = "ClientDB"
    $instance = ".\SQLEXPRESS"

    # see if DB is already there
    Write-Host "-- peeking at sys.databases for $dbName"
    $exists = Invoke-Sqlcmd -ServerInstance $instance -Database master `
             -Query "SELECT name FROM sys.databases WHERE name='$dbName'" -ErrorAction Stop

    if ($exists) {
        Write-Host "Found it.  Yeeting $dbName..." -ForegroundColor Yellow
        Invoke-Sqlcmd -ServerInstance $instance -Query "DROP DATABASE [$dbName]" -ErrorAction Stop
        Write-Host "DB nuked." -ForegroundColor Green
    } else {
        Write-Host "No DB.  Cool—fresh build coming up."
    }

    # spin up blank DB
    Write-Host "-- making $dbName"
    Invoke-Sqlcmd -ServerInstance $instance -Query "CREATE DATABASE [$dbName]" -ErrorAction Stop
    Write-Host "DB ready." -ForegroundColor Green

    # table schema
    $tableDDL = @"
USE [$dbName];
CREATE TABLE dbo.Client_A_Contacts (
    FirstName   NVARCHAR(50),
    LastName    NVARCHAR(50),
    DisplayName NVARCHAR(101),
    City        NVARCHAR(50),
    County      NVARCHAR(50),
    PostalCode  NVARCHAR(20),
    OfficePhone NVARCHAR(20),
    MobilePhone NVARCHAR(20)
);
"@
    Write-Host "-- slapping down table Client_A_Contacts"
    Invoke-Sqlcmd -ServerInstance $instance -Query $tableDDL -ErrorAction Stop
    Write-Host "Table done." -ForegroundColor Green

    # suck in the CSV and shove rows into table
    $csvPath = Join-Path $PSScriptRoot "NewClientData.csv"
    Write-Host "-- pulling rows from $csvPath"
    $rows = Import-Csv -Path $csvPath -ErrorAction Stop

    foreach ($r in $rows) {
        $display = "$($r.first_name) $($r.last_name)"
        $insert = @"
USE [$dbName];
INSERT INTO dbo.Client_A_Contacts
    (FirstName, LastName, DisplayName, City, County, PostalCode, OfficePhone, MobilePhone)
VALUES
    (N'$($r.first_name)', N'$($r.last_name)', N'$display',
     N'$($r.city)', N'$($r.county)', N'$($r.zip)',
     N'$($r.officePhone)', N'$($r.mobilePhone)');
"@
        Invoke-Sqlcmd -ServerInstance $instance -Query $insert -ErrorAction Stop
        Write-Host "• added $display"
    }

    Write-Host "CSV import = ✅" -ForegroundColor Green
}
catch {
    Write-Error "[SQL Script Error] $($_.Exception.Message)"
}

Invoke-Sqlcmd -Database ClientDB -ServerInstance .\SQLEXPRESS -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt
