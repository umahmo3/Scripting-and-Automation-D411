# Umer Mahmood
# Student ID: 001224010
# Restore SQL DB and import data

# Ensure SqlServer module is available and current user can create databases
# NOTE: Run PowerShell as a user with SQL Server sysadmin or dbcreator rights.
Import-Module SqlServer -ErrorAction Stop

try {
    $dbName = "ClientDB"
    $instance = ".\SQLEXPRESS"

    Write-Host "-- Checking for existing database: $dbName"
    $exists = Invoke-Sqlcmd -ServerInstance $instance -Database master \
        -Query "SELECT name FROM sys.databases WHERE name='$dbName'" -ErrorAction Stop
    if ($exists) {
        Write-Host "Database exists. Dropping $dbName..." -ForegroundColor Yellow
        Invoke-Sqlcmd -ServerInstance $instance -Query "DROP DATABASE [$dbName]" -ErrorAction Stop
        Write-Host "Database dropped." -ForegroundColor Green
    } else {
        Write-Host "Database not found. Creating new $dbName."
    }

    Write-Host "-- Creating database: $dbName"
    Invoke-Sqlcmd -ServerInstance $instance -Query "CREATE DATABASE [$dbName]" -ErrorAction Stop
    Write-Host "Database created." -ForegroundColor Green

    $tableDDL = @"
USE [$dbName];
CREATE TABLE dbo.Client_A_Contacts (
    FirstName NVARCHAR(50),
    LastName NVARCHAR(50),
    DisplayName NVARCHAR(101),
    PostalCode NVARCHAR(20),
    OfficePhone NVARCHAR(20),
    MobilePhone NVARCHAR(20)
);
"@
    Write-Host "-- Creating table Client_A_Contacts"
    Invoke-Sqlcmd -ServerInstance $instance -Query $tableDDL -ErrorAction Stop
    Write-Host "Table created." -ForegroundColor Green

    # Import CSV data
    $csvPath = Join-Path $PSScriptRoot "NewClientData.csv"
    Write-Host "-- Importing data from: $csvPath"
    $rows = Import-Csv -Path $csvPath -ErrorAction Stop
    foreach ($r in $rows) {
        $insert = @"
USE [$dbName];
INSERT INTO dbo.Client_A_Contacts (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)
VALUES (N'$($r.FirstName)', N'$($r.LastName)', N'$($r.DisplayName)', N'$($r.PostalCode)', N'$($r.OfficePhone)', N'$($r.MobilePhone)');
"@
        Invoke-Sqlcmd -ServerInstance $instance -Query $insert -ErrorAction Stop
        Write-Host "Inserted: $($r.DisplayName)"
    }
    Write-Host "All client records imported." -ForegroundColor Green

    # Export SQL results
    $sqlOutput = Join-Path $PSScriptRoot "SqlResults.txt"
    Write-Host "-- Exporting SQL results to: $sqlOutput"
    Invoke-Sqlcmd -ServerInstance $instance -Database $dbName \
        -Query "SELECT * FROM dbo.Client_A_Contacts" -ErrorAction Stop |
        Out-File -FilePath $sqlOutput -Encoding UTF8
    Write-Host "SQL export complete." -ForegroundColor Green
}
catch {
    Write-Error "[SQL Script Error] $($_.Exception.Message)"
}
