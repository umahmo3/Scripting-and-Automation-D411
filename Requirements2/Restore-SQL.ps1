<##
    Umer Mahmood
    Student ID: 001224010
    Task 2: Restore SQL Database
##>

# Must be run by a user with dbcreator or sysadmin rights
Import-Module SqlServer -ErrorAction Stop

try {
    $dbName = "ClientDB"
    $instance = ".\SQLEXPRESS"

    Write-Host "-- Checking for existing database: $dbName"
    $exists = Invoke-Sqlcmd -ServerInstance $instance -Database master `
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

    # Define and create the required table
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

    # Import data from the provided CSV
    $csvPath = Join-Path $PSScriptRoot "NewClientData.csv"
    Write-Host "-- Importing data from: $csvPath"
    $rows = Import-Csv -Path $csvPath -ErrorAction Stop

    foreach ($r in $rows) {
        $display = "$($r.first_name) $($r.last_name)"
        $insert = @"
USE [$dbName];
INSERT INTO dbo.Client_A_Contacts (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)
VALUES (N'$($r.first_name)', N'$($r.last_name)', N'$display', N'$($r.zip)', N'$($r.officePhone)', N'$($r.mobilePhone)');
"@
        Invoke-Sqlcmd -ServerInstance $instance -Query $insert -ErrorAction Stop
        Write-Host "Inserted: $display"
    }

    Write-Host "All client records imported." -ForegroundColor Green
}
catch {
    Write-Error "[SQL Script Error] $($_.Exception.Message)"
}

# Export table contents for submission
Invoke-Sqlcmd -Database $dbName -ServerInstance $instance -Query 'SELECT * FROM dbo.Client_A_Contacts' > "$PSScriptRoot\SqlResults.txt"
