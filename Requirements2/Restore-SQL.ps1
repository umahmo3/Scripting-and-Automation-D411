<##
    Umer Mahmood
    Student ID: 001224010
    Task 2: Restore SQL Database
##>

# Must be run by a user with dbcreator or sysadmin rights
Import-Module SqlServer -ErrorAction Stop
function Get-Field {
    param($row, [string[]]$names)
    foreach ($n in $names) {
        $val = $row.$n
        if ($val) { return $val }
    }
    return $null
}


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

    $inserted = 0
    foreach ($r in $rows) {
        # Handle different header variations
        $first  = Get-Field $r @("FirstName","first_name","First Name")
        $last   = Get-Field $r @("LastName","last_name","Last Name")
        $postal = Get-Field $r @("PostalCode","zip","Postal Code")
        $office = Get-Field $r @("OfficePhone","officePhone","Office Phone")
        $mobile = Get-Field $r @("MobilePhone","mobilePhone","Mobile Phone")
        $display = $r.DisplayName
        if (-not $display) { $display = "$first $last" }

        $insert = @"
USE [$dbName];
INSERT INTO dbo.Client_A_Contacts (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)
VALUES (N'$first', N'$last', N'$display', N'$postal', N'$office', N'$mobile');
"@
        Invoke-Sqlcmd -ServerInstance $instance -Query $insert -ErrorAction Stop
        Write-Host "Inserted: $display"
        $inserted++
    }

    Write-Host "$inserted rows inserted." -ForegroundColor Green
}
catch {
    Write-Error "[SQL Script Error] $($_.Exception.Message)"
}

# Export table contents for submission
# Required command for task verification
Invoke-Sqlcmd -Database ClientDB -ServerInstance .\SQLEXPRESS -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt
