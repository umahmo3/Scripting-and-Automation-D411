<##
    Umer Mahmood
    Student ID: 001224010
    AUN1 Task 2: Restored AD and SQL Automation Scripts
    NOTE: Run PowerShell as a user with Domain Admin rights and elevated privileges.
    If you see a security warning, use `Unblock-File -Path Restore-AD.ps1`.
##>

# Ensure Active Directory module is available and run under a Domain Admin account
# NOTE: This script must be executed in a PowerShell session where your user is a member of the *Domain Admins* group.
# If you see a security warning, run:
#   Unblock-File -Path Restore-AD.ps1
# Then launch PowerShell as your domain user (Domain Admin) with elevated rights.
Import-Module ActiveDirectory -ErrorAction Stop

# --------------------------------------------------
try {
    $dbName = "ClientDB"
    $serverInstance = ".\SQLEXPRESS"

    Write-Host "Checking for database '$dbName'..."
    $checkDb = Invoke-Sqlcmd -ServerInstance $serverInstance -Database master `
        -Query "SELECT name FROM sys.databases WHERE name = '$dbName'" -ErrorAction Stop
    if ($checkDb) {
        Write-Host "Database '$dbName' exists. Dropping it..." -ForegroundColor Yellow
        Invoke-Sqlcmd -ServerInstance $serverInstance -Query "DROP DATABASE [$dbName]" -ErrorAction Stop
        Write-Host "Database dropped." -ForegroundColor Green
    } else {
        Write-Host "Database '$dbName' does not exist. Continuing..."
    }

    Write-Host "Creating database '$dbName'..."
    Invoke-Sqlcmd -ServerInstance $serverInstance -Query "CREATE DATABASE [$dbName]" -ErrorAction Stop
    Write-Host "Database created." -ForegroundColor Green

    $createTableQuery = @"
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
    Write-Host "Creating table Client_A_Contacts..."
    Invoke-Sqlcmd -ServerInstance $serverInstance -Query $createTableQuery -ErrorAction Stop
    Write-Host "Table created." -ForegroundColor Green

    $csvPath = Join-Path $PSScriptRoot "NewClientData.csv"
    Write-Host "Importing CSV from $csvPath..."
    $clients = Import-Csv -Path $csvPath -ErrorAction Stop
    foreach ($c in $clients) {
        $insertQuery = @"
USE [$dbName];
INSERT INTO dbo.Client_A_Contacts (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)
VALUES (N'$($c.FirstName)', N'$($c.LastName)', N'$($c.DisplayName)', N'$($c.PostalCode)', N'$($c.OfficePhone)', N'$($c.MobilePhone)');
"@
        Invoke-Sqlcmd -ServerInstance $serverInstance -Query $insertQuery -ErrorAction Stop
        Write-Host "Inserted client: $($c.DisplayName)"
    }
    Write-Host "All client data inserted." -ForegroundColor Green

    $exportPath = Join-Path $PSScriptRoot "SqlResults.txt"
    Write-Host "Exporting SQL results to $exportPath..."
    Invoke-Sqlcmd -ServerInstance $serverInstance -Database $dbName `
        -Query "SELECT * FROM dbo.Client_A_Contacts" -ErrorAction Stop `
        | Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "Export completed." -ForegroundColor Green
}
catch {
    Write-Error "Error in SQL script: $($_.Exception.Message)"
}
