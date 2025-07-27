<##
    Umer Mahmood
    Student ID: 001224010
    Restore-SQL.ps1 - recreate ClientDB and import contacts
##>

try {
    $server = '.\SQLEXPRESS'
    $database = 'ClientDB'

    # Check if ClientDB exists
    $checkDb = Invoke-Sqlcmd -ServerInstance $server -Query "SELECT name FROM sys.databases WHERE name = '$database'" -ErrorAction SilentlyContinue
    if ($checkDb) {
        Write-Host "$database already exists. Deleting it..."
        Invoke-Sqlcmd -ServerInstance $server -Query "DROP DATABASE [$database]" -ErrorAction Stop
        Write-Host "$database deleted."
    } else {
        Write-Host "$database does not exist."
    }

    # Create new database
    Invoke-Sqlcmd -ServerInstance $server -Query "CREATE DATABASE [$database]" -ErrorAction Stop
    Write-Host "$database created."

    # Create table Client_A_Contacts
    $createTable = @"
CREATE TABLE dbo.Client_A_Contacts (
    first_name NVARCHAR(50),
    last_name NVARCHAR(50),
    city NVARCHAR(100),
    county NVARCHAR(100),
    zip NVARCHAR(10),
    officePhone NVARCHAR(20),
    mobilePhone NVARCHAR(20)
);
"@
    Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $createTable -ErrorAction Stop
    Write-Host "Client_A_Contacts table created."

    # Insert data from CSV
    $csvPath = Join-Path $PSScriptRoot 'NewClientData.csv'
    $rows = Import-Csv -Path $csvPath
    foreach ($r in $rows) {
        $query = @"INSERT INTO dbo.Client_A_Contacts (first_name,last_name,city,county,zip,officePhone,mobilePhone) VALUES (N'$($r.first_name)',N'$($r.last_name)',N'$($r.city)',N'$($r.county)',N'$($r.zip)',N'$($r.officePhone)',N'$($r.mobilePhone)');"@
        Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $query -ErrorAction Stop
    }
    Write-Host "CSV data inserted into Client_A_Contacts."

    Invoke-Sqlcmd -Database $database -ServerInstance $server -Query "SELECT * FROM dbo.Client_A_Contacts" > "$PSScriptRoot\SqlResults.txt"
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}
