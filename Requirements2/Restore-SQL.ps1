# Umer Mahmood
# Student ID: 001224010
# Restore SQL DB and import data

Import-Module SqlServer

$server = ".\SQLEXPRESS"
$db = "ClientDB"
$table = "Client_A_Contacts"

try {
    # Check and drop existing DB
    $exists = Invoke-Sqlcmd -ServerInstance $server -Query "SELECT name FROM sys.databases WHERE name = '$db'"
    if ($exists) {
        Write-Host "Database $db exists. Dropping..."
        Invoke-Sqlcmd -ServerInstance $server -Query "DROP DATABASE [$db]"
        Write-Host "Database $db dropped."
    } else {
        Write-Host "Database $db does not exist."
    }

    # Create DB
    Invoke-Sqlcmd -ServerInstance $server -Query "CREATE DATABASE [$db]"
    Write-Host "Database $db created."

    # Create table
    Invoke-Sqlcmd -ServerInstance $server -Database $db -Query @"
    CREATE TABLE $table (
        [First Name] NVARCHAR(50),
        [Last Name] NVARCHAR(50),
        [Email] NVARCHAR(100),
        [Phone] NVARCHAR(25)
    )
"@
    Write-Host "Table $table created."

    # Import from CSV and insert into table
    $data = Import-Csv ".\NewClientData.csv"
    foreach ($row in $data) {
        $q = "INSERT INTO $table ([First Name],[Last Name],[Email],[Phone]) VALUES (N'$($row.'First Name')',N'$($row.'Last Name')',N'$($row.Email)',N'$($row.Phone)')"
        Invoke-Sqlcmd -ServerInstance $server -Database $db -Query $q
    }

    # Export to file
    Invoke-Sqlcmd -ServerInstance $server -Database $db -Query "SELECT * FROM $table" | Out-File -FilePath ".\SqlResults.txt"
    Write-Host "SQL export to SqlResults.txt done."
}
catch {
    Write-Error "Error: $_"
}
