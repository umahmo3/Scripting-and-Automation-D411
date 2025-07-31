<##
Umer Mahmood  |  Student ID 001224010
Restore-AD.ps1 – final version
##>

# This script is part of the WGU performance assessment.  Its purpose is to
# recreate the Finance organizational unit (OU) and populate it with user
# accounts taken from a provided CSV file.  The script is designed to run
# unattended and can be executed repeatedly.  If an existing Finance OU is
# detected, it will be removed and recreated to ensure a clean state.  At
# the end of execution, a file named AdResults.txt will be generated that
# contains a list of the new AD users and selected properties.

# NOTE: Make sure the Active Directory module is available on the system
# running this script.  The module is loaded explicitly below.

try {
    # Import the ActiveDirectory module if it is not already loaded.  If the module
    # cannot be imported, the script will throw an exception and jump to the
    # catch block.  This explicit import avoids ambiguous behaviour if the
    # module auto-loading is disabled in the host environment.
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "The ActiveDirectory module is not available on this system."
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    # Define the domain components for the lab environment.  These values are
    # combined to form distinguished names (DNs) used when creating and
    # locating objects in Active Directory.  If your domain differs from the
    # consultingfirm.com domain used in the assessment, adjust the values
    # accordingly.
    $domainDN = "dc=consultingfirm,dc=com"
    $financeOUName = "Finance"
    $financeOUDN   = "ou=$financeOUName,$domainDN"

    # Step 1: Check for an existing OU named 'Finance'.  We search within the
    # domain using an LDAP filter.  If found, we remove it recursively to
    # clean up any child objects.  Removal is done without prompting the
    # operator (-Confirm:$false) so that the script can run unattended.
    $existingOU = Get-ADOrganizationalUnit -LDAPFilter "(ou=$financeOUName)" -SearchBase $domainDN -ErrorAction SilentlyContinue
    if ($null -ne $existingOU) {
        Write-Host "An existing '$financeOUName' OU was found and will be deleted..." -ForegroundColor Yellow
        Remove-ADOrganizationalUnit -Identity $existingOU.DistinguishedName -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "The old '$financeOUName' OU has been removed." -ForegroundColor Yellow
    } else {
        Write-Host "No existing '$financeOUName' OU was detected." -ForegroundColor Cyan
    }

    # Step 2: Create the Finance OU.  The Path parameter specifies where in
    # the domain hierarchy the OU should be placed.  We intentionally
    # separate this step from the removal above to clearly communicate the
    # creation action to anyone reviewing the script.
    New-ADOrganizationalUnit -Name $financeOUName -Path $domainDN -ErrorAction Stop
    Write-Host "A fresh '$financeOUName' OU has been created." -ForegroundColor Green

    # Step 3: Import user data from the CSV file and create new AD accounts.
    # The CSV file is expected to reside in the same folder as this script.
    # To locate the file reliably, we compute the directory of the running
    # script via $MyInvocation.MyCommand.Path and build an absolute path to
    # financePersonnel.csv.  This approach allows the script to be run from
    # any working directory without breaking the file lookup.
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $csvFile   = Join-Path -Path $scriptDir -ChildPath "financePersonnel.csv"
    
    if (-not (Test-Path -Path $csvFile)) {
        throw "Input CSV file not found at '$csvFile'.  Please ensure financePersonnel.csv is in the same directory as this script."
    }

    # Read all records from the CSV.  Each record represents an employee in
    # the Finance department.  Column names in the CSV are used as
    # properties on the resulting objects (e.g., First_Name, Last_Name).  By
    # default, Import-Csv returns strings for each field; no additional
    # type conversion is required here.
    $staffList = Import-Csv -Path $csvFile

    foreach ($staff in $staffList) {
        # Build common attributes for the AD user based on the CSV record.
        # DisplayName concatenates the first and last names with a space.  We
        # also form the user principal name (UPN) by appending the domain name
        # to the samAccount identifier, which is stored in the CSV as
        # samAccount.  If your domain uses a different UPN suffix, adjust
        # accordingly.
        $givenName        = $staff.First_Name
        $surname          = $staff.Last_Name
        $displayName      = "${givenName} ${surname}"
        $samAccount       = $staff.samAccount
        $userPrincipal    = "${samAccount}@consultingfirm.com"

        # Informative output for each user being created.  Using Write-Host
        # with different colours makes it easier to distinguish steps when
        # watching the script run interactively.  This line can be removed
        # safely if minimal output is desired.
        Write-Host "Creating user: $displayName" -ForegroundColor DarkCyan

        # Create the user account.  A default password is set to satisfy the
        # New-ADUser cmdlet requirements.  In a real-world scenario, consider
        # generating unique passwords or prompting for secure input.  The
        # account is enabled immediately and flagged to require a password
        # change at next logon for security.  Additional properties (postal
        # code, office phone and mobile phone) are mapped directly from the
        # CSV record.
        New-ADUser `
            -Name $displayName `
            -GivenName $givenName `
            -Surname $surname `
            -DisplayName $displayName `
            -SamAccountName $samAccount `
            -UserPrincipalName $userPrincipal `
            -Path $financeOUDN `
            -PostalCode $staff.PostalCode `
            -OfficePhone $staff.OfficePhone `
            -MobilePhone $staff.MobilePhone `
            -AccountPassword (ConvertTo-SecureString -String "WguP@sswd2025!" -AsPlainText -Force) `
            -Enabled $true `
            -ChangePasswordAtLogon $true
    }

    # Step 4: Generate an output file with the newly created users.  The
    # provided specification requires the output to include DisplayName,
    # PostalCode, OfficePhone and MobilePhone properties.  Selecting the
    # properties explicitly via Select-Object ensures the file contains only
    # the requested data and is easy to read.
    # Generate the results file as specified in the assessment.  Rather than
    # formatting the output ourselves, we mirror the example line from the
    # assignment by piping the cmdlet output directly to a file.  The file
    # will reside in the same directory as this script.
    $resultsPath = Join-Path -Path $scriptDir -ChildPath "AdResults.txt"
    Get-ADUser -Filter * -SearchBase $financeOUDN -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > $resultsPath
    Write-Host "User import complete.  Results have been written to '$($resultsPath)'." -ForegroundColor Green
}
catch {
    # Any terminating errors in the try block are handled here.  The built-in
    # variable $_ holds the current error record.  Write-Host is used to
    # display a concise message along with the exception details so that
    # troubleshooting information is available.  Using a coloured output
    # helps draw attention to the error in the console.
    Write-Host "An unexpected error occurred:`n$($_.ToString())" -ForegroundColor Red
}
