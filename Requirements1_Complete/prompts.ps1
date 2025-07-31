# Umer Mahmood - Student ID: 001224010

Set-StrictMode -Version Latest

# finds all .log files and adds them to a text file
function Get-LogFiles {
    param(
        [string]$path
    )
    $logOutput = Join-Path -Path $path -ChildPath "DailyLog.txt"

    # add a timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logOutput -Value "--- Log check on $timestamp ---"

    # get .log files and add to the file
    $logs = Get-ChildItem -Path $path -Filter *.log | Select-Object -ExpandProperty Name
    
    if ($logs) {
        Add-Content -Path $logOutput -Value $logs
        Write-Host "Log names saved to '$logOutput'."
    } else {
        Add-Content -Path $logOutput -Value "No .log files found."
        Write-Host "No .log files found in '$path'."
    }
}

# lists everything in the folder
function Show-FolderContents {
    param(
        [string]$path
    )
    $contentsFile = Join-Path -Path $path -ChildPath "C916contents.txt"

    # list stuff, sort it, and save
    Get-ChildItem -Path $path |
        Sort-Object Name |
        Format-Table Name, Length, LastWriteTime -AutoSize |
        Out-File -FilePath $contentsFile

    Write-Host "Saved folder contents to '$contentsFile'."
}

function Show-SystemUsage {
    # Get CPU load as a simple percentage
    $cpu = Get-CimInstance Win32_Processor | Select-Object -ExpandProperty LoadPercentage
    Write-Host "CPU Load: $cpu %"

    # Get memory in MB
    $os  = Get-CimInstance Win32_OperatingSystem
    $totalMB = [int]($os.TotalVisibleMemorySize / 1024)
    $freeMB  = [int]($os.FreePhysicalMemory      / 1024)
    $usedMB  = $totalMB - $freeMB

    Write-Host "Total RAM: $totalMB MB"
    Write-Host "Used  RAM: $usedMB MB"
    Write-Host "Free  RAM: $freeMB MB"
}

# show running processes in a grid
function Show-RunningProcesses {
    Get-Process |
        Sort-Object -Property VirtualMemorySize64 |
        Out-GridView -Title "Running Processes"

    Write-Host "Showing processes in new window."
    Write-Host "(take a screenshot of the grid view window.)"
}


# Main script body
try {
    # Loop until user exits
    while ($true) {
        Write-Host "
PowerShell Script Menu:
1) Find .log files
2) List folder contents
3) Show CPU/RAM usage
4) Show running processes
5) Exit
"
        $choice = Read-Host "Enter an option [1-5]"

        if ($choice -eq '1') {
            Write-Host "Finding .log files..."
            Get-LogFiles -path $PSScriptRoot
        }
        elseif ($choice -eq '2') {
            Write-Host "Listing files..."
            Show-FolderContents -path $PSScriptRoot
        }
        elseif ($choice -eq '3') {
            Write-Host "Getting system usage..."
            Show-SystemUsage
        }
        elseif ($choice -eq '4') {
            Write-Host "Getting processes..."
            Show-RunningProcesses
        }
        elseif ($choice -eq '5') {
            Write-Host "Done."
            break # Exit the while loop
        }
        else {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
        }
    }

    # calculate hash after exiting
    $scriptFile = $MyInvocation.MyCommand.Path
    $hash = Get-FileHash -Path $scriptFile -Algorithm SHA256
    Write-Host "Script Hash (SHA256): $($hash.Hash)"

}
catch {
    # generic error handler
    Write-Error "An error occurred: $($_.Exception.Message)"
}