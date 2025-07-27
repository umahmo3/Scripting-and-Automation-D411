# Umer Mahmood - Student ID: 001224010


try {

    # do-while loop keeps the menu running until the user picks option 5.
    do {
        # Display menu choices.
        Write-Host "" 
        Write-Host "PowerShell Script Menu:"
        Write-Host "1. List .log files and add to DailyLog.txt" # Requirement B1
        Write-Host "2. List all files in this folder and save to C916contents.txt" # Requirement B2
        Write-Host "3. Show current CPU and RAM (memory) usage" # Requirement B3
        Write-Host "4. Show running processes (sorted by memory)" # Requirement B4
        Write-Host "5. Exit the script" # Requirement B5
        Write-Host ""

        # Ask user to enter their choice.
        $userChoiceInput = Read-Host "Please enter a number from 1 to 5"

        # Switch statement will perform an action based on the user's input.
        switch ($userChoiceInput) {
            '1' {
                Write-Host "Option 1: Finding .log files..."
                # The script is inside the "Requirements1" folder, so $PSScriptRoot is that folder's path.

                $currentScriptFolder = $PSScriptRoot
                $logFileOutputPath = Join-Path -Path $currentScriptFolder -ChildPath "DailyLog.txt"

                # Get the current date and time for the log entry. 
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                
                # Add the timestamp to DailyLog.txt.
                Add-Content -Path $logFileOutputPath -Value "--- Log check on $timestamp ---"

                $logFilesFound = Get-ChildItem -Path $currentScriptFolder -File | Where-Object {$_.Name -match '\.log$'} | Select-Object -ExpandProperty Name
                
                if ($logFilesFound) {

                    # Add the list of found .log file names to DailyLog.txt.
                    Add-Content -Path $logFileOutputPath -Value $logFilesFound
                    Write-Host ".log file names have been appended to '$logFileOutputPath'."
                } else {
                    Add-Content -Path $logFileOutputPath -Value "No .log files were found in this folder at this time."
                    Write-Host "No .log files found in '$currentScriptFolder'."
                }
            }
            '2' {
                
                Write-Host "Option 2: Listing all files in this folder..."

                $currentScriptFolder = $PSScriptRoot
                $contentsFileOutputPath = Join-Path -Path $currentScriptFolder -ChildPath "C916contents.txt"

                # Get all items, sort by name, format as table, then save to file.
                Get-ChildItem -Path $currentScriptFolder | Sort-Object Name | Format-Table Name, Length, LastWriteTime -AutoSize | Out-File -FilePath $contentsFileOutputPath
                
                Write-Host "The list of files has been saved to '$contentsFileOutputPath'."
            }
            '3' {
 
                Write-Host "Option 3: Displaying current CPU and Memory (RAM) usage..."

                # Get CPU Load. Using Get-Counter for this.
                # '\Processor(_Total)\% Processor Time' gets the overall CPU usage.
                $cpuInfo = Get-Counter '\Processor(_Total)\% Processor Time'
                # The 'CookedValue' is the actual percentage value.
                $currentCpuLoad = $cpuInfo.CounterSamples[0].CookedValue 
                Write-Host "Current CPU Usage: $($currentCpuLoad) %"

                # Get Memory (RAM) Usage. Get-CimInstance Win32_OperatingSystem has this info.
                $memInfo = Get-CimInstance -ClassName Win32_OperatingSystem
                # Values are in Kilobytes, so I'll convert to Megabytes (MB) by dividing by 1024.
                $totalRamMB = [math]::Round($memInfo.TotalVisibleMemorySize / 1024)
                $freeRamMB = [math]::Round($memInfo.FreePhysicalMemory / 1024)
                $usedRamMB = $totalRamMB - $freeRamMB
                Write-Host "Total RAM: $totalRamMB MB"
                Write-Host "Used RAM: $usedRamMB MB"
                Write-Host "Free RAM: $freeRamMB MB"
                Write-Host "(take a screenshot of these CPU and Memory results.)"
            }
            '4' {

                Write-Host "Option 4: Listing all running processes..."

                Get-Process | Sort-Object VM | Out-GridView -Title "Running Processes (Sorted by Virtual Memory)"
                
                Write-Host "A list of running processes has been displayed in a new window."
                Write-Host "(take a screenshot of the grid view window.)"
            }
            '5' {

                Write-Host "Exiting the script now. Goodbye!"
                # 'break' command exits the current loop (the 'do-while' loop).
            }
            default {
                Write-Warning "Invalid choice. Please enter a number between 1 and 5."
            }
        }

    } while ($userChoiceInput -ne '5') 
}
catch [System.OutOfMemoryException] {
    # If the script tries to use too much memory, this error message will be shown.
    Write-Error "A critical error occurred: The system is out of memory."
    Write-Error "Please close some applications and try running the script again."
}
catch {
    Write-Error "An unexpected error happened in the script: $($_.Exception.Message)"
    Write-Host "The script had to stop due to this error."
}
