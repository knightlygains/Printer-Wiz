Write-Host " _____     _     _              _ _ _ _     
|  _  |___|_|___| |_ ___ ___   | | | |_|___ 
|   __|  _| |   |  _| -_|  _|  | | | | |- _|
|__|  |_| |_|_|_|_| |___|_|    |_____|_|___|
"
#Check for PsTools
if (-not(Test-Path "c:\windows\system32\psexec.exe" -PathType leaf) -OR -not(Test-Path "c:\windows\system32\psservice.exe" -PathType leaf)) {
    Write-Host "PsExec or PsService was not detected in your system32 folder. These tools will be necessary to invoke commands on the remote computer."
    Write-Host "These tools can be downloaded from https://learn.microsoft.com/en-us/sysinternals/downloads/psexec"
    $openLink = Read-Host "Would you like to go there now? (Y/N)"
    if ($openLink -eq "y") {
        Start-Process "https://learn.microsoft.com/en-us/sysinternals/downloads/psexec"
        exit
    }
    else {
        exit
    }
}

Function EnableWinRM {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Computers
    )
    #accept eula for psexec so it works
    Start-Process cmd -ArgumentList "/c psexec -accepteula" -WindowStyle hidden
    #Enable WinRM to use Get-CimInstance and Invoke-Command
    foreach ($computer in $Computers) {
        #check if winrm running on computer
        $result = winrm id -r:$Computers 2>$null
        if ($LastExitCode -eq 0) {
            Write-Host "WinRM already enabled on" $computer "..." -ForegroundColor green
        }
        else {
            Write-Host "Enabling WinRM on" $computer "..." -ForegroundColor red
            psexec.exe \\$computer -s C:\Windows\System32\winrm.cmd qc -quiet
        
            if ($LastExitCode -eq 0) {
                psservice.exe \\$computer restart WinRM
                $result = winrm id -r:$computer 2>$null
        
                if ($LastExitCode -eq 0) { Write-Host "WinRM successfully enabled!" -ForegroundColor green }
                else {
                    exit 1
                }
            }
            else {
                Write-Host "Couldn't enable WinRM on $computer."
            }
        }
    }
    Write-Host ""
}

Function getPrinters {
    param($errors)
    #Variable that allows us to loop through and get all printers on a remote computer.
    $printers = Get-CimInstance -Class Win32_Printer -ComputerName $Computer | Select-Object Name, PrinterStatus, LastErrorCode

    $variableNumber = 1
    #Loop through printers and create/update variables for each one.
    Write-Host "Printers:" -ForegroundColor Green
    foreach ($printer in $printers) {
        [string]$printerName = $printer.Name 
        $printerName = $printerName -replace "\*", ""
        Set-Variable -Name "Printer_$($variableNumber)_$($printerName)" -Value "$($printer.Name)" -Scope script #create variable of each printer with the name as the value
        
        $printerStatus = ""
        #Convert Printer Status code
        switch ($printer.PrinterStatus) {
            1 { $printerStatus = "Other" }
            2 { $printerStatus = "Unknown" }
            3 { $printerStatus = "Idle" }
            4 { $printerStatus = "Printing" }
            5 { $printerStatus = "Warmup" }
            6 { $printerStatus = "Stopped Printing" }
            7 { $printerStatus = "Offline" }
        }

        if ($errors -eq $true) {
            foreach ($printer in $printers) {
                $errorCode = $($printer.LastErrorCode)
                if ($null -eq $errorCode) {
                    $errorCode = "No error code"
                }
                Write-Host "Printer_$($variableNumber) $($printer.Name) " -NoNewline
                Write-Host "Error Code: $errorCode." -ForegroundColor Yellow
                $variableNumber += 1
            }
            return
        }

        Write-Host "Printer_$($variableNumber) $($printer.Name) Status: $printerStatus."
        $variableNumber += 1
    }
}

Function getComp {
    Write-Host "What's the computer hostname?" -ForegroundColor Yellow
    $comp = Read-Host
    return $comp
}

$Computer = getComp

if (Test-Connection $Computer -Count 1) {
    EnableWinRM -Computers $Computer
    Do {
        #Start of printer modification
        getPrinters #Call function to create variables of the printers

        Do {
            Write-Host ""
            Write-Host "Which printer do you want to change? (e.g. '1')" -ForegroundColor Yellow
            Write-Host "Type 'other' to modify spooler and other options"
            $answer = Read-Host
            if ($answer -eq "other") {
                break
            }
            else {
                $printerSelection = Get-Variable -Name "Printer_$($answer)_*"
                Write-Host "Printer selected: $($printerSelection.Value)."
            }
        } Until (-not($null -eq $printerSelection) -OR $answer -eq "other")

        $options = "uninstall", "rename", "testpage", "stop", "restart", "errors"
        $optionsPass = $false

        Do {
            #Get answer for what we will do with the printer
            if ($answer -eq "other") {
                Write-Host "What will we do? (e.g. 'restart')" -ForegroundColor Yellow
                Write-Host "Stop (Stops the print spooler)"
                Write-Host "Restart (restarts the print spooler)"
                Write-Host "Errors (refreshes printer list with last error codes)"
                $options = "stop", "restart", "errors"
            }
            else {
                Write-Host "What will we do? (e.g. 'rename')" -ForegroundColor Yellow
                Write-Host "Uninstall (uninstalls printer)"
                Write-Host "Rename (renames a printer)"
                Write-Host "TestPage (prints a test page)"
                Write-Host "Stop (Stops the print spooler)"
                Write-Host "Restart (restarts the print spooler)"
                Write-Host "Errors (refreshes printer list with last error codes)"
            }
            $answer2 = Read-Host
            foreach ($option in $options) {
                if ($answer2 -match $option) {
                    $optionsPass = $true
                    Write-Host ""
                    Continue
                }
            }
        }Until($optionsPass -eq $true)

        $printerToChange = $printerSelection.Value

        switch ($answer2) {
            "rename" {
                #Rename printer
                Write-Host "What will the new name be?" -ForegroundColor Yellow
                $newName = Read-Host #Get new name

                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    param($printerToChange, $newName)
                    Start-Process powershell -Wait -ArgumentList "Rename-Printer", "-Name", "'$printerToChange'", "-NewName", "'$newName'"
                } -ArgumentList ($printerToChange, $newName)

                Write-Host "Changed printer $printerToChange name to: $newName."
            }
            "testpage" {
                #Print a test page
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    param($printerToChange)
                    $printer = Get-WmiObject Win32_Printer | Where-Object { $_.name -eq "$printerToChange" }
                    Write-Host "$printerToChange, $printer"
                    $printer.PrintTestPage()
                } -ArgumentList ($printerToChange)
                Write-Host "Test page sent from printer $printerToChange on computer $Computer."
            }
            "uninstall" {
                #Uninstall printer
                Write-Host "Are you sure you want to uninstall $($printerSelection.Value) from $Computer? (Y/N)" -ForegroundColor Yellow
                $areYouSure = Read-Host
                if ($areYouSure -eq "y") {
                    Invoke-Command -ComputerName $Computer -ScriptBlock {
                        param($printerToChange)
                        Start-Process powershell -Wait -ArgumentList "Remove-Printer", "-Name", "'$printerToChange'"
                    } -ArgumentList ($printerToChange)
                    Write-Host "Removed printer: $printerToChange."
                }
                else {
                    Write-Host "Cancelled removal."
                }
            }
            "restart" {
                #Restart spooler
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $spoolerStatus = Get-Service Spooler | Select-Object Status
                    Get-Service Spooler | Stop-Service
                    if ($spoolerStatus -eq "stopped") {
                        Write-Host "Spooler stopped."
                        Start-Sleep 6
                    }
                    else {
                        Write-Host "Spooler failed to stop." -ForegroundColor Red
                        break
                    }
                    Write-Host "Starting spooler..."
                    Get-Service Spooler | Start-Service
                    if ($spoolerStatus -eq "Running") {
                        Write-Host "Spooler started."
                    }
                    else {
                        Write-Host "Spooler failed to start" -ForegroundColor Red
                    }
                }
                Write-Host "Print spooler has been restarted."
            }
            "stop" {
                #Stop spooler
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $spoolerStatus = Get-Service Spooler | Select-Object Status
                    Get-Service Spooler | Stop-Service
                    if ($spoolerStatus -eq "Stopped") {
                        Write-Host "Spooler stopped."
                    }
                    else {
                        Write-Host "Couldn't stop Spooler." -ForegroundColor Red
                    }
                }

                Read-Host "The spooler will need to be started again to continue. Press ENTER to restart" -ForegroundColor Yellow

                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    Get-Service Spooler | Start-Service
                    if ($spoolerStatus -eq "Running") {
                        Write-Host "Spooler started."
                    }
                    else {
                        Write-Host "Couldn't start Spooler." -ForegroundColor Red
                    }
                }
            }
            "errors" {
                #Get printer error codes
                getPrinters($errors = $true)
            }
        }

        Do {
            #Get answer for if we will modify another printer
            Write-Host "Would you like to modify another printer? (Y/N)" -ForegroundColor Yellow
            $continue = Read-Host
            if (-not($continue -eq "y" -OR $continue -eq "n")) {
                Write-Host "Invalid answer."
            }
        }Until($continue -eq "y" -OR $continue -eq "n")

        Write-Host ""
        
    }Until($continue -eq "n") #End of printer modification and script
}
else {
    Write-Host "Couldn't contact $Computer."
}