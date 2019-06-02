function Date {
    Get-Date -Format "yyyy.MM.dd HH:mm:ss.fff"
    }

$ServerInstance = "VLADPOPOV-NB"
$Database = "OBD2"

# (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where {$_.Name -Match "COM\d+"}).name

$port_ind = "8" #Index of COM Port
$baudrate = "38400"
$parity = [System.IO.Ports.Parity]::None  
$databits = 8  
$stopbits = [System.IO.Ports.StopBits]::One
  
$port = New-Object System.IO.Ports.SerialPort "COM$port_ind", $baudrate, $parity, $databits, $stopbits
# $port. = $true
$port.NewLine = "`r"
# $port.DTREnable =  "true
# $port.RtsEnable = "true"

try {
    Write-Output "$(Date) Opening COM port #$port_ind"
    $port.Open()
    # $port
    $commands = @("ATI", "ATL1", "ATH1", "ATS1", "ATAL", "ATMR11")
    
    $commands_track_next = @("3D 11 20", "00") # 9B, Left Up Button
    $commands_track_prev = @("3D 11 10", "00") # 5A, Left Down Button
    $commands_track_pause = @("3D 11 00", "80") # C8, Left Center Button
    
    $commands_volume_up = @("3D 11 04", "00") # C3, Right Up Button
    $commands_volume_down = @("3D 11 02", "00") # 76, Right Down Button 3D 11 20 1-2 step 3D 11 00 up to 3 steps
    $commands_volume_mute = @("3D 11 00", "02") # D4, Right Center Button

    $glass_driver_up = @("25 A0 00", "80") # Under construction 

    foreach ($command in $commands)
        {


        if ($command.Length -eq "8")
            {   
            $command = "ATSH $command"
            }
        
        Write-Output "$(Date) Sending $command command"
        $port.Write("$command`r")
        
        Start-Sleep -Milliseconds 500

        # $answer = $port.ReadLine().Trim()
        # Write-Output "$(Date) Answer: $answer"
 
        if ($command -like "ATM*")
            {
            Write-Output "$(Date) Monitor $command"

            while ($port.IsOpen)
                {
                $answer = $port.ReadLine().Trim()

                if ($answer -eq "3D 11 00 00 EE")
                    {
                    Continue # Noise
                    }

                $answer
                if ($answer -like "$($commands_track_next[0]) $($commands_track_next[1])*")
                    {
                    Write-Output "$(Date) Detected NextTrack button"
                    & "$PSScriptRoot\Next_Track.ps1"
                    }
                elseif ($answer -like "$($commands_track_prev[0]) $($commands_track_prev[1])*")
                    {
                    Write-Output "$(Date) Detected PrevTrack button"
                    & "$PSScriptRoot\Previous_Track.ps1"
                    }
                elseif ($answer -like "$($commands_track_pause[0]) $($commands_track_pause[1])*")
                    {
                    Write-Output "$(Date) Detected PlayPause button"
                    & "$PSScriptRoot\Play_Pause.ps1"
                    }
                }
            }
        else
            {
            $port.ReadLine()    
            }
           
        # $answer = $port.ReadExisting() -replace ([System.Environment]::NewLine, "")
        
        try {
            Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query "INSERT INTO [dbo].[Values] (COMPort, Value) VALUES ('$port_ind', '$answer')" -ErrorAction Stop  
            }
        catch {
            Write-Output "$(Date) $($_.Exception.Message)"
            }
        }
    }
catch {
    Write-Output "$(Date) $($_.Exception.Message)"
    }
finally
    {
    Write-Output "$(Date) Closing COM port #$port_ind"

    $port.Close()
    }
