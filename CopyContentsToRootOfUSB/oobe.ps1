# ===========================
#
# By Chris Barry
#
# This is the main OOBE script for the custom recovery solution.
# It gets called by Unattend.xml during the Sysprep Process.
#
# ===========================
using namespace System.Security.Principal

# ==========
# Iterate through C:\Recovery\OEM\.env file for variables and add them
# ==========
function Read_Env_File {
    Write-Host "========= Begin Environment Variables ==========" -ForegroundColor Cyan
    $envPath = "C:\Recovery\OEM\.env"
    $envFile = Get-Content $envPath 2>$null
    if (-not $envFile) { return }

    foreach ($line in $envFile) 
	{
        $line = $line.Trim()
        if ($line -eq "" -or $line.StartsWith("#")) { continue }

        $parts = $line.Split('=', 2)
        if ($parts.Count -lt 2) { Write-Host "Issue parsing Environment Variable, skipping '$line'" -ForegroundColor Red; continue }

        $name  = $parts[0].Trim()
        $value = $parts[1].Trim()

        if ([string]::IsNullOrWhiteSpace($name) -or $name.Contains('#')) {
            Write-Host "Issue parsing Environment Variable, skipping '$line'" -ForegroundColor Red
            continue
        }

        Set-Content env:\$name $value
        $writtenValue = (Get-Content -Path ("env:\" + $name))

        if ($writtenValue -eq $value) {
            Write-Host "Writing Environment Variable: $name = $value" -ForegroundColor Green
        } else {
            Write-Host "Verification failed for environment variable $name. Expected: '$value' Actual: '$writtenValue' " -ForegroundColor Red
        }
    }
    Write-Host "========= End Environment Variables ==========" -ForegroundColor Cyan
}

# ==========
# Execute each reg file in C:\Recovery\OEM\reg folder in order by name, then date modified.
# ==========
function Write_Reg_Files
{
    Write-Host "========== Begin Registry ==========" -ForegroundColor Cyan
    $regFiles = Get-ChildItem -Path "C:\Recovery\OEM\reg" -Filter '*.reg' -File | 
                Sort-Object -Property Name, LastWriteTime

    foreach ($file in $regFiles) 
    {
        $filePath = $file.FullName
        Write-Host "Attempting to import: $filePath" -ForegroundColor Yellow

        $exitCode = Start-Process -FilePath 'reg.exe' `
                              -ArgumentList @('import', "`"$filePath`"") `
                              -Wait -NoNewWindow -PassThru

        if ($exitCode.ExitCode -eq 0) 
		{
			Write-Host "Successfully imported: $filePath (exited with ExitCode: $($exitCode.ExitCode))" -ForegroundColor Green
		}
		else # error
		{
            Write-Host "Failed to import: $filePath (ExitCode: $($exitCode.ExitCode))" -ForegroundColor Red
		}
    }
    Write-Host "========== End Registry ==========" -ForegroundColor Cyan
}

# ==========
# Execute each script in C:\Recovery\OEM\oobe_scripts folder in order by name, then date modified.
# ==========
function Run_OOBE_Scripts
{
    Write-Host "========== Begin OOBE Scripts==========" -ForegroundColor Cyan
    $executableExts = @('.ps1', '.exe', '.bat', '.cmd')

    $oobeScriptsFolder = Get-ChildItem -Path "C:\Recovery\OEM\oobe_scripts" -File |
						 Where-Object { $executableExts -contains $_.Extension.ToLower() } |
						 Sort-Object -Property Name, LastWriteTime

    foreach ($file in $oobeScriptsFolder) 
	{
        $filePath = $file.FullName
        $fileExt  = $file.Extension.ToLower()

        try 
	    {
			Write-Host "Attempting to execute: $filePath" -ForegroundColor Yellow

            switch ($fileExt) 
		    {
                '.ps1' 
			    {
                    $exitCode = Start-Process -FilePath 'powershell.exe' `
											  -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File', "`"$filePath`"" `
                                              -Wait -PassThru
                }
                '.exe' 
			    {
                    $exitCode = Start-Process -FilePath $filePath -Wait -PassThru
                }
                '.bat' 
			    {
                    $exitCode = Start-Process -FilePath $filePath -Wait -PassThru
                }
                '.cmd' 
			    {
                    $exitCode = Start-Process -FilePath $filePath -Wait -PassThru
                }
            }

            if ($exitCode.ExitCode -eq 0) 
		    {
                Write-Host "Successfully executed: $filePath (exited with ExitCode: $($exitCode.ExitCode))" -ForegroundColor Green
            } 
			else 
		    {
                Write-Host "Failed to Start-Process: $filePath (exited with code $($exitCode.ExitCode): $filePath" -ForegroundColor Red
            }
        } 
		catch 
	    {
            Write-Warning "There was an issue with executing $filePath : $_"
        }
    }
    Write-Host "========== End OOBE Scripts==========" -ForegroundColor Cyan
}

# Main entry point
function Main
{
	# initialize log folder (if not already) to prevent errors 
	New-Item -Path "C:\Recovery\OEM\logs" -ItemType Directory -Force | Out-Null
	
    # Start logging everything to a transcript file
    Start-Transcript -IncludeInvocationHeader -Path "C:\Recovery\OEM\logs\oobe.ps1_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
	Write-Host "Press enter to begin oobe customizations..... " -ForegroundColor Yellow -NoNewline; Read-Host
    Read_Env_File
    Write_Reg_Files
    Run_OOBE_Scripts
	
	Stop-Transcript
	
	Write-Host "Would you like to copy all log files from C:\Recovery\OEM\logs to C:\users\admin\desktop\logs? (Y to continue/N to exit): " -ForegroundColor Yellow -NoNewline
	$userResponse = Read-Host
	if($userResponse.ToLower() -eq "y")
	{
		$dest = "C:\Users\admin\Desktop\logs"
		
		Write-Host "Attempting to copy relevant log files to: $dest" -ForegroundColor Yellow
		New-Item -ItemType Directory -Force -Path $dest | Out-Null
		Copy-Item -Path "C:\Recovery\OEM\logs\*" -Destination $dest -Recurse -Force
		Copy-Item -Path "C:\Windows\Panther\setupact.log" -Destination "$dest\setupact_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log" -Force
		Copy-Item -Path "C:\Windows\Panther\setuperr.log" -Destination "$dest\setuperr_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log" -Force
		wevtutil.exe epl System "$dest\system_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').evtx"
		wevtutil.exe epl Application "$dest\application_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').evtx"
		Write-Host "Copied relevant log files to: $dest" -ForegroundColor Green
		Write-Host "Press enter to exit..." -ForegroundColor Yellow
		Read-Host
	}
	Write-Host "Goodbye." -Foregroundcolor Green
}

Main # main entry point
