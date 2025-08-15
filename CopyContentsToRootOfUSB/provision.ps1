# ===========================
#
# This script wipes out C:\Recovery
# and injects the custom oobe_scripts,
# reg, and .env files into two places:
# 1. Into the Local C:\ (unattended answer file calls oobe.ps1)
# 2. Into the "Reset This PC" Push Button Reset configuration for future recoveries.
# 
# NOTE: This script runs without confirmation prompts,
#		you should only use this if you want to overwrite
# 		the OEM Recovery Solution on your machine. You
#		may need to reinstall Windows or modify the
#		recovery solution on your PC after running this.
#
# ===========================
$recoveryRootFolder = "C:\Recovery"
$logsFolder   = Join-Path $recoveryRootFolder "OEM\logs"

# Remove all contents inside C:\Recovery without deleting the folder itself
Get-ChildItem -Path $recoveryRootFolder -Force -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# Make sure C:\Recovery exists, create if not
New-Item -ItemType Directory -Path $recoveryRootFolder -Force | Out-Null

# Make sure logs folder exists, create if not
New-Item -ItemType Directory -Path $logsFolder -Force | Out-Null

# Start logging (transcript) to the logs folder, useful for debugging
$logPath = Join-Path $logsFolder ("provision.ps1_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log")
Start-Transcript -IncludeInvocationHeader -Path $logPath

# Temporary test admin user
New-LocalUser -Name admin -Password (ConvertTo-SecureString "admin" -AsPlainText -Force)
Add-LocalGroupMember -Group "Administrators" -Member "admin"

# Set permissions for C:\Recovery so we can write files.
# Per Microsoft documentation (https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/recovery-strategy-for-common-customizations?view=windows-11)
icacls "C:\Recovery" /inheritance:r
icacls "C:\Recovery" /grant:r "SYSTEM:(OI)(CI)(F)"
icacls "C:\Recovery" /grant:r "*S-1-5-32-544:(OI)(CI)(F)"
takeown /f "C:\Recovery" /a
attrib +H "C:\Recovery"

# Copy over our customization files to work with 
# the "Reset This PC" Push Button Reset Feature
$currentDirectory=[System.IO.Path]::GetPathRoot($PSScriptRoot)
Copy-Item -Path "$currentDirectory\ResetConfig.xml" -Destination "C:\Recovery\OEM\ResetConfig.xml" -Force
Copy-Item -Path "$currentDirectory\runcustomizations.cmd" -Destination "C:\Recovery\OEM\runcustomizations.cmd" -Force
Copy-Item -Path "$currentDirectory\unattend.xml" -Destination "C:\Recovery\OEM\unattend.xml" -Force
Copy-Item -Path "$currentDirectory\oobe.ps1" -Destination "C:\Recovery\OEM\oobe.ps1" -Force
Copy-Item -Path "$currentDirectory\.env" -Destination "C:\Recovery\OEM\.env" -Force
Copy-Item -Path "$currentDirectory\oobe_scripts" "C:\Recovery\OEM\oobe_scripts" -Recurse -Force
Copy-Item -Path "$currentDirectory\reg" "C:\Recovery\OEM\reg" -Recurse -Force

# Tell Sysprep.exe to use the unattended.xml file on the current dir (likely removeable media)
# This will apply the configurations to the C:\ drive via C:\Recovery\OEM\oobe.ps1
C:\Windows\System32\Sysprep\sysprep.exe /oobe /reboot /unattend:"$currentDirectory\unattend.xml"

Stop-Transcript
