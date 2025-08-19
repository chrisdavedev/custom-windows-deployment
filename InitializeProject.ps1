# ===========================
#
# By Chris Barry 
#
# This script self elevates, prompts for a required 
# dependency installation of Windows Configuration 
# Designer (WCD), finds the DAT file required to
# build the PPKG and uses ICD.exe to build the PPKG. 
# This script also prints out some useful information 
# for using this project.
# 
# ===========================

# ============================
# Elevate to an administrator shell if not Already
# ============================
function Self_Elevate
{
    Write-Host "Checking for administrator status....." -ForegroundColor Cyan
	$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $IsAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	if (-not $IsAdmin) {
        Write-Host "Elevating to administrator powershell session....." -ForegroundColor Cyan
		$psi = New-Object System.Diagnostics.ProcessStartInfo
		$psi.FileName = (Get-Command powershell.exe).Source
		$psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
		$psi.Verb = "runas"
		try { [Diagnostics.Process]::Start($psi) | Out-Null } catch { exit 1 }
		Write-Host "Couldn't elevate privileges, try again." -ForegroundColor Red
		exit
	}
}

# ============================
# Install WCD from Windows Store
# ============================
function Install_WCD
{
    Write-Host "Checking for Windows Configuration Designer (WCD)....." -ForegroundColor Cyan

    $pkgName = "Microsoft.WindowsConfigurationDesigner"

	# see if the package is already installed
    $wcdPackageExists = Get-AppxPackage | Where-Object { $_.Name -like "$pkgName*" }
    if ($wcdPackageExists) 
    {
        Write-Host "Windows Configuration Designer already installed, continuing....." -ForegroundColor Green
        return $true
    }

	Write-Host "Windows Configuration Designer (WCD) is required to generate the PPKG file, would you like to install it now? [Y/N]" -ForegroundColor Cyan -NoNewline
    $response = Read-Host

    if ($response.ToLower() -eq 'y') 
    {
		# try winget
		$wingetExists = Get-Command winget.exe -ErrorAction SilentlyContinue
		if ($wingetExists) 
		{
			try 
			{
				Write-Host "Using winget to install Windows Configuration Designer....." -ForegroundColor Cyan
				winget install --id 9NBLGGH4TX22 --accept-package-agreements --accept-source-agreements -e
				$wcdPackageExists = Get-AppxPackage | Where-Object { $_.Name -like "$pkgName*" }
				if ($wcdPackageExists) 
				{
					Write-Host "WCD installed successfully via winget....." -ForegroundColor Green
					return
				} 
				else 
				{
					Write-Host "Winget did not confirm installation, falling back to Microsoft Store....." -ForegroundColor Red
				}
			} 
			catch 
			{
				Write-Host "Winget installation attempt failed, falling back to Microsoft Store....." -ForegroundColor Red
			}
		}
		else 
		{
			Write-Host "winget not found on this system, falling back to Microsoft Store....." -ForegroundColor Red
		}
	}

	# open the store or web page if all else fails
	$storeUri = "ms-windows-store://pdp/?ProductId=9NBLGGH4TX22"
	$webUrl   = "https://www.microsoft.com/store/apps/9NBLGGH4TX22"

	try 
	{
		Write-Host "Attempting to open Windows Store to WCD page....." -ForegroundColor Cyan
		Start-Process -FilePath $storeUri -ErrorAction Stop
		Write-Host "Press enter to continue...."
		Read-Host
	} 
	catch 
	{
		Write-Warning "Could not open the Microsoft Store. Opening the web page instead....."
		Start-Process $webUrl
		Write-Host "Press enter to continue...."
		Read-Host
	}
}

# ============================
# Uninstall WCD if user requests
# ============================
function Uninstall_WCD
{
    Write-Host "Would you like to uninstall Windows Configuration Designer (WCD)? [Y/N]" -ForegroundColor Cyan -NoNewline
    $response = Read-Host

    if ($response.ToLower() -eq 'y') 
    {
        Write-Host "Attempting to remove Windows Configuration Designer....." -ForegroundColor Cyan
        
        # Correct package name for WCD
    	$pkgName = "Microsoft.WindowsConfigurationDesigner"

        # Look for installed package
        $existing = Get-AppxPackage | Where-Object { $_.Name -eq $pkgName }
        if ($existing) 
        {
            try 
            {
                Write-Host "Found WCD package, uninstalling....." -ForegroundColor Cyan
                Remove-AppxPackage -Package $existing.PackageFullName -AllUsers -ErrorAction Stop
                Write-Host "Windows Configuration Designer removed successfully....." -ForegroundColor Green
            } 
            catch 
            {
                Write-Host "Failed to uninstall Windows Configuration Designer....." -ForegroundColor Red
            }
        } 
        else 
        {
            Write-Host "No Windows Configuration Designer installation detected, nothing to uninstall....." -ForegroundColor Yellow
        }
    }
    elseif ($response.ToLower() -eq 'n') 
    {
        Write-Host "Uninstall canceled by user....." -ForegroundColor Cyan
    }
    else 
    {
        Write-Host "Invalid input. Please enter Y or N next time....." -ForegroundColor Red
    }
}


# ============================
# Find and copy Microsoft-Desktop-Provisioning.dat from Windows Configuration Designer (WCD)
# which is a required file for creating the PPKG.
# ============================
function Copy_WCD_DAT_File 
{
    $microsoftDatFile = "Microsoft-Desktop-Provisioning.dat"
    $destPath    = Join-Path -Path $PSScriptRoot -ChildPath "$microsoftDatFile"
	Write-Host "Searching for $microsoftDatFile in common WCD locations....." -ForegroundColor Cyan
    if (Test-Path $destPath) 
	{
        Write-Host "$destPath already present, continuing....." -ForegroundColor Green
        return $destPath
    }

    $commonLocations = @(
        "C:\Program Files\Windows Configuration Designer",
        "C:\Program Files (x86)\Windows Configuration Designer",
        "C:\ProgramData",
        "$env:LOCALAPPDATA\Microsoft\Windows Configuration Designer",
        "$env:ProgramFiles\Microsoft\Windows Configuration Designer",
        "$env:ProgramFiles(x86)\Microsoft\Windows Configuration Designer",
		"C:\Program Files\WindowsApps\",
        $env:ProgramFiles, $env:ProgramFilesx86
    ) | Where-Object { $_ -and (Test-Path $_) }

    # check each location in the above list
    foreach ($base in $commonLocations) 
	{
		Write-Host "Searching in $base....." -ForegroundColor Cyan
        $cand = Get-ChildItem -Path $base -Recurse -Filter $microsoftDatFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cand) 
		{
            Write-Host "Found candidate at $($cand.FullName)" -ForegroundColor Cyan
            Copy-Item -Path $cand.FullName -Destination $destPath -Force
            Write-Host "Copied '$($cand.FullName)' -> '$destPath'" -ForegroundColor Green
            return $destPath
        }
    }

    # check some higher order locations
    foreach ($root in @("$env:ProgramFiles","$env:ProgramFiles(x86)","$env:ProgramData","$env:LOCALAPPDATA")) 
	{
		Write-Host "Searching in $root....." -ForegroundColor Cyan
        if (-not (Test-Path $root)) { continue }
        $cand = Get-ChildItem -Path $root -Recurse -Filter $microsoftDatFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cand) 
		{
			Write-Host "Found candidate at $($cand.FullName)" -ForegroundColor Cyan
            Copy-Item -Path $cand.FullName -Destination $destPath -Force
            Write-Host "Copied '$($cand.FullName)' -> '$destPath'" -ForegroundColor Green
			return $destPath
		}
    }

    Write-Host "Unable to locate $microsoftDatFile in common WCD locations. Exiting..." -ForegroundColor Red
	Read-Host
	exit 1
}

# ============================
# Create PPKG file
# ============================
function Create_PPKG
{
    $xmlPath    = Join-Path -Path $PSScriptRoot -ChildPath 'Customizations.xml'
    $copyPath   = Join-Path -Path $PSScriptRoot -ChildPath 'CopyContentsToRootOfUSB'
    $ppkgFile   = Join-Path -Path $copyPath -ChildPath 'Provisioning.ppkg'
    $datFile    = Join-Path -Path $PSScriptRoot -ChildPath 'Microsoft-Desktop-Provisioning.dat'

    if (-not (Test-Path -Path $copyPath -PathType Container)) 
	{
        Write-Host "Could not find $copyPath, please make sure the project directory has all contents..." -ForegroundColor Red
		Read-Host
		exit 1
    }

    if (-not (Test-Path -Path $xmlPath)) 
	{
        Write-Host "Customizations.xml not found at $xmlPath, please make sure the project directory has all contents..." -ForegroundColor Red
		Read-Host
		exit 1
    }

    Write-Host "Attempting to build provisioning package using icd.exe...." -ForegroundColor Cyan
    icd.exe /Build-ProvisioningPackage /CustomizationXML:"$xmlPath" /PackagePath:"$ppkgFile" /StoreFile:"$datFile" +Overwrite
}

# ============================
# Prompt the user to copy project 
# files to a flash drive now
# ============================
function Copy_Files_To_USB {
    $response = Read-Host "Would you like to copy files to USB flash drive (will overwrite existing files with same names)? (Y/N)"
    
    if ($response.ToLower() -eq "y") 
	{
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select the destination folder on the USB flash drive"
        $folderDialog.ShowNewFolderButton = $true

        $result = $folderDialog.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) 
		{
            $destinationPath = $folderDialog.SelectedPath
			$sourcePath = Join-Path -Path $PSScriptRoot -ChildPath 'CopyContentsToRootOfUSB'

			Write-Host "Copying files to $destinationPath..." -ForegroundColor Cyan
			
			Copy-Item -Path "$sourcePath\*" -Destination $destinationPath -Recurse -Force
        } 
		else 
		{
            Write-Host "No destination selected. Aborting copy operation." -ForegroundColor Red
        }
    } 
	else 
	{
        Write-Host "User chose not to copy files to USB." -ForegroundColor Red
    }
}

# ============================
# Instructions on the files in this project
# ============================
function Print_File_Instructions
{
	Clear-Host
	Write-Host "The Provisioning Package (Provisioning.ppkg) has been created using ICD.exe.`n"
    Write-Host "If you navigate in Explorer to the project directory and look at the folder CopyContentsToRootOfUSB, you will see several files... Described below.`n"
	Write-Host "1. The .\CopyContentsToRootOfUSB\provision.ps1 file." -ForegroundColor Yellow 
	Write-Host "`tThis is a custom file and it does a few things."
	Write-Host "`n`tIt can be called from the PPKG or directly using .\provision.ps1,"
	Write-Host "`tit will copy the required files over to C:\Recovery\OEM for use with"
	Write-Host "`tfuture resets using the Reset This PC feature.`n"
	Write-Host "`tThis script will also invoke sysprep.exe using our custom unattend file to"
	Write-Host "`trun customizations on the target machine immediately.`n"
	Write-Host "2. The .\CopyContentsToRootOfUSB\ResetConfig.xml file." -ForegroundColor Yellow 
	Write-Host "`tA file used in the Reset This PC feature to call runcustomizations.cmd.`n"
	Write-Host "3. The .\CopyContentsToRootOfUSB\runcustomizations.cmd file." -ForegroundColor Yellow 
	Write-Host "`tThis gets executed AFTER the PC is reset, and copies the"
	Write-Host "`tcustom unattend file over to C:\Windows\Panther to be used"
	Write-Host "`tin the OOBE process.`n"
	Write-Host "4. The .\CopyContentsToRootOfUSB\unattend.xml file." -ForegroundColor Yellow 
	Write-Host "`tThis is the file that Sysprep uses to run customizations,"
	Write-Host "`tand in our case it has a FirstLogonCommand for the admin user"
	Write-Host "`tthat calls oobe.ps1.`n"
	Write-Host "5. The .\CopyContentsToRootOfUSB\Provisioning.ppkg file." -ForegroundColor Yellow 
	Write-Host "`tThis is the package created earlier in this script. It is to be placed at"
	Write-Host "`tthe root of your removable media. When this file is recognized on removable media"
	Write-Host "`tduring Windows Setup, the Setup will execute it. In our setup, this PPKG calls"
	Write-Host "`tthe provision.ps1 file.`n"
	Write-Host "6. The .\CopyContentsToRootOfUSB\Provisioning.cat file." -ForegroundColor Yellow 
	Write-Host "`tThis is a required file generated by ICD.exe it is"
	Write-Host "`twhat allows the PPKG to run on a new system.`n"
	Write-Host "7. The .\CopyContentsToRootOfUSB\oobe.ps1 file." -ForegroundColor Green 
	Write-Host "`tThis script contains the logic for recursively running the custom"
	Write-Host "`tOOBE scripts, registry files, and environment variables from their"
	Write-Host "`trespective folders in this project directory. It is called by Sysprep"
	Write-Host "`tin the unattend.xml answer file during the OOBE system pass as"
	Write-Host "`ta FirstLogonCommand`n"
	Write-Host "8. The .\CopyContentsToRootOfUSB\oobe_scripts folder." -ForegroundColor Green 
	Write-Host "`tEach script (.EXE, .CMD, .BAT, & .PS1 supported) in this folder is executed"
	Write-Host "`tby the oobe.ps1 during the FirstLogon in the Out Of Box Experience (OOBE)."
	Write-Host "`tYou should place the custom files you would like executed during FirstLogon in this folder.`n"
	Write-Host "9. The .\CopyContentsToRootOfUSB\reg folder." -ForegroundColor Green 
	Write-Host "`tEach REG file placed in this directory will be executed during First Logon."
	Write-Host "`tYou should place the custom files you would like executed at OOBE in this folder.`n"
	
}


function Print_Next_Steps
{
	Clear-Host
	Write-Host "Quick Start Information:`n" -ForegroundColor Green
	Write-Host "- This script overwrites C:\Recovery, which may cause your OEM recovery" -ForegroundColor Red
	Write-Host "  solution to break. Please make a backup of C:\Recovery before running" -ForegroundColor Red
	Write-Host "  this tool if reinstalling Windows from scratch is not an option for you." -ForegroundColor Red
	Write-Host "`nStarting The Tool:" -ForegroundColor Green
	Write-Host "Method 1:`n`tPlug the removable media into a PC that is on the language screen in Windows Setup, the Provisioning"
	Write-Host "`tPackage will be picked up automatically by setup and will execute Provision.ps1."
	Write-Host "`tYou should end up on a desktop with oobe.ps1 waiting for your input."
	Write-Host "Method 2:`n`tPlug the removable media into a PC that has already completed Windows Setup."
	Write-Host "`tThen from an elevated powershell window, execute '.\provision.ps1'."
	Write-Host "`n`tUsing Method 2 requires the extra step of running the 'Reset This PC' feature"
	Write-Host "`tby going to Settings -> Update & Security -> Recovery -> Reset This PC"
}



# Main entry point
function Main
{
	Add-Type -AssemblyName System.Windows.Forms
	
	Self_Elevate
	Install_WCD
	Copy_WCD_DAT_File
	Create_PPKG
	Copy_Files_To_USB
	Uninstall_WCD
	
	Write-Host "`n`nInitialization complete, press enter for further instructions...`n`n" -ForegroundColor Green
	Read-Host
	Clear-Host
	Print_File_Instructions
	Write-Host "Press enter for next steps..." -NoNewline -ForegroundColor Green
	Read-Host
	Print_Next_Steps
	Write-Host "`nPress enter to exit program..." -NoNewline -ForegroundColor Green
	Read-Host
}

Main