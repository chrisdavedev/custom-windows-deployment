# ============================
# Elevate to an administrator shell if not Already
# ============================
function Self_Elevate
{
	$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $IsAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	if (-not $IsAdmin) {
		$psi = New-Object System.Diagnostics.ProcessStartInfo
		$psi.FileName = (Get-Command powershell.exe).Source
		$psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
		$psi.Verb = "runas"
		try { [Diagnostics.Process]::Start($psi) | Out-Null } catch { exit 1 }
		exit
	}
}

# ============================
# Install WCD from Windows Store
# ============================
function Install_WCD
{
	$storeUri = "ms-windows-store://pdp/?ProductId=9NBLGGH4TX22"
	$webUrl   = "https://www.microsoft.com/store/apps/9NBLGGH4TX22"

	[System.Windows.Forms.MessageBox]::Show("This script will now open the Windows Store, please install the Windows Configuration Designer (WCD) to use this tool.")
	# Try to open the Store app page, fall back to web if that fails
	try 
	{
		Write-Host "Attempting to open Windows Store to Windows Configuration Designer page"
		Start-Process -FilePath $storeUri -ErrorAction Stop
	} 
    catch 
    {
		Write-Warning "Could not open the Microsoft Store. Opening the web page instead."
		Start-Process $webUrl
	}
	
	Write-Host "Waiting for user to press OK on dialog"
	[System.Windows.Forms.MessageBox]::Show("Press OK on this dialog once WCD is installed.")
}

# ============================
# Find and copy Microsoft-Desktop-Provisioning.dat from Windows Configuration Designer (WCD)
# which is a required file for creating the PPKG.
# ============================
function Copy_WCD_DAT_File 
{
    $microsoftDatFile = "Microsoft-Desktop-Provisioning.dat"
    $destPath    = Join-Path -Path $PSScriptRoot -ChildPath "$microsoftDatFile"

    if (Test-Path $destPath) 
	{
        Write-Host "Already present: $destPath" -ForegroundColor Green
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

    foreach ($base in $commonLocations) 
	{
        $cand = Get-ChildItem -Path $base -Recurse -Filter $microsoftDatFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cand) 
		{
            Copy-Item -Path $cand.FullName -Destination $destPath -Force
            Write-Host "Copied '$($cand.FullName)' -> '$destPath'" -ForegroundColor Cyan
            return $destPath
        }
    }

    foreach ($root in @("$env:ProgramFiles","$env:ProgramFiles(x86)","$env:ProgramData","$env:LOCALAPPDATA")) 
	{
        if (-not (Test-Path $root)) { continue }
        $cand = Get-ChildItem -Path $root -Recurse -Filter $microsoftDatFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cand) 
		{
            Copy-Item -Path $cand.FullName -Destination $destPath -Force
            Write-Host "Copied '$($cand.FullName)' -> '$destPath'" -ForegroundColor Cyan
            return $destPath
        }
    }

    throw "Unable to locate '$microsoftDatFile' in common WCD locations. Exiting..."
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
        New-Item -Path $copyPath -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path -Path $xmlPath)) 
	{
        throw "Customizations.xml not found at $xmlPath"
    }

    # Prefer direct invocation (no string building)
    icd.exe /Build-ProvisioningPackage /CustomizationXML:"$xmlPath" /PackagePath:"$ppkgFile" /StoreFile:"$datFile" +Overwrite
}

# Main entry point
function Main
{
	Add-Type -AssemblyName System.Windows.Forms
	
	Self_Elevate
	Install_WCD
	Copy_WCD_DAT_File
	Create_PPKG
	
	
	Write-Host "Press enter to exit..." -NoNewline 
	Read-Host
}

Main