param (
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
   # We are running "as Administrator" - so change the title and background color to indicate this
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   $Host.UI.RawUI.BackgroundColor = "DarkBlue"
   clear-host
   }
else
   {
   # We are not running "as Administrator" - so relaunch as administrator

   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";

   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;

   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";

   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);

   # Exit from the current, unelevated, process
   exit
}

# This script will manually rip out all VMware Tools registry entries and files for Windows 2008-2019
# Tested for 2019, 2016, and probably works on 2012 R2 after the 2016 fixes.
function Get-VMwareToolsInstallerID {
    foreach ($item in $(Get-ChildItem Registry::HKEY_CLASSES_ROOT\Installer\Products)) {
        If ($item.GetValue('ProductName') -eq 'VMware Tools') {
            return @{
                RegistryID = $item.PSChildName;
                MSIId = [Regex]::Match($item.GetValue('ProductIcon'), '(?<={)(.*?)(?=})') | Select-Object -ExpandProperty Value
            }
        }
    }
}

$VMWareToolsIDs = Get-VMwareToolsInstallerID

# File Path Targets
$FileTargets = @()
$FileTargets += "C:\Program Files\VMware" # VMware Tools directory
$FileTargets += "C:\Program Files\Common Files\VMware" # VMware Common Files directory
$FileTargets += "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\VMware" # VMware Start Menu folder

# Registry Targets
$RegistryTargets = @()
$RegistryTargets += "HKLM:\SOFTWARE\VMware, Inc."

if ($VMWareToolsIDs) {

    # Common Registry
    $RegistryTargets += "Registry::HKEY_CLASSES_ROOT\Installer\Features\$($VMWareToolsIDs.RegistryID)"
    $RegistryTargets += "Registry::HKEY_CLASSES_ROOT\Installer\Products\$($VMWareToolsIDs.RegistryID)"
    $RegistryTargets += "HKLM:\SOFTWARE\Classes\Installer\Features\$($VMWareToolsIDs.RegistryID)"
    $RegistryTargets += "HKLM:\SOFTWARE\Classes\Installer\Products\$($VMWareToolsIDs.RegistryID)"
    $RegistryTargets += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$($VMWareToolsIDs.RegistryID)"

    # MSI Registry
    $RegistryTargets += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{$($VMWareToolsIDs.MSIId)}"

}

# This is a bit of a shotgun approach, but if we are at a version less than 2016, add the Uninstaller entries we don't try to automatically determine.
If ([Environment]::OSVersion.Version.Major -lt 10) {
    $RegistryTargets += "HKCR:\CLSID\{D86ADE52-C4D9-4B98-AA0D-9B0C7F1EBBC8}"
    $RegistryTargets += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9709436B-5A41-4946-8BE7-2AA433CAF108}"
    $RegistryTargets += "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}"
}


# Service Targets
$ServiceTargets = @()
$ServiceTargets += Get-Service -DisplayName "VMware*" -ErrorAction SilentlyContinue
$ServiceTargets += Get-Service -DisplayName "GISvc" -ErrorAction SilentlyContinue

# Check paths and registry keys for existence and remove them if they don't exist. Present the user with a list of items to remove.
Write-Host "The following registry keys, files/folders, and services have been found:" -ForegroundColor Yellow


# Create an array to store the targets
$Targets = @()

foreach ($item in $FileTargets) {
    If (Test-Path $item) {

        Write-Host "Path: $item" -ForegroundColor Green

        $Targets += [PSCustomObject]@{
            Type = "Path"
            Path = $item
        }
    }
}

foreach ($item in $RegistryTargets) {
    If (Test-Path $item) {

        Write-Host "Registry: $item" -ForegroundColor Green

        $Targets += [PSCustomObject]@{
            Type = "Registry"
            Path = $item
        }
    }
}

foreach ($item in $ServiceTargets) {
    If ($item) {

        Write-Host "Service: $($item.DisplayName)" -ForegroundColor Green

        $Targets += [PSCustomObject]@{
            Type = "Service"
            Name = $item.Name
        }
    }
}


if ($Targets.Count -eq 0) {
    Write-Host "No VMware Tools registry entries, files/folders, or services found." -ForegroundColor Green
    Exit
}


# Prompt the user for confirmation

if (!$Force.IsPresent) {
    Write-Warning "Would you like to proceed with the removal of the above items?" -WarningAction Inquire
} else {
    Write-Warning "-Force flag detected. Proceeding with removal of the above items."
}


# Remove the services
$services = $targets | Where-Object { $_.Type -eq "Service" }

if ($services.length -gt 0) {
    Write-Host "Removing Services"

    $HasRemoveServiceCmdlet = Get-Command Remove-Service -ErrorAction SilentlyContinue

    foreach ($Service in $Services) {

        Write-Host "Service: $($Service.Name)..." -NoNewline 
        
        # Stop the service via taskkill
        if ($Service.Name -eq "VMTools") {
            Start-Process -FilePath "taskkill" -ArgumentList "/F /FI `"SERVICES eq $($Service.Name)`"" -NoNewWindow -Wait
        } else {
            try {
                Stop-Service -Name $Service.Name -Force -ErrorAction Stop
            } catch {
                Write-Host "Error occurred while stopping service: $_" -ForegroundColor Red
            }
        }

        if ($HasRemoveServiceCmdlet) {
            try {
                Remove-Service -Name $Service.Name -Force -ErrorAction Stop
                Write-Host "Removed Service" -ForegroundColor Green
            }
            catch {
                <#Do this if a terminating exception happens#>
                Write-Host "Error occurred while removing service: $_" -ForegroundColor Red
            }
        }
        else {
            sc.exe DELETE $Service.Name | Out-Null
            Write-Host "Removed Service" -ForegroundColor Green
        }

        

    }
}

# Remove registry entries
$registry = $targets | Where-Object { $_.Type -eq "Registry" }

if ($registry.length -gt 0) {

    foreach ($item in $registry) {

        # Test if the registry key exists
        If (!(Test-Path $item.Path)) {
            Write-Host "Registry $($item.Path) does not exist. Skipping..." -ForegroundColor Yellow
            Continue
        }

        Write-Host "Removing Registry $($item.Path)..." -NoNewline

        try {
            Remove-Item -Path $item.Path -Recurse -Force -ErrorAction Stop 
            Write-Host "Removed Registry" -ForegroundColor Green
        } catch {
            Write-Host "Error occurred while removing registry: $_" -ForegroundColor Red
        }

        Write-Host "Removed Registry"
    }

}

# Remove paths
$paths = $targets | Where-Object { $_.Type -eq "Path" }

# Stop the Windows Event Log service
Write-Host "Stopping Windows Event Log service..." -NoNewline

try {
    Stop-Service -Name "EventLog" -Force -ErrorAction Stop 
    Write-Host "Stopped service" -ForegroundColor Green
} catch {
    Write-Host "Failed to stop service: $_" -ForegroundColor Red
}

foreach ($path in $paths) {

    Write-Host "Removing Path $($path.path)..." -NoNewline

    try {
        Remove-Item -Path $path.path -Recurse -Force -ErrorAction Stop 
        Write-Host "Removed Path" -ForegroundColor Green
    } catch {
        Write-Host "Error occurred while removing path: $_" -ForegroundColor Red
    }

}

Write-Host "Starting Windows Event Log service..." -NoNewline
try {
    Start-Service -Name "EventLog"
    Write-Host "Started service" -ForegroundColor Green
} catch {
    Write-Host "Failed to start service: $_" -ForegroundColor Red
}

Write-Host "It is recommended to restart your computer to ensure all VMware Tools components are removed. Re-run this script to check that all components are removed." -ForegroundColor Yellow