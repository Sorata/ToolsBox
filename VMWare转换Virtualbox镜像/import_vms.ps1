# Set VBoxManage path
$vboxManagePath = "D:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$logFile = Join-Path $PSScriptRoot "import_log.txt"

function Log-Message {
    param([string]$Message, [ConsoleColor]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

# Clear log
if (Test-Path $logFile) { Remove-Item $logFile }

# Check VBoxManage
if (-not (Test-Path $vboxManagePath)) {
    Log-Message "Error: VBoxManage not found at: $vboxManagePath" "Red"
    exit 1
}

$rootDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($rootDir)) { $rootDir = Get-Location }
$ovfSourceDir = Join-Path $rootDir "_OVF_Exports"

if (-not (Test-Path $ovfSourceDir)) {
    Log-Message "Error: OVF export directory not found: $ovfSourceDir" "Red"
    exit 1
}

$ovfFiles = @(Get-ChildItem -Path $ovfSourceDir -Filter "*.ovf" -Recurse)

if ($ovfFiles.Count -eq 0) {
    Log-Message "No .ovf files found in $ovfSourceDir" "Yellow"
    exit
}

Log-Message "Found $($ovfFiles.Count) OVF templates. Starting import..." "Cyan"

foreach ($ovf in $ovfFiles) {
    Log-Message "`n----------------------------------------"
    Log-Message "Importing: $($ovf.Name)"
    
    # Use explicit VM name to avoid conflicts
    $vmName = $ovf.BaseName
    $ovfPath = '"{0}"' -f $ovf.FullName
    $argString = "import $ovfPath --vsys 0 --vmname ""$vmName"""
    
    Log-Message "Command: VBoxManage $argString"
    
    try {
        $process = Start-Process -FilePath $vboxManagePath -ArgumentList $argString -NoNewWindow -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Log-Message "Success: $vmName" "Green"
        } else {
            Log-Message "Failed: $vmName (Exit Code: $($process.ExitCode))" "Red"
        }
    } catch {
        Log-Message "Exception: $_" "Red"
    }
}

Log-Message "`n----------------------------------------"
Log-Message "All import tasks completed." "Cyan"
