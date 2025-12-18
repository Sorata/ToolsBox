# Set ovftool path
$ovfToolPath = "D:\Program Files\ovftool\ovftool.exe"

# Check if ovftool exists
if (-not (Test-Path $ovfToolPath)) {
    Write-Error "Error: ovftool not found at: $ovfToolPath"
    exit 1
}

# Set root dir
$rootDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($rootDir)) { $rootDir = Get-Location }

# Set export root dir
$exportRootDir = Join-Path $rootDir "_OVF_Exports"

# Create export dir
if (-not (Test-Path $exportRootDir)) {
    New-Item -ItemType Directory -Path $exportRootDir | Out-Null
    Write-Host "Creating export dir: $exportRootDir" -ForegroundColor Cyan
}

# Find all .vmx files, exclude _OVF_Exports, strict extension check
$vmxFiles = @(Get-ChildItem -Path $rootDir -Filter "*.vmx" -Recurse | Where-Object { $_.Extension -eq ".vmx" -and $_.FullName -notlike "*_OVF_Exports*" })

if ($vmxFiles.Count -eq 0) {
    Write-Warning "No .vmx files found."
    exit
}

Write-Host "Found $($vmxFiles.Count) VMs. Starting conversion..." -ForegroundColor Cyan

foreach ($vmx in $vmxFiles) {
    $vmName = $vmx.BaseName
    
    # Output dir
    $outputDir = Join-Path $exportRootDir $vmName
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    # Output OVF path
    $outputOvf = Join-Path $outputDir "$vmName.ovf"

    Write-Host "`n----------------------------------------"
    Write-Host "Converting: $($vmx.Name)"
    Write-Host "Source: $($vmx.FullName)"
    Write-Host "Target: $outputOvf"
    
    # ovftool args - manually quote paths to handle spaces
    $sourcePath = '"{0}"' -f $vmx.FullName
    $targetPath = '"{0}"' -f $outputOvf
    
    # Execute
    try {
        # Note: When using Start-Process with ArgumentList as array, it handles quoting, 
        # but sometimes manual quoting is needed if the receiver is picky.
        # However, adding quotes manually AND passing as array might double quote.
        # Let's try passing arguments as a single string to be safe and explicit.
        $argString = "$sourcePath $targetPath"
        
        $process = Start-Process -FilePath $ovfToolPath -ArgumentList $argString -NoNewWindow -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Success: $vmName" -ForegroundColor Green
        } else {
            Write-Host "Failed: $vmName (Exit Code: $($process.ExitCode))" -ForegroundColor Red
        }
    } catch {
        Write-Error "Exception executing ovftool: $_"
    }
}

Write-Host "`n----------------------------------------"
Write-Host "All tasks completed." -ForegroundColor Cyan
Write-Host "Export directory: $exportRootDir"
