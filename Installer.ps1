<#
.SYNOPSIS
    Microsoft Visual C++ All-In-One Runtimes Installer
.DESCRIPTION
    Downloads and installs all Visual C++ Redistributable packages from 2005 to 2022
.NOTES
    Automatically requests Administrator privileges
    Usage: irm "YOUR_GITHUB_RAW_URL/Install_all_VC_Runtimes.ps1" | iex
#>

# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host ""
    Write-Host "[!] Requesting Administrator privileges..." -ForegroundColor Yellow
    Write-Host ""
    
    # Get the script path
    if ($MyInvocation.MyCommand.Path) {
        # Script is running from a file
        $ScriptPath = $MyInvocation.MyCommand.Path
        Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" -Verb RunAs
    } else {
        # Script is running from pipeline (irm | iex)
        $ScriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
        $TempScript = Join-Path $env:TEMP "VC_Install_$(Get-Date -Format 'yyyyMMddHHmmss').ps1"
        $ScriptContent | Out-File -FilePath $TempScript -Encoding UTF8 -Force
        Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempScript`"" -Verb RunAs
    }
    exit
}

# Set console title and colors
$Host.UI.RawUI.WindowTitle = "Microsoft Visual C++ Runtimes Installer"
Clear-Host

# GitHub release URL
$GitHubRelease = "https://github.com/ZodiacTeamOS/VC-redist-C-ALL-IN-ONE/releases/download/x64_x86"

# Define packages
$Packages = @(
    @{Name = "Visual C++ 2005 x86"; File = "1-vcredist_x86_C++2005.exe"; Args = "/q"; Arch = "x86"},
    @{Name = "Visual C++ 2005 x64"; File = "1-vcredist_x64_C++2005.exe"; Args = "/q"; Arch = "x64"},
    @{Name = "Visual C++ 2008 x86"; File = "2-vcredist_x86_C++2008.exe"; Args = "/qb"; Arch = "x86"},
    @{Name = "Visual C++ 2008 x64"; File = "2-vcredist_x64_C++2008.exe"; Args = "/qb"; Arch = "x64"},
    @{Name = "Visual C++ 2010 x86"; File = "3-vcredist_x86_C++2010.exe"; Args = "/passive /norestart"; Arch = "x86"},
    @{Name = "Visual C++ 2010 x64"; File = "3-vcredist_x64_C++2010.exe"; Args = "/passive /norestart"; Arch = "x64"},
    @{Name = "Visual C++ 2012 x86"; File = "4-vcredist_x86_C++2012.exe"; Args = "/passive /norestart"; Arch = "x86"},
    @{Name = "Visual C++ 2012 x64"; File = "4-vcredist_x64_C++2012.exe"; Args = "/passive /norestart"; Arch = "x64"},
    @{Name = "Visual C++ 2013 x86"; File = "5-vcredist_x86_C++2013.exe"; Args = "/passive /norestart"; Arch = "x86"},
    @{Name = "Visual C++ 2013 x64"; File = "5-vcredist_x64_C++2013.exe"; Args = "/passive /norestart"; Arch = "x64"},
    @{Name = "Visual C++ 2015-2022 x86"; File = "6-VC_redist.x86_C++2015-2022.exe"; Args = "/passive /norestart"; Arch = "x86"},
    @{Name = "Visual C++ 2015-2022 x64"; File = "6-VC_redist.x64_C++2015-2022.exe"; Args = "/passive /norestart"; Arch = "x64"}
)

function Show-Header {
    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "   Microsoft Visual C++ All-In-One Runtimes Installer" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host ""
}



Show-Header

# Determine system architecture
$Is64Bit = [Environment]::Is64BitOperatingSystem

if ($Is64Bit) {
    Write-Host "System Architecture: 64-bit (x64)" -ForegroundColor Green
    $PackagesToInstall = $Packages
} else {
    Write-Host "System Architecture: 32-bit (x86)" -ForegroundColor Green
    $PackagesToInstall = $Packages | Where-Object { $_.Arch -eq "x86" }
}

Write-Host ""
Write-Host "Total packages to install: $($PackagesToInstall.Count)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Starting installation..." -ForegroundColor Yellow
Write-Host ""
Start-Sleep -Seconds 1

# Create temp directory
$TempDir = Join-Path $env:TEMP "VC_Runtimes_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $TempDir -ItemType Directory -Force | Out-Null

$TotalSteps = $PackagesToInstall.Count
$CurrentStep = 0
$SuccessCount = 0
$FailCount = 0

# Initial progress bar position
$ProgressLine = $Host.UI.RawUI.CursorPosition.Y

foreach ($Package in $PackagesToInstall) {
    $CurrentStep++
    
    # Update progress bar in place
    $Percent = [math]::Round(($CurrentStep / $TotalSteps) * 100)
    $FilledChars = [math]::Floor($Percent / 5)
    $EmptyChars = 20 - $FilledChars
    $ProgressBar = "█" * $FilledChars + "░" * $EmptyChars
    
    # Move cursor to progress bar line
    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates(0, $ProgressLine)
    Write-Host "[$ProgressBar] $Percent% - Installing packages...                    " -ForegroundColor Cyan
    
    $DownloadUrl = "$GitHubRelease/$($Package.File)"
    $LocalFile = Join-Path $TempDir $Package.File
    
    try {
        # Download file silently
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $LocalFile -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'
        
        # Install package silently
        $ProcessArgs = @{
            FilePath = $LocalFile
            ArgumentList = $Package.Args
            Wait = $true
            PassThru = $true
            NoNewWindow = $true
        }
        $Process = Start-Process @ProcessArgs
        
        # Check exit code and show result below progress bar
        if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
            Write-Host ""
            Write-Host "[✓] $($Package.Name) installed successfully" -ForegroundColor Green
            $SuccessCount++
        } elseif ($Process.ExitCode -eq 1638) {
            Write-Host ""
            Write-Host "[!] $($Package.Name) already installed (newer version)" -ForegroundColor Yellow
            $SuccessCount++
        } else {
            Write-Host ""
            Write-Host "[!] $($Package.Name) - Exit code: $($Process.ExitCode)" -ForegroundColor Yellow
            $FailCount++
        }
        
        # Clean up
        Remove-Item $LocalFile -Force -ErrorAction SilentlyContinue
        
    } catch {
        Write-Host ""
        Write-Host "[✗] $($Package.Name) - Failed: $($_.Exception.Message)" -ForegroundColor Red
        $FailCount++
    }
}

# Clean up temp directory
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

# Final summary
Clear-Host
Write-Host ""
Write-Host "===============================================================" -ForegroundColor Green
Write-Host "                  Installation Completed!" -ForegroundColor Green
Write-Host "===============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  ✓ Successful: $SuccessCount" -ForegroundColor Green
if ($FailCount -gt 0) {
    Write-Host "  ✗ Failed: $FailCount" -ForegroundColor Red
}
Write-Host ""
Write-Host "All Visual C++ Runtime packages have been processed." -ForegroundColor White
Write-Host ""

if ($FailCount -gt 0) {
    Write-Host "Note: Some packages failed to install. This might be normal if:" -ForegroundColor Yellow
    Write-Host "  - Newer versions are already installed" -ForegroundColor Yellow
    Write-Host "  - Windows Update handles these packages" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Press any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
