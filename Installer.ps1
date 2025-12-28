#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Microsoft Visual C++ All-In-One Runtimes Installer
.DESCRIPTION
    Downloads and installs all Visual C++ Redistributable packages from 2005 to 2022
.NOTES
    Run as Administrator
    Usage: irm "YOUR_GITHUB_RAW_URL/Install_all_VC_Runtimes.ps1" | iex
#>

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

function Show-ProgressBar {
    param(
        [int]$Current,
        [int]$Total,
        [string]$PackageName
    )
    
    $Percent = [math]::Round(($Current / $Total) * 100)
    $FilledChars = [math]::Floor($Percent / 5)
    $EmptyChars = 20 - $FilledChars
    $ProgressBar = "█" * $FilledChars + "░" * $EmptyChars
    
    Write-Host ""
    Write-Host "[$ProgressBar] $Percent% - $PackageName..." -ForegroundColor Cyan
}

# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host ""
    Write-Host "[!] This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "[*] Please right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    Start-Sleep -Seconds 3
    exit 1
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
Start-Sleep -Seconds 2

# Create temp directory
$TempDir = Join-Path $env:TEMP "VC_Runtimes_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
Write-Host "Temp directory: $TempDir" -ForegroundColor DarkGray

$TotalSteps = $PackagesToInstall.Count
$CurrentStep = 0
$SuccessCount = 0
$FailCount = 0

foreach ($Package in $PackagesToInstall) {
    $CurrentStep++
    
    Show-ProgressBar -Current $CurrentStep -Total $TotalSteps -PackageName $Package.Name
    
    $DownloadUrl = "$GitHubRelease/$($Package.File)"
    $LocalFile = Join-Path $TempDir $Package.File
    
    try {
        # Download file
        Write-Host "  → Downloading..." -ForegroundColor Yellow -NoNewline
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $LocalFile -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'
        Write-Host " Done" -ForegroundColor Green
        
        # Install package
        Write-Host "  → Installing..." -ForegroundColor Yellow -NoNewline
        $ProcessArgs = @{
            FilePath = $LocalFile
            ArgumentList = $Package.Args
            Wait = $true
            PassThru = $true
            NoNewWindow = $true
        }
        $Process = Start-Process @ProcessArgs
        
        # Check exit code
        if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
            Write-Host " Done" -ForegroundColor Green
            Write-Host "  [✓] $($Package.Name) installed successfully" -ForegroundColor Green
            $SuccessCount++
        } elseif ($Process.ExitCode -eq 1638) {
            Write-Host " Already Installed" -ForegroundColor Yellow
            Write-Host "  [!] $($Package.Name) is already installed (newer version)" -ForegroundColor Yellow
            $SuccessCount++
        } else {
            Write-Host " Warning" -ForegroundColor Yellow
            Write-Host "  [!] Exit code: $($Process.ExitCode)" -ForegroundColor Yellow
            $FailCount++
        }
        
        # Clean up
        Remove-Item $LocalFile -Force -ErrorAction SilentlyContinue
        
    } catch {
        Write-Host " Failed" -ForegroundColor Red
        Write-Host "  [✗] Error: $($_.Exception.Message)" -ForegroundColor Red
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
