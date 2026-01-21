<#
.SYNOPSIS
    Build the CANoe Logging Tool executable with PyInstaller.

.DESCRIPTION
    Wrapper around the PyInstaller invocation that is typically entered
    manually. Ensures the command is executed from the repository root,
    points PyInstaller to the src folder, and collects the required
    win32com/customtkinter assets.

.PARAMETER PyInstaller
    Optional override for the pyinstaller executable. Defaults to
    the instance on PATH; you can pass the full path if needed, e.g.
    `.\build_exe.ps1 -PyInstaller .\venv\Scripts\pyinstaller.exe`.
#>
param(
    [string]$PyInstaller = "pyinstaller",
    [string]$Python = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $repoRoot

$srcPath = Join-Path $repoRoot "src"
$distPath = Join-Path $repoRoot "dist"
$buildPath = Join-Path $repoRoot "build"
$exeName = "anSWer Logging Hub.exe"
$finalExePath = Join-Path $srcPath $exeName
$requirementsPath = Join-Path $repoRoot "requirements.txt"
$venvDir = Join-Path $repoRoot ".venv"
if (-not (Test-Path $venvDir)) {
    $legacyVenv = Join-Path $repoRoot "venv"
    if (Test-Path $legacyVenv) {
        $venvDir = $legacyVenv
    }
}
$venvPython = Join-Path $venvDir "Scripts\python.exe"

$arguments = @(
    "--name", "anSWer Logging Hub",
    "--onefile",
    "--windowed",
    "--icon", (Join-Path $repoRoot "src\ico\CANoe_Logging.ico"),
    "--paths", $srcPath,
    "--collect-submodules", "win32com",
    "--hidden-import=win32com.client",
    "--hidden-import=pythoncom",
    "--collect-data", "customtkinter",
    "src/app.py"
)
$pyInstallerCommand = $null
$pyInstallerArgs = $arguments

if ($PyInstaller -eq "pyinstaller") {
    if (-not (Test-Path $venvPython)) {
        $pyLauncher = $null
        if ($Python) {
            $pyLauncher = $Python
        } elseif (Get-Command py -ErrorAction SilentlyContinue) {
            $pyLauncher = "py -3"
        } elseif (Get-Command python -ErrorAction SilentlyContinue) {
            $pyLauncher = "python"
        }
        if (-not $pyLauncher) {
            throw "Python not found. Install Python 3 or pass -Python with a full path."
        }

        Write-Host "Creating venv at $venvDir..."
        & $pyLauncher -m venv $venvDir
    }

    if (-not (Test-Path $venvPython)) {
        throw "Virtual environment python not found at $venvPython"
    }

    if (Test-Path $requirementsPath) {
        Write-Host "Installing dependencies from requirements.txt..."
        & $venvPython -m pip install --upgrade pip
        & $venvPython -m pip install -r $requirementsPath
    }

    $pyInstallerCommand = $venvPython
    $pyInstallerArgs = @("-m", "PyInstaller") + $arguments
} else {
    $pyInstallerCommand = $PyInstaller
}

Write-Host "Building anSWer Logging Hub executable..."
Write-Host "$pyInstallerCommand $($pyInstallerArgs -join ' ')"

& $pyInstallerCommand @pyInstallerArgs
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller exited with code $LASTEXITCODE"
}

$sourceExePath = Join-Path $distPath $exeName
if (-not (Test-Path $sourceExePath)) {
    throw "Expected executable not found at $sourceExePath"
}

Write-Host "Moving $exeName to src folder..."
Move-Item -LiteralPath $sourceExePath -Destination $finalExePath -Force

foreach ($path in @($buildPath, $distPath)) {
    if (Test-Path $path) {
        Write-Host "Removing $path ..."
        Remove-Item -LiteralPath $path -Recurse -Force
    }
}

Write-Host "Executable ready at $finalExePath"
