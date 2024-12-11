# Define the paths for the installers directories
$installersPath = "C:\Installers"
$pythonWorkshopPath = "C:\PythonWorkshop"
$wrhWorkshopPath = "C:\WRHWorkshop"

# Create the directories if they don't exist
if (-not (Test-Path $installersPath)) {
    New-Item -Path $installersPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $installersPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $installersPath"
}

if (-not (Test-Path $pythonWorkshopPath)) {
    New-Item -Path $pythonWorkshopPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $pythonWorkshopPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $pythonWorkshopPath"
}


if (-not (Test-Path $wrhWorkshopPath)) {
    New-Item -Path $wrhWorkshopPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $wrhWorkshopPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $wrhWorkshopPath"
}

# Define the download URLs
$sevenZipUrl = "https://www.7-zip.org/a/7z2408-x64.exe"
$pythonUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"

# Define the local installer paths
$InstallerPath1 = "$installersPath\7z2408-x64.exe"
$InstallerPath5 = "$installersPath\python-3.11.4-amd64.exe"

# Function to download files using Invoke-WebRequest in the background (parallel download)
function Start-Download {
    param (
        [string]$url,
        [string]$outputPath,
        [string]$name
    )

    Write-Host -ForegroundColor Yellow "Starting download for $name..."

    # Run Invoke-WebRequest in a background job
    Start-Job -ScriptBlock {
        param ($url, $outputPath, $name)
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
            if (Test-Path $outputPath) {
                Write-Host -ForegroundColor Green "$name downloaded successfully."
            } else {
                Write-Host -ForegroundColor Red "$name download failed."
            }
        } catch {
            Write-Host -ForegroundColor Red "$name download failed with error: $_"
        }
    } -ArgumentList $url, $outputPath, $name
}

# Download installers with verification (in parallel)
Start-Download -url $sevenZipUrl -outputPath $InstallerPath1 -name "7z2408-x64.exe"
Start-Download -url $pythonUrl -outputPath $InstallerPath5 -name "python-3.11.4-amd64.exe"

# Wait for all background jobs to finish
Get-Job | Wait-Job

# Removing completed jobs
Get-Job | Remove-Job

# Installing 7-Zip
Write-host -f Yellow "Installing 7-Zip..."
Start-Process -FilePath $InstallerPath1 -Args "/S" -Verb RunAs -Wait

# Installing Python with Admin privileges and specific settings
Write-host -f Yellow "Installing Python..."
Start-Process -FilePath $InstallerPath5 -ArgumentList `
"/quiet InstallAllUsers=1 PrependPath=1 Include_doc=1 Include_pip=1 Include_tcltk=1 Include_test=1 Include_launcher=1 InstallLauncherAllUsers=1 InstallDir=""C:\Program Files\Python311"" Include_symbols=1 Include_debug=1 AssociateFiles=1 Shortcuts=1" -Wait

Write-Host -ForegroundColor Green "All installations and setups completed successfully."

# Define the sequence of commands
$commands = @"
SET PATH=%PATH%;C:\Program Files\Python311\Scripts;C:\Program Files\Python311\
cd C:\PythonWorkshop
python -m venv .venv
start "Environment setup" cmd /k "C:\PythonWorkshop\.venv\Scripts\activate.bat && pip install pandas==2.1.2 && pip install xlrd==2.0.1 && pip install openpyxl==3.1.2 && pip install matplotlib==3.8.1 && pip install pyarrow==13.0.0 && pip install jupyter==1.0.0"
"@

# Write the commands to a temporary batch file
$tempBatchFile = "$env:TEMP\run_python_setup.bat"
Set-Content -Path $tempBatchFile -Value $commands

# Run the batch file using cmd.exe in sequence
Start-Process -FilePath "cmd.exe" -ArgumentList "/k", $tempBatchFile -Wait

# Clean up - Remove the temporary batch file after execution
Remove-Item -Path $tempBatchFile


# -------------------------------------------
# 1. Create the directory for warehouse files
# -------------------------------------------
$directoryPath = "C:\PythonWorkshop\wsbpythonfiles"
if (!(Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath
    Write-Host "Created directory at $directoryPath."
} else {
    Write-Host "Directory $directoryPath already exists."
}

# -------------------------------------------
# 2. Download files from GitHub into the directory using curl
# -------------------------------------------
$filesToDownload = @(
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.001",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.002",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.003",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.004",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.005",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.006",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.007",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.008",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.009",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.010",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.011",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - zadania kontekst danych.ipynb",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - wprowadzenie.pdf",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - wprowadzenie - Jupyter.ipynb",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - sredni - kontekst danych.ipynb",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy.pdf",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy zadania.pdf",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy zadania - Jupyter.ipynb",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy - Jupyter.ipynb",
    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Zajecia.7z"
)

foreach ($url in $filesToDownload) {
    $fileName = [System.IO.Path]::GetFileName($url)
    $destination = Join-Path $directoryPath $fileName
    Write-Host "Downloading $fileName to $destination using curl..."
    # Using curl to download the file
    curl -o $destination $url
}

# -------------------------------------------
# 3. Extract files using 7z
# -------------------------------------------
# Path to the 7z executable (assumes 7z is in PATH)
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Extract ZadaniaGitHub.7z.001 and related parts
$zadaniaFilePath = Join-Path $directoryPath "ZadaniaGitHub.7z.001"
if (Test-Path -Path $zadaniaFilePath) {
    Write-Host "Extracting ZadaniaGitHub.7z.001 and related parts..."
    & $sevenZipPath x $zadaniaFilePath -o"$directoryPath\Zadania" -y
    Write-Host "Extraction of ZadaniaGitHub completed."
} else {
    Write-Host "File ZadaniaGitHub.7z.001 not found. Skipping extraction."
}

# Extract Zajecia.7z
$zajeciaFilePath = Join-Path $directoryPath "Zajecia.7z"
if (Test-Path -Path $zajeciaFilePath) {
    Write-Host "Extracting Zajecia.7z..."
    & $sevenZipPath x $zajeciaFilePath -o"$directoryPath\Zajecia" -y
    Write-Host "Extraction of Zajecia completed."
} else {
    Write-Host "File Zajecia.7z not found. Skipping extraction."
}
