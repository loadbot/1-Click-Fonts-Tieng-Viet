# Bypass execution policy for the current session
Set-ExecutionPolicy Bypass -Scope Process -Force
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted

# Prompt the user for confirmation
$confirmation = Read-Host "This script will install fonts for all users. Proceed? (Y/N)"
if ($confirmation -ne 'Y') {
    Write-Host "Fonts installation aborted!"
    exit
}

# URL of the fonts zip file
$FontsZipURL = "https://github.com/loadbot/1-Click-Fonts-Tieng-Viet/releases/download/Font/Fonts-Tieng-Viet.zip"

# Destination directory to extract fonts
$TempExtractPath = "$env:TEMP\Fonts"

# Check if the fonts zip file already exists in the TEMP directory
$FontsZipPath = "$env:TEMP\Fonts.zip"
if (-not (Test-Path $FontsZipPath)) {
    # Download the fonts zip file if it doesn't exist
    Invoke-WebRequest -Uri $FontsZipURL -OutFile $FontsZipPath
}

# Extract fonts from the zip file to a temporary directory
Expand-Archive -Path $FontsZipPath -DestinationPath $TempExtractPath -Force

# Set the path to the extracted fonts directory
$Path = $TempExtractPath

# Get the directory of the script
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Set the full path to the Install-Font.ps1 script
$InstallFontScriptPath = Join-Path -Path $ScriptDirectory -ChildPath "Install-Font.ps1"

# Dot-source the Install-Font.ps1 script
. $InstallFontScriptPath -Path $Path -Scope System -Method Manual -UninstallExisting

# Define the font installation function
Function Install-Fonts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('FullName')]
        [String]$Path,
        [String]$Scope = 'System'
    )
    
    # Add the font installation logic here
    # Use the provided Install-Font.ps1 script to install fonts
    
    # Call the Install-Font.ps1 script with appropriate parameters
    . $InstallFontScriptPath -Path $Path -Scope System -Method Manual -UninstallExisting
}

# Loop through each font file and install it for all users
$FontsToInstall = Get-ChildItem $TempExtractPath -Include '*.ttf','*.ttc','*.otf'
foreach ($Font in $FontsToInstall) {
    Install-Fonts -Path $Font.FullName
}

# Log file path
$LogFilePath = Join-Path -Path $ScriptDirectory -ChildPath "fonts.log"
Write-Output "Fonts installation completed!"

# Check if the log file exists before opening it
if (Test-Path $LogFilePath) {
    Start-Process -FilePath $LogFilePath
} else {
    Write-Host ""
}
