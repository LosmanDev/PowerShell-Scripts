# Define the source and destination paths
$sourcePath = "$env:APPDATA\Microsoft\Signatures"
$destinationPath = "$env:USERPROFILE\"

# Check if the destination folder exists and create it if not
if (!(Test-Path -Path $destinationPath)) {
    New-Item -Path $destinationPath -ItemType Directory
}

# Copy files from source to destination
Copy-Item -Path $sourcePath\* -Destination $destinationPath -Recurse -Force

Write-Host "Files copied to $destinationPath"



########################## Reverese Script ###################################################

# Define the original destination and source paths, now reversed
$sourcePath = "$env:USERPROFILE\"
$destinationPath = "$env:APPDATA\Microsoft\Signatures"

# Check if the destination folder exists and create it if not
if (!(Test-Path -Path $destinationPath)) {
    New-Item -Path $destinationPath -ItemType Directory
}

# Copy files from original destination back to source
Move-Item -Path $sourcePath\* -Destination $destinationPath -Recurse -Force

Write-Host "Files copied back to $destinationPath"
