$folderPath = "C:\DriveH\GitHub\KimaiAutoInput"
$newTime = Get-Date  # Get current system time

# Recursively update all files and folders
Get-ChildItem -Path $folderPath -Recurse | ForEach-Object {
    $_.CreationTime = $newTime
    $_.LastWriteTime = $newTime
    $_.LastAccessTime = $newTime
}

Write-Output "Timestamps updated to the current time for all files and folders in $folderPath"
