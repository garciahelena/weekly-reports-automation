$ErrorActionPreference = "Stop"

# Directories
$SourceDir = "C:\Users\helen\Desktop\Reports"
$DestDir = "C:\Users\helen\Desktop"

# Get latest Excel file
$file = Get-ChildItem $SourceDir -Filter *.xlsx |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

# Get today's date
$today = Get-Date

# Calculate most recent Sunday
$daysUntilSunday = (7 - [int]$today.DayOfWeek) % 7

$nextSunday = $today.AddDays($daysUntilSunday)

# Format date
$weekEnd = $nextSunday.ToString("dd.MM.yy")

# New file name
$newName = "KEV EON Reports WE $weekEnd.xlsx"

# Destination path
$destPath = Join-Path $DestDir $newName

# Safety check
if ($file.FullName -eq $destPath) {
    Write-Host "Source and destination are the same. Skipping copy."
    exit
}

# Copy file
Copy-Item $file.FullName -Destination $destPath

# Call Python script
python edit_excel.py "$destPath" "$weekEnd"

Write-Host "Report created and updated: $newName"