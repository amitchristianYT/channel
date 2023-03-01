# 
# Check out YouTube Channel for Video Tutorial
# https://www.youtube.com/amitchristian
# Please consider to Like or Subscribe
#
# Prerequisite
# Enable Powershell Execution Policy in Powershell Administrator Shell
#   set-executionpolicy remotesigned
#
# 1. Specify Input Folder Path
# 2. Specify Output Folder Path
#
###################################################################################

# Get the files which should be moved, without folders
$files = Get-ChildItem 'Insert Your Input Folder Path' -Recurse | where {!$_.PsIsContainer}
 
# List Files which will be moved
$files
 
# Target Filder where files should be moved to. The script will automatically create a folder for the year and month.
$targetPath = 'Insert Your Output Folder Path'
 
foreach ($file in $files)
{
# Get year and Month of the file
# I used LastWriteTime since this are synced files and the creation day will be the date when it was synced
$year = $file.LastWriteTime.Year.ToString()
$month = $file.LastWriteTime.Month.ToString()
 
# Out FileName, year and month
$file.Name
$year
$month
 
# Set Directory Path
$Directory = $targetPath + "\" + $year + "\" + $month
# Create directory if it doesn't exist
if (!(Test-Path $Directory))
{
New-Item $directory -type directory
}
 
# Move File to new location
$file | Move-Item -Destination $Directory
}
