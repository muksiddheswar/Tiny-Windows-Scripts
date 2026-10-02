
# Define the two locations
$location2 = "C:\Title\work\docs\CENELEC\tp"
$location1 = "C:\Title\work\docs\CENELEC\tp\IS"

# Get .doc files from both locations (file names only)
$files1 = Get-ChildItem -Path $location1 -Filter *.doc -File | Select-Object -ExpandProperty Name
$files2 = Get-ChildItem -Path $location2 -Filter *.doc -File | Select-Object -ExpandProperty Name

# Compare and find files in location1 not in location2
$uniqueFiles = Compare-Object -ReferenceObject $files1 -DifferenceObject $files2 -PassThru | Where-Object { $_ -in $files1 }

# Output the result
if ($uniqueFiles) {
    Write-Host "Files present in Location 1 but not in Location 2:"
    $uniqueFiles
} else {
    Write-Host "All .doc files in Location 1 are also present in Location 2."
}
