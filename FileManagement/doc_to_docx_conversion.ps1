# Find all .doc files in the current directory and then create .docx copy and a .zip copy

$docPath = [Environment]::CurrentDirectory 
$files = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -ieq ".doc" } | ForEach-Object { $_.Name }

foreach ($file in $files) 
{
	$oldFileName = $docPath + "\" + $file
	$newFileName = [System.IO.Path]::ChangeExtension($oldFileName, ".docx")
	$zipFileName = [System.IO.Path]::ChangeExtension($oldFileName, ".zip")

	if (-Not (Test-Path $newFileName)) 
	{
		Write-Host "Converted: $oldFileName"

		# Create a Word application COM object
		$word = New-Object -ComObject Word.Application
		$word.Visible = $false

		# Open the .doc file
		$document = $word.Documents.Open($oldFileName)

		# Save as .docx (FileFormat = 16 for docx)
		$document.SaveAs([ref] $newFileName, [ref] 16)

		# Close the document and quit Word
		$document.Close()
		$word.Quit()
		
		cp $newFileName $zipFileName
	}
	else
	{
		Write-Host "Skipped: $newFileName already exists."
	}
}

