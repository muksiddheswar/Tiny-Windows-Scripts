# Unzip multiple zip files in current location.

$zipFiles = Get-ChildItem . -Filter *.zip

foreach ($zipFile in $zipFiles) {

	$zipOutPutFolderExtended = $zipFile.BaseName    
	Expand-Archive -Path $zipFile.FullName -DestinationPath $zipOutPutFolderExtended 

    }

