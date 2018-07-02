Function SaveImagesFromDirectory
{
	 Param
    (
        [String]$OriginFolder =$(throw "Parameter missing: -OriginFolder OriginFolder"),
        [String]$DestinationFolder =$(throw "Parameter missing: -DestinationFolder DestinationFolder")
    )
	Add-Type -Assembly “system.io.compression.filesystem”
	
	$array = Get-ChildItem -Path $OriginFolder -Recurse -Include *.docx

	foreach($element in $array) {
		#"$DestinationFolder$($element.name).jpeg"
		#Continue

		#copy and renaming the doc to .zip file
		$ZipFile = $env:temp + "\" + [guid]::NewGuid().ToString() + ".zip"
		Copy-Item -Path "$element" -Destination $ZipFile
		
		#extract all file to tmp folder
		$TmpFolder = $env:temp + "\" + [guid]::NewGuid().ToString()
		[io.compression.zipfile]::ExtractToDirectory($ZipFile, $TmpFolder)

		#copy image files from media folder to destination
		Get-ChildItem "$TmpFolder\word\media\" -recurse | Copy-Item -destination "$DestinationFolder$($element.name).jpeg"

	}
}

Function ExtractImages {

Param
(
	[String]$rtfpath =$(throw "Parameter missing: -rtfpath rtfpath"),
	[String]$DestinationFolder =$(throw "Parameter missing: -DestinationFolder DestinationFolder")
)

$tempDestinationFolder = "C:\Temp\DOCX"
$srcfiles = Get-ChildItem -Path $rtfpath -Recurse -Include *.doc, *.rtf | where { $_.Length -gt 60KB}

# SaveFormat to be .DOCX
$saveFormat =[Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault");

$word = new-object -comobject word.application
$word.Visible = $False

ForEach ($rtf in $srcfiles)
     {
	Write-Host "Processing:" $rtf.FullName
	
  	$opendoc = $word.documents.open($rtf.FullName);
  	$filename = $rtf.basename + ".docx"
  	$opendoc.saveas("$tempDestinationFolder\$filename", $saveFormat);
  	$opendoc.close();	
	SaveImagesFromDirectory -OriginFolder "$tempDestinationFolder" -DestinationFolder $DestinationFolder
	Remove-Item -path "$tempDestinationFolder\$filename"
     }
$word.quit();
}

ExtractImages -rtfpath 'Memo Tags 2' -DestinationFolder 'C:/Temp/images/'