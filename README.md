# extract-images-from-word-docs

Function to be used in PowerShell

What it does:
1) Searches for .DOC and .RTF files in $filesdirectory with file size < 60KB
2) Temporary saves the file to $tempDestinationFolder = "C:\Temp\DOCX" as a .DOCX
3) Extracs images and saves them to $DestinationFolder
4) Temporary .DOCX file is deleted

Requirements:
MS Word library
