$TestDocumentFilename = "HugeDocument.xlsx"
$OutputDirectory = "HugeDocument"
$TestDocumentPath = Join-Path $PSScriptRoot $TestDocumentFilename
$OutputDirectoryPath = Join-Path $PSScriptRoot $OutputDirectory

if (!(Test-Path $TestDocumentPath -PathType Leaf)) {
    Write-Error "$TestDocumentFilename does not exist"
    exit 1
}

if (Test-Path $OutputDirectoryPath) {
    Remove-Item -Force -Recurse "$OutputDirectoryPath\**"
}

$TempArchiveFilename = $TestDocumentFilename -replace "\.xlsx$",".zip"

# Copy the .xlsx to a .zip
Copy-Item $TestDocumentFilename -Destination $TempArchiveFilename -Force

# Extract it
Expand-Archive $TempArchiveFilename -DestinationPath $OutputDirectoryPath -Force

# Delete the .zip file
Remove-Item $TempArchiveFilename -Force