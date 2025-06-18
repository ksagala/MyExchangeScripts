# Requires -Modules: ImportExcel, PSWriteWord

# Parameters
$SourceFolder = "C:\Path\To\Your\Folder"
$OutputDoc = "C:\Path\To\Output\ExchangeDocs.docx"
$DocumentTitle = "Exchange Documentation"

# Load PSWriteWord module
Import-Module PSWriteWord

# Create new Word document
$doc = New-WordDocument -FilePath $OutputDoc

# Title Page
Add-WordText -WordDocument $doc -Text $DocumentTitle -Size 36 -Bold -Alignment Center
Add-WordParagraph -WordDocument $doc
Add-WordText -WordDocument $doc -Text "Generated on: $(Get-Date -Format 'yyyy-MM-dd')" -Size 14 -Alignment Center
Add-WordParagraph -WordDocument $doc
Add-WordPageBreak -WordDocument $doc

# Table of Contents
Add-WordTOC -WordDocument $doc -Title "Table of Contents"
Add-WordPageBreak -WordDocument $doc

# Get files from source folder
$files = Get-ChildItem -Path $SourceFolder -File | Sort-Object Name

foreach ($file in $files) {
    # Chapter Title
    Add-WordText -WordDocument $doc -Text $file.BaseName -Size 24 -Bold -Style Heading1
    Add-WordParagraph -WordDocument $doc

    # Add file content
    $content = Get-Content $file.FullName -Raw
    Add-WordText -WordDocument $doc -Text $content -Size 12 -Style Normal
    Add-WordPageBreak -WordDocument $doc
}

# Save document
Close-WordDocument -WordDocument $doc

Write-Host "Document created at $OutputDoc"