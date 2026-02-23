# Define the staging ground
$TargetFolder = "C:\Path\To\Your\Batch\Folder"

# Engage Word in stealth mode (headless)
$Word = New-Object -ComObject Word.Application
$Word.Visible = $False 

# Word's internal code for the PDF format
$wdFormatPDF = 17 

# Gather the raw materials
$Docs = Get-ChildItem -Path $TargetFolder -Filter *.docx

Write-Host "Initiating batch conversion protocol..."

foreach ($Doc in $Docs) {
    try {
        # Construct the exact output path for the new PDF
        $PdfPath = $Doc.FullName.Replace(".docx", ".pdf")
        
        Write-Host "Synthesizing PDF: $($Doc.Name) -> $($Doc.BaseName).pdf"
        
        # Open the document silently
        $OpenDoc = $Word.Documents.Open($Doc.FullName)
        
        # Force the conversion and save the output
        $OpenDoc.SaveAs([ref]$PdfPath, [ref]$wdFormatPDF)
        
        # Close the document, discarding any accidental local changes
        $OpenDoc.Close($False)
        
    }
    catch {
        Write-Host "Dissonance detected on $($Doc.Name): $_"
    }
}

# Terminate the Word process and sweep the memory footprint
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null

Write-Host "Batch rendering complete. The spooler is safe."
