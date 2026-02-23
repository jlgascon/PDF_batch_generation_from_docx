# ==============================================================================
# DOCX to PDF Batch Transmutation Protocol
# ==============================================================================

$TargetFolder = "C:\Path\To\Your\Batch\Folder"

# Pre-flight check: Ensure the staging ground actually exists
if (-not (Test-Path -Path $TargetFolder)) {
    Write-Error "Write-error detected: Target folder '$TargetFolder' does not exist. Aborting sequence."
    exit
}

$Docs = Get-ChildItem -Path $TargetFolder -Filter *.docx
if ($Docs.Count -eq 0) {
    Write-Host "Staging ground is empty. No raw materials found."
    exit
}

Write-Host "Initiating batch conversion pipeline for $($Docs.Count) files..."

# Word COM Object Constants, 17 is the hardcode for PDF in word
$wdFormatPDF = 17 
$wdAlertsNone = 0
$msoAutomationSecurityForceDisable = 3

$Word = $null

try {
    # Spin up the headless rendering engine
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $False 
    
    # Best Practice: Suppress all modal pop-ups (e.g., "Document contains unreadable content")
    # If a pop-up triggers in a hidden window, the script hangs eternally.
    $Word.DisplayAlerts = $wdAlertsNone
    
    # Best Practice: Quarantine. Force disable macros in the target documents to prevent rogue execution
    $Word.AutomationSecurity = $msoAutomationSecurityForceDisable

    foreach ($Doc in $Docs) {
        $PdfPath = $Doc.FullName.Replace(".docx", ".pdf")
        Write-Host "Synthesizing: $($Doc.Name) -> $($Doc.BaseName).pdf"
        
        $OpenDoc = $null
        try {
            # Open silently and strictly in read-only mode
            $OpenDoc = $Word.Documents.Open($Doc.FullName, $null, $True)
            
            # Forge the PDF
            $OpenDoc.SaveAs([ref]$PdfPath, [ref]$wdFormatPDF)
        }
        catch {
            Write-Error "Rendering failure on $($Doc.Name): $_"
        }
        finally {
            # Best Practice: Always close the document within a finally block
            # wdDoNotSaveChanges = 0
            if ($null -ne $OpenDoc) {
                $OpenDoc.Close(0)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OpenDoc) | Out-Null
            }
        }
    }
}
catch {
    Write-Error "Catastrophic pipeline failure: $_"
}
finally {
    # ==========================================================================
    # NUCLEAR CLEANUP: This block executes even if the script crashes or is killed
    # ==========================================================================
    if ($null -ne $Word) {
        Write-Host "Terminating Word engine and sweeping memory footprint..."
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
    }
    
    # Force the .NET garbage collector to purge the released COM pointers
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Pipeline secured."
}
