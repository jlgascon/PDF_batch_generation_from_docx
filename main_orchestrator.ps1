# ==============================================================================
# DOCX -> PDF Heavy Transmutation Pipeline (Production Grade)
# ==============================================================================

# --- 0. Global Configuration ---
# Set your master directory here. The script will build the rest.
$BaseLogisticsHub = "C:\Path\To\Your\LogisticsHub" 

$QueueDir   = Join-Path $BaseLogisticsHub "01_Queue"    
$OutboxDir  = Join-Path $BaseLogisticsHub "02_Outbox"   
$ArchiveDir = Join-Path $BaseLogisticsHub "03_Archive"  
$LogFile    = Join-Path $BaseLogisticsHub "Transmutation_Audit.log"

# Word COM Object Constants
$wdFormatPDF = 17 
$wdAlertsNone = 0
$msoAutomationSecurityForceDisable = 3

# --- 1. Helper Functions ---

function Write-Audit {
    param ([string]$Message, [string]$Level="INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry
}

function Test-KineticAccess {
    param ([string]$TargetDirectory)
    $ProbeFile = Join-Path $TargetDirectory "recon_probe_$([guid]::NewGuid()).tmp"
    try {
        New-Item -Path $ProbeFile -ItemType File -Force -ErrorAction Stop | Out-Null
        Remove-Item -Path $ProbeFile -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Start-WordEngine {
    Write-Audit "Spinning up headless Word rendering engine..." "INFO"
    $Engine = New-Object -ComObject Word.Application
    $Engine.Visible = $False 
    $Engine.DisplayAlerts = $wdAlertsNone
    $Engine.AutomationSecurity = $msoAutomationSecurityForceDisable
    return $Engine
}

# --- 2. Infrastructure & Reconnaissance ---

# Forge the directories if they don't exist
foreach ($Dir in @($QueueDir, $OutboxDir, $ArchiveDir)) {
    if (-not (Test-Path $Dir)) { 
        New-Item -ItemType Directory -Path $Dir | Out-Null 
    }
}

Write-Audit "Initiating kinetic access probes on logistics hubs..." "INFO"

foreach ($Dir in @($QueueDir, $OutboxDir, $ArchiveDir)) {
    if (-not (Test-KineticAccess -TargetDirectory $Dir)) {
        Write-Audit "ACCESS DENIED on $Dir. Dissonance detected in network permissions. Aborting deployment." "CRITICAL"
        exit
    }
}
Write-Audit "Perimeter secure. Full read/write network access confirmed." "SUCCESS"

# --- 3. Target Acquisition ---

$Docs = Get-ChildItem -Path $QueueDir -Filter *.docx
if ($Docs.Count -eq 0) {
    Write-Audit "Queue is empty. Standing down." "INFO"
    exit
}
Write-Audit "Target acquired: $($Docs.Count) documents in the Queue." "INFO"

# --- 4. The Forge (Main Execution Loop) ---

$Word = Start-WordEngine
# LETS COME BACK TO THIS AND CLEAN UP THE OUTPUTS FOR CONSOLE CLARITY AND LOGGING EFFICIENCY

try {
    foreach ($Doc in $Docs) {
        $PdfName = "$($Doc.BaseName).pdf"
        $PdfPath = Join-Path $OutboxDir $PdfName
        $ArchivePath = Join-Path $ArchiveDir $Doc.Name

        # Idempotency Bypass
        if (Test-Path $PdfPath) {
            Write-Audit "Bypassing $($Doc.Name) - PDF already forged in Outbox." "WARN"
            Move-Item -Path $Doc.FullName -Destination $ArchivePath -Force
            continue
        }

        Write-Audit "Synthesizing: $($Doc.Name)" "INFO"

        # Engine Health Check (Defibrillator)
        try {
            $HealthCheck = $Word.Version
        } catch {
            Write-Audit "Word engine flatlined (RPC Server Unavailable). Initiating hard restart..." "ERROR"
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
            $Word = Start-WordEngine
        }

        $OpenDoc = $null
        try {
            # Open silently, read-only
            $OpenDoc = $Word.Documents.Open($Doc.FullName, $null, $True)
            
            # Forge the PDF
            # $OpenDoc.SaveAs([ref]$PdfPath, [ref]$wdFormatPDF)
            # Issues with object being passed instead of string literal, trying a hard coded 
	        $OpenDoc.SaveAs([string]$PdfPath, 17)
            
            # Close immediately
            $OpenDoc.Close(0)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OpenDoc) | Out-Null
            $OpenDoc = $null
            
            # Archive the raw ore
            Move-Item -Path $Doc.FullName -Destination $ArchivePath -Force
            Write-Audit "Render complete. Archived raw ore: $($Doc.Name)" "SUCCESS"
        }
        catch {
            Write-Audit "Catastrophic dissonance rendering $($Doc.Name): $_" "ERROR"
        }
        finally {
            if ($null -ne $OpenDoc) {
                try { $OpenDoc.Close(0) } catch {}
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OpenDoc) | Out-Null
            }
        }
    }
}
catch {
    Write-Audit "Critical Pipeline Failure: $_" "CRITICAL"
}
finally {
    # --- 5. Nuclear Cleanup ---
    Write-Audit "Terminating Word engine and sweeping memory footprint..." "INFO"
    if ($null -ne $Word) {
        try { $Word.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Audit "Pipeline secured. The spooler is safe." "INFO"
}
