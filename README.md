# DOCX to PDF Heavy Transmutation Pipeline

## Version: 1.1 (Prototype Grade)
## Author: Jake Gascon

Core Technology: PowerShell, Microsoft Word COM Interop

# Overview

The DOCX -> PDF Heavy Transmutation Pipeline is an automated, hardened PowerShell script designed to reliably batch-convert Microsoft Word documents (.docx) into fixed-format PDF files.

Unlike standard conversion scripts, this pipeline is built for high-throughput, unattended environments. It features built-in idempotency, network permission validation, self-healing COM object recovery, and aggressive memory management to prevent zombie WINWORD.EXE processes.

# Core Features

Automated Logistics Hub: Automatically builds its required folder infrastructure (Queue, Outbox, Archive).

Kinetic Access Probes: Pre-flights target directories with read/write/delete probes to ensure full network permissions before beginning operations.

Headless Word Engine: Spins up MS Word entirely in the background (Visible = $False), suppresses all UI alerts, and forcefully disables macros for maximum security.

Idempotency Bypass: Detects if a document has already been processed (PDF exists in Outbox) to save processing cycles, archiving the original immediately.

Defibrillator Check: Monitors the Word RPC server mid-flight. If the Word engine flatlines or crashes, the script performs a hard restart of the COM object and resumes.

Nuclear Cleanup: Aggressive garbage collection ([System.GC]) and specific COM object release protocols guarantee that system memory is swept and no background applications are left running.

Comprehensive Audit Logging: Writes timestamped logs to the console and a local Transmutation_Audit.log file, categorizing outputs by [INFO], [SUCCESS], [WARN], [ERROR], and [CRITICAL].

# Prerequisites

To deploy the Transmutation Pipeline, the host machine must have:

Windows OS with PowerShell 5.1 or higher.

Microsoft Office / Microsoft Word installed locally. The script relies on the Word COM Object (Word.Application) to ensure 100% native formatting accuracy.

# Folder Architecture (The Logistics Hub)

The script relies on a central "Logistics Hub." When executed, it guarantees the existence of the following hierarchy:

LogisticsHub/
│
├── 01_Queue/                 <-- Place target .docx files here.
├── 02_Outbox/                <-- Generated .pdf files appear here.
├── 03_Archive/               <-- Original .docx files are moved here after success.
└── Transmutation_Audit.log   <-- Continuous rolling log of all pipeline events.


# Setup & Execution

# 1. Configuration

Open PDF_from_DOCX.ps1 in a text editor and update the $BaseLogisticsHub variable to point to your desired master directory:

# Example:
$BaseLogisticsHub = "C:\Users\jlgascon\Desktop\Test_Print_Batch\LogisticsHub"


# 2. Deployment

Drop your .docx files into the 01_Queue folder.

Open PowerShell and execute the script:

.\PDF_from_DOCX.ps1


Watch the console or the Transmutation_Audit.log to monitor the process.

Retrieve your forged PDFs from the 02_Outbox.

To execute without needing to change directies:

powershell.exe -ExecutionPolicy Bypass -File [ABSOLUTE REF]

To execute in Powershell go to directory for LogisticsHub
Set-ExecutionPolicy Bypass -Scope Process -Force .\PDF_from_DOCX.ps1

# Troubleshooting & Known Issues

Error: Cannot convert the value of type "psobject" to type "Object"

Fix: Ensure the $OpenDoc.SaveAs() command passes a strict [string] path. (Fixed in current build via $OpenDoc.SaveAs([string]$PdfPath, 17)).

Error: ACCESS DENIED on [Directory]. Dissonance detected in network permissions.

Fix: The script's kinetic probe failed. Check Windows folder permissions or ensure your antivirus isn't blocking script-based file generation in the target path.

Issue: Word keeps crashing on a specific document.

Fix: Check the source document for corrupt elements, password protection, or active DRM.

Pipeline secured. The spooler is safe.
