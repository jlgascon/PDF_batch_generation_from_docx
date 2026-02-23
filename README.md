# DOCX-to-PDF Heavy Transmutation Pipeline (v1.0)

## Overview
A headless, idempotent PowerShell engine designed to batch-convert Microsoft Word (`.docx`) files into lightweight `.pdf` formats. 

This pipeline was engineered to bypass the native Windows Explorer GUI limitations. By instantiating a hidden Word COM object, it shifts the rendering math to the background, preventing system-wide UI lockups and protecting the local print spooler from catastrophic RAM thrashing during mass administrative deployments.

## Architecture: The Logistics Hub
The script enforces strict data hygiene. It requires a master directory containing three distinct sub-folders to segregate raw materials from processed assets.



```text
C:\Path\To\Your\LogisticsHub\
│
├── 01_Queue\                  # Drop raw .docx files here.
├── 02_Outbox\                 # Rendered .pdf files are minted here.
├── 03_Archive\                # Processed .docx files are buried here.
│
├── Transmutation_Audit.log    # Rolling system log.
└── Transmutation-Engine.ps1   # The master script.
