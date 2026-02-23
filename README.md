DOCX-to-PDF Heavy Transmutation Pipeline (v1.0)
Overview
A headless, idempotent PowerShell engine designed to batch-convert Microsoft Word (.docx) files into lightweight .pdf formats.

This pipeline was engineered to bypass the native Windows Explorer GUI limitations. By instantiating a hidden Word COM object, it shifts the rendering math to the background, preventing system-wide UI lockups and protecting the local print spooler from catastrophic RAM thrashing during mass administrative deployments.

Architecture: The Logistics Hub
The script enforces strict data hygiene. It requires a master directory containing three distinct sub-folders to segregate raw materials from processed assets.

Plaintext
C:\Path\To\Your\LogisticsHub\
│
├── 01_Queue\      # Drop raw .docx files here.
├── 02_Outbox\     # Rendered .pdf files are minted here.
├── 03_Archive\    # Processed .docx files are buried here.
│
├── Transmutation_Audit.log    # Rolling system log.
└── Transmutation-Engine.ps1   # The master script.
Note: The script will automatically forge these directories on its first run if they do not exist.

Core Engineering Features
Kinetic Reconnaissance: Before spinning up the rendering engine, the script drops a microscopic, mathematically unique probe file into all three network directories and instantly deletes it. If write-access is denied, it hard-aborts to prevent ghost files and corrupted states.

Headless Execution: Suppresses the winword.exe UI (Visible = $False) and neutralizes modal pop-ups to prevent the loop from hanging on broken margins or unreadable content.

Idempotency Checks: Prevents wasted compute cycles. If a requested PDF already exists in 02_Outbox, the engine bypasses rendering, routes the raw file to 03_Archive, and moves to the next target.

COM Object Resiliency: Includes a pre-render heartbeat check. If a deeply corrupted document hard-crashes the background Word process, the script catches the RPC disconnect, aggressively releases the dead pointer, and spins up a fresh rendering engine mid-loop.

Macro Quarantine: Forces $msoAutomationSecurityForceDisable to prevent rogue VBA execution from third-party .docx files.

Nuclear Garbage Collection: Guarantees that the Word COM object is violently terminated and swept by the .NET Garbage Collector at the end of the run, even if the user manually kills the terminal window. No zombie RAM leaks.

Deployment: The Detonator (For End-Users)
Windows Enterprise policies natively block double-click execution of .ps1 files. To deploy this to an end-user without requiring command-line access, create a Windows Shortcut that temporarily bypasses the local execution policy.

Right-click the Desktop -> New > Shortcut.

Inject the following target string (update the absolute path to match your environment):

Plaintext
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File "C:\Path\To\Your\LogisticsHub\Transmutation-Engine.ps1"
Name it "Initialize PDF Pipeline" and assign an appropriate icon.

Troubleshooting & Dissonance Resolution
If the pipeline jams, do not rely on user testimony. Open Transmutation_Audit.log.

[INFO]: Standard operational milestones.

[SUCCESS]: File successfully rendered and archived.

[WARN]: Non-fatal bypass (e.g., file already existed in Outbox).

[ERROR]: A specific file failed to render or hard-crashed the COM object. The engine will attempt to auto-recover and proceed to the next file.

[CRITICAL]: Catastrophic script failure or Network Access Denied. The pipeline has shut down and initiated the memory sweep.
