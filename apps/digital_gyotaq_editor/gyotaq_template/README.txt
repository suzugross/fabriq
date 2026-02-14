========================================
Digital Gyotaku Template Module
========================================

Usage:

  1. Copy this entire folder to a new folder name.
     Example: modules/standard/gyotaku_vpn_setup/

  2. Edit module.csv:
     - Change MenuName to your task set name
     - Adjust Order as needed
     - Category should remain "Evidence"

  3. Edit task_list.csv:
     - Add your manual task items
     - Set Enabled=1 for active tasks, 0 for disabled

  4. Restart Fabriq. The framework will auto-detect
     the new module and display it in the menu.

----------------------------------------
CSV Column Reference (task_list.csv)
----------------------------------------

  Enabled      1=active, 0=disabled
  TaskID       Unique ID (used in screenshot filename)
  TaskTitle    Short title (max ~40 chars recommended)
  Instruction  Full instruction text for the worker
  OpenCommand  Auto-open command (optional, leave empty
               if not needed)
               Examples:
                 ms-settings:windowsupdate
                 ms-settings:network
                 control
                 C:\Windows\System32\mmc.exe
  OpenArgs     Arguments for OpenCommand (optional)
               Example: /name Microsoft.BitLockerDriveEncryption

----------------------------------------
Output
----------------------------------------

  Screenshots: logs\gyotaku\{PCName}\
  Manifest:    logs\gyotaku\{PCName}\{date}_manifest.csv

  Filename format:
    {yyyy_MM_dd}_{TaskID}_{TaskTitle}_{PCName}.png

----------------------------------------
Keyboard Shortcuts (in the dialog)
----------------------------------------

  Enter / F5   Take evidence screenshot & complete
  Escape       Skip current task

========================================
