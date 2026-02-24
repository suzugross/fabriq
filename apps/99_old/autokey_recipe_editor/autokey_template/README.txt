========================================
Auto Keyboard Template - Usage Guide
========================================

This is a template module for automating keyboard input.
Copy this directory and customize recipe.csv for your automation needs.

HOW TO CREATE A NEW AUTOMATION MODULE
========================================

1. Copy the entire "autokey_template" directory
   Example: autokey_template -> autokey_vpn_setup

2. Edit module.csv: Change MenuName
   Example: "Auto VPN Setup,Automation,autokey_config.ps1,51,1"

3. Edit recipe.csv: Define your automation steps

4. Restart Fabriq - the new module will appear in the menu

RECIPE CSV FORMAT
========================================

Columns:
  Step   - Step number (display only, execution follows CSV row order)
  Action - Action type (see below)
  Value  - Action-specific value
  Wait   - Post-step wait in ms (for WaitWin: timeout in ms)
  Note   - Description (display only, not used in processing)

AVAILABLE ACTIONS
========================================

Open
  Opens an application, file, or URL.
  Value: Command to execute. Arguments separated by space.
  Examples:
    notepad
    notepad C:\temp\test.txt
    https://www.google.com
    ms-settings:windowsupdate
    control /name Microsoft.BitLockerDriveEncryption

WaitWin
  Waits for a window to appear (partial title match).
  Activates the window once found.
  Value: Window title (partial match)
  Wait:  Timeout in milliseconds (default: 10000)
  Example: WaitWin,Notepad,10000,Wait for Notepad

AppFocus
  Switches focus to an existing window.
  Value: Window title (passed to WScript.Shell.AppActivate)
  Example: AppFocus,Notepad,500,Focus Notepad

Type
  Types a string into the focused window.
  Value: Text to type
  Note:  Uses SendKeys - special characters may need escaping.
  Example: Type,Hello World,500,Type greeting

Key
  Sends special key combinations.
  Value: SendKeys notation
  Common keys:
    {ENTER}    - Enter key
    {TAB}      - Tab key
    {ESC}      - Escape key
    {BACKSPACE} - Backspace
    {DELETE}   - Delete
    {UP}{DOWN}{LEFT}{RIGHT} - Arrow keys
    {F1}~{F12} - Function keys
    %{F4}      - Alt+F4
    ^c         - Ctrl+C
    ^v         - Ctrl+V
    ^a         - Ctrl+A
    +{TAB}     - Shift+Tab
    ^{ESC}     - Ctrl+Esc (Start menu)
  Full reference: https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys

Wait
  Pauses execution for a specified duration.
  Value: Wait time in milliseconds
  Example: Wait,2000,0,Wait 2 seconds

TIPS
========================================

- WaitWin timeout: If a window doesn't appear within the timeout,
  the step is marked as failed but execution continues.

- Post-step wait: The Wait column adds a pause AFTER the action
  completes (except for WaitWin and Wait actions which handle
  their own timing).

- Japanese input: Type action uses SendKeys which works with
  IME-converted text. For direct Japanese input, ensure IME
  is in the correct mode first.

- URL opening: Use Open action with a URL to open in default browser.

- Settings pages: Use "ms-settings:" URIs to open Windows Settings.
  Example: ms-settings:network, ms-settings:windowsupdate
