// ========================================
// Fabriq Launcher
// Tiny C# wrapper that launches kernel\main.ps1 via conhost + powershell.
// Purpose: provide an app-like entry point (custom icon, UAC auto-elevation,
// product metadata) without modifying the fabriq runtime.
// ========================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("Fabriq")]
[assembly: AssemblyProduct("Fabriq")]
[assembly: AssemblyDescription("Windows kitting automation framework")]
[assembly: AssemblyCompany("Fabriq Project")]
[assembly: AssemblyVersion("2.1.0.0")]
[assembly: AssemblyFileVersion("2.1.0.0")]

namespace Fabriq
{
    internal static class Launcher
    {
        // Win32 MessageBox for error reporting (avoids pulling in WinForms).
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        private const uint MB_ICONERROR = 0x00000010;
        private const uint MB_OK = 0x00000000;

        [STAThread]
        private static int Main()
        {
            try
            {
                // Resolve the directory where Fabriq.exe lives.
                // Use Assembly location rather than AppContext.BaseDirectory to be
                // compatible with .NET Framework 4.x.
                string exePath = Assembly.GetExecutingAssembly().Location;
                string baseDir = Path.GetDirectoryName(exePath);

                if (string.IsNullOrEmpty(baseDir))
                {
                    ShowError("Failed to resolve launcher directory.");
                    return 1;
                }

                // Fabriq main.ps1 relies on relative paths like .\kernel\common.ps1,
                // so the current directory must be pinned to the fabriq root.
                Directory.SetCurrentDirectory(baseDir);

                string mainPs1 = Path.Combine(baseDir, "kernel", "main.ps1");
                if (!File.Exists(mainPs1))
                {
                    ShowError("kernel\\main.ps1 was not found under:\n" + baseDir);
                    return 2;
                }

                // Launch via conhost.exe to match the behavior of Fabriq.bat
                // (reliable window-size control; Windows Terminal ignores mode con).
                // Fire-and-forget: the launcher exits immediately after starting
                // the child so UAC flow feels like a native app launch.
                // UseShellExecute = true ensures the child is launched via
                // ShellExecuteEx and gets its own fresh console window, matching
                // the experience of double-clicking conhost.exe from Explorer.
                // This is required when the launcher itself is a Windows
                // subsystem binary (/target:winexe), so the child console is
                // not accidentally tied to a non-existent parent console.
                var psi = new ProcessStartInfo
                {
                    FileName = "conhost.exe",
                    Arguments = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \".\\kernel\\main.ps1\"",
                    WorkingDirectory = baseDir,
                    UseShellExecute = true,
                };

                Process.Start(psi);
                return 0;
            }
            catch (Exception ex)
            {
                ShowError("Unexpected launcher error:\n" + ex.Message);
                return 99;
            }
        }

        private static void ShowError(string message)
        {
            MessageBoxW(IntPtr.Zero, message, "Fabriq Launcher", MB_ICONERROR | MB_OK);
        }
    }
}
