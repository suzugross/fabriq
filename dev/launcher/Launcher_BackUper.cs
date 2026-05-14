// ========================================
// Fabriq BackUper Launcher
// Tiny C# wrapper that launches apps\fabriq_backuper\fabriq_backuper.ps1
// via conhost + powershell. Standalone entry point for the
// fabriq_backuper satellite app (backup/restore tool).
// ========================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("Fabriq BackUper")]
[assembly: AssemblyProduct("Fabriq BackUper")]
[assembly: AssemblyDescription("Standalone backup/restore satellite for the Fabriq kitting framework")]
[assembly: AssemblyCompany("Fabriq Project")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]

namespace FabriqBackUper
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
                string exePath = Assembly.GetExecutingAssembly().Location;
                string baseDir = Path.GetDirectoryName(exePath);

                if (string.IsNullOrEmpty(baseDir))
                {
                    ShowError("Failed to resolve launcher directory.");
                    return 1;
                }

                // fabriq_backuper.ps1 dot-sources kernel\common.ps1 via
                // relative path, so the current directory must be pinned to
                // the fabriq root (where Fabriq_BackUper.exe lives).
                Directory.SetCurrentDirectory(baseDir);

                string entryPs1 = Path.Combine(baseDir, "apps", "fabriq_backuper", "fabriq_backuper.ps1");
                if (!File.Exists(entryPs1))
                {
                    ShowError("apps\\fabriq_backuper\\fabriq_backuper.ps1 was not found under:\n" + baseDir);
                    return 2;
                }

                // Launch via conhost.exe so we get a fresh console window
                // independent of any parent. fabriq_backuper.ps1 will then
                // self-spawn into an isolated subprocess (FABRIQ_BACKUPER_SUBPROCESS
                // sentinel) so PSReadLine handlers / env vars / global
                // state never leak back into the launcher process.
                var psi = new ProcessStartInfo
                {
                    FileName = "conhost.exe",
                    Arguments = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \".\\apps\\fabriq_backuper\\fabriq_backuper.ps1\"",
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
            MessageBoxW(IntPtr.Zero, message, "Fabriq BackUper Launcher", MB_ICONERROR | MB_OK);
        }
    }
}
