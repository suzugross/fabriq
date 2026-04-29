// ========================================
// Fabriq IOS Launcher
// Tiny C# wrapper that launches apps\fabriq_ios\fabriq_ios.ps1 via
// conhost + powershell. Standalone entry point for the fabriq_ios
// sub-project (no operator dashboard, no profile system, no host
// list - just the joke shell).
// ========================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("Fabriq IOS")]
[assembly: AssemblyProduct("Fabriq IOS")]
[assembly: AssemblyDescription("Cisco IOS-style joke shell over the Fabriq kitting framework")]
[assembly: AssemblyCompany("Fabriq Project")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]

namespace FabriqIos
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

                // fabriq_ios.ps1 dot-sources kernel\common.ps1 via relative
                // path, so the current directory must be pinned to the fabriq
                // root (where Fabriq_IOS.exe lives).
                Directory.SetCurrentDirectory(baseDir);

                string entryPs1 = Path.Combine(baseDir, "apps", "fabriq_ios", "fabriq_ios.ps1");
                if (!File.Exists(entryPs1))
                {
                    ShowError("apps\\fabriq_ios\\fabriq_ios.ps1 was not found under:\n" + baseDir);
                    return 2;
                }

                // Launch via conhost.exe so we get a fresh console window
                // independent of any parent. fabriq_ios.ps1 will then self-
                // spawn into an isolated subprocess (FABRIQ_IOS_SUBPROCESS
                // sentinel) so PSReadLine handlers / env vars / global
                // state never leak back into the launcher process.
                var psi = new ProcessStartInfo
                {
                    FileName = "conhost.exe",
                    Arguments = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \".\\apps\\fabriq_ios\\fabriq_ios.ps1\"",
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
            MessageBoxW(IntPtr.Zero, message, "Fabriq IOS Launcher", MB_ICONERROR | MB_OK);
        }
    }
}
