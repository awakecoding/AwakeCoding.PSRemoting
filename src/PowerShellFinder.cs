using System;
using System.IO;
using System.Runtime.InteropServices;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal static class PowerShellFinder
    {
        public static string? GetPowerShellPath()
        {
            string? currentExePath = Environment.ProcessPath;

            if (currentExePath != null &&
                Path.GetFileName(currentExePath).Contains("pwsh", StringComparison.OrdinalIgnoreCase))
            {
                return currentExePath;
            }

            string executableName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "pwsh.exe" : "pwsh";
            return FindInPath(executableName);
        }

        public static string? GetWindowsPowerShellPath()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                return null;

            string? currentExePath = Environment.ProcessPath;

            if (currentExePath != null &&
                Path.GetFileName(currentExePath).Contains("powershell", StringComparison.OrdinalIgnoreCase))
            {
                return currentExePath;
            }

            return FindInPath("powershell.exe");
        }

        public static string? GetExecutablePath(bool useWindowsPowerShell)
        {
            return useWindowsPowerShell
                ? GetWindowsPowerShellPath()
                : GetPowerShellPath();
        }

        private static string? FindInPath(string executableName)
        {
            string[]? paths = Environment.GetEnvironmentVariable("PATH")?.Split(Path.PathSeparator);
            if (paths == null)
                return null;

            foreach (string path in paths)
            {
                string fullPath = Path.Combine(path, executableName);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }
    }
}
