using System;
using System.Diagnostics;
using System.Globalization;
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

        /// <summary>
        /// Constructs the PowerShell host named pipe name for a given process ID.
        /// Formula: PSHost.{timestamp}.{pid}.{appDomain}.{processName}
        /// On Windows, appDomain defaults to "DefaultAppDomain" (or uses provided value)
        /// On Unix, appDomain is always "None" (the appDomainName parameter is ignored)
        /// </summary>
        public static string GetProcessPipeName(int processId, string? appDomainName = null)
        {
            try
            {
                using (var process = Process.GetProcessById(processId))
                {
                    string processName = process.ProcessName;
                    string timestamp;
                    string effectiveAppDomainName;

                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        // Windows: use FileTime format and DefaultAppDomain (or provided value)
                        timestamp = process.StartTime.ToFileTime().ToString(CultureInfo.InvariantCulture);
                        effectiveAppDomainName = appDomainName ?? "DefaultAppDomain";
                    }
                    else
                    {
                        // Unix: use hex-encoded seconds (8 characters from FileTime) and always "None" for appDomain
                        // PowerShell on Unix does not use AppDomains, so the pipe name always uses "None"
                        timestamp = process.StartTime.ToFileTime().ToString("X8").Substring(1, 8);
                        effectiveAppDomainName = "None";
                    }

                    return $"PSHost.{timestamp}.{processId}.{effectiveAppDomainName}.{processName}";
                }
            }
            catch (ArgumentException)
            {
                throw new InvalidOperationException($"Process with ID {processId} not found");
            }
        }

        /// <summary>
        /// Gets the full named pipe path with platform-specific prefix.
        /// Windows: \\.\pipe\{pipeName}
        /// Unix: /tmp/CoreFxPipe_{pipeName}
        /// </summary>
        public static string GetPipeNameWithPrefix(string pipeName)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return $@"\\.\pipe\{pipeName}";
            }
            else
            {
                return Path.Combine(Path.GetTempPath(), $"CoreFxPipe_{pipeName}");
            }
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
