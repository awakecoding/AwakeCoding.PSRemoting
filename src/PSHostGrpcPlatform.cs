using System;
using System.Runtime.InteropServices;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal static class PSHostGrpcPlatform
    {
        public static bool IsSupported =>
            !(RuntimeInformation.IsOSPlatform(OSPlatform.OSX) &&
              RuntimeInformation.ProcessArchitecture == Architecture.Arm64);

        public static void EnsureSupported()
        {
            if (!IsSupported)
            {
                throw new PlatformNotSupportedException("The Grpc.Core native runtime used by this transport does not support macOS arm64.");
            }
        }
    }
}
