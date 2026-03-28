namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Transport type for PSHost server connections
    /// </summary>
    public enum PSHostTransportType
    {
        /// <summary>
        /// TCP transport
        /// </summary>
        TCP,

        /// <summary>
        /// WebSocket transport
        /// </summary>
        WebSocket,

        /// <summary>
        /// Named Pipe transport
        /// </summary>
        NamedPipe,

        /// <summary>
        /// WinRM (WSMan over HTTP/HTTPS) transport
        /// </summary>
        WinRM
    }
}
