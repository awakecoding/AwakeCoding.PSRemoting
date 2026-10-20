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
        WinRM,

        /// <summary>
        /// gRPC transport
        /// </summary>
        Grpc
    }

    /// <summary>
    /// PSObject payload encoding mode for the gRPC transport.
    /// </summary>
    public enum PSHostGrpcEncoding
    {
        /// <summary>
        /// Protobuf v3 PSObject encoding (PSProtoSerializer).
        /// </summary>
        Protobuf = 0,

        /// <summary>
        /// CBOR RFC 8949 PSObject encoding (PSCborSerializer).
        /// </summary>
        Cbor = 1
    }
}
