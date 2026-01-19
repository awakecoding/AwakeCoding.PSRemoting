using System;
using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal sealed class PSHostClientInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public string Executable { get; set; }

        public string Arguments { get; set; }

        public override PSCredential? Credential
        {
            get { return null; }
            set { throw new NotImplementedException(); }
        }

        public override AuthenticationMechanism AuthenticationMechanism
        {
            get { return AuthenticationMechanism.Default; }
            set { throw new NotImplementedException(); }
        }

        public override string CertificateThumbprint
        {
            get { return string.Empty; }
            set { throw new NotImplementedException(); }
        }

        public PSHostClientInfo(string computerName, string executable, string arguments)
        {
            ComputerName = computerName;
            Executable = executable;
            Arguments = arguments;
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostClientInfo(ComputerName, Executable, Arguments);
            return connectionInfo;
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    internal sealed class PSHostClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostClientInfo _connectionInfo;
        private Process? _process = null;

        internal PSHostClientSessionTransportMgr(
            PSHostClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            _process = new Process();
            _process.StartInfo.FileName = _connectionInfo.Executable;
            _process.StartInfo.Arguments = _connectionInfo.Arguments;
            _process.StartInfo.RedirectStandardInput = true;
            _process.StartInfo.RedirectStandardOutput = true;
            _process.StartInfo.RedirectStandardError = true;
            _process.StartInfo.UseShellExecute = false;
            _process.StartInfo.CreateNoWindow = true;

            _process.ErrorDataReceived += ErrorDataReceived;
            _process.OutputDataReceived += OutputDataReceived;
            _process.Start();
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            SetMessageWriter(_process.StandardInput);
            SendOneItem();
        }

        private void ErrorDataReceived(object? sender, DataReceivedEventArgs args)
        {
            HandleErrorDataReceived(args.Data);
        }

        private void OutputDataReceived(object? sender, DataReceivedEventArgs args)
        {
            HandleOutputDataReceived(args.Data);
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                CleanupConnection();
            }

            base.Dispose(isDisposing);
        }

        protected override void CleanupConnection()
        {
            if (_process != null)
            {
                try
                {
                    if (!_process.HasExited)
                    {
                        _process.Kill();
                        _process.WaitForExit(500);
                    }
                }
                catch { }
                finally
                {
                    _process?.Dispose();
                }
            }
            _process = null;
        }
    }
}
