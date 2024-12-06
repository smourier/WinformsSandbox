using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSTSCLib;
using Windows.Management.Deployment;

namespace WinformsSandbox;

public partial class Main : Form
{
    private ClientProxy? _grpcClient;
    private ConfigProxy? _sandboxConfig;
    private AxMSTSCLib.AxMsRdpClient8NotSafeForScripting? _rdpClient;

    public Main()
    {
        InitializeComponent();
        Icon = Resources.Resources.MainIcon;
    }

    protected override void Dispose(bool disposing) // removed from Main.Designer.cs
    {
        if (disposing)
        {
            components?.Dispose();
            if (_grpcClient != null)
            {
                if (_sandboxConfig != null)
                {
                    _grpcClient.ShutdownSandbox(_sandboxConfig.SandboxId);
                    _sandboxConfig = null;
                }
                _grpcClient.Dispose();
                _grpcClient = null;
            }
            _rdpClient?.Dispose();
        }
        base.Dispose(disposing);
    }

    protected override void CreateHandle()
    {
        base.CreateHandle();
        _ = Task.Run(() =>
        {
            BeginInvoke(async () =>
            {
                try
                {
                    GetSandboxClient();
                }
                catch (Exception ex)
                {
                    Controls.Add(new Label { Text = $"The Windows Sandbox API is not available. {ex.Message}", AutoSize = true, Padding = new Padding(10) });
                    return;
                }

                var connector = NamedPipeConnector.Create();
                if (connector == null)
                {
                    Controls.Add(new Label { Text = "The RDP API is not available.", AutoSize = true, Padding = new Padding(10) });
                    return;
                }

                await StartAndConnect(connector);
            });
        });
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        if (_rdpClient != null)
        {
            var size = ClientSize;
            if (_rdpClient.Connected == 1)
            {
                var ocx = (IMsRdpClient9)_rdpClient.GetOcx();
                ocx.UpdateSessionDisplaySettings((uint)size.Width, (uint)size.Height, (uint)size.Width, (uint)size.Height, 0u, 100, 100);
            }
            else
            {
                _rdpClient.DesktopHeight = size.Height;
                _rdpClient.DesktopWidth = size.Width;
            }
        }
    }

    private async Task StartAndConnect(NamedPipeConnector connector)
    {
        // start the sandbox
        _sandboxConfig = await _grpcClient!.StartSandboxAsync();

        // get a named pipe endpoint
        var pipeName = $@"\\.\pipe\{_sandboxConfig.VMId}";
        var endpoint = await connector.GetEndpoint(pipeName);

        // connect the RDP client
        _rdpClient = new AxMSTSCLib.AxMsRdpClient8NotSafeForScripting { Dock = DockStyle.Fill };
        Controls.Add(_rdpClient);
        _rdpClient.AdvancedSettings9.EnableCredSspSupport = false;
        _rdpClient.AdvancedSettings9.NegotiateSecurityLayer = true;
        _rdpClient.UserName = _sandboxConfig.Username;
        _rdpClient.AdvancedSettings9.ClearTextPassword = _sandboxConfig.Password;
        _rdpClient.AdvancedSettings9.set_ConnectWithEndpoint(ref endpoint);
        _rdpClient.Connect();
    }

    private void GetSandboxClient()
    {
        var pm = new PackageManager();
        var packageName = "Windows Sandbox";
        var package = pm.FindPackagesForUser(string.Empty).FirstOrDefault(p => p.DisplayName == packageName);
        if (package == null)
            throw new Exception($"Cannot find '{packageName}' package.");

        var file = Path.Combine(package.InstalledPath, "SandboxCommon.dll");
        if (!File.Exists(file))
            throw new Exception($"Cannot find '{file}' file.");

        var bytes = File.ReadAllBytes(file);
        var asm = Assembly.Load(bytes);
        var typeName = "SandboxCommon.Grpc.GrpcClient";
        var clientType = asm.GetType(typeName);
        if (clientType == null)
            throw new Exception($"Cannot find '{typeName}' type.");

        // resolve all dlls from the package
        AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
        {
            var name = e.Name.Split(',')[0] + ".dll";
            var assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.FullName == e.Name);
            if (assembly != null)
                return assembly;

            var file = Path.Combine(package.InstalledPath, name);
            if (File.Exists(file))
            {
                var bytes = File.ReadAllBytes(file);
                var asm = Assembly.Load(bytes);
                return asm;
            }
            return null;
        };

        if (Activator.CreateInstance(clientType, [null]) is not IDisposable client)
            throw new Exception($"Cannot find create instance of '{clientType.FullName}' type.");

        _grpcClient = new ClientProxy(client);
    }

    // reflection-built proxies, we could use the SandboxCommon.dll directly too.
    private sealed class ClientProxy(IDisposable client) : IDisposable
    {
        private IDisposable? _client = client;

        public void ShutdownSandbox(Guid sandboxId)
        {
            ObjectDisposedException.ThrowIf(_client == null, this);
            _client.GetType().InvokeMember(nameof(ShutdownSandbox), BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod, null, _client,
                [sandboxId]);
        }

        public async Task<ConfigProxy> StartSandboxAsync()
        {
            ObjectDisposedException.ThrowIf(_client == null, this);
            dynamic task = _client.GetType().InvokeMember(nameof(StartSandboxAsync), BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod, null, _client,
                [string.Empty, null, null])!;
            try
            {
                await task;
            }
            catch (COMException ex)
            {
                const int CO_E_APPSINGLEUSE = unchecked((int)0x800401f6);
                if (ex.HResult == CO_E_APPSINGLEUSE)
                    throw new Exception("Cannot start a new sandbox, the main one may be opened, or if it's not, try to kill running instances of WindowsSandboxServer.exe and ManagedWindowsVM.exe processes.");

                throw;
            }

            var config = task.GetAwaiter().GetResult();
            return new ConfigProxy(config);
        }

        public void Dispose() => Interlocked.Exchange(ref _client, null)?.Dispose();
    }

    private sealed class ConfigProxy
    {
        public ConfigProxy(object config)
        {
            var type = config.GetType();
            SandboxId = (Guid)type.GetProperty(nameof(SandboxId))!.GetValue(config)!;
            VMId = (Guid)type.GetProperty(nameof(VMId))!.GetValue(config)!;
            Username = (string)type.GetProperty(nameof(Username))!.GetValue(config)!;
            Password = (string)type.GetProperty(nameof(Password))!.GetValue(config)!;
        }

        public Guid SandboxId { get; }
        public Guid VMId { get; }
        public string Username { get; }
        public string Password { get; }

        public override string ToString() => $"{SandboxId}";
    }

    // undocumented RDP interfaces
    private sealed class NamedPipeConnector : NamedPipeConnector.IRDPENCNamedPipeDirectConnectorCallbacks, IDisposable
    {
        private IRDPENCNamedPipeDirectConnector? _connector;
        private readonly TaskCompletionSource<object> _source;

        public static NamedPipeConnector? Create()
        {
            var CLSID_RDPRuntimeSTAContext = new Guid("fb332ae7-0055-4208-92b7-20410ca8382b");
            _ = RDPBASE_CreateInstance(0, CLSID_RDPRuntimeSTAContext, typeof(IRDPENCPlatformContext).GUID, out var ctx);
            if (ctx == null)
                return null;

            var hr = ((IRDPENCPlatformContext)ctx).InitializeInstance();
            if (hr < 0)
                return null;

            var CLSID_RDPENCNamedPipeDirectConnector = new Guid("fb332ae7-0088-4208-92b7-20410ca8382b");
            _ = RDPBASE_CreateInstance(Marshal.GetIUnknownForObject(ctx), CLSID_RDPENCNamedPipeDirectConnector, typeof(IRDPENCNamedPipeDirectConnector).GUID, out var connector);
            if (connector == null)
                return null;

            return new NamedPipeConnector((IRDPENCNamedPipeDirectConnector)connector);
        }

        private NamedPipeConnector(IRDPENCNamedPipeDirectConnector connector)
        {
            _connector = connector;
            _source = new TaskCompletionSource<object>();
            var hr = _connector.InitializeInstance(this);
            if (hr < 0)
            {
                _source.SetException(Marshal.GetExceptionForHR(hr)!);
            }
        }

        public async Task<object> GetEndpoint(string pipeName)
        {
            ObjectDisposedException.ThrowIf(_connector == null, this);
            _connector.StartConnect(pipeName);
            return await _source.Task;
        }

        public void OnConnectionCompleted(object stream) => _source.TrySetResult(stream);
        public void OnConnectorError(int hr) => _source.TrySetException(Marshal.GetExceptionForHR(hr)!);

        public void Dispose()
        {
            try
            {
                Interlocked.Exchange(ref _connector, null)?.TerminateInstance();
            }
            catch
            {
                // do nothing
            }
        }

        [DllImport("RdpBase")]
        private static extern int RDPBASE_CreateInstance(nint platformContext, in Guid rclsid, in Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

        [ComImport, Guid("4ACF942D-EADC-45bf-8EA8-793FE3CE31E8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IRDPENCNamedPipeDirectConnector
        {
            [PreserveSig]
            int InitializeInstance(IRDPENCNamedPipeDirectConnectorCallbacks callbackInstance);

            [PreserveSig]
            int TerminateInstance();

            [PreserveSig]
            int StartConnect(string pipeName);
        }

        [ComImport, Guid("FB332AE7-000E-4208-92B7-20410CA8382B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IRDPENCPlatformContext
        {
            void _VtblGap0_9(); // skip 9 methods
            int InitializeInstance();
        }

        [ComImport, Guid("D923EFE9-0A6D-4344-92B3-164229DB8D2D"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IRDPENCNamedPipeDirectConnectorCallbacks
        {
            [PreserveSig]
            void OnConnectionCompleted([MarshalAs(UnmanagedType.IUnknown)] object endpoint);

            [PreserveSig]
            void OnConnectorError(int hr);
        }
    }
}
