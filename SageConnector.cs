using System.Runtime.InteropServices;

namespace SageWebAPI;

/// <summary>
/// Wraps the Sage 100 Contractor COM DLL (Sage.SMB.Api.dll).
/// Must run as x86 (32-bit) process — COM interop requirement.
/// </summary>
public class SageConnector : IDisposable
{
    // This is the COM interface the Sage DLL exposes.
    // We define it here so C# knows what methods to expect.
    // The Guid must match exactly what Sage registered in Windows COM.
    [ComImport]
    [Guid("6C9B6F2E-B680-11D0-A3A5-00AA0060D93B")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    private interface IMBXML
    {
        int InitializeAPI();
        int SetDataSource(string sqlInstance);
        int EnableRequests();
        void DisableRequests();
        void DeInitializeAPI();
        string SubmitXML(string company, string username, string password, string xmlRequest);
        string apiVersion();
    }

    // The actual COM object instance
    private IMBXML? _api;
    private bool _initialized = false;
    private bool _disposed = false;

    // These come from your appsettings.json — we'll wire that up next
    private readonly string _sqlInstance;
    private readonly string _company;
    private readonly string _username;
    private readonly string _password;

    public SageConnector(string sqlInstance, string company, string username, string password)
    {
        _sqlInstance = sqlInstance;
        _company = company;
        _username = username;
        _password = password;
    }

    /// <summary>
    /// Starts up the Sage API. Call this once when the app starts.
    /// </summary>
    public void Initialize()
    {
        // Activate the COM object from the registered DLL
        var comType = Type.GetTypeFromProgID("Sage.SMB.Api")
            ?? throw new InvalidOperationException(
                "Could not find Sage.SMB.Api COM object. " +
                "Is Sage 100 Contractor installed on this machine?");

        _api = (IMBXML)Activator.CreateInstance(comType)!;

        // Step 1: Initialize
        int result = _api.InitializeAPI();
        if (result != 0)
            throw new InvalidOperationException($"InitializeAPI failed with code {result}");

        // Step 2: Point it at the SQL Server instance
        result = _api.SetDataSource(_sqlInstance);
        if (result != 0)
            throw new InvalidOperationException($"SetDataSource failed with code {result}");

        // Step 3: Open for requests
        result = _api.EnableRequests();
        if (result != 0)
            throw new InvalidOperationException($"EnableRequests failed with code {result}");

        _initialized = true;
        Console.WriteLine($"Sage API initialized. Version: {_api.apiVersion()}");
    }

    /// <summary>
    /// Sends an mbXML request to Sage and returns the XML response string.
    /// </summary>
    public string SubmitXML(string xmlRequest)
    {
        if (!_initialized || _api == null)
            throw new InvalidOperationException("SageConnector is not initialized.");

        return _api.SubmitXML(_company, _username, _password, xmlRequest);
    }

    /// <summary>
    /// Cleanup — called automatically when the app shuts down.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;

        if (_api != null)
        {
            _api.DisableRequests();
            _api.DeInitializeAPI();

            // Release the COM object back to Windows
            Marshal.ReleaseComObject(_api);
            _api = null;
        }

        _disposed = true;
    }
}