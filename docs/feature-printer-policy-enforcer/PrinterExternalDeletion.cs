using System;
using System.Diagnostics.CodeAnalysis;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using Microsoft.Extensions.Logging;
using ndd.SharedKernel.Result;
using nddSmartPrintQueueAgent.Domain.Contracts.PrintQueue;
using nddSmartPrintQueueAgent.Infra.Windows.Win32;
using Vanara.PInvoke;
using static Vanara.PInvoke.Kernel32;
using static Vanara.PInvoke.WinSpool;

namespace nddSmartPrintQueueAgent.Infra.Windows.PrinterExternalDeletion;

[SupportedOSPlatform("windows")]
[ExcludeFromCodeCoverage(Justification = "Class access windows methods that cannot be mocked")]
public class PrinterExternalDeletion : IPrinterExternalDeletion
{
    private readonly IPrinterUtils _printerUtils;
    private readonly ISystemPrinterDatabase _systemPrinterDatabase;
    private readonly ILogger<PrinterExternalDeletion> _logger;

    private SafeHPRINTERCHANGENOTIFICATION _hChange;
    private SafeHPRINTER _hPrinter;
    private nint _pNotifyOptions;
    private nint _pTypes;
    private nint _pFields;

    public PrinterExternalDeletion(IPrinterUtils printerUtils,
                                   ISystemPrinterDatabase systemPrinterDatabase,
                                   ILogger<PrinterExternalDeletion> logger)                                  
    {
        _printerUtils = printerUtils;
        _systemPrinterDatabase = systemPrinterDatabase;
        _logger = logger;
    }

    public async Task DeleteExternalPrintersByIpAsync(List<PrinterMetadata> printerList)
    {
        var localPrinterPorts = GetLocalPrinterPorts();

        var printersAllowedToDelete = GetPrintersAllowedToDelete(localPrinterPorts);

        if (printersAllowedToDelete.Count == 0)
            return;

        foreach (var printerFromCloud in printerList)
        {
            if (!printerFromCloud.ExternalQueueDeletionEnabled)
            {
                _logger.LogInformation("Printer {printerName} doesn't have externalQueueDeletionEnabled", printerFromCloud.Name);
                continue;
            }

            var localPrintersDeleted = new List<PRINTER_INFO_2>();
            foreach (var localPrinter in printersAllowedToDelete)
            {
                var localPrinterAddress = localPrinterPorts.GetValueOrDefault(localPrinter.pPortName);

                if (localPrinterAddress is null)
                {
                    continue;
                }

                GetIpAddressesFromCloudAndLocal(printerFromCloud.PortAddress,
                                              localPrinterAddress.HostAddress!,
                                              out var cloudAddress,
                                              out var localAddress);

                if (cloudAddress == IPAddress.Any || localAddress == IPAddress.Any)
                {
                    _logger.LogError("Failed to parse ip address, continuing process...");
                    continue;
                }

                if (cloudAddress.AddressFamily == localAddress.AddressFamily)
                {
                    if (cloudAddress.Equals(localAddress))
                    {
                        await DeletePrinterAsync(localPrinter.pPrinterName);
                        localPrintersDeleted.Add(localPrinter);
                    }

                    continue;
                }

                // IPV4
                if (cloudAddress.AddressFamily == AddressFamily.InterNetwork)
                {
                    var ipv4FromIpv6 = localAddress.MapToIPv4();
                    if (cloudAddress.Equals(ipv4FromIpv6))
                    {
                        await DeletePrinterAsync(localPrinter.pPrinterName);
                        localPrintersDeleted.Add(localPrinter);
                    }

                    continue;
                }

                // IPV6
                if (cloudAddress.AddressFamily == AddressFamily.InterNetworkV6)
                {
                    var ipv6FromIpv4 = localAddress.MapToIPv6();
                    if (cloudAddress.Equals(ipv6FromIpv4))
                    {
                        await DeletePrinterAsync(localPrinter.pPrinterName);
                        localPrintersDeleted.Add(localPrinter);
                    }
                }
            }

            if (localPrintersDeleted.Count > 0)
            {
                localPrintersDeleted.ForEach(x => printersAllowedToDelete.Remove(x));
                localPrintersDeleted.Clear();
            }

            _logger.LogInformation("Finished deleting external printers by ip");
        }
    } 

    public async Task DeleteAllExternalPrintersAsync()
    {
        _logger.LogInformation("Starting to delete all external printers...");

        var localPrinterPorts = GetLocalPrinterPorts();

        var printersAllowedToDelete = GetPrintersAllowedToDelete(localPrinterPorts);

        if (printersAllowedToDelete.Count == 0)
        {
            _logger.LogInformation("There are no exterrnal printers to delete, finishing...");
            return;
        }

        _logger.LogInformation("Found {printersAllowedToDeleteCount} local printers to delete.", printersAllowedToDelete.Count);

        foreach (var printer in printersAllowedToDelete)
        {
            await DeletePrinterAsync(printer.pPrinterName);
        }

        _logger.LogInformation("Finished deleting all external printers");
    }

    /// <summary>
    /// Recupera um dicionário de portas de impressoras locais.
    /// </summary>
    /// <returns>
    /// Um dicionário onde a chave é o nome da porta e o valor é o objeto correspondente <see cref="Win32TcpipPrinterPort"/>.
    /// </returns>
    private Dictionary<string, Win32TcpipPrinterPort> GetLocalPrinterPorts()
    {
        return _printerUtils.GetTcpipPrinterPorts()
                     .Where(x => x.Name is not null)
                     .ToDictionary(x => x.Name!);
    }

    /// <summary>
    /// Obtém uma lista de impressoras locais que podem ser excluídas.
    /// </summary>
    /// <param name="localPrinterPorts">Dicionário de portas de impressoras locais.</param>
    /// <returns>Lista de impressoras que podem ser excluídas.</returns>
    private List<PRINTER_INFO_2> GetPrintersAllowedToDelete(Dictionary<string, Win32TcpipPrinterPort> localPrinterPorts)
    {
        var localPrinters = EnumPrinters<PRINTER_INFO_2>();
        var localTcpPrinters = localPrinters.Where(x => localPrinterPorts.ContainsKey(x.pPortName));
        var localUsbPrinters = _printerUtils.FilterOnlyUsbPrinters(localPrinters);

        var allPrinters = localTcpPrinters.Concat(localUsbPrinters);

        return GetDeletablePrinters(allPrinters);
    }

    /// <summary>
    /// Obtém uma lista de impressoras que podem ser excluídas.
    /// </summary>
    /// <param name="localPrinters">Coleção de impressoras locais.</param>
    /// <returns>Lista de impressoras que podem ser excluídas.</returns>
    private List<PRINTER_INFO_2> GetDeletablePrinters(IEnumerable<PRINTER_INFO_2> localPrinters)
    {
        var printersToDelete = new List<PRINTER_INFO_2>();

        foreach (var localPrinter in localPrinters)
        {
            if (ShouldDeletePrinter(localPrinter))
                printersToDelete.Add(localPrinter);
        }

        return printersToDelete;
    }

    public Result<Exception, Unit> InitializePrinterMonitor()
    {
        byte[] bytes = [0x01, 0x00];
        IntPtr pFields = Marshal.AllocHGlobal(bytes.Length);
        Marshal.Copy(bytes, 0, pFields, bytes.Length);

        PRINTER_NOTIFY_OPTIONS_TYPE sNotifyType = new()
        {
            Type = NOTIFY_TYPE.PRINTER_NOTIFY_TYPE,
            Reserved0 = 0,
            Reserved1 = 0,
            Reserved2 = 0,
            pFields = pFields,
            Count = 1
        };

        nint pTypes = Marshal.AllocHGlobal(Marshal.SizeOf(sNotifyType));
        Marshal.StructureToPtr(sNotifyType, pTypes, false);

        PRINTER_NOTIFY_OPTIONS sNotifyOpts = new()
        {
            Version = 2,
            Flags = 0,
            Count = 1,
            pTypes = pTypes
        };

        nint pNotifyOptions = Marshal.AllocHGlobal(Marshal.SizeOf(sNotifyOpts));
        Marshal.StructureToPtr(sNotifyOpts, pNotifyOptions, false);
        SafeHPRINTER? hPrinter = null;

        try
        {
            PRINTER_DEFAULTS defaults = new()
            {
                DesiredAccess = (int)WinSpool.AccessRights.SERVER_ACCESS_ADMINISTER | (int)WinSpool.AccessRights.SERVER_ACCESS_ENUMERATE,
                pDatatype = null,
                pDevMode = null
            };

            if (!OpenPrinter(null, out hPrinter, defaults) || hPrinter == null || hPrinter.IsInvalid)
            {
               _logger.LogError("Error while starting to monitor printer changes. Failed to open printer. Error code {ErrorCode}",GetLastError());
                return new Exception("Failed to open printer");
            }

            SafeHPRINTERCHANGENOTIFICATION hChange = FindFirstPrinterChangeNotification(hPrinter, PRINTER_CHANGE.PRINTER_CHANGE_ADD_PRINTER, 0, pNotifyOptions);
            if (!hChange.IsNull)
            {
                _hChange = hChange;
                _hPrinter = hPrinter;
                _pNotifyOptions = pNotifyOptions;
                _pTypes = pTypes;
                _pFields = pFields;
            }
            else
            {
                _logger.LogError("Error while starting to monitor printer changes. Error code: {ErrorCode}", GetLastError());
                return new Exception("Failed to initialize");
            }
        }
        catch (Exception ex) 
        {
            _logger.LogError(ex, "Failed to initialize printer monitor");
        }       

        return Unit.Successful;
    
    }

    public bool HasPrinterChanges(CancellationToken stoppingToken)
    {
        SafeEventHandle cancelEvent = CreateEvent(null, true, false, null);
        List<ISyncHandle> syncHandles =
        [
            _hChange,
            cancelEvent
        ];
    
        if (WaitForMultipleObjects(syncHandles, false, INFINITE) == WAIT_STATUS.WAIT_OBJECT_0 + 1)
        {
            _logger.LogInformation("PrinterChangesMonitor - Cancellation requested");
            return false;
        }

        if (!FindNextPrinterChangeNotification(_hChange, out PRINTER_CHANGE change, _pNotifyOptions, out SafePRINTER_NOTIFY_INFO notifyInfo))
        {
            _logger.LogError("Printer changes monitor stopped. Error code:{ErrorCode}", GetLastError());
            return false;
        }

        try
        {
            if (change.HasFlag(PRINTER_CHANGE.PRINTER_CHANGE_ADD_PRINTER))
            {
                return true;
            }
        }
        finally
        {
            notifyInfo.Close();
        }

        return false;
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        if (_pTypes != default) Marshal.FreeHGlobal(_pTypes);
        if (_pFields != default) Marshal.FreeHGlobal(_pFields);
        if (_pNotifyOptions != default) Marshal.FreeHGlobal(_pNotifyOptions);

        _hPrinter?.Dispose();
        _hChange?.Dispose();
    }

    /// <summary>
    /// Obtém os endereços IP da impressora do orbix e local.
    /// </summary>
    /// <param name="cloudPrinterAddress">Endereço da impressora no orbix.</param>
    /// <param name="localPrinterAddress">Endereço da impressora local.</param>
    /// <param name="cloudAddress">Endereço IP da impressora no orbix.</param>
    /// <param name="localAddress">Endereço IP da impressora local.</param>
    private void GetIpAddressesFromCloudAndLocal(string cloudPrinterAddress,
                                               string localPrinterAddress,
                                               out IPAddress cloudAddress,
                                               out IPAddress localAddress)
    {
        var cloudAddressType = Uri.CheckHostName(cloudPrinterAddress);
        var localAddressType = Uri.CheckHostName(localPrinterAddress);

        cloudAddress = IPAddress.Any;
        localAddress = IPAddress.Any;

        if (cloudAddressType == UriHostNameType.Dns && localAddressType == UriHostNameType.Dns)
        {
            cloudAddress = GetIpAddressFromHostName(cloudPrinterAddress, AddressFamily.InterNetwork);
            localAddress = GetIpAddressFromHostName(localPrinterAddress, AddressFamily.InterNetwork);

            if (cloudAddress == IPAddress.Any || localAddress == IPAddress.Any)
            {
                cloudAddress = GetIpAddressFromHostName(cloudPrinterAddress, AddressFamily.InterNetworkV6);
                localAddress = GetIpAddressFromHostName(localPrinterAddress, AddressFamily.InterNetworkV6);
            }

            return;
        }

        if (cloudAddressType == UriHostNameType.Dns && localAddressType != UriHostNameType.Dns)
        {
            var addressFamily = localAddressType switch
            {
                UriHostNameType.IPv6 => AddressFamily.InterNetworkV6,
                UriHostNameType.IPv4 => AddressFamily.InterNetwork,
                _ => AddressFamily.Unknown,
            };

            cloudAddress = GetIpAddressFromHostName(cloudPrinterAddress, addressFamily);
            localAddress = GetIpAddress(localPrinterAddress);
            return;
        }

        if (cloudAddressType != UriHostNameType.Dns && localAddressType == UriHostNameType.Dns)
        {
            var addressFamily = cloudAddressType switch
            {
                UriHostNameType.IPv6 => AddressFamily.InterNetworkV6,
                UriHostNameType.IPv4 => AddressFamily.InterNetwork,
                _ => AddressFamily.Unknown,
            };

            localAddress = GetIpAddressFromHostName(localPrinterAddress, addressFamily);
            cloudAddress = GetIpAddress(cloudPrinterAddress);
            return;
        }

        if (cloudAddressType != UriHostNameType.Dns && localAddressType != UriHostNameType.Dns)
        {
            cloudAddress = GetIpAddress(cloudPrinterAddress);
            localAddress = GetIpAddress(localPrinterAddress);
        }
    }

    /// <summary>
    /// Exclui uma impressora de forma assíncrona.
    /// </summary>
    /// <param name="printerName">Nome da impressora a ser excluída.</param>
    /// <param name="tries">Número de tentativas de exclusão. O padrão é 1.</param>
    private async Task DeletePrinterAsync(string printerName, int tries = 1)
    {
        if (tries > 3)
        {
            _logger.LogInformation("Exceeded tries to delete printer {printerName}", printerName);
            return;
        }

        try
        {
            await _printerUtils.DeletePrinter(printerName);
            _logger.LogInformation("Printer {printerName} deleted successfully", printerName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to delete printer with name {printerName}... trying again", printerName);
            await DeletePrinterAsync(printerName, ++tries);
        }

    }

    private IPAddress GetIpAddress(string portAddress)
    {
        IPAddress address;

        if (IPAddress.TryParse(portAddress, out address))
        {
            return address;
        }

        _logger.LogError("Failed to parse ip address: {ipAddress}", portAddress);

        return IPAddress.Any;
    }

    private IPAddress GetIpAddressFromHostName(string hostName, AddressFamily addressFamily)
    {
        try
        {
            return Dns.GetHostEntry(hostName, addressFamily).AddressList[0];
        }
        catch
        {
            _logger.LogError("Failed to resolve hostname {hostName}", hostName);
            return IPAddress.Any;
        }
    }

    /// <summary>
    /// Verifica se a impressora deve ser excluída.
    /// </summary>
    /// <param name="printer">Informações da impressora.</param>
    /// <returns>Retorna verdadeiro se a impressora deve ser excluída, caso contrário, falso.</returns>
    private bool ShouldDeletePrinter(PRINTER_INFO_2 printer)
    {
        if (_systemPrinterDatabase.GetPrinterUniqueId(printer.pPrinterName) != Guid.Empty)
        {
            _logger.LogInformation("Printer {printerName} came from cloud portal, not deleting...", printer.pPrinterName);
            return false;
        }

        if (printer.pComment is not null &&
            (printer.pComment.Contains("#Allowed", StringComparison.OrdinalIgnoreCase) ||
            printer.pComment.Contains("#Alowed", StringComparison.OrdinalIgnoreCase)))
        {
            _logger.LogInformation("Printer {printerName} has #Allowed tag, not deleting...", printer.pPrinterName);
            return false;
        }

        return true;
    }
}
