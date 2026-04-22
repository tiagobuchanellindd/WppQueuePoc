using NDDPrintAgent.Domain.ExternalServices.WinSpool;
using NDDPrintAgent.Domain.Features.JobProcessing;
using NDDPrintAgent.Domain.Features.PrinterMonitoring;
using NDDPrintAgent.Domain.Features.PrinterMonitoring.Sync;
using NDDPrintAgent.Domain.Features.PrinterMonitoring.WinSpool;
using NDDPrintAgent.Domain.Models;
using NDDPrintAgent.Infra.ExternalServices.WinSpool.PrintProcessors;
using Serilog;
using System.Runtime.InteropServices;
using Vanara.PInvoke;
using static Vanara.PInvoke.Kernel32;
using static Vanara.PInvoke.WinSpool;

namespace NDDPrintAgent.Infra.ExternalServices.WinSpool.Monitors;

#pragma warning disable CA1416 // Validate platform compatibility

public enum PrinterMonitorTypes
{
    AddOrEdit = 0,
    Remove = 1
}

public class WinPrinterChangesMonitor
{
    private readonly Dictionary<string, string> printerChecksums = [];
    private readonly object _removalLock = new();
    private HashSet<string> previousPrinters = [];

    private readonly IWinJobNotificationProcessor winJobNotificationProcessor;
    private readonly IWinSpoolFailureMonitor winSpoolMonitorFailureMonitor;
    private readonly IPrinterLister printerLister;
    private readonly IPartialPrinterSync partialPrinterSync;
    private readonly IRemovedPrintersSync removedPrintersSync;
    private readonly IWinPrintProcessorInstaller winPrintProcessorInstaller;
    private readonly PrinterMonitorTypes type;

    public WinPrinterChangesMonitor(
        IWinJobNotificationProcessor winJobNotificationProcessor,
        IWinSpoolFailureMonitor winSpoolMonitorFailureMonitor,
        IPrinterLister printerLister,
        IPartialPrinterSync partialPrinterSync,
        IRemovedPrintersSync removedPrintersSync,
        IWinPrintProcessorInstaller winPrintProcessorInstaller,
        PrinterMonitorTypes type)
    {
        this.winJobNotificationProcessor = winJobNotificationProcessor;
        this.winSpoolMonitorFailureMonitor = winSpoolMonitorFailureMonitor;
        this.printerLister = printerLister;
        this.partialPrinterSync = partialPrinterSync;
        this.removedPrintersSync = removedPrintersSync;
        this.winPrintProcessorInstaller = winPrintProcessorInstaller;
        this.type = type;

        if (type == PrinterMonitorTypes.AddOrEdit)
        {
            foreach (PRINTER_INFO_2 printer in EnumPrinters<PRINTER_INFO_2>())
                printerChecksums[printer.pPrinterName] = printer.GetCheckSum();
        }
        else if (type == PrinterMonitorTypes.Remove)
        {
            foreach (PRINTER_INFO_1 printer in EnumPrinters<PRINTER_INFO_1>())
                previousPrinters.Add(printer.pName);
        }
    }

    public void DoMonitor(CancellationToken cToken)
    {
        Log.Debug("Starting to monitor printer changes");

        ushort[] printerFields =
        {
            (ushort)PRINTER_NOTIFY_FIELD.PRINTER_NOTIFY_FIELD_PRINTER_NAME,
            (ushort)PRINTER_NOTIFY_FIELD.PRINTER_NOTIFY_FIELD_STATUS,
        };

        nint pFields = Marshal.AllocHGlobal(sizeof(ushort) * printerFields.Length);

        for (int i = 0; i < printerFields.Length; i++)
        {
            Marshal.WriteInt16(pFields, i * sizeof(ushort), (short)printerFields[i]);
        }

        ushort[] jobFields =
        {
            (ushort)JOB_NOTIFY_FIELD.JOB_NOTIFY_FIELD_REMOTE_JOB_ID,
            (ushort)JOB_NOTIFY_FIELD.JOB_NOTIFY_FIELD_STATUS
        };

        nint pJobFields = Marshal.AllocHGlobal(sizeof(ushort) * jobFields.Length);
        for (int i = 0; i < jobFields.Length; i++)
        {
            Marshal.WriteInt16(pJobFields, i * sizeof(ushort), (short)jobFields[i]);
        }

        PRINTER_NOTIFY_OPTIONS_TYPE[] notifyTypes = [
            new PRINTER_NOTIFY_OPTIONS_TYPE
            {
                Type = NOTIFY_TYPE.PRINTER_NOTIFY_TYPE,
                Reserved0 = 0,
                Reserved1 = 0,
                Reserved2 = 0,
                pFields = pFields,
                Count = (uint)printerFields.Length
            },

            new PRINTER_NOTIFY_OPTIONS_TYPE
            {
                Type = NOTIFY_TYPE.JOB_NOTIFY_TYPE,
                Reserved0 = 0,
                Reserved1 = 0,
                Reserved2 = 0,
                pFields = pJobFields,
                Count = (uint)jobFields.Length
            }
        ];

        int printNotifyOptionsTypeSize = Marshal.SizeOf<PRINTER_NOTIFY_OPTIONS_TYPE>();
        nint pTypes = Marshal.AllocHGlobal(printNotifyOptionsTypeSize * notifyTypes.Length);

        for (int i = 0; i < notifyTypes.Length; i++)
        {
            Marshal.StructureToPtr(notifyTypes[i], pTypes + i * printNotifyOptionsTypeSize, false);
        }

        PRINTER_NOTIFY_OPTIONS sNotifyOpts = new()
        {
            Version = 2,
            Flags = 0,
            Count = (uint)notifyTypes.Length,
            pTypes = pTypes
        };

        nint pNotifyOptions = Marshal.AllocHGlobal(Marshal.SizeOf(sNotifyOpts));
        Marshal.StructureToPtr(sNotifyOpts, pNotifyOptions, false);
        SafeHPRINTER? hPrinter = null;

        try
        {
            PRINTER_DEFAULTS defaults = new()
            {
                DesiredAccess = (int)(WinPrintServerPermissions.Use | WinPrintServerPermissions.Administer),
                pDatatype = null,
                pDevMode = null
            };

            if (!OpenPrinter(null, out hPrinter, defaults) || hPrinter == null || hPrinter.IsInvalid)
            {
                Log.Error("Error while starting to monitor printer changes. Failed to open printer. Error code: {0}", GetLastError());
                return;
            }

            PRINTER_CHANGE flags = 0;
            if (type == PrinterMonitorTypes.AddOrEdit)
            {
                flags = PRINTER_CHANGE.PRINTER_CHANGE_SET_PRINTER | PRINTER_CHANGE.PRINTER_CHANGE_ADD_PRINTER | PRINTER_CHANGE.PRINTER_CHANGE_ADD_JOB
                    | PRINTER_CHANGE.PRINTER_CHANGE_SET_JOB | PRINTER_CHANGE.PRINTER_CHANGE_DELETE_JOB;
            }
            else if (type == PrinterMonitorTypes.Remove)
            {
                flags = PRINTER_CHANGE.PRINTER_CHANGE_DELETE_PRINTER;
            }

            SafeHPRINTERCHANGENOTIFICATION hChange = FindFirstPrinterChangeNotification(hPrinter, flags, 0, pNotifyOptions);
            if (!hChange.IsNull)
            {
                MonitorLoop(hChange, pNotifyOptions, cToken);
            }
            else
            {
                Log.Error("Error while starting to monitor printer changes. Error code: {0}", GetLastError());
            }
        }
        finally
        {
            Marshal.FreeHGlobal(pTypes);
            Marshal.FreeHGlobal(pNotifyOptions);
            Marshal.FreeHGlobal(pFields);
            Marshal.FreeHGlobal(pJobFields);
            hPrinter?.Dispose();
        }
    }

    void MonitorLoop(SafeHPRINTERCHANGENOTIFICATION hChange, nint pNotifyOptions, CancellationToken cToken)
    {
        using SafeEventHandle cancelEvent = CreateEvent(null, true, false, null);

        List<ISyncHandle> syncHandles =
        [
            hChange,
            cancelEvent,
            winSpoolMonitorFailureMonitor.GetNotificationHandle() //Esse evento é disparado quando o Spool morre
        ];

        cToken.Register(() => SetEvent(cancelEvent));

        while (true)
        {
            WAIT_STATUS waitResult = WaitForMultipleObjects(syncHandles, false, INFINITE);

            if (waitResult == WAIT_STATUS.WAIT_OBJECT_0 + 1)
            {
                Log.Debug("PrinterChangesMonitor - Cancellation requested");
                break;
            }

            if (waitResult == WAIT_STATUS.WAIT_OBJECT_0 + 2)
            {
                Log.Warning("PrinterChangesMonitor - Spooler failure detected, stopping monitor");
                break;
            }

            if (!FindNextPrinterChangeNotification(hChange, out PRINTER_CHANGE change, pNotifyOptions, out SafePRINTER_NOTIFY_INFO notifyInfo))
            {
                Log.Error("Printer changes monitor stopped. Error code: {0}", GetLastError());
                break;
            }

            try
            {
                if (!notifyInfo.IsNull && change.HasFlag(PRINTER_CHANGE.PRINTER_CHANGE_ADD_JOB))
                {
                    PRINTER_NOTIFY_INFO notifyData = notifyInfo;

                    if (notifyData.Count > 0)
                    {
                        foreach (var data in notifyData.aData)
                        {
                            if (data.Type != NOTIFY_TYPE.JOB_NOTIFY_TYPE)
                                continue;

                            uint jobID = data.Id;
                            if (jobID == 0)
                                continue;

                            Log.Debug("The print job {JobID} was captured by the Print Notifier.", jobID);
                            _ = Task.Run(async () =>
                            {
                                try
                                {
                                    winJobNotificationProcessor.PauseJob(jobID);
                                    await winJobNotificationProcessor.ProcessNewJob(jobID, cToken);
                                    winJobNotificationProcessor.ResumeJob(jobID);
                                }
                                catch (Exception ex)
                                {
                                    Log.Error(ex, "Error processing print job {JobID}", jobID);
                                }
                            });
                        }
                    }
                    else
                    {
                        Log.Warning("PRINTER_CHANGE_ADD_JOB received but PRINTER_NOTIFY_INFO.aData is empty (Count=0)");
                    }
                }
                else if (type == PrinterMonitorTypes.AddOrEdit &&
                    (change.HasFlag(PRINTER_CHANGE.PRINTER_CHANGE_ADD_PRINTER) || change.HasFlag(PRINTER_CHANGE.PRINTER_CHANGE_SET_PRINTER)))
                {
                    _ = Task.Run(async () =>
                    {
                        try { await CalculatePrinterChanges(cToken); }
                        catch (Exception ex) { Log.Error(ex, "Error processing printer changes"); }
                    });
                }
                else if (type == PrinterMonitorTypes.Remove && change.HasFlag(PRINTER_CHANGE.PRINTER_CHANGE_DELETE_PRINTER))
                {
                    _ = Task.Run(async () =>
                    {
                        try { await CalculatePrinterRemoval(cToken); }
                        catch (Exception ex) { Log.Error(ex, "Error processing printer removal"); }
                    });
                }
            }
            finally
            {
                notifyInfo.Close();
            }
        }
    }

    private async Task CalculatePrinterChanges(CancellationToken cToken)
    {
        try
        {
            List<PRINTER_INFO_2> changedPrinters = [];
            List<PORT_INFO_2> localPorts = EnumPorts<PORT_INFO_2>(null).ToList();
            List<MONITOR_INFO_2> monitors = EnumMonitors<MONITOR_INFO_2>().ToList();

            lock (printerChecksums)
            {
                foreach (PRINTER_INFO_2 printer in EnumPrinters<PRINTER_INFO_2>())
                {
                    string newChecksum = printer.GetCheckSum();
                    if (!printerChecksums.TryGetValue(printer.pPrinterName, out string? previousChecksum) || previousChecksum != newChecksum)
                    {
                        changedPrinters.Add(printer);
                        printerChecksums[printer.pPrinterName] = newChecksum;
                    }
                }
            }

            if (changedPrinters.Count == 0)
                return;

            List<EnumeratedPrinter> enumeratedPrinters = [];
            await Parallel.ForEachAsync(changedPrinters, async (printer, _) =>
            {
                if (printerLister.TryLoadPrinter(printer, localPorts, monitors, out EnumeratedPrinter? enumeratedPrinter))
                {
                    lock (enumeratedPrinters)
                        enumeratedPrinters.Add(enumeratedPrinter);
                }
            });

            await partialPrinterSync.PartialSync(enumeratedPrinters, cToken);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error calculation printer changes");
        }
    }

    private async Task CalculatePrinterRemoval(CancellationToken cToken)
    {
        try
        {
            HashSet<string> latestPrinters = [];
            foreach (var printer in EnumPrinters<PRINTER_INFO_1>())
            {
                latestPrinters.Add(printer.pName);
            }

            List<string> removedPrinters = [];
            lock (_removalLock)
            {
                removedPrinters.AddRange(previousPrinters.Where(p => !latestPrinters.Contains(p)));
                previousPrinters = latestPrinters;
            }

            await removedPrintersSync.RemovePrinter(removedPrinters, cToken);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error calculating printer removal");
        }
    }
}

#pragma warning restore CA1416 // Validate platform compatibility