using System;
using System.Threading;
using System.Threading.Tasks;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Interop;

namespace WppQueuePoc.Services
{
    /// <summary>
    /// Monitora propriedades de impressoras e restaura valores conforme política definida.
    /// Usa Win32 API para eventos, integrando-se com PrintTicket.
    /// </summary>
    public class PrinterPolicyEnforcer : IDisposable
    {
        /// <summary>
        /// Configuração de políticas (quais propriedades, valores).
        /// </summary>
        public class Policy
        {
            public bool EnforceDuplex { get; set; }
            public string? RequiredDuplexValue { get; set; }
            public bool EnforceColor { get; set; }
            public string? RequiredColorValue { get; set; }
            public bool EnforceOrientation { get; set; }
            public string? RequiredOrientationValue { get; set; }
        }

        public event EventHandler<string>? StatusChanged;
        public event EventHandler<string>? EnforcementLog;
        public event EventHandler<Exception>? Error;

        private CancellationTokenSource? _enforcementWorkerCts;
        private Task? _enforcementWorkerTask;
        private volatile bool _flagsEnforcementPending = false;
        private readonly object _enforcementLock = new();

        private readonly IPrintTicketService _printTicketService;
        private CancellationTokenSource? _cts;
        private Task? _monitorTask;
        private DateTime? _lastEnforcementUtc = null;
        private readonly TimeSpan _debounceInterval = TimeSpan.FromSeconds(3);

        public bool IsRunning { get; private set; } = false;

        private string _printerName = "Microsoft Print to PDF";

        public PrinterPolicyEnforcer(Policy policy, IPrintTicketService printTicketService)
        {
            _printTicketService = printTicketService;
        }

        public void Start(string? printerName)
        {
            if (IsRunning) return;
            if (printerName != null) _printerName = printerName;

            _cts = new CancellationTokenSource();
            _monitorTask = Task.Run(() => MonitorLoop(_cts.Token));
            _enforcementWorkerCts = new CancellationTokenSource();
            _enforcementWorkerTask = Task.Run(() => EnforcementWorkerLoop(_enforcementWorkerCts.Token));
            IsRunning = true;
            StatusChanged?.Invoke(this, "PrinterPolicyEnforcer started.");
        }

        public void Stop()
        {
            if (!IsRunning) return;
            _cts?.Cancel();
            IsRunning = false;
            StatusChanged?.Invoke(this, "PrinterPolicyEnforcer stopped.");

            _enforcementWorkerCts?.Cancel();
            try { _enforcementWorkerTask?.Wait(1000); } catch { }
            _enforcementWorkerCts?.Dispose();
            _enforcementWorkerTask = null;
            _enforcementWorkerCts = null;
        }

        private async Task MonitorLoop(CancellationToken ct)
        {
            IntPtr hPrinter = IntPtr.Zero;
            IntPtr hNotify = IntPtr.Zero;
            try
            {
                var defaults = new NativeMethods.PRINTER_DEFAULTS
                {
                    pDatatype = null,
                    pDevMode = IntPtr.Zero,
                    DesiredAccess = NativeMethods.PRINTER_ACCESS_USE
                };
                if (!NativeMethods.OpenPrinter(_printerName, out hPrinter, ref defaults) || hPrinter == IntPtr.Zero)
                {
                    EnforcementLog?.Invoke(this, $"[Monitor] Falha ao abrir a impressora '{_printerName}'.");
                    return;
                }

                hNotify = PrinterChangeNotificationNative.FindFirstPrinterChangeNotification(
                    hPrinter,
                    PrinterChangeNotificationNative.PRINTER_CHANGE_SET_PRINTER,
                    0,
                    IntPtr.Zero);

                if (hNotify == IntPtr.Zero || hNotify == new IntPtr(-1))
                {
                    EnforcementLog?.Invoke(this, "[Monitor] Falha ao criar handle de notificação para impressora.");
                    return;
                }

                EnforcementLog?.Invoke(this, $"[Monitor] Aguardando eventos em '{_printerName}'.");
                while (!ct.IsCancellationRequested)
                {
                    int waitResult = PrinterChangeNotificationNative.WaitForSingleObject(hNotify, 2000);
                    if (waitResult == PrinterChangeNotificationNative.WAIT_OBJECT_0)
                    {
                        if (PrinterChangeNotificationNative.FindNextPrinterChangeNotification(hNotify, out int change, IntPtr.Zero, out IntPtr notifyInfo))
                        {
                            EnforcementLog?.Invoke(this, $"[Monitor] Mudança capturada na impressora: 0x{change:X}.");
                            try
                            {
                                var now = DateTime.UtcNow;
                                if (_lastEnforcementUtc.HasValue && now - _lastEnforcementUtc.Value < _debounceInterval)
                                {
                                    EnforcementLog?.Invoke(this, $"[Enforcement] Ignorado (debounce de {_debounceInterval.TotalSeconds}s).");
                                }
                                else
                                {
                                    lock (_enforcementLock)
                                    {
                                        _flagsEnforcementPending = true;
                                    }
                                    EnforcementLog?.Invoke(this, "[Monitor] Enforcement pendente.");
                                }
                            }
                            catch (Exception enfEx)
                            {
                                EnforcementLog?.Invoke(this, $"[Enforcement][ERROR] {enfEx.Message}");
                            }
                        }
                        else
                        {
                            EnforcementLog?.Invoke(this, "[Monitor] Erro ao processar notificação de mudança.");
                        }
                    }
                    else if (waitResult == PrinterChangeNotificationNative.WAIT_TIMEOUT)
                    {
                        continue;
                    }
                    else
                    {
                        EnforcementLog?.Invoke(this, "[Monitor] Erro inesperado ao aguardar evento.");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Error?.Invoke(this, ex);
            }
            finally
            {
                if (hNotify != IntPtr.Zero && hNotify != new IntPtr(-1))
                {
                    PrinterChangeNotificationNative.FindClosePrinterChangeNotification(hNotify);
                }
                if (hPrinter != IntPtr.Zero)
                {
                    NativeMethods.ClosePrinter(hPrinter);
                }
                EnforcementLog?.Invoke(this, "[Monitor] Monitoramento finalizado.");
            }
        }

        private async Task EnforcementWorkerLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                bool shouldEnforce = false;
                lock (_enforcementLock)
                {
                    if (_flagsEnforcementPending)
                    {
                        _flagsEnforcementPending = false;
                        shouldEnforce = true;
                    }
                }
                if (shouldEnforce)
                {
                    try
                    {
                        var enforcement = PrintTicketEnforcementHelper.EnforceDefaultTicketPolicy(
                            _printTicketService, _printerName);
                        _lastEnforcementUtc = DateTime.UtcNow;
                        EnforcementLog?.Invoke(this, "[Enforcement] " + enforcement.Details);
                    }
                    catch (Exception ex)
                    {
                        EnforcementLog?.Invoke(this, "[Enforcement][ERRO] " + ex.Message);
                    }
                }
                await Task.Delay(500, ct);
            }
        }

        public void Dispose()
        {
            Stop();
            _cts?.Dispose();
        }
    }
}