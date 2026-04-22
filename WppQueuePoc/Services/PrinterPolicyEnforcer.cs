using System;
using System.Threading;
using System.Threading.Tasks;
using WppQueuePoc.Interop;

namespace WppQueuePoc.Services
{
    /// <summary>
    /// Monitora propriedades de impressoras e restaura valores conforme política definida.
    /// Usa Win32 API para eventos, integrando-se com PrintTicket.
    /// </summary>
    public class PrinterPolicyEnforcer : IDisposable
    {
        // Configuração de políticas (quais propriedades, valores)
        public class Policy
        {
            public bool EnforceDuplex { get; set; }
            public string? RequiredDuplexValue { get; set; }
            public bool EnforceColor { get; set; }
            public string? RequiredColorValue { get; set; }
            public bool EnforceOrientation { get; set; }
            public string? RequiredOrientationValue { get; set; }
            // Adicione mais propriedades conforme a extensão da Print Schema
        }

        public event EventHandler<string>? StatusChanged;
        public event EventHandler<string>? EnforcementLog;
        public event EventHandler<Exception>? Error;

        private readonly Policy _policy;
        private readonly WppQueuePoc.Abstractions.IPrintTicketService _printTicketService;
        private CancellationTokenSource? _cts;
        private Task? _monitorTask;
        private DateTime? _lastEnforcementUtc = null;
        private readonly TimeSpan _debounceInterval = TimeSpan.FromSeconds(3);

        public bool IsRunning { get; private set; } = false;

        // Por ora, fila fixa "Microsoft Print to PDF"; pode ser parametrizada futuramente
        private const string DefaultPrinterName = "Microsoft Print to PDF";

        public PrinterPolicyEnforcer(Policy policy, WppQueuePoc.Abstractions.IPrintTicketService printTicketService)
        {
            _policy = policy;
            _printTicketService = printTicketService;
        }

        public void Start()
        {
            if (IsRunning) return;
            _cts = new CancellationTokenSource();
            _monitorTask = Task.Run(() => MonitorLoop(_cts.Token));
            IsRunning = true;
            StatusChanged?.Invoke(this, "PrinterPolicyEnforcer started.");
        }

        public void Stop()
        {
            if (!IsRunning) return;
            _cts?.Cancel();
            IsRunning = false;
            StatusChanged?.Invoke(this, "PrinterPolicyEnforcer stopped.");
        }

        private async Task MonitorLoop(CancellationToken ct)
        {
            IntPtr hPrinter = IntPtr.Zero;
            IntPtr hNotify = IntPtr.Zero;
            try
            {
                // Abrir handle para a impressora fixa
                var defaults = new NativeMethods.PRINTER_DEFAULTS
                {
                    pDatatype = null,
                    pDevMode = IntPtr.Zero,
                    DesiredAccess = NativeMethods.PRINTER_ACCESS_USE
                };
                if (!NativeMethods.OpenPrinter(DefaultPrinterName, out hPrinter, ref defaults) || hPrinter == IntPtr.Zero)
                {
                    EnforcementLog?.Invoke(this, $"[Monitor] Falha ao abrir a impressora '{DefaultPrinterName}'.");
                    return;
                }

                // Monitorar mudanças (SET_PRINTER)
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

                EnforcementLog?.Invoke(this, $"[Monitor] Aguardando eventos em '{DefaultPrinterName}'.");

                while (!ct.IsCancellationRequested)
                {
                    int waitResult = PrinterChangeNotificationNative.WaitForSingleObject(hNotify, 2000);
                    if (waitResult == PrinterChangeNotificationNative.WAIT_OBJECT_0)
                    {
                        // Evento de mudança capturado
                        if (PrinterChangeNotificationNative.FindNextPrinterChangeNotification(hNotify, out int change, IntPtr.Zero, out IntPtr notifyInfo))
                        {
                                EnforcementLog?.Invoke(this, $"[Monitor] Mudança capturada na impressora: 0x{change:X} (provável alteração de propriedade)." );
                                try
                                {
                                    var now = DateTime.UtcNow;
                                    if (_lastEnforcementUtc.HasValue && now - _lastEnforcementUtc.Value < _debounceInterval)
                                    {
                                        EnforcementLog?.Invoke(this, $"[Enforcement] Ignorado devido ao debounce (aguardando intervalo de {_debounceInterval.TotalSeconds} segundos).");
                                    }
                                    else
                                    {
                                        var enforcement = PrintTicketEnforcementHelper.EnforceDefaultTicketPolicy(_printTicketService, DefaultPrinterName, _policy);
                                        _lastEnforcementUtc = now;
                                        EnforcementLog?.Invoke(this, $"[Enforcement] {enforcement.Details}");
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
                        continue; // Timeout normal (verifica ct e repete)
                    }
                    else
                    {
                        EnforcementLog?.Invoke(this, "[Monitor] Erro inesperado ao aguardar evento de notificação.");
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
                EnforcementLog?.Invoke(this, "[Monitor] Monitoramento finalizado e recursos liberados.");
            }
        }

        public void Dispose()
        {
            Stop();
            _cts?.Dispose();
        }
    }
}
