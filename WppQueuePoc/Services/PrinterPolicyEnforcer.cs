using System;
using System.Threading;
using System.Threading.Tasks;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Interop;

namespace WppQueuePoc.Services
{
    /// <summary>
    /// Monitor de politicas de impressao que observa alterações na fila e força restauração de valores.

    /// Este componente implementa um patron de "enforcement ativo": ele monitora eventos
    /// da impressora via Win32 API (FindFirstPrinterChangeNotification) e, quando detecta
    /// modificacoes, executa o fluxo de comparacao e restauração do PrintTicket. O objetivo
    /// e manter a impressora em conformidade com as politicas organizacionais mesmo quando
    /// usuarios ou drivers alteram configurações manualmente.
    ///
    /// O fluxo funciona assim:
    /// 1. <see cref="Start(string?)"/> abre handles Win32 e cria um loop de monitoramentoasync
    /// 2. O loop utiliza WaitForSingleObject para aguardar eventos de mudança
    /// 3. Quando evento ocorre, marca flag _flagsEnforcementPending
    /// 4. Um segundo loop (EnforcementWorkerLoop) processa a flag e aplica o enforcement
    /// 5. O debounce de 3s evita loops de re-enforcement consecutivos muito rapidos
    ///
    /// Esta abordagem separando monitoring (bloqueio de I/O) de enforcement (trabalho)
    /// permite que o componente seja responsivo sem bloquear o monitoramento de novos eventos.
    /// </summary>
    /// <remarks>
    /// Importante: O enforcement usa <see cref="IPrintTicketService"/> para ler e atualizar
    /// configuracoes. O servico precisa de permissao de gerenciamento (Manage Printers) na fila
    /// para atualizar o DefaultPrintTicket. Sem permissao, o enforcement falhara silenciosamente
    /// (apenas loga erro). O componente tambem espera que o nome da impressora
    /// seja valido e que o driver suporte as propriedades sendo monitoradas.
    /// </remarks>
    public class PrinterPolicyEnforcer : IDisposable
    {
        /// <summary>
        /// Define quais propriedades devem ser forçadas e seus valores-alvo.
        ///
        /// A classe Policy age como contratos: cada propriedade booleana (EnforceX)
        /// indica se essa dimensiao deve ser monitorada/forcada, e o valor string
        ///对应的 (RequiredXValue) e o valor que deve ser restaurado quando a politica
        ///estiver ativa.
        ///
        /// Exemplo de uso:
        /// <code>
        /// var policy = new Policy {
        ///     EnforceDuplex = true,
        ///     RequiredDuplexValue = "TwoSidedLongEdge",
        ///     EnforceColor = true,
        ///     RequiredColorValue = "Monochrome"
        /// };
        /// </code>
        /// </summary>
        public class Policy
        {
            /// <summary>
            /// Indica se o parametro Duplexing (frente e verso) deve ser forcado.
            /// Quando true, o enforcement verificara e corrigira o valor atual.
            /// </summary>
            public bool EnforceDuplex { get; set; }

            /// <summary>
            /// Valor obrigatorio para Duplexing quando EnforceDuplex e true.
            /// Valores típicos: "TwoSidedLongEdge", "TwoSidedShortEdge", "OneSided".
            /// </summary>
            public string? RequiredDuplexValue { get; set; }

            /// <summary>
            /// Indica se a cor (OutputColor) deve ser forcada.
            /// </summary>
            public bool EnforceColor { get; set; }

            /// <summary>
            /// Valor obrigatorio para OutputColor quando EnforceColor e true.
            /// Valores típicos: "Monochrome", "Color".
            /// </summary>
            public string? RequiredColorValue { get; set; }

            /// <summary>
            /// Indica se a orientação da página deve ser forcada.
            /// </summary>
            public bool EnforceOrientation { get; set; }

            /// <summary>
            /// Valor obrigatorio para PageOrientation quando EnforceOrientation.
            /// Valores típicos: "Portrait", "Landscape".
            /// </summary>
            public string? RequiredOrientationValue { get; set; }
        }

        /// <summary>
        /// Ocorre quando o estado geral do enforcement muda (start/stop).
        /// O argumento e uma mensagem descritiva do novo estado.
        /// </summary>
        public event EventHandler<string>? StatusChanged;

        /// <summary>
        /// Ocorre a cada operacao de log relevante do component.
        /// Útil para diagnostico e debugging em produção.
        /// </summary>
        public event EventHandler<string>? EnforcementLog;

        /// <summary>
        /// Ocorre quando uma exceção não tratada ocorre nos loops internos.
        /// </summary>
        public event EventHandler<Exception>? Error;

        private CancellationTokenSource? _enforcementWorkerCts;
        private Task? _enforcementWorkerTask;
        private volatile bool _flagsEnforcementPending = false;
        private readonly object _enforcementLock = new();

        private readonly Policy _policy;
        private readonly IPrintTicketService _printTicketService;
        private CancellationTokenSource? _cts;
        private Task? _monitorTask;
        private DateTime? _lastEnforcementUtc = null;
        private readonly TimeSpan _debounceInterval = TimeSpan.FromSeconds(3);

        /// <summary>
        /// Indica se o componente esta atualmente monitorando a impressora.
        /// </summary>
        public bool IsRunning { get; private set; } = false;

        private string _printerName = "Microsoft Print to PDF";

        /// <summary>
        /// Construtor que recebe a politica de enforcement e o servico de ticket.
        /// </summary>
        /// <param name="policy">Politica com parametros a serem forçados.</param>
        /// <param name="printTicketService">Instancia de IPrintTicketService para ler/atualizar tickets.</param>
        public PrinterPolicyEnforcer(Policy policy, IPrintTicketService printTicketService)
        {
            _policy = policy;
            _printTicketService = printTicketService;
        }

        /// <summary>
        /// Inicia o monitoramento de eventos da impressora.
        ///
        /// Este metodo cria dois loops asynconos: um para monitorar eventos via Win32 API
        /// (MonitorLoop) e outro para processar enforcement quando necessario (EnforcementWorkerLoop).
        /// A separacao permite que o monitoramento seja continuo e responsivo.
        /// </summary>
        /// <param name="printerName">Nome da impressora a ser monitorada. Se null, usa o padrao.</param>
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

        /// <summary>
        /// Para o monitoramento e limpa recursos alocados.
        /// </summary>
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

        /// <summary>
        /// Loop principal de monitoramento de eventos da impressora.
        ///
        /// Este metodo usa a API Win32 FindFirstPrinterChangeNotification para registrar um
        /// handle de notificacao na impressora. Entao entra em loop esperando (WaitForSingleObject)
        /// por ate 2 segundos por eventos. Quando um evento de mudanca e detectado,
        /// marca a flag _flagsEnforcementPending (protegida por lock) para que o
        /// EnforcementWorkerLoop processe a acao.
        ///
        /// O debounce de 3 segundos (via _lastEnforcementUtc) evita que multiplas mudanças
        /// consecutivas disparm mútiplos enforcements em sequencia rapida, o que poderia
        /// causar loops infinitos de re-enforcement.
        /// </summary>
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
                // Loop principal: aguarda eventos de mudança da impressora via WaitForSingleObject.
                // Cada iteração esperara ate 2 segundos. Se timeout, apenas continuam (polling passivo).
                // Se evento: processa e marca flag para enforcement se passou no debounce.
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

        /// <summary>
        /// Loop de trabalho que processa enforcement pendente.
        ///
        /// Este loop executa periodicamente (a cada 500ms) e verifica se ha flag
        /// _flagsEnforcementPending marcada. Se sim, limpa a flag e chama
        /// o helper de enforcement para comparar valores atuais com a politica
        /// e aplicar mudancas se necessario.
        ///
        /// A separacao entre MonitorLoop (bloqueante em I/O de espera de eventos)
        /// e EnforcementWorkerLoop (executando em pool de threads) permite que
        /// o componente terusponsivo a novos eventos mesmo durante o processamento
        /// de enforcement.
        /// </summary>
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
                            _printTicketService, _printerName, _policy);
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

        /// <summary>
        /// Libera recursos mantidos pelo componente.
        /// </summary>
        public void Dispose()
        {
            Stop();
            _cts?.Dispose();
        }
    }
}