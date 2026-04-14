using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Interop;
using WppQueuePoc.Models;

namespace WppQueuePoc.Services;

/// <summary>
/// Serviço de administração de impressão baseado em Winspool (API nativa do Windows).
///
/// Esta classe centraliza os casos de uso da POC para filas de impressão:
/// criação, atualização, exclusão, enumeração de recursos (filas, portas,
/// drivers, processadores e datatypes) e inspeção heurística de aderência a WPP.
///
/// Em termos técnicos, o serviço encapsula chamadas nativas via P/Invoke,
/// gerencia buffers não gerenciados, converte estruturas Win32 para modelos
/// de domínio e padroniza erros em exceções mais explicativas para o aplicativo.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class PrintSpoolerService(IWppStatusProvider wppStatusProvider) : IPrintSpoolerService
{
    /// <summary>
    /// Tenta criar uma porta WSD no spooler usando o canal administrativo XcvMonitor.
    ///
    /// O método primeiro verifica idempotência (porta já existente) e, se necessário,
    /// envia o comando nativo <c>AddPort</c> ao monitor "WSD Port". Mesmo quando a
    /// chamada Win32 retorna sucesso, também valida o código de status devolvido pelo
    /// monitor para distinguir sucesso real de rejeições funcionais.
    ///
    /// Em ambientes onde a criação direta não é suportada (por exemplo,
    /// <c>ERROR_NOT_SUPPORTED</c>), o método lança erro orientando o uso de fluxo de
    /// descoberta/reuso de portas existentes.
    /// </summary>
    public void AddWsdPort(string portName)
    {
        // Validação de idempotência: evita tentar criar uma porta que já existe.
        if (ListPorts().Any(p => string.Equals(p, portName, StringComparison.OrdinalIgnoreCase)))
        {
            Console.WriteLine($"Port '{portName}' already exists.");
            return;
        }

        var defaults = new NativeMethods.PRINTER_DEFAULTS
        {
            pDatatype = null,
            pDevMode = IntPtr.Zero,
            DesiredAccess = NativeMethods.SERVER_ACCESS_ADMINISTER
        };

        // Validação de acesso nativo: sem handle Xcv não há como enviar AddPort.
        if (!NativeMethods.OpenPrinter(",XcvMonitor WSD Port", out var hXcv, ref defaults))
        {
            ThrowLastWin32("OpenPrinter for XcvMonitor WSD Port failed.");
        }

        try
        {
            var payload = Encoding.Unicode.GetBytes(portName + '\0');
            // Validação da chamada nativa: erro aqui indica falha do canal XcvData.
            if (!NativeMethods.XcvData(
                    hXcv,
                    "AddPort",
                    payload,
                    payload.Length,
                    IntPtr.Zero,
                    0,
                    out _,
                    out var status))
            {
                ThrowLastWin32("XcvData(AddPort) call failed.");
            }

            // Validação de status de operação: mesmo com chamada bem-sucedida, o monitor pode rejeitar.
            if (status != NativeMethods.ERROR_SUCCESS)
            {
                // Validação de cenário conhecido: ambiente sem suporte direto ao AddPort.
                if (status == NativeMethods.ERROR_NOT_SUPPORTED)
                {
                    throw new Win32Exception(
                        (int)status,
                        "XcvData(AddPort) returned ERROR_NOT_SUPPORTED for WSD monitor. This environment likely requires device-discovery flow for WSD port creation. Use 'list-ports' and reuse an existing WSD port.");
                }

                throw new Win32Exception((int)status, $"XcvData(AddPort) returned status {status}.");
            }

            Console.WriteLine($"XcvData(AddPort) dwStatus={status}.");
        }
        finally
        {
            NativeMethods.ClosePrinter(hXcv);
        }
    }

    /// <summary>
    /// Cria uma nova fila de impressão no spooler com os parâmetros informados.
    ///
    /// Monta uma estrutura <c>PRINTER_INFO_2</c> com os metadados da fila
    /// (nome, porta, driver, processador, datatype e propriedades descritivas),
    /// aloca memória não gerenciada para interoperabilidade e chama <c>AddPrinter</c>.
    ///
    /// Se o spooler recusar a operação, converte o último erro Win32 em exceção.
    /// Ao final, garante liberação de handle e memória alocada, mesmo em caso de falha.
    /// </summary>
    public void CreateQueue(
        string queueName,
        string driverName,
        string portName,
        string printProcessor,
        string dataType,
        string comment,
        string location)
    {
        var info = new NativeMethods.PRINTER_INFO_2
        {
            pServerName = null,
            pPrinterName = queueName,
            pShareName = null,
            pPortName = portName,
            pDriverName = driverName,
            pComment = comment,
            pLocation = location,
            pDevMode = IntPtr.Zero,
            pSepFile = null,
            pPrintProcessor = printProcessor,
            pDatatype = dataType,
            pParameters = null,
            pSecurityDescriptor = IntPtr.Zero,
            Attributes = NativeMethods.PRINTER_ATTRIBUTE_QUEUED,
            Priority = 1,
            DefaultPriority = 1,
            StartTime = 0,
            UntilTime = 0,
            Status = 0,
            cJobs = 0,
            AveragePPM = 0
        };

        var infoPtr = IntPtr.Zero;
        try
        {
            infoPtr = Marshal.AllocHGlobal(Marshal.SizeOf<NativeMethods.PRINTER_INFO_2>());
            Marshal.StructureToPtr(info, infoPtr, false);
            var handle = NativeMethods.AddPrinter(null, 2, infoPtr);
            // Validação de criação: handle nulo significa que o spooler recusou a nova fila.
            if (handle == IntPtr.Zero)
            {
                ThrowLastWin32("AddPrinter failed.");
            }

            NativeMethods.ClosePrinter(handle);
        }
        finally
        {
            if (infoPtr != IntPtr.Zero)
            {
                Marshal.DestroyStructure<NativeMethods.PRINTER_INFO_2>(infoPtr);
                Marshal.FreeHGlobal(infoPtr);
            }
        }
    }

    /// <summary>
    /// Lista filas locais e conexões de impressora com campos úteis para operação.
    ///
    /// Internamente enumera dados no nível <c>PRINTER_INFO_2</c>, projeta para o
    /// modelo <see cref="QueueInfo"/>, normaliza valores nulos para string vazia
    /// e ordena alfabeticamente por nome para facilitar consumo em CLI/UI.
    /// </summary>
    public IReadOnlyList<QueueInfo> ListQueues()
    {
        var printers = EnumeratePrinterInfo2();
        return printers
            .Select(p => new QueueInfo(
                p.pPrinterName ?? string.Empty,
                p.pPortName ?? string.Empty,
                p.pDriverName ?? string.Empty,
                (p.Attributes & NativeMethods.PRINTER_ATTRIBUTE_SHARED) != 0,
                p.pComment ?? string.Empty,
                p.pLocation ?? string.Empty))
            .OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Lista portas de impressão registradas no host.
    ///
    /// Executa o padrão Win32 de duas chamadas: primeiro consulta o tamanho de buffer
    /// necessário, depois lê o bloco com <c>PORT_INFO_1</c>. Em seguida, converte cada
    /// item para nome de porta, filtra entradas inválidas e retorna a coleção ordenada.
    /// </summary>
    public IReadOnlyList<string> ListPorts()
    {
        // Validação de tamanho de buffer: primeira chamada descobre bytes necessários.
        if (!NativeMethods.EnumPorts(
                null,
                1,
                IntPtr.Zero,
                0,
                out var needed,
                out _))
        {
            var error = Marshal.GetLastWin32Error();
            // Validação de erro esperado: fora desses códigos, trata como falha real.
            if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
            {
                throw new Win32Exception(error, "EnumPorts buffer query failed.");
            }
        }

        // Validação de conteúdo: sem bytes necessários, não há portas para retornar.
        if (needed == 0)
        {
            return [];
        }

        var buffer = Marshal.AllocHGlobal((int)needed);
        try
        {
            // Validação da leitura efetiva: segunda chamada precisa preencher o buffer.
            if (!NativeMethods.EnumPorts(
                    null,
                    1,
                    buffer,
                    needed,
                    out _,
                    out var returned))
            {
                ThrowLastWin32("EnumPorts failed.");
            }

            var structSize = Marshal.SizeOf<NativeMethods.PORT_INFO_1>();
            var ports = new List<string>((int)returned);
            for (var i = 0; i < returned; i++)
            {
                var ptr = IntPtr.Add(buffer, i * structSize);
                var info = Marshal.PtrToStructure<NativeMethods.PORT_INFO_1>(ptr);
                // Validação de dado: ignora nomes nulos/vazios para não poluir saída.
                if (!string.IsNullOrWhiteSpace(info.pName))
                {
                    ports.Add(info.pName);
                }
            }

            ports.Sort(StringComparer.OrdinalIgnoreCase);
            return ports;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Lista drivers de impressão instalados no sistema.
    ///
    /// Usa <c>EnumPrinterDrivers</c> (nível 2) para obter <c>DRIVER_INFO_2</c>,
    /// extrai os nomes válidos e devolve uma lista ordenada, pronta para seleção
    /// em fluxos de criação/atualização de fila.
    /// </summary>
    public IReadOnlyList<string> ListDrivers()
    {
        // Validação de tamanho de buffer para enumeração de drivers.
        if (!NativeMethods.EnumPrinterDrivers(
                null,
                null,
                2,
                IntPtr.Zero,
                0,
                out var needed,
                out _))
        {
            var error = Marshal.GetLastWin32Error();
            // Validação de erro esperado na fase de descoberta de buffer.
            if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
            {
                throw new Win32Exception(error, "EnumPrinterDrivers buffer query failed.");
            }
        }

        // Validação de conteúdo: host sem drivers retornáveis.
        if (needed == 0)
        {
            return [];
        }

        var buffer = Marshal.AllocHGlobal((int)needed);
        try
        {
            // Validação da leitura efetiva no buffer alocado.
            if (!NativeMethods.EnumPrinterDrivers(
                    null,
                    null,
                    2,
                    buffer,
                    needed,
                    out _,
                    out var returned))
            {
                ThrowLastWin32("EnumPrinterDrivers failed.");
            }

            var structSize = Marshal.SizeOf<NativeMethods.DRIVER_INFO_2>();
            var drivers = new List<string>((int)returned);
            for (var i = 0; i < returned; i++)
            {
                var ptr = IntPtr.Add(buffer, i * structSize);
                var info = Marshal.PtrToStructure<NativeMethods.DRIVER_INFO_2>(ptr);
                // Validação de dado: adiciona apenas nomes válidos.
                if (!string.IsNullOrWhiteSpace(info.pName))
                {
                    drivers.Add(info.pName);
                }
            }

            drivers.Sort(StringComparer.OrdinalIgnoreCase);
            return drivers;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Lista processadores de impressão registrados no spooler.
    ///
    /// O método enumera estruturas <c>PRINTPROCESSOR_INFO_1</c>, extrai apenas
    /// nomes não vazios e organiza o retorno em ordem alfabética, servindo de base
    /// para descobrir quais datatypes cada processador suporta.
    /// </summary>
    public IReadOnlyList<string> ListPrintProcessors()
    {
        // Validação de tamanho de buffer para enumeração de processadores.
        if (!NativeMethods.EnumPrintProcessors(
                null,
                null,
                1,
                IntPtr.Zero,
                0,
                out var needed,
                out _))
        {
            var error = Marshal.GetLastWin32Error();
            // Validação de erro esperado na etapa de descoberta de buffer.
            if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
            {
                throw new Win32Exception(error, "EnumPrintProcessors buffer query failed.");
            }
        }

        // Validação de conteúdo: nenhum processador retornado.
        if (needed == 0)
        {
            return [];
        }

        var buffer = Marshal.AllocHGlobal((int)needed);
        try
        {
            // Validação da leitura efetiva no buffer alocado.
            if (!NativeMethods.EnumPrintProcessors(
                    null,
                    null,
                    1,
                    buffer,
                    needed,
                    out _,
                    out var returned))
            {
                ThrowLastWin32("EnumPrintProcessors failed.");
            }

            // O buffer já foi preenchido pelo EnumPrintProcessors com um "array" nativo de structs.
            // Aqui calculamos o tamanho de cada item para navegar posição a posição no bloco de memória.
            // Em cada posição, convertemos os bytes para PRINTPROCESSOR_INFO_1 e extraímos o nome.
            var structSize = Marshal.SizeOf<NativeMethods.PRINTPROCESSOR_INFO_1>();
            var processors = new List<string>((int)returned);
            for (var i = 0; i < returned; i++)
            {
                var ptr = IntPtr.Add(buffer, i * structSize);
                var info = Marshal.PtrToStructure<NativeMethods.PRINTPROCESSOR_INFO_1>(ptr);
                // Validação de dado: ignora entradas sem nome.
                if (!string.IsNullOrWhiteSpace(info.pName))
                {
                    processors.Add(info.pName);
                }
            }

            processors.Sort(StringComparer.OrdinalIgnoreCase);
            return processors;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Lista os datatypes disponíveis para um processador de impressão específico.
    ///
    /// Valida a entrada, consulta o spooler com
    /// <c>EnumPrintProcessorDatatypes</c> (nível 1), converte as estruturas
    /// <c>DATATYPES_INFO_1</c> para strings e retorna os nomes ordenados.
    ///
    /// Este método é útil para montar combinações válidas de
    /// processador + datatype durante criação ou atualização de filas.
    /// </summary>
    public IReadOnlyList<string> ListDataTypes(string printProcessor)
    {
        // Validação de entrada: datatype depende de processador explicitamente informado.
        if (string.IsNullOrWhiteSpace(printProcessor))
        {
            throw new InvalidOperationException("Print processor name is required.");
        }

        // Validação de tamanho de buffer para enumeração dos datatypes.
        if (!NativeMethods.EnumPrintProcessorDatatypes(
                null,
                printProcessor,
                1,
                IntPtr.Zero,
                0,
                out var needed,
                out _))
        {
            var error = Marshal.GetLastWin32Error();
            // Validação de erro esperado na fase de descoberta de buffer.
            if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
            {
                throw new Win32Exception(error, $"EnumPrintProcessorDatatypes buffer query failed for '{printProcessor}'.");
            }
        }

        // Validação de conteúdo: processador sem datatypes retornáveis.
        if (needed == 0)
        {
            return [];
        }

        var buffer = Marshal.AllocHGlobal((int)needed);
        try
        {
            // Validação da leitura efetiva no buffer alocado.
            if (!NativeMethods.EnumPrintProcessorDatatypes(
                    null,
                    printProcessor,
                    1,
                    buffer,
                    needed,
                    out _,
                    out var returned))
            {
                ThrowLastWin32($"EnumPrintProcessorDatatypes failed for '{printProcessor}'.");
            }

            var structSize = Marshal.SizeOf<NativeMethods.DATATYPES_INFO_1>();
            var dataTypes = new List<string>((int)returned);
            for (var i = 0; i < returned; i++)
            {
                var ptr = IntPtr.Add(buffer, i * structSize);
                var info = Marshal.PtrToStructure<NativeMethods.DATATYPES_INFO_1>(ptr);
                // Validação de dado: adiciona apenas nomes de datatype válidos.
                if (!string.IsNullOrWhiteSpace(info.pName))
                {
                    dataTypes.Add(info.pName);
                }
            }

            dataTypes.Sort(StringComparer.OrdinalIgnoreCase);
            return dataTypes;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Atualiza propriedades de uma fila existente de forma parcial.
    ///
    /// Lê o estado atual da fila, aplica somente os campos informados
    /// (rename, driver, porta, comentário e localização), valida se há mudanças
    /// reais e persiste no spooler via <c>SetPrinter</c> com acesso administrativo.
    ///
    /// Também zera ponteiros sensíveis da estrutura para evitar reutilização indevida
    /// de dados nativos e garante liberação de recursos em todas as saídas.
    /// </summary>
    public void UpdateQueue(string queueName, string? newQueueName, string? newDriverName, string? newPortName, string? comment, string? location)
    {
        var info = GetQueueInfo(queueName);
        info.pDevMode = IntPtr.Zero;
        info.pSecurityDescriptor = IntPtr.Zero;
        var hasChanges = false;

        if (newQueueName is not null)
        {
            info.pPrinterName = newQueueName;
            hasChanges = true;
        }

        if (newDriverName is not null)
        {
            info.pDriverName = newDriverName;
            hasChanges = true;
        }

        if (newPortName is not null)
        {
            info.pPortName = newPortName;
            hasChanges = true;
        }

        if (comment is not null)
        {
            info.pComment = comment;
            hasChanges = true;
        }

        if (location is not null)
        {
            info.pLocation = location;
            hasChanges = true;
        }

        // Validação de mudanças: update sem campos altera nada e gera erro de uso.
        if (!hasChanges)
        {
            throw new InvalidOperationException("Provide at least one update field.");
        }

        var defaults = new NativeMethods.PRINTER_DEFAULTS
        {
            pDatatype = null,
            pDevMode = IntPtr.Zero,
            DesiredAccess = NativeMethods.PRINTER_ACCESS_ADMINISTER
        };

        // Validação de acesso administrativo: necessário para SetPrinter.
        if (!NativeMethods.OpenPrinter(queueName, out var handle, ref defaults))
        {
            ThrowLastWin32($"OpenPrinter failed for queue '{queueName}'.");
        }

        try
        {
            var infoPtr = IntPtr.Zero;
            try
            {
                infoPtr = Marshal.AllocHGlobal(Marshal.SizeOf<NativeMethods.PRINTER_INFO_2>());
                Marshal.StructureToPtr(info, infoPtr, false);
                // Validação de persistência: garante que a alteração foi aplicada no spooler.
                if (!NativeMethods.SetPrinter(handle, 2, infoPtr, 0))
                {
                    ThrowLastWin32($"SetPrinter failed for queue '{queueName}'.");
                }
            }
            finally
            {
                if (infoPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<NativeMethods.PRINTER_INFO_2>(infoPtr);
                    Marshal.FreeHGlobal(infoPtr);
                }
            }
        }
        finally
        {
            NativeMethods.ClosePrinter(handle);
        }
    }

    /// <summary>
    /// Exclui uma fila do spooler usando handle com permissões elevadas.
    ///
    /// O método abre a fila com <c>PRINTER_ALL_ACCESS</c> e chama
    /// <c>DeletePrinter</c>. Em caso de falha, traduz o erro Win32 para mensagem
    /// mais acionável (por exemplo, acesso negado por falta de elevação ou console
    /// de gerenciamento aberto bloqueando a operação).
    /// </summary>
    public void DeleteQueue(string queueName)
    {
        var defaults = new NativeMethods.PRINTER_DEFAULTS
        {
            pDatatype = null,
            pDevMode = IntPtr.Zero,
            DesiredAccess = NativeMethods.PRINTER_ALL_ACCESS
        };

        // Validação de acesso total: exclusão exige permissões elevadas no handle.
        if (!NativeMethods.OpenPrinter(queueName, out var handle, ref defaults))
        {
            ThrowLastWin32($"OpenPrinter failed for queue '{queueName}'.");
        }

        try
        {
            // Validação de exclusão: sem sucesso, converte erro nativo para mensagem operável.
            if (!NativeMethods.DeletePrinter(handle))
            {
                var error = Marshal.GetLastWin32Error();
                // Validação de erro comum: feedback explícito para o caso de acesso negado.
                if (error == NativeMethods.ERROR_ACCESS_DENIED)
                {
                    throw new Win32Exception(
                        error,
                        $"DeletePrinter failed for queue '{queueName}'. Access denied. Run elevated and close any open queue property dialogs/management consoles.");
                }

                throw new Win32Exception(error, $"DeletePrinter failed for queue '{queueName}'.");
            }
        }
        finally
        {
            NativeMethods.ClosePrinter(handle);
        }
    }

    /// <summary>
    /// Inspeciona uma fila e aplica heurística para classificar aderência provável a WPP.
    ///
    /// A classificação combina três fontes de evidência:
    /// estado global de WPP, padrão do nome de porta (ex.: WSD*) e dados opcionais
    /// de APMON (protocolo/URL) quando disponíveis. O resultado é retornado em
    /// <see cref="QueueInspectionResult"/> com diagnóstico textual para auditoria.
    ///
    /// Importante: trata-se de inferência orientada por sinais operacionais,
    /// não de prova definitiva de conformidade.
    /// </summary>
    public QueueInspectionResult InspectQueue(string queueName)
    {
        var queue = GetQueueInfo(queueName);
        var port = queue.pPortName ?? string.Empty;
        var isWsdPort = port.StartsWith("WSD", StringComparison.OrdinalIgnoreCase);
        var global = wppStatusProvider.GetWppStatus().Status;

        var detailsParts = new List<string>();
        var apPort = TryGetApPortInfo(port);
        // Validação de evidência opcional: APMON enriquece a classificação quando disponível.
        if (apPort is not null)
        {
            detailsParts.Add($"APMON Protocol={ProtocolToString(apPort.Protocol)}({apPort.Protocol})");
            // Validação de detalhe adicional: URL pode existir para reforçar rastreabilidade.
            if (!string.IsNullOrWhiteSpace(apPort.DeviceOrServiceUrl))
            {
                detailsParts.Add($"APMON Url={apPort.DeviceOrServiceUrl}");
            }
        }

        // Regra da POC: status global + evidências de porta/protocolo determinam a classificação.
        var classification = WppQueueClassification.Indeterminate;
        // Validação principal: WPP global habilitado + sinais modernos -> provável WPP.
        if (global == WppStatus.Enabled && (isWsdPort || (apPort?.Protocol is 1 or 2)))
        {
            classification = WppQueueClassification.LikelyWpp;
            detailsParts.Insert(0, "Global WPP is enabled and queue indicates modern monitored port.");
        }
        // Validação principal: WPP global desabilitado prevalece como provável não-WPP.
        else if (global == WppStatus.Disabled)
        {
            classification = WppQueueClassification.LikelyNotWpp;
            detailsParts.Insert(0, "Global WPP is disabled.");
        }
        // Validação de inconsistência: global ativo sem sinal de porta moderna.
        else if (global == WppStatus.Enabled && !isWsdPort)
        {
            detailsParts.Insert(0, "Global WPP is enabled but queue port does not look like WSD.");
        }
        else
        {
            detailsParts.Insert(0, "Queue classification needs more evidence.");
        }

        return new QueueInspectionResult(
            queue.pPrinterName ?? queueName,
            port,
            global,
            classification,
            string.Join(" | ", detailsParts));
    }

    /// <summary>
    /// Converte o código numérico de protocolo APMON para rótulo legível.
    ///
    /// Mapeia os códigos conhecidos usados pela POC (WSD e IPP) e retorna
    /// "Unknown" para valores fora do catálogo atual.
    /// </summary>
    private static string ProtocolToString(uint protocol)
    {
        return protocol switch
        {
            1 => "WSD",
            2 => "IPP",
            _ => "Unknown"
        };
    }

    /// <summary>
    /// Tenta coletar metadados APMON da porta (protocolo e URL do dispositivo/serviço).
    ///
    /// Abre um canal XcvPort para a porta informada e executa o comando
    /// <c>GetAPPortInfo</c>. Se qualquer etapa falhar (porta inválida, monitor
    /// indisponível, erro de comando), retorna <see langword="null"/> para manter a
    /// inspeção resiliente, sem interromper o fluxo principal.
    /// </summary>
    private static ApPortData? TryGetApPortInfo(string portName)
    {
        // Validação de entrada: sem porta não há consulta APMON.
        if (string.IsNullOrWhiteSpace(portName))
        {
            return null;
        }

        var defaults = new NativeMethods.PRINTER_DEFAULTS
        {
            pDatatype = null,
            pDevMode = IntPtr.Zero,
            DesiredAccess = NativeMethods.SERVER_ACCESS_ADMINISTER
        };

        var xcvName = $",XcvPort {portName}";
        // Validação de acesso: se XcvPort não abre, a inspeção segue sem APMON.
        if (!NativeMethods.OpenPrinter(xcvName, out var hXcv, ref defaults))
        {
            return null;
        }

        try
        {
            var dataSize = Marshal.SizeOf<NativeMethods.AP_PORT_DATA_1>();
            var pData = Marshal.AllocHGlobal(dataSize);
            try
            {
                // Validação da chamada: falhas em GetAPPortInfo não devem quebrar inspeção.
                if (!NativeMethods.XcvData(
                        hXcv,
                        "GetAPPortInfo",
                        IntPtr.Zero,
                        0,
                        pData,
                        (uint)dataSize,
                        out _,
                        out var xcvStatus))
                {
                    return null;
                }

                // Validação de status de comando: só usa evidência quando o monitor confirma sucesso.
                if (xcvStatus != NativeMethods.ERROR_SUCCESS)
                {
                    return null;
                }

                var nativeData = Marshal.PtrToStructure<NativeMethods.AP_PORT_DATA_1>(pData);
                return new ApPortData(nativeData.Version, nativeData.Protocol, nativeData.DeviceOrServiceUrl ?? string.Empty);
            }
            finally
            {
                Marshal.FreeHGlobal(pData);
            }
        }
        finally
        {
            NativeMethods.ClosePrinter(hXcv);
        }
    }

    /// <summary>
    /// Lê a configuração completa de uma fila pelo nível 2 do Winspool.
    ///
    /// Abre a fila, consulta o tamanho necessário de buffer e recupera uma
    /// estrutura <c>PRINTER_INFO_2</c> com metadados completos (driver, porta,
    /// atributos, comentários etc.). É a base para cenários de inspeção e update.
    /// </summary>
    private static NativeMethods.PRINTER_INFO_2 GetQueueInfo(string queueName)
    {
        var defaults = new NativeMethods.PRINTER_DEFAULTS
        {
            pDatatype = null,
            pDevMode = IntPtr.Zero,
            DesiredAccess = NativeMethods.PRINTER_ACCESS_USE
        };

        // Validação de existência/acesso: sem abrir a fila, não há inspeção/atualização confiável.
        if (!NativeMethods.OpenPrinter(queueName, out var handle, ref defaults))
        {
            ThrowLastWin32($"OpenPrinter failed for queue '{queueName}'.");
        }

        try
        {
            // Validação de tamanho de buffer para leitura de PRINTER_INFO_2.
            if (!NativeMethods.GetPrinter(handle, 2, IntPtr.Zero, 0, out var needed))
            {
                var error = Marshal.GetLastWin32Error();
                // Validação de erro esperado: fora INSUFFICIENT_BUFFER é falha real.
                if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER)
                {
                    throw new Win32Exception(error, $"GetPrinter buffer query failed for queue '{queueName}'.");
                }
            }

            var buffer = Marshal.AllocHGlobal((int)needed);
            try
            {
                // Validação da leitura efetiva dos metadados da fila.
                if (!NativeMethods.GetPrinter(handle, 2, buffer, needed, out _))
                {
                    ThrowLastWin32($"GetPrinter failed for queue '{queueName}'.");
                }

                return Marshal.PtrToStructure<NativeMethods.PRINTER_INFO_2>(buffer);
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        finally
        {
            NativeMethods.ClosePrinter(handle);
        }
    }

    /// <summary>
    /// Enumera filas locais e conexões de impressora no nível <c>PRINTER_INFO_2</c>.
    ///
    /// Executa a estratégia de dupla chamada para obter buffer nativo,
    /// converte cada entrada para estrutura gerenciada e devolve a lista bruta,
    /// que será posteriormente projetada para modelos de domínio.
    /// </summary>
    private static List<NativeMethods.PRINTER_INFO_2> EnumeratePrinterInfo2()
    {
        // Validação de tamanho de buffer para enumeração de filas.
        if (!NativeMethods.EnumPrinters(
                NativeMethods.PRINTER_ENUM_LOCAL | NativeMethods.PRINTER_ENUM_CONNECTIONS,
                null,
                2,
                IntPtr.Zero,
                0,
                out var needed,
                out _))
        {
            var error = Marshal.GetLastWin32Error();
            // Validação de erro esperado na fase de descoberta de buffer.
            if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
            {
                throw new Win32Exception(error, "EnumPrinters buffer query failed.");
            }
        }

        // Validação de conteúdo: sem dados necessários, retorna lista vazia.
        if (needed == 0)
        {
            return [];
        }

        var buffer = Marshal.AllocHGlobal((int)needed);
        try
        {
            // Validação da leitura efetiva no buffer alocado.
            if (!NativeMethods.EnumPrinters(
                    NativeMethods.PRINTER_ENUM_LOCAL | NativeMethods.PRINTER_ENUM_CONNECTIONS,
                    null,
                    2,
                    buffer,
                    needed,
                    out _,
                    out var returned))
            {
                ThrowLastWin32("EnumPrinters failed.");
            }

            var list = new List<NativeMethods.PRINTER_INFO_2>((int)returned);
            var structSize = Marshal.SizeOf<NativeMethods.PRINTER_INFO_2>();
            for (var i = 0; i < returned; i++)
            {
                var ptr = IntPtr.Add(buffer, i * structSize);
                list.Add(Marshal.PtrToStructure<NativeMethods.PRINTER_INFO_2>(ptr));
            }

            return list;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Lança uma <see cref="Win32Exception"/> usando o último erro nativo da thread.
    ///
    /// Este helper padroniza o tratamento de falhas de interoperabilidade,
    /// anexando um contexto funcional à mensagem para facilitar diagnóstico.
    /// </summary>
    private static void ThrowLastWin32(string context)
    {
        throw new Win32Exception(Marshal.GetLastWin32Error(), context);
    }
}
