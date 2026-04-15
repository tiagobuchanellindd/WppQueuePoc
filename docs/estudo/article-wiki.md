# Filas WPP no Windows: estudo tecnico de funcionamento e operacao

Este estudo apresenta uma camada tecnica para operacao de filas de impressao no Windows com foco em previsibilidade operacional.
Ele concentra, em um fluxo unico, o ciclo de vida de filas no spooler, a leitura do contexto global de WPP, a inspecao heuristica de aderencia por fila e o diagnostico/ajuste de PrintTicket.

## 1) Objetivo do estudo

Em ambientes Windows, o comportamento de impressao muda conforme o contexto de politica global, driver e porta.
O estudo foi estruturado para transformar esse cenario em um fluxo operacional previsivel e auditavel.

Escopo implementado:

1. ciclo de vida de fila no spooler: criar, listar, atualizar e remover;
2. descoberta de recursos do host: portas, drivers, print processors e datatypes;
3. administracao de porta WSD via Xcv (`AddPort`) com validacao de retorno;
4. deteccao de status global WPP (`Enabled`, `Disabled`, `Unknown`);
5. inspecao heuristica de fila (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`);
6. leitura e atualizacao de PrintTicket em dois escopos:
   - default da fila,
   - user do usuario atual.

Em resumo, a abordagem proposta encapsula APIs nativas complexas em servicos diretos para operacao, suporte e diagnostico.

## 2) Visao tecnica da solucao

A implementacao foi separada em quatro blocos simples:

- **Contratos (Abstractions)**: definem o que pode ser feito.
- **Services**: executam os fluxos de negocio/operacao.
- **Interop (Win32)**: concentra chamadas nativas (`winspool.drv`).
- **Models**: padronizam entradas, saidas e classificacoes.

Essa separacao evita espalhar detalhes de ponteiro/memoria Win32 por todo o codigo e melhora manutencao.

## 3) Capacidades publicas da solucao

### 3.1 Operacoes de spooler

- `AddWsdPort(portName)`
- `CreateQueue(queueName, driverName, portName, printProcessor, dataType, comment, location)`
- `ListQueues()`
- `ListPorts()`
- `ListDrivers()`
- `ListPrintProcessors()`
- `ListDataTypes(printProcessor)`
- `UpdateQueue(queueName, newQueueName, newDriverName, newPortName, comment, location)`
- `DeleteQueue(queueName)`
- `InspectQueue(queueName)`

### 3.2 Status global WPP

- `GetWppStatus()` retorna status, origem da evidencia, detalhes e valor bruto quando existir.

### 3.3 PrintTicket

- `GetDefaultTicketInfo(queueName)`
- `GetUserTicketInfo(queueName)`
- `UpdateDefaultTicket(queueName, request)`
- `UpdateUserTicket(queueName, request)`

## 4) Fluxo de uso ponta a ponta

O passo a passo abaixo mostra a sequencia recomendada de uso.

### Passo 1: descobrir estado global WPP

```csharp
var wpp = wppStatusProvider.GetWppStatus();
```

Por que comecar aqui:

- WPP e politica global do Windows;
- isso influencia interpretacao de filas e portas;
- melhora a qualidade do diagnostico desde o inicio.

### Passo 2: mapear opcoes validas do host

```csharp
var ports = spooler.ListPorts();
var drivers = spooler.ListDrivers();
var processors = spooler.ListPrintProcessors();
var dataTypes = spooler.ListDataTypes("WinPrint");
```

Esse passo reduz erro de combinacao invalida (driver/processador/datatype).

### Passo 3: criar fila

```csharp
spooler.CreateQueue(
    queueName: "WPP-Queue-01",
    driverName: "Microsoft IPP Class Driver",
    portName: "WSD-12345678-ABCD",
    printProcessor: "WinPrint",
    dataType: "RAW",
    comment: "Fila criada para estudo WPP",
    location: "Lab");
```

Internamente, o servico monta `PRINTER_INFO_2`, aloca memoria nao gerenciada e chama `AddPrinter`.

### Passo 4: inspecionar fila no contexto WPP

```csharp
var inspection = spooler.InspectQueue("WPP-Queue-01");
```

Classificacoes possiveis:

- `LikelyWpp`
- `LikelyNotWpp`
- `Indeterminate`

A decisao combina sinais globais e sinais tecnicos da fila (porta e evidencia APMON quando disponivel).

### Passo 5: ler e ajustar PrintTicket

Leitura:

```csharp
var defaultInfo = ticketService.GetDefaultTicketInfo("WPP-Queue-01");
var userInfo = ticketService.GetUserTicketInfo("WPP-Queue-01");
```

Update:

```csharp
var request = new PrintTicketUpdateRequest(
    Duplexing: "TwoSidedLongEdge",
    OutputColor: "Monochrome",
    PageOrientation: "Portrait");

var result = ticketService.UpdateUserTicket("WPP-Queue-01", request);
```

Regra pratica: sempre comparar `Requested` e `AppliedValues` para validar o que o driver realmente aceitou.

### Passo 6: manter ou encerrar ciclo

Update parcial de fila:

```csharp
spooler.UpdateQueue(
    queueName: "WPP-Queue-01",
    newQueueName: "WPP-Queue-01-Renamed",
    newDriverName: null,
    newPortName: null,
    comment: "Fila ajustada",
    location: "Sala 2");
```

Delete:

```csharp
spooler.DeleteQueue("WPP-Queue-01-Renamed");
```

## 5) Funcionamento interno dos servicos

### 5.1 PrintSpoolerService

E o nucleo operacional de fila. Centraliza criacao, atualizacao, exclusao, enumeracao de recursos (filas, portas, drivers, processadores e datatypes) e inspecao heuristica de aderencia a WPP.

Padroes tecnicos importantes:

- APIs de enumeracao usam duas chamadas (descobrir tamanho, depois ler);
- erros Win32 sao convertidos para mensagens acionaveis;
- update de fila protege ponteiros sensiveis (`pDevMode`, `pSecurityDescriptor`);
- em Xcv, sucesso de transporte e sucesso funcional sao validados separadamente.

#### 5.1.1 Padrao de buffer dupla-chamada

Quase todas as APIs de enumeracao do spooler (`EnumPorts`, `EnumPrinters`, `EnumPrinterDrivers`, etc.) seguem o mesmo contrato: a primeira chamada descobre quantos bytes sao necessarios; a segunda preenche o buffer real. Esse padrao se repete em `ListPorts`, `ListDrivers`, `ListPrintProcessors`, `ListDataTypes` e `ListQueues`.

```csharp
// 1a chamada: descobre tamanho necessario
if (!NativeMethods.EnumPorts(null, 1, IntPtr.Zero, 0, out var requiredBufferSize, out _))
{
    var error = Marshal.GetLastWin32Error();
    // ERROR_INSUFFICIENT_BUFFER e esperado aqui — nao e falha real
    if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
    {
        throw new Win32Exception(error, "EnumPorts buffer query failed.");
    }
}

// Sem bytes necessarios = sem portas
if (requiredBufferSize == 0)
{
    return [];
}

// 2a chamada: aloca e preenche o buffer
var portsBuffer = Marshal.AllocHGlobal((int)requiredBufferSize);
try
{
    if (!NativeMethods.EnumPorts(null, 1, portsBuffer, requiredBufferSize, out _, out var returnedPortCount))
    {
        ThrowLastWin32("EnumPorts failed.");
    }

    // Itera as structs PORT_INFO_1 no bloco de memoria
    var portInfoStructSize = Marshal.SizeOf<NativeMethods.PORT_INFO_1>();
    var ports = new List<string>((int)returnedPortCount);
    for (var index = 0; index < returnedPortCount; index++)
    {
        var currentPortInfoPtr = IntPtr.Add(portsBuffer, index * portInfoStructSize);
        var portInfo = Marshal.PtrToStructure<NativeMethods.PORT_INFO_1>(currentPortInfoPtr);

        if (!string.IsNullOrWhiteSpace(portInfo.pName))
        {
            ports.Add(portInfo.pName);
        }
    }

    ports.Sort(StringComparer.OrdinalIgnoreCase);
    return ports;
}
finally
{
    Marshal.FreeHGlobal(portsBuffer);
}
```

O mesmo padrao e aplicado para drivers (com `DRIVER_INFO_2`), processadores (`PRINTPROCESSOR_INFO_1`), datatypes (`DATATYPES_INFO_1`) e filas (`PRINTER_INFO_2` via `EnumPrinters`).

#### 5.1.2 Criacao de fila no spooler

A criacao monta uma `PRINTER_INFO_2`, aloca memoria nao gerenciada, chama `AddPrinter` e garante liberacao de recursos em qualquer saida.

```csharp
public void CreateQueue(
    string queueName, string driverName, string portName,
    string printProcessor, string dataType, string comment, string location)
{
    var printerInfo = new NativeMethods.PRINTER_INFO_2
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

    var printerInfoPtr = IntPtr.Zero;
    try
    {
        printerInfoPtr = Marshal.AllocHGlobal(Marshal.SizeOf<NativeMethods.PRINTER_INFO_2>());
        Marshal.StructureToPtr(printerInfo, printerInfoPtr, false);
        var printerHandle = NativeMethods.AddPrinter(null, 2, printerInfoPtr);

        // Handle nulo = spooler recusou a nova fila
        if (printerHandle == IntPtr.Zero)
        {
            ThrowLastWin32("AddPrinter failed.");
        }

        NativeMethods.ClosePrinter(printerHandle);
    }
    finally
    {
        if (printerInfoPtr != IntPtr.Zero)
        {
            Marshal.DestroyStructure<NativeMethods.PRINTER_INFO_2>(printerInfoPtr);
            Marshal.FreeHGlobal(printerInfoPtr);
        }
    }
}
```

Pontos de atencao:
- `pDevMode` e `pSecurityDescriptor` sao `IntPtr.Zero` porque a criacao delega esses detalhes ao driver/sistema;
- `PRINTER_ATTRIBUTE_QUEUED` ativa spooling na fila;
- `DestroyStructure` e necessario antes de `FreeHGlobal` para liberar corretamente os ponteiros internos da struct.

#### 5.1.3 Criacao de porta WSD via Xcv

O fluxo de `AddWsdPort` usa o canal administrativo `XcvMonitor` e faz validacao dupla: sucesso da chamada Win32 e sucesso funcional do monitor.

```csharp
public void AddWsdPort(string portName)
{
    // Idempotencia: evita criar porta que ja existe
    if (ListPorts().Any(p => string.Equals(p, portName, StringComparison.OrdinalIgnoreCase)))
    {
        Console.WriteLine($"Port '{portName}' already exists.");
        return;
    }

    var printerDefaults = new NativeMethods.PRINTER_DEFAULTS
    {
        pDatatype = null,
        pDevMode = IntPtr.Zero,
        DesiredAccess = NativeMethods.SERVER_ACCESS_ADMINISTER
    };

    // Abre canal Xcv para o monitor WSD
    if (!NativeMethods.OpenPrinter(",XcvMonitor WSD Port", out var xcvHandle, ref printerDefaults))
    {
        ThrowLastWin32("OpenPrinter for XcvMonitor WSD Port failed.");
    }

    try
    {
        var addPortPayload = Encoding.Unicode.GetBytes(portName + '\0');

        // Envia comando AddPort ao monitor
        if (!NativeMethods.XcvData(
                xcvHandle,
                "AddPort",
                addPortPayload,
                addPortPayload.Length,
                IntPtr.Zero, 0,
                out _,
                out var xcvCommandStatus))
        {
            ThrowLastWin32("XcvData(AddPort) call failed.");
        }

        // Validacao funcional: mesmo com chamada OK, o monitor pode rejeitar
        if (xcvCommandStatus != NativeMethods.ERROR_SUCCESS)
        {
            if (xcvCommandStatus == NativeMethods.ERROR_NOT_SUPPORTED)
            {
                throw new Win32Exception(
                    (int)xcvCommandStatus,
                    "XcvData(AddPort) returned ERROR_NOT_SUPPORTED for WSD monitor. "
                    + "This environment likely requires device-discovery flow for WSD port creation. "
                    + "Use 'list-ports' and reuse an existing WSD port.");
            }

            throw new Win32Exception(
                (int)xcvCommandStatus,
                $"XcvData(AddPort) returned status {xcvCommandStatus}.");
        }
    }
    finally
    {
        NativeMethods.ClosePrinter(xcvHandle);
    }
}
```

O ponto central e a distincao entre **sucesso de transporte** (`XcvData` retornou `true`) e **sucesso funcional** (`xcvCommandStatus == ERROR_SUCCESS`). Isso evita falsos positivos em ambientes onde o monitor aceita o comando mas rejeita a operacao.

#### 5.1.4 Update parcial de fila

O update le o estado atual da fila, zera ponteiros sensiveis, aplica apenas campos informados e persiste via `SetPrinter`.

```csharp
public void UpdateQueue(string queueName, string? newQueueName, string? newDriverName,
    string? newPortName, string? comment, string? location)
{
    // Le estado atual completo (PRINTER_INFO_2 nivel 2)
    var queueInfo = GetQueueInfo(queueName);

    // Zera ponteiros sensiveis para evitar reutilizacao indevida de dados nativos
    queueInfo.pDevMode = IntPtr.Zero;
    queueInfo.pSecurityDescriptor = IntPtr.Zero;

    var hasChanges = false;

    if (newQueueName is not null) { queueInfo.pPrinterName = newQueueName; hasChanges = true; }
    if (newDriverName is not null) { queueInfo.pDriverName = newDriverName; hasChanges = true; }
    if (newPortName is not null) { queueInfo.pPortName = newPortName; hasChanges = true; }
    if (comment is not null) { queueInfo.pComment = comment; hasChanges = true; }
    if (location is not null) { queueInfo.pLocation = location; hasChanges = true; }

    if (!hasChanges)
    {
        throw new InvalidOperationException("Provide at least one update field.");
    }

    var printerDefaults = new NativeMethods.PRINTER_DEFAULTS
    {
        pDatatype = null,
        pDevMode = IntPtr.Zero,
        DesiredAccess = NativeMethods.PRINTER_ACCESS_ADMINISTER
    };

    // Abre com acesso administrativo para SetPrinter
    if (!NativeMethods.OpenPrinter(queueName, out var printerHandle, ref printerDefaults))
    {
        ThrowLastWin32($"OpenPrinter failed for queue '{queueName}'.");
    }

    try
    {
        var printerInfoPtr = IntPtr.Zero;
        try
        {
            printerInfoPtr = Marshal.AllocHGlobal(Marshal.SizeOf<NativeMethods.PRINTER_INFO_2>());
            Marshal.StructureToPtr(queueInfo, printerInfoPtr, false);

            if (!NativeMethods.SetPrinter(printerHandle, 2, printerInfoPtr, 0))
            {
                ThrowLastWin32($"SetPrinter failed for queue '{queueName}'.");
            }
        }
        finally
        {
            if (printerInfoPtr != IntPtr.Zero)
            {
                Marshal.DestroyStructure<NativeMethods.PRINTER_INFO_2>(printerInfoPtr);
                Marshal.FreeHGlobal(printerInfoPtr);
            }
        }
    }
    finally
    {
        NativeMethods.ClosePrinter(printerHandle);
    }
}
```

Pontos de atencao:
- `pDevMode` e `pSecurityDescriptor` sao zerados propositalmente: o `PRINTER_INFO_2` lido contem ponteiros nativos que ficam invalidos fora do buffer original. Repassa-los ao `SetPrinter` causaria corrupcao de memoria;
- o update e parcial: campos `null` nao alteram o valor existente;
- o handle e aberto com `PRINTER_ACCESS_ADMINISTER`, necessario para `SetPrinter`.

#### 5.1.5 Inspecao heuristica de fila

A classificacao WPP combina tres fontes de evidencia: estado global de WPP, padrao do nome de porta e dados APMON (quando disponiveis).

```csharp
public QueueInspectionResult InspectQueue(string queueName)
{
    var queueInfo = GetQueueInfo(queueName);
    var portName = queueInfo.pPortName ?? string.Empty;
    var isWsdPort = portName.StartsWith("WSD", StringComparison.OrdinalIgnoreCase);
    var globalWppStatus = wppStatusProvider.GetWppStatus().Status;

    var diagnosticDetails = new List<string>();

    // Evidencia opcional: APMON enriquece a classificacao quando disponivel
    var apPortInfo = TryGetApPortInfo(portName);
    if (apPortInfo is not null)
    {
        diagnosticDetails.Add($"APMON Protocol={ProtocolToString(apPortInfo.Protocol)}({apPortInfo.Protocol})");
        if (!string.IsNullOrWhiteSpace(apPortInfo.DeviceOrServiceUrl))
        {
            diagnosticDetails.Add($"APMON Url={apPortInfo.DeviceOrServiceUrl}");
        }
    }

    var classification = WppQueueClassification.Indeterminate;

    // WPP global habilitado + sinais de porta moderna -> provavel WPP
    if (globalWppStatus == WppStatus.Enabled && (isWsdPort || (apPortInfo?.Protocol is 1 or 2)))
    {
        classification = WppQueueClassification.LikelyWpp;
        diagnosticDetails.Insert(0, "Global WPP is enabled and queue indicates modern monitored port.");
    }
    // WPP global desabilitado prevalece como provavel nao-WPP
    else if (globalWppStatus == WppStatus.Disabled)
    {
        classification = WppQueueClassification.LikelyNotWpp;
        diagnosticDetails.Insert(0, "Global WPP is disabled.");
    }
    // Global ativo sem sinal de porta moderna -> inconsistencia
    else if (globalWppStatus == WppStatus.Enabled && !isWsdPort)
    {
        diagnosticDetails.Insert(0, "Global WPP is enabled but queue port does not look like WSD.");
    }
    else
    {
        diagnosticDetails.Insert(0, "Queue classification needs more evidence.");
    }

    return new QueueInspectionResult(
        queueInfo.pPrinterName ?? queueName,
        portName,
        globalWppStatus,
        classification,
        string.Join(" | ", diagnosticDetails));
}
```

A arvore de decisao resumida:

| Status Global | Porta WSD ou APMON WSD/IPP | Classificacao |
|---|---|---|
| `Enabled` | Sim | `LikelyWpp` |
| `Enabled` | Nao | `Indeterminate` |
| `Disabled` | Qualquer | `LikelyNotWpp` |
| `Unknown` | Qualquer | `Indeterminate` |

#### 5.1.6 Coleta de evidencias APMON via XcvPort

O `TryGetApPortInfo` e um helper resiliente: abre um canal `XcvPort` para a porta e executa `GetAPPortInfo`. Se qualquer etapa falhar, retorna `null` sem interromper a inspecao.

```csharp
private static ApPortData? TryGetApPortInfo(string portName)
{
    if (string.IsNullOrWhiteSpace(portName))
    {
        return null;
    }

    var printerDefaults = new NativeMethods.PRINTER_DEFAULTS
    {
        pDatatype = null,
        pDevMode = IntPtr.Zero,
        DesiredAccess = NativeMethods.SERVER_ACCESS_ADMINISTER
    };

    var xcvPortName = $",XcvPort {portName}";

    // Se XcvPort nao abre, a inspecao segue sem evidencia APMON
    if (!NativeMethods.OpenPrinter(xcvPortName, out var xcvHandle, ref printerDefaults))
    {
        return null;
    }

    try
    {
        var apPortDataSize = Marshal.SizeOf<NativeMethods.AP_PORT_DATA_1>();
        var apPortDataPtr = Marshal.AllocHGlobal(apPortDataSize);
        try
        {
            if (!NativeMethods.XcvData(
                    xcvHandle,
                    "GetAPPortInfo",
                    IntPtr.Zero, 0,
                    apPortDataPtr,
                    (uint)apPortDataSize,
                    out _,
                    out var xcvCommandStatus))
            {
                return null; // Falha no canal — segue sem APMON
            }

            if (xcvCommandStatus != NativeMethods.ERROR_SUCCESS)
            {
                return null; // Monitor rejeitou — segue sem APMON
            }

            var apPortData = Marshal.PtrToStructure<NativeMethods.AP_PORT_DATA_1>(apPortDataPtr);
            return new ApPortData(
                apPortData.Version,
                apPortData.Protocol,
                apPortData.DeviceOrServiceUrl ?? string.Empty);
        }
        finally
        {
            Marshal.FreeHGlobal(apPortDataPtr);
        }
    }
    finally
    {
        NativeMethods.ClosePrinter(xcvHandle);
    }
}
```

A struct `AP_PORT_DATA_1` retornada pelo monitor traz:
- `Version`: versao da estrutura;
- `Protocol`: 1 = WSD, 2 = IPP;
- `DeviceOrServiceUrl`: URL do dispositivo/servico (260 chars fixos no layout nativo).

### 5.2 WppRegistryService

Resolve o estado global WPP lendo valores de politica no Registry do Windows (HKLM).

#### 5.2.1 Definicao dos probes e regra de decisao

O servico define uma lista de "probes" — combinacoes de caminho de registro + nome de valor — com os valores esperados para habilitado e desabilitado. A iteracao percorre todos os probes e aplica uma estrategia de fallback com prioridade para `Enabled`.

```csharp
private static readonly PolicyStatusCheck[] StatusChecks =
[
    new(
        @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
        "WindowsProtectedPrintGroupPolicyState",
        EnabledWhen: 1,
        DisabledWhen: 0),
    new(
        @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
        "WindowsProtectedPrintMode",
        EnabledWhen: 1,
        DisabledWhen: 0)
];

private sealed record PolicyStatusCheck(
    string RegistryPath,
    string PolicyValueName,
    int EnabledWhen,
    int DisabledWhen);
```

Essa estrutura permite adicionar novos probes sem alterar a logica de decisao.

#### 5.2.2 Fluxo completo de `GetWppStatus`

A rotina percorre os probes em sequencia. A regra e: `Enabled` retorna imediatamente, `Disabled` e guardado como candidato (porque um probe posterior pode indicar `Enabled`), e valores fora do mapeamento retornam `Unknown`.

```csharp
public WppStatusResult GetWppStatus()
{
    WppStatusResult? disabledCandidate = null;

    foreach (var statusCheck in StatusChecks)
    {
        using var key = Registry.LocalMachine.OpenSubKey(statusCheck.RegistryPath, writable: false);

        // Se a chave nao existe, tenta o proximo probe
        if (key is null)
        {
            continue;
        }

        var rawValue = key.GetValue(statusCheck.PolicyValueName);

        // Se o valor nao existe neste probe, tenta o proximo
        if (rawValue is null)
        {
            continue;
        }

        // So aceita tipos que conseguimos converter com seguranca
        if (!TryConvertToInt(rawValue, out var numericValue))
        {
            return new WppStatusResult(
                WppStatus.Unknown,
                $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                $"Unsupported value type: {rawValue.GetType().Name}.",
                null);
        }

        // Habilitado -> retorna imediatamente
        if (numericValue == statusCheck.EnabledWhen)
        {
            return new WppStatusResult(
                WppStatus.Enabled,
                $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                "Matched known enabled value.",
                numericValue);
        }

        // Desabilitado -> guarda candidato e continua
        if (numericValue == statusCheck.DisabledWhen)
        {
            disabledCandidate ??= new WppStatusResult(
                WppStatus.Disabled,
                $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                "Matched known disabled value.",
                numericValue);
            continue;
        }

        // Valor fora do mapeamento conhecido -> Unknown
        return new WppStatusResult(
            WppStatus.Unknown,
            $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
            "Value found but does not match known enabled/disabled mapping.",
            numericValue);
    }

    // Se encontrou candidato Disabled, retorna ele
    if (disabledCandidate is not null)
    {
        return disabledCandidate;
    }

    // Nenhum probe encontrou valor -> assume Disabled por padrao
    return new WppStatusResult(
        WppStatus.Disabled,
        "No known WPP Registry value found.",
        "No known enabled/disabled value was found in mapped probes. Treated as disabled by default.",
        null);
}
```

Cenarios de saida resumidos:

| Probe encontrado | Valor lido | Resultado |
|---|---|---|
| Sim | Igual a `EnabledWhen` | `Enabled` (retorno imediato) |
| Sim | Igual a `DisabledWhen` | Guarda candidato `Disabled`, continua |
| Sim | Outro inteiro | `Unknown` (retorno imediato) |
| Sim | Tipo nao conversivel | `Unknown` (retorno imediato) |
| Nenhum probe com valor | — | `Disabled` por padrao |

#### 5.2.3 Conversao segura de valor do Registry

O Registry pode armazenar o valor como `int` (DWORD) ou como `string` (texto numerico). O helper normaliza os dois formatos.

```csharp
private static bool TryConvertToInt(object rawValue, out int numericValue)
{
    // Tipo nativo comum no Registry (DWORD)
    if (rawValue is int integerValue)
    {
        numericValue = integerValue;
        return true;
    }

    // Fallback: alguns ambientes armazenam como texto numerico
    if (rawValue is string stringValue && int.TryParse(stringValue, out var parsedValue))
    {
        numericValue = parsedValue;
        return true;
    }

    numericValue = default;
    return false;
}
```

Se o tipo nao for `int` nem `string` parseable (por exemplo, `byte[]` de um valor REG_BINARY), `TryConvertToInt` retorna `false` e o chamador trata como `Unknown`.

### 5.3 PrintTicketService

Le e atualiza ticket com estrategia de compatibilidade por reflection em `System.Printing`.

Por que essa escolha ajuda:

- reduz acoplamento forte a variacoes de runtime/driver;
- permite fallback mais suave quando o ambiente nao suporta um caminho especifico;
- retorna resultado funcional com contexto ao inves de quebrar todo o fluxo.

#### 5.3.1 Leitura de ticket (Default e User)

O fluxo de leitura verifica disponibilidade de `System.Printing` por reflection, abre `LocalPrintServer`, localiza a fila e extrai propriedades do ticket. A mesma estrutura e usada para `DefaultPrintTicket` e `UserPrintTicket`.

```csharp
public PrintTicketInfoResult GetDefaultTicketInfo(string queueName)
{
    var ticketSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    // Verifica se System.Printing esta disponivel no runtime
    var printServerType = Type.GetType(
        "System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
    if (printServerType is null)
    {
        return new PrintTicketInfoResult(
            queueName,
            Available: false,
            Details: "System.Printing assembly is not available in this runtime.",
            Attributes: ticketSettings);
    }

    object? printServer = null;
    object? queue = null;
    try
    {
        printServer = CreateLocalPrintServer(printServerType, "AdministrateServer");
        if (printServer is null)
        {
            return new PrintTicketInfoResult(
                queueName, false, "Unable to create LocalPrintServer.", ticketSettings);
        }

        queue = GetPrintQueue(printServerType, printServer, queueName,
            "DefaultPrintTicket", "AdministratePrinter");
        if (queue is null)
        {
            return new PrintTicketInfoResult(
                queueName, false, $"Queue '{queueName}' not found.", ticketSettings);
        }

        var queueType = queue.GetType();
        var defaultTicketPropertyInfo = queueType.GetProperty("DefaultPrintTicket");
        var defaultTicket = defaultTicketPropertyInfo?.GetValue(queue);
        if (defaultTicket is null)
        {
            return new PrintTicketInfoResult(
                queueName, false, "DefaultPrintTicket is not available.", ticketSettings);
        }

        // Extrai propriedades conhecidas do ticket por reflection (best-effort)
        ReadTicketAttribute(defaultTicket, "OutputColor", ticketSettings);
        ReadTicketAttribute(defaultTicket, "PageMediaSize", ticketSettings);
        ReadTicketAttribute(defaultTicket, "PageOrientation", ticketSettings);
        ReadTicketAttribute(defaultTicket, "InputBin", ticketSettings);
        ReadTicketAttribute(defaultTicket, "Duplexing", ticketSettings);
        ReadTicketAttribute(defaultTicket, "CopyCount", ticketSettings);
        ReadTicketAttribute(defaultTicket, "Collation", ticketSettings);
        ReadTicketAttribute(defaultTicket, "Stapling", ticketSettings);

        return new PrintTicketInfoResult(
            queueName,
            Available: true,
            Details: "Default print ticket information captured.",
            Attributes: ticketSettings);
    }
    catch (Exception ex)
    {
        return new PrintTicketInfoResult(
            queueName,
            Available: false,
            Details: $"Failed to read default print ticket information: {ex.Message}",
            Attributes: ticketSettings);
    }
    finally
    {
        DisposeIfPossible(queue);
        DisposeIfPossible(printServer);
    }
}
```

A versao para `GetUserTicketInfo` e identica, substituindo apenas `"DefaultPrintTicket"` por `"UserPrintTicket"` na obtencao da property e na chamada a `GetPrintQueue`.

#### 5.3.2 Fluxo interno de update (`UpdatePrintTicketInternal`)

Os metodos publicos `UpdateDefaultTicket` e `UpdateUserTicket` delegam para um fluxo interno comum que: prepara atributos, abre servidor/fila, aplica diferencas, faz `Commit()` e retorna pedido vs aplicado.

```csharp
public PrintTicketUpdateResult UpdateDefaultTicket(string queueName, PrintTicketUpdateRequest request)
    => UpdatePrintTicketInternal(queueName, request, "DefaultPrintTicket", "Default");

public PrintTicketUpdateResult UpdateUserTicket(string queueName, PrintTicketUpdateRequest request)
    => UpdatePrintTicketInternal(queueName, request, "UserPrintTicket", "User");

private PrintTicketUpdateResult UpdatePrintTicketInternal(
    string queueName, PrintTicketUpdateRequest request,
    string ticketPropertyName, string ticketScope)
{
    // 1) Prepara atributos solicitados
    var requestedSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    void TryAdd(string settingName, string? settingValue)
    {
        if (!string.IsNullOrWhiteSpace(settingValue))
            requestedSettings[settingName] = settingValue.Trim();
    }
    TryAdd("Duplexing", request.Duplexing);
    TryAdd("OutputColor", request.OutputColor);

    var printServerType = Type.GetType(
        "System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
    if (printServerType is null)
    {
        return new PrintTicketUpdateResult(
            queueName, ticketScope, false,
            "System.Printing not available.",
            requestedSettings, new Dictionary<string, string>());
    }

    object? printServer = null;
    object? queue = null;
    try
    {
        // 2) Abre servidor e fila
        printServer = CreateLocalPrintServer(printServerType, "AdministrateServer");
        queue = GetPrintQueue(printServerType, printServer, queueName,
            ticketPropertyName, "AdministratePrinter");

        if (queue is null)
            return new PrintTicketUpdateResult(
                queueName, ticketScope, false,
                $"Queue '{queueName}' not found.",
                requestedSettings, new Dictionary<string, string>());

        var queueType = queue.GetType();
        var ticketPropertyInfo = queueType.GetProperty(ticketPropertyName);
        if (ticketPropertyInfo == null)
            return new PrintTicketUpdateResult(
                queueName, ticketScope, false,
                $"Print ticket property '{ticketPropertyName}' not found.",
                requestedSettings, new Dictionary<string, string>());

        // 3) Le o ticket alvo
        var targetTicket = ticketPropertyInfo.GetValue(queue);
        if (targetTicket is null)
            return new PrintTicketUpdateResult(
                queueName, ticketScope, false,
                $"Print ticket '{ticketPropertyName}' not found.",
                requestedSettings, new Dictionary<string, string>());

        // 4) Aplica apenas diferencas reais
        bool hasChanges = false;
        hasChanges |= WriteTicketAttribute(targetTicket, "Duplexing", request.Duplexing);
        hasChanges |= WriteTicketAttribute(targetTicket, "OutputColor", request.OutputColor);
        hasChanges |= WriteTicketAttribute(targetTicket, "PageOrientation", request.PageOrientation);

        string? saveErrorDetails = null;
        if (hasChanges)
        {
            // 5a) Seta ticket de volta na fila
            try
            {
                ticketPropertyInfo.SetValue(queue, targetTicket);
            }
            catch (Exception setTicketException)
            {
                var error = GetInnermostMessage(setTicketException);
                saveErrorDetails = $"Failed to set updated ticket back on '{ticketPropertyName}': {error}";
            }

            // 5b) Commit para persistir
            try
            {
                var commitMethod = queueType.GetMethod("Commit");
                commitMethod?.Invoke(queue, null);
            }
            catch (Exception commitException)
            {
                var error = GetInnermostMessage(commitException);
                saveErrorDetails = $"Commit failed: {error}";
            }
        }

        // 6) Le valores efetivos apos commit e retorna pedido x aplicado
        var appliedSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ReadTicketAttribute(targetTicket, "Duplexing", appliedSettings);
        ReadTicketAttribute(targetTicket, "OutputColor", appliedSettings);
        ReadTicketAttribute(targetTicket, "PageOrientation", appliedSettings);

        string statusMessage = hasChanges
            ? (saveErrorDetails == null
                ? "Ticket updated successfully."
                : $"Ticket update attempted. {saveErrorDetails}")
            : "No changes needed: the provided values matched current settings.";

        return new PrintTicketUpdateResult(
            queueName, ticketScope, hasChanges, statusMessage,
            requestedSettings, appliedSettings);
    }
    catch (Exception ex)
    {
        var error = GetInnermostMessage(ex);
        return new PrintTicketUpdateResult(
            queueName, ticketScope, false,
            $"Exception during PrintTicket update: {error}.",
            requestedSettings, new Dictionary<string, string>());
    }
    finally
    {
        DisposeIfPossible(queue);
        DisposeIfPossible(printServer);
    }
}
```

Regra pratica: sempre comparar `Requested` e `AppliedValues` no resultado para validar o que o driver realmente aceitou.

#### 5.3.3 Escrita de atributo com deteccao de mudanca (`WriteTicketAttribute`)

O metodo compara o valor atual com o solicitado e so aplica quando ha diferenca real, evitando commits desnecessarios.

```csharp
private static bool WriteTicketAttribute(object? ticket, string propertyName, string? requestedValue)
{
    try
    {
        if (ticket == null || string.IsNullOrWhiteSpace(requestedValue))
            return false;

        var ticketType = ticket.GetType();
        var propertyInfo = ticketType.GetProperty(propertyName);
        if (propertyInfo == null || !propertyInfo.CanWrite) return false;

        var currentValue = propertyInfo.GetValue(ticket);
        var propertyType = propertyInfo.PropertyType;
        object? convertedValue = ConvertIfPossible(propertyType, requestedValue);

        // Compara valor atual vs solicitado
        bool isDifferent;
        if (currentValue is null && convertedValue is not null) isDifferent = true;
        else if (currentValue is not null && convertedValue is null) isDifferent = true;
        else if (currentValue is null && convertedValue is null) isDifferent = false;
        else if (propertyType == typeof(string) || propertyType.IsEnum)
            isDifferent = !string.Equals(
                currentValue?.ToString()?.Trim(),
                convertedValue?.ToString()?.Trim(),
                StringComparison.OrdinalIgnoreCase);
        else
            isDifferent = !Equals(currentValue, convertedValue);

        if (!isDifferent)
            return false;

        propertyInfo.SetValue(ticket, convertedValue);
        return true;
    }
    catch { return false; }
}
```

O retorno `bool` e acumulado via `|=` no fluxo principal para determinar se houve pelo menos uma mudanca que justifique o `Commit()`.

#### 5.3.4 Conversao de tipo para propriedades do ticket (`ConvertIfPossible`)

O ticket de `System.Printing` usa enums (como `Duplexing`, `OutputColor`), tipos primitivos e `Nullable<T>`. O helper converte a string do request para o tipo real da propriedade.

```csharp
private static object? ConvertIfPossible(Type targetType, string rawValue)
{
    try
    {
        var isNullable = targetType.IsGenericType
            && targetType.GetGenericTypeDefinition() == typeof(Nullable<>);
        var underlyingType = isNullable ? Nullable.GetUnderlyingType(targetType) : null;

        // Nullable sem valor -> null
        if (isNullable && string.IsNullOrWhiteSpace(rawValue))
            return null;

        // Enum: converte por nome (case-insensitive)
        if ((underlyingType ?? targetType).IsEnum)
        {
            var enumType = (underlyingType ?? targetType);
            return Enum.Parse(enumType, rawValue, ignoreCase: true);
        }

        // Tipo primitivo: usa ChangeType
        var realType = underlyingType ?? targetType;
        return System.Convert.ChangeType(rawValue, realType);
    }
    catch
    {
        // Fallback para string: preserva valor original
        if (targetType == typeof(string)
            || (Nullable.GetUnderlyingType(targetType) == typeof(string)))
            return rawValue;
        throw;
    }
}
```

Isso permite que o consumidor envie `"TwoSidedLongEdge"` como string e o metodo resolva para o enum `System.Printing.Duplexing.TwoSidedLongEdge` automaticamente.

#### 5.3.5 Criacao de `LocalPrintServer` e obtencao de fila por reflection

Como `System.Printing` e carregado por reflection, a criacao do servidor e a obtencao da fila usam estrategias de fallback para maximizar compatibilidade.

**Criacao do servidor:**

```csharp
private static object? CreateLocalPrintServer(Type printServerType, string requestedAccessName)
{
    try
    {
        // Tenta construtor com PrintSystemDesiredAccess
        var desiredAccessType = Type.GetType(
            "System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);
        if (desiredAccessType is null)
            return Activator.CreateInstance(printServerType);

        var desiredAccess = Enum.Parse(desiredAccessType, requestedAccessName, ignoreCase: true);
        var ctor = printServerType.GetConstructor(new[] { desiredAccessType });
        if (ctor is null)
            return Activator.CreateInstance(printServerType); // Fallback: construtor padrao

        return ctor.Invoke(new[] { desiredAccess });
    }
    catch
    {
        return Activator.CreateInstance(printServerType); // Fallback final
    }
}
```

**Obtencao da fila (3 estrategias em cascata):**

```csharp
private static object? GetPrintQueue(
    Type printServerType, object? printServer, string queueName,
    string requiredTicketPropertyName, string requestedAccessName)
{
    if (printServer is null) return null;

    var desiredAccessType = Type.GetType(
        "System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);

    // Estrategia 1: new PrintQueue(printServer, queueName, desiredAccess)
    if (desiredAccessType is not null)
    {
        var desiredAccess = Enum.Parse(desiredAccessType, requestedAccessName, ignoreCase: true);
        var printQueueType = Type.GetType(
            "System.Printing.PrintQueue, System.Printing", throwOnError: false);
        var basePrintServerType = Type.GetType(
            "System.Printing.PrintServer, System.Printing", throwOnError: false);

        if (printQueueType is not null && basePrintServerType is not null
            && basePrintServerType.IsInstanceOfType(printServer))
        {
            var ctor = printQueueType.GetConstructor(
                new[] { basePrintServerType, typeof(string), desiredAccessType });
            if (ctor is not null)
                return ctor.Invoke(new object[] { printServer, queueName, desiredAccess });
        }
    }

    // Estrategia 2: printServer.GetPrintQueue(queueName, propertiesFilter)
    var getQueueWithProperties = printServerType.GetMethod(
        "GetPrintQueue", new[] { typeof(string), typeof(string[]) });
    if (getQueueWithProperties is not null)
    {
        var properties = new[] { requiredTicketPropertyName };
        return getQueueWithProperties.Invoke(printServer, new object[] { queueName, properties });
    }

    // Estrategia 3: printServer.GetPrintQueue(queueName)
    var getQueueDefault = printServerType.GetMethod(
        "GetPrintQueue", new[] { typeof(string) });
    return getQueueDefault?.Invoke(printServer, new object[] { queueName });
}
```

As 3 estrategias em ordem de prioridade:

1. **Construtor direto com acesso explicito** — garante permissao adequada para leitura/escrita;
2. **`GetPrintQueue` com filtro de propriedades** — carrega apenas as propriedades necessarias;
3. **`GetPrintQueue` simples** — fallback final quando as sobrecargas anteriores nao existem.

Essa cascata torna o servico resiliente a diferencas entre versoes de runtime e configuracoes de ambiente.

## 6) Como interpretar os resultados

### `WppStatusResult`

- `Status`: estado final (`Enabled`, `Disabled`, `Unknown`)
- `Source`: origem da leitura
- `Details`: explicacao tecnica
- `RawValue`: valor bruto, quando aplicavel

### `QueueInspectionResult`

- `QueueName`, `PortName`
- `GlobalWppStatus`
- `Classification`
- `Details`

Leitura pratica: bom para inventario e triagem rapida de filas.

### `PrintTicketInfoResult`

- `Available`: leitura foi possivel
- `Details`: contexto da tentativa
- `Attributes`: snapshot de propriedades relevantes

### `PrintTicketUpdateResult`

- `Applied`: houve alteracao efetiva
- `Requested`: intencao enviada
- `AppliedValues`: resultado observado
- `Details`: mensagem final da tentativa

Leitura pratica: valida diferenca entre o que foi solicitado e o que o driver realmente sustentou.

## 7) Exemplo unico de execucao

```csharp
using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;
using WppQueuePoc.Services;

IWppStatusProvider wppStatusProvider = new WppRegistryService();
IPrintSpoolerService spooler = new PrintSpoolerService(wppStatusProvider);
IPrintTicketService ticketService = new PrintTicketService();

var queue = "WPP-Queue-01";

var wpp = wppStatusProvider.GetWppStatus();

var ports = spooler.ListPorts();
var drivers = spooler.ListDrivers();

spooler.CreateQueue(queue, drivers[0], ports[0], "WinPrint", "RAW", "WPP Study", "Lab");

var inspection = spooler.InspectQueue(queue);

var before = ticketService.GetUserTicketInfo(queue);
var req = new PrintTicketUpdateRequest("TwoSidedLongEdge", "Monochrome", "Portrait");
var after = ticketService.UpdateUserTicket(queue, req);
```

Observacao: em cenario real, valide listas antes de indexar e trate excecoes operacionais.

## 8) Erros comuns e leitura correta

- `ERROR_ACCESS_DENIED`: falta privilegio para operacao administrativa.
- `ERROR_INSUFFICIENT_BUFFER`: comportamento esperado na primeira chamada de enumeracao.
- `ERROR_NOT_SUPPORTED`: comando nao suportado no monitor/driver/porta daquele ambiente.

## 9) Limites conhecidos

- classificacao WPP por fila e heuristica, nao certificacao formal;
- resultado depende de versao de Windows, driver e permissao local;
- nem todo update de ticket e aceito integralmente pelo driver;
- ambientes diferentes podem acionar fallbacks diferentes em `System.Printing`.

## 10) Checklist rapido de validacao

1. consultar status global WPP;
2. listar recursos disponiveis (ports, drivers, processors, datatypes);
3. criar fila com combinacao valida;
4. inspecionar fila;
5. ler ticket default e user;
6. aplicar update e comparar pedido x aplicado;
7. atualizar metadados da fila;
8. remover fila e confirmar limpeza.

## 11) POC

Link da POC: https://github.com/tiagobuchanellindd/WppQueuePoc

