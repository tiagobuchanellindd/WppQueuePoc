# Filas WPP no Windows: estudo técnico de funcionamento e operação

Este estudo apresenta uma camada técnica para operação de filas de impressão no Windows com foco em previsibilidade operacional.
Ele concentra, em um fluxo único, o ciclo de vida de filas no spooler, a leitura do contexto global de WPP, a inspeção heurística de aderência por fila e o diagnóstico/ajuste de PrintTicket.

## 1) Objetivo do estudo

Em ambientes Windows, o comportamento de impressão muda conforme o contexto de política global, driver e porta.
O estudo foi estruturado para transformar esse cenário em um fluxo operacional previsível e auditável.

Escopo implementado:

1. ciclo de vida de fila no spooler: criar, listar, atualizar e remover;
2. descoberta de recursos do host: portas, drivers, print processors e datatypes;
3. administração de porta WSD via Xcv (`AddPort`) com validação de retorno;
4. detecção de status global WPP (`Enabled`, `Disabled`, `Unknown`);
5. inspeção heurística de fila (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`);
6. leitura e atualização de PrintTicket em dois escopos:
   - default da fila,
   - usuário do usuário atual.

Em resumo, a abordagem proposta encapsula APIs nativas complexas em serviços diretos para operação, suporte e diagnóstico.

## 2) Visão técnica da solução

A implementação foi separada em quatro blocos simples:

- **Contratos (Abstractions)**: definem o que pode ser feito.
- **Services**: executam os fluxos de negócio/operação.
- **Interop (Win32)**: concentra chamadas nativas (`winspool.drv`).
- **Models**: padronizam entradas, saídas e classificações.

Essa separacao evita espalhar detalhes de ponteiro/memória Win32 por todo o código e melhora manutenção.

## 3) Capacidades públicas da solução

### 3.1 Operações de spooler

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

- `GetWppStatus()` retorna status, origem da evidência, detalhes e valor bruto quando existir.

### 3.3 PrintTicket

- `GetDefaultTicketInfo(queueName)`
- `GetUserTicketInfo(queueName)`
- `UpdateDefaultTicket(queueName, request)`
- `UpdateUserTicket(queueName, request)`

## 4) Fluxo de uso ponta a ponta

O passo a passo abaixo mostra a sequência recomendada de uso.

### Passo 1: descobrir estado global WPP

```csharp
var wpp = wppStatusProvider.GetWppStatus();
```

Por que começar aqui:

- WPP é política global do Windows;
- isso influência interpretação de filas e portas;
- melhora a qualidade do diagnóstico desde o início.

### Passo 2: mapear opções válidas do host

```csharp
var ports = spooler.ListPorts();
var drivers = spooler.ListDrivers();
var processors = spooler.ListPrintProcessors();
var dataTypes = spooler.ListDataTypes("WinPrint");
```

Esse passo reduz erro de combinação inválida (driver/processador/datatype).

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

Internamente, o serviço monta `PRINTER_INFO_2`, aloca memória não gerenciada e chama `AddPrinter`.

### Passo 4: inspecionar fila no contexto WPP

```csharp
var inspection = spooler.InspectQueue("WPP-Queue-01");
```

Classificações possíveis:

- `LikelyWpp`
- `LikelyNotWpp`
- `Indeterminate`

A decisão combina sinais globais e sinais técnicos da fila (porta e evidência APMON quando disponível).

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

Regra prática: sempre comparar `Requested` e `AppliedValues` para validar o que o driver realmente aceitou.

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

## 5) Funcionamento interno dos serviços

### 5.1 PrintSpoolerService

É o núcleo operacional de fila. Centraliza criação, atualização, exclusão, enumeração de recursos (filas, portas, drivers, processadores e datatypes) e inspeção heurística de aderência a WPP.

Padrões técnicos importantes:

- APIs de enumeração usam duas chamadas (descobrir tamanho, depois ler);
- erros Win32 são convertidos para mensagens acionáveis;
- update de fila protege ponteiros sensíveis (`pDevMode`, `pSecurityDescriptor`);
- em Xcv, sucesso de transporte e sucesso funcional são validados separadamente.

#### 5.1.1 Padrão de buffer dupla-chamada

Quase todas as APIs de enumeração do spooler (`EnumPorts`, `EnumPrinters`, `EnumPrinterDrivers`, etc.) seguem o mesmo contrato: a primeira chamada descobre quantos bytes são necessários; a segunda preenche o buffer real. Esse padrão se repete em `ListPorts`, `ListDrivers`, `ListPrintProcessors`, `ListDataTypes` e `ListQueues`.

```csharp
// 1a chamada: descobre tamanho necessário
if (!NativeMethods.EnumPorts(null, 1, IntPtr.Zero, 0, out var requiredBufferSize, out _))
{
    var error = Marshal.GetLastWin32Error();
    // ERROR_INSUFFICIENT_BUFFER é esperado aqui — não é falha real
    if (error != NativeMethods.ERROR_INSUFFICIENT_BUFFER && error != NativeMethods.ERROR_SUCCESS)
    {
        throw new Win32Exception(error, "EnumPorts buffer query failed.");
    }
}

// Sem bytes necessários = sem portas
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

    // Itera as structs PORT_INFO_1 no bloco de memória
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

O mesmo padrão é aplicado para drivers (com `DRIVER_INFO_2`), processadores (`PRINTPROCESSOR_INFO_1`), datatypes (`DATATYPES_INFO_1`) e filas (`PRINTER_INFO_2` via `EnumPrinters`).

#### 5.1.2 Criação de fila no spooler

A criação monta uma `PRINTER_INFO_2`, aloca memória não gerenciada, chama `AddPrinter` e garante liberação de recursos em qualquer saída.

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

Pontos de atenção:
- `pDevMode` e `pSecurityDescriptor` são `IntPtr.Zero` porque a criação delega esses detalhes ao driver/sistema;
- `PRINTER_ATTRIBUTE_QUEUED` ativa spooling na fila;
- `DestroyStructure` é necessário antes de `FreeHGlobal` para liberar corretamente os ponteiros internos da struct.

#### 5.1.3 Criação de porta WSD via Xcv

O fluxo de `AddWsdPort` usa o canal administrativo `XcvMonitor` e faz validação dupla: sucesso da chamada Win32 e sucesso funcional do monitor.

```csharp
public void AddWsdPort(string portName)
{
    // Idempotência: evita criar porta que já existe
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

        // Validação funcional: mesmo com chamada OK, o monitor pode rejeitar
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

O ponto central é a distinção entre **sucesso de transporte** (`XcvData` retornou `true`) e **sucesso funcional** (`xcvCommandStatus == ERROR_SUCCESS`). Isso evita falsos positivos em ambientes onde o monitor aceita o comando mas rejeita a operação.

#### 5.1.4 Update parcial de fila

O update lê o estado atual da fila, zera ponteiros sensíveis, aplica apenas campos informados e persiste via `SetPrinter`.

```csharp
public void UpdateQueue(string queueName, string? newQueueName, string? newDriverName,
    string? newPortName, string? comment, string? location)
{
    // Le estado atual completo (PRINTER_INFO_2 nível 2)
    var queueInfo = GetQueueInfo(queueName);

    // Zera ponteiros sensíveis para evitar reutilização indevida de dados nativos
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

Pontos de atenção:
- `pDevMode` e `pSecurityDescriptor` são zerados propositalmente: o `PRINTER_INFO_2` lido contém ponteiros nativos que ficam inválidos fora do buffer original. Repassa-los ao `SetPrinter` causaria corrupção de memória;
- o update é parcial: campos `null` não alteram o valor existente;
- o handle é aberto com `PRINTER_ACCESS_ADMINISTER`, necessário para `SetPrinter`.

#### 5.1.5 Inspeção heurística de fila

A classificação WPP combina três fontes de evidência: estado global de WPP, padrão do nome de porta e dados APMON (quando disponíveis).

```csharp
public QueueInspectionResult InspectQueue(string queueName)
{
    var queueInfo = GetQueueInfo(queueName);
    var portName = queueInfo.pPortName ?? string.Empty;
    var isWsdPort = portName.StartsWith("WSD", StringComparison.OrdinalIgnoreCase);
    var globalWppStatus = wppStatusProvider.GetWppStatus().Status;

    var diagnosticDetails = new List<string>();

    // Evidência opcional: APMON enriquece a classificação quando disponível
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

    // WPP global habilitado + sinais de porta moderna -> provável WPP
    if (globalWppStatus == WppStatus.Enabled && (isWsdPort || (apPortInfo?.Protocol is 1 or 2)))
    {
        classification = WppQueueClassification.LikelyWpp;
        diagnosticDetails.Insert(0, "Global WPP is enabled and queue indicates modern monitored port.");
    }
    // WPP global desabilitado prevalece como provável não-WPP
    else if (globalWppStatus == WppStatus.Disabled)
    {
        classification = WppQueueClassification.LikelyNotWpp;
        diagnosticDetails.Insert(0, "Global WPP is disabled.");
    }
    // Global ativo sem sinal de porta moderna -> inconsistência
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

A árvore de decisão resumida:

| Status Global | Porta WSD ou APMON WSD/IPP | Classificação |
|---|---|---|
| `Enabled` | Sim | `LikelyWpp` |
| `Enabled` | Não | `Indeterminate` |
| `Disabled` | Qualquer | `LikelyNotWpp` |
| `Unknown` | Qualquer | `Indeterminate` |

#### 5.1.6 Coleta de evidências APMON via XcvPort

O `TryGetApPortInfo` é um helper resiliente: abre um canal `XcvPort` para a porta e executa `GetAPPortInfo`. Se qualquer etapa falhar, retorna `null` sem interromper a inspeção.

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

    // Se XcvPort não abre, a inspeção segue sem evidência APMON
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
- `Version`: versão da estrutura;
- `Protocol`: 1 = WSD, 2 = IPP;
- `DeviceOrServiceUrl`: URL do dispositivo/serviço (260 chars fixos no layout nativo).

### 5.2 WppRegistryService

Resolve o estado global WPP lendo valores de política no Registry do Windows (HKLM).

#### 5.2.1 Definição dos probes e regra de decisão

O serviço define uma lista de "probes" — combinações de caminho de registro + nome de valor — com os valores esperados para habilitado e desabilitado. A iteração percorre todos os probes e aplica uma estratégia de fallback com prioridade para `Enabled`.

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

Essa estrutura permite adicionar novos probes sem alterar a lógica de decisão.

#### 5.2.2 Fluxo completo de `GetWppStatus`

A rotina percorre os probes em sequência. A regra é: `Enabled` retorna imediatamente, `Disabled` é guardado como candidato (porque um probe posterior pode indicar `Enabled`), e valores fora do mapeamento retornam `Unknown`.

```csharp
public WppStatusResult GetWppStatus()
{
    WppStatusResult? disabledCandidate = null;

    foreach (var statusCheck in StatusChecks)
    {
        using var key = Registry.LocalMachine.OpenSubKey(statusCheck.RegistryPath, writable: false);

        // Se a chave não existe, tenta o próximo probe
        if (key is null)
        {
            continue;
        }

        var rawValue = key.GetValue(statusCheck.PolicyValueName);

        // Se o valor não existe neste probe, tenta o próximo
        if (rawValue is null)
        {
            continue;
        }

        // Só aceita tipos que conseguimos converter com segurança
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

    // Nenhum probe encontrou valor -> assume Disabled por padrão
    return new WppStatusResult(
        WppStatus.Disabled,
        "No known WPP Registry value found.",
        "No known enabled/disabled value was found in mapped probes. Treated as disabled by default.",
        null);
}
```

Cenários de saída resumidos:

| Probe encontrado | Valor lido | Resultado |
|---|---|---|
| Sim | Igual a `EnabledWhen` | `Enabled` (retorno imediato) |
| Sim | Igual a `DisabledWhen` | Guarda candidato `Disabled`, continua |
| Sim | Outro inteiro | `Unknown` (retorno imediato) |
| Sim | Tipo não conversível | `Unknown` (retorno imediato) |
| Nenhum probe com valor | — | `Disabled` por padrão |

#### 5.2.3 Conversão segura de valor do Registry

O Registry pode armazenar o valor como `int` (DWORD) ou como `string` (texto numérico). O helper normaliza os dois formatos.

```csharp
private static bool TryConvertToInt(object rawValue, out int numericValue)
{
    // Tipo nativo comum no Registry (DWORD)
    if (rawValue is int integerValue)
    {
        numericValue = integerValue;
        return true;
    }

    // Fallback: alguns ambientes armazenam como texto numérico
    if (rawValue is string stringValue && int.TryParse(stringValue, out var parsedValue))
    {
        numericValue = parsedValue;
        return true;
    }

    numericValue = default;
    return false;
}
```

Se o tipo não for `int` nem `string` parseável (por exemplo, `byte[]` de um valor REG_BINARY), `TryConvertToInt` retorna `false` e o chamador trata como `Unknown`.

### 5.3 PrintTicketService

Lê e atualiza ticket com estratégia de compatibilidade por reflexão em `System.Printing`.

Por que essa escolha ajuda:

- reduz acoplamento forte a variações de runtime/driver;
- permite fallback mais suave quando o ambiente não suporta um caminho específico;
- retorna resultado funcional com contexto ao invés de quebrar todo o fluxo.

#### 5.3.1 Leitura de ticket (Default e User)

O fluxo de leitura verifica disponibilidade de `System.Printing` por reflexão, abre `LocalPrintServer`, localizationa a fila e extrai propriedades do ticket. A mesma estrutura é usada para `DefaultPrintTicket` e `UserPrintTicket`.

```csharp
public PrintTicketInfoResult GetDefaultTicketInfo(string queueName)
{
    var ticketSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    // Verifica se System.Printing está disponível no runtime
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

        // Extrai propriedades conhecidoas do ticket por reflexão (best-effort)
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

A versão para `GetUserTicketInfo` é idêntica, substituindo apenas `"DefaultPrintTicket"` por `"UserPrintTicket"` na obtenção da property e na chamada a `GetPrintQueue`.

#### 5.3.2 Fluxo interno de update (`UpdatePrintTicketInternal`)

Os métodos públicos `UpdateDefaultTicket` e `UpdateUserTicket` delegam para um fluxo interno comum que: prepara atributos, abre servidor/fila, aplica diferenças, faz `Commit()` e retorna pedido vs aplicado.

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
                requestedSettings, new Dictionary<string, string());

        var queueType = queue.GetType();
        var ticketPropertyInfo = queueType.GetProperty(ticketPropertyName);
        if (ticketPropertyInfo == null)
            return new PrintTicketUpdateResult(
                queueName, ticketScope, false,
                $"Print ticket property '{ticketPropertyName}' not found.",
                requestedSettings, new Dictionary<string, string());

        // 3) Lê o ticket alvo
        var targetTicket = ticketPropertyInfo.GetValue(queue);
        if (targetTicket is null)
            return new PrintTicketUpdateResult(
                queueName, ticketScope, false,
                $"Print ticket '{ticketPropertyName}' not found.",
                requestedSettings, new Dictionary<string, string());

        // 4) Aplica apenas diferenças reais
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

        // 6) Lê valores efetivos após commit e retorna pedido x aplicado
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
            requestedSettings, new Dictionary<string, string());
    }
    finally
    {
        DisposeIfPossible(queue);
        DisposeIfPossible(printServer);
    }
}
```

Regra prática: sempre comparar `Requested` e `AppliedValues` no resultado para validar o que o driver realmente aceitou.

#### 5.3.3 Escrita de atributo com detecção de mudança (`WriteTicketAttribute`)

O método compara o valor atual com o solicitado e só aplica quando há diferença real, evitando commits desnecessários.

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

O retorno `bool` é acumulado via `|=` no fluxo principal para determinar se houve pelo menos uma mudança que justifique o `Commit()`.

#### 5.3.4 Conversão de tipo para propriedades do ticket (`ConvertIfPossible`)

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

Isso permite que o consumidor envie `"TwoSidedLongEdge"` como string e o método resolva para o enum `System.Printing.Duplexing.TwoSidedLongEdge` automaticamente.

#### 5.3.5 Criação de `LocalPrintServer` e obtenção de fila por reflexão

Como `System.Printing` é carregado por reflexão, a criação do servidor e a obtenção da fila usam estratégias de fallback para maximizar compatibilidade.

**Criação do servidor:**

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
            return Activator.CreateInstance(printServerType); // Fallback: construtor padrão

        return ctor.Invoke(new[] { desiredAccess });
    }
    catch
    {
        return Activator.CreateInstance(printServerType); // Fallback final
    }
}
```

**Obtenção da fila (3 estratégias em cascata):**

```csharp
private static object? GetPrintQueue(
    Type printServerType, object? printServer, string queueName,
    string requiredTicketPropertyName, string requestedAccessName)
{
    if (printServer is null) return null;

    var desiredAccessType = Type.GetType(
        "System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);

    // Estratégia 1: new PrintQueue(printServer, queueName, desiredAccess)
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

    // Estratégia 2: printServer.GetPrintQueue(queueName, propertiesFilter)
    var getQueueWithProperties = printServerType.GetMethod(
        "GetPrintQueue", new[] { typeof(string), typeof(string[]) });
    if (getQueueWithProperties is not null)
    {
        var properties = new[] { requiredTicketPropertyName };
        return getQueueWithProperties.Invoke(printServer, new object[] { queueName, properties });
    }

    // Estratégia 3: printServer.GetPrintQueue(queueName)
    var getQueueDefault = printServerType.GetMethod(
        "GetPrintQueue", new[] { typeof(string) });
    return getQueueDefault?.Invoke(printServer, new object[] { queueName });
}
```

As 3 estratégias em ordem de prioridade:

1. **Construtor direto com acesso explícito** — garante permissão adequada para leitura/escrita;
2. **`GetPrintQueue` com filtro de propriedades** — carrega apenas as propriedades necessárias;
3. **`GetPrintQueue` simples** — fallback final quando as sobrecargas anteriores não existem.

Essa cascata torna o serviço resiliente a diferenças entre versões de runtime e configurações de ambiente.

## 6) Como interpretar os resultados

### `WppStatusResult`

- `Status`: estado final (`Enabled`, `Disabled`, `Unknown`)
- `Source`: origem da leitura
- `Details`: explicação técnica
- `RawValue`: valor bruto, quando aplicável

### `QueueInspectionResult`

- `QueueName`, `PortName`
- `GlobalWppStatus`
- `Classification`
- `Details`

Leitura prática: bom para inventário e triagem rápida de filas.

### `PrintTicketInfoResult`

- `Available`: leitura foi possível
- `Details`: contexto da tentativa
- `Attributes`: snapshot de propriedades relevantes

### `PrintTicketUpdateResult`

- `Applied`: houve alteração efetiva
- `Requested`: intenção enviada
- `AppliedValues`: resultado observado
- `Details`: mensagem final da tentativa

Leitura prática: valida diferença entre o que foi solicitado e o que o driver realmente sustentou.

## 7) Exemplo único de execução

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

Observação: em cenário real, valide listas antes de indexar e trate exceções operacionais.

## 8) Erros comuns e leitura correta

- `ERROR_ACCESS_DENIED`: falta privilégio para operação administrativa.
- `ERROR_INSUFFICIENT_BUFFER`: comportamento esperado na primeira chamada de enumeração.
- `ERROR_NOT_SUPPORTED`: comando não suportado no monitor/driver/porta daquele ambiente.

## 9) Limites conhecidos

- classificação WPP por fila é heurística, não certificação formal;
- resultado depende de versão de Windows, driver e permissão local;
- nem todo update de ticket é aceito integralmente pelo driver;
- ambientes diferentes podem acionar fallbacks diferentes em `System.Printing`.

## 10) Checklist rápido de validação

1. consultar status global WPP;
2. listar recursos disponíveis (ports, drivers, processors, datatypes);
3. criar fila com combinação válida;
4. inspecionar fila;
5. ler ticket default e user;
6. aplicar update e comparar pedido x aplicado;
7. atualizar metadados da fila;
8. remover fila e confirmar limpeza.

## 11) POC

Link da POC: https://github.com/tiagobuchanellindd/WppQueuePoc