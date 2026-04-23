# Enforcement ativo de PrintTicket via Win32 API: arquitetura de monitoramento reativo

Este estudo apresenta uma solucao de enforcement ativo para configuracoes de impressao no Windows. Ele combina monitoramento de eventos via Win32 API com atualizacao automatica de PrintTicket para manter configuracoes alinhadas com politicas organizacionais.

## 1) Contexto e problema

Em ambientes corporativos, ha necessidade de garantir que impressoras mantenham configuracoesfixas:

- duplex forcado (frente e verso)
- cor forcada (monochrome)
- orientacao fixa (portrait ou landscape)

Usuarios podem alterar essas configuracoes manualmente via driver ou dialogo de impressao, e a organizacao precisa reverter essas mudancas para garantir conformidade.

Abordagens tradicionais:

- GPO (Group Policy): funciona para politicas de driver, mas nao para todas as propriedades
- Logon script: executa uma vez, nao reage a mudancas subsequentes
- Monitoramento manual: dependencia de operacao constante

Solucao proposta: **Enforcement Ativo** monitore eventos da impressora e restaure configuracoes automaticamente.

## 2) Visao geral da solucao

A solucao e composta por dois componentes principais:

- **PrinterPolicyEnforcer**: monitora eventos de mudanca via Win32 API e coordena o fluxo de enforcement
- **PrintTicketEnforcementHelper**: helper estatico que compara configuracoes atuais com a politica e aplica correcoes

### Fluxo conceitual

```
[Impressora] --(evento de mudanca)--> [MonitorLoop]
                                          |
                                          V
                                  [Flag pendente]
                                          |
                                          V
                              [EnforcementWorkerLoop]
                                          |
                                          V
                              [PrintTicketEnforcementHelper]
                                          |
                                          V
                              [IPrintTicketService]
                                          |
                                          V
                              [PrintTicket atualizado]
```

### Caracteristicas principais

- Reage a eventos ativos (nao apenas polling periodico)
- Dois loops separados para responsividade
- Debounce de 3 segundos para evitar loops infinitos
- Falhas resilientes com log descritivo

## 3) PrinterPolicyEnforcer: arquitetura e componentes

### 3.1 Classe Policy

A Policy define o contrato de enforcement:

```csharp
public class Policy
{
    public bool EnforceDuplex { get; set; }
    public string? RequiredDuplexValue { get; set; }  // "TwoSidedLongEdge"

    public bool EnforceColor { get; set; }
    public string? RequiredColorValue { get; set; }  // "Monochrome"

    public bool EnforceOrientation { get; set; }
    public string? RequiredOrientationValue { get; set; }  // "Portrait"
}
```

Cada dimensiao tem:

- flag (EnforceX): habilita monitoramento forcar
- valor (RequiredXValue): valor que deve ser restaurado

Valores null ou EnforceX=false significa "ignorar esta propriedade".

### 3.2 Loop duplo

A arquitetura usa dois loops asynconos separados:

**MonitorLoop** (thread 1):

- Espera eventos via Win32 API (bloqueante em WaitForSingleObject)
- Quando detecta mudanca, marca flag pendente
- Nao pode ser bloqueado por operacoes lentas

**EnforcementWorkerLoop** (thread 2):

- Poll a cada 500ms
- verifica flag pendente
- executa enforcement (leitura, comparacao, escrita)
- Atualiza timestamp para debounce

Beneficio: monitoramento permanece responsivo mesmo durante processamento de enforcement.

### 3.3 Flag e lock thread-safe

Comunicacao entre loops:

```csharp
private volatile bool _flagsEnforcementPending = false;
private readonly object _enforcementLock = new();
```

- `volatile`: garante visibilidade entre threads
- `lock`: protege acesso atomico

```csharp
lock (_enforcementLock)
{
    if (_flagsEnforcementPending)
    {
        _flagsEnforcementPending = false;
        shouldEnforce = true;
    }
}
```

### 3.4 Debounce

Evita loops infinitos de re-enforcement:

```csharp
private readonly TimeSpan _debounceInterval = TimeSpan.FromSeconds(3);

var now = DateTime.UtcNow;
if (_lastEnforcementUtc.HasValue && now - _lastEnforcementUtc.Value < _debounceInterval)
{
    // Ignora (debounce)
}
```

Ciclo evitado:

```
Mudanca detecta -> Enforcement aplicado -> Driver reporta Nova mudanca -> Enforcement aplicado -> ...
```

## 4) PrintTicketEnforcementHelper: logica de comparacao e aplicacao

### 4.1 Helper estatico para testabilidade

```csharp
public static class PrintTicketEnforcementHelper
{
    public static PrintTicketEnforcementResult EnforceDefaultTicketPolicy(
        IPrintTicketService service,
        string queueName,
        PrinterPolicyEnforcer.Policy policy)
    {
        // ...
    }
}
```

Beneficios:

- Sem estado, sem necessidade de instanciacao
- Pode ser testado com mock de IPrintTicketService
- Reuso em diferentes contextos

### 4.2 Fluxo de comparacao

```csharp
var info = service.GetDefaultTicketInfo(queueName);
if (!info.Available)
    return failure("Nao foi possivel ler: " + info.Details);

var changes = new Dictionary<string, string?>();
bool requiresUpdate = false;

if (policy.EnforceDuplex && policy.RequiredDuplexValue != null)
{
    info.Attributes.TryGetValue("Duplexing", out var curValue);
    if (!string.Equals(curValue, policy.RequiredDuplexValue, StringComparison.OrdinalIgnoreCase))
    {
        changes["Duplexing"] = policy.RequiredDuplexValue;
        requiresUpdate = true;
    }
}
```

Comparacao case-insensitive: drivers podem retornar valores em diferentes capitalizacoes.

### 4.3 Acumulo de mudancas

```csharp
if (!requiresUpdate)
    return success("Ja em conformidade", attemptedChange: false);

var request = new PrintTicketUpdateRequest(
    changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
    changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
    changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);

var update = service.UpdateDefaultTicket(queueName, request);
```

Unica chamada de update para multiplas propriedades (melhor performance).

### 4.4 Resultado estruturado

```csharp
public sealed record PrintTicketEnforcementResult(
    bool Success,        // Operacao global
    string Details,    // Mensagem descritiva
    bool AttemptedChange);  // Tentou mudar?
```

Casos:

```
Success=true, AttemptedChange=false  -> Ja estava conforme
Success=true, AttemptedChange=true  -> Mudou com sucesso
Success=false, AttemptedChange=true -> Tentou e falhou
```

## 5) Integracao com Win32 API

### 5.1 FindFirstPrinterChangeNotification

```cpp
HANDLE FindFirstPrinterChangeNotification(
    HANDLE hPrinter,
    DWORD  fdwFlags,     // PRINTER_CHANGE_SET_PRINTER
    DWORD  fdwOptions,
    LPVOID pNotifyInfo);
```

Registra handle que sinaliza quando evento ocorre.

### 5.2 WaitForSingleObject

```cpp
DWORD WaitForSingleObject(
    HANDLE hHandle,
    DWORD  dwMilliseconds);  // timeout
```

Espera ate timeout ou sinal.

Retornos:

- WAIT_OBJECT_0: evento ocorreu
- WAIT_TIMEOUT: timeout expirou
- WAIT_FAILED: erro

### 5.3 FindNextPrinterChangeNotification

```cpp
BOOL FindNextPrinterChangeNotification(
    HANDLE hNotify,
    PDWORD pfdwChange,
    ...);
```

Avanca estado e retorna flags do evento.

### 5.4 Flags de monitoramento

```csharp
PrinterChangeNotificationNative.PRINTER_CHANGE_SET_PRINTER
```

Detecta qualquer alteracao em configuracoes da impressora.

## 6) Fluxo de uso

### Exemplo: forcar duplex e monochrome

```csharp
var policy = new PrinterPolicyEnforcer.Policy
{
    EnforceDuplex = true,
    RequiredDuplexValue = "TwoSidedLongEdge",
    EnforceColor = true,
    RequiredColorValue = "Monochrome"
};

var printTicketService = new PrintTicketService();
var enforcer = new PrinterPolicyEnforcer(policy, printTicketService);

enforcer.EnforcementLog += (s, msg) => Console.WriteLine(msg);
enforcer.Start("MinhaImpressora");

Console.ReadLine();

enforcer.Stop();
enforcer.Dispose();
```

### Exemplo: monitorar eventos e verificar enforcement

```csharp
enforcer.EnforcementLog += (s, msg) =>
{
    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {msg}");
};

// Log esperado:
// [Monitor] Aguardando eventos em 'MinhaImpressora'.
// [Monitor] Mudanca capturada na impressora: 0x4.
// [Monitor] Enforcement pendente.
// [Enforcement] Enforcement realizado com sucesso.
//   Duplexing: OneSided -> TwoSidedLongEdge
//   OutputColor: Color -> Monochrome
```

## 7) Tratamento de erros

### Falta de permissao

Se o processo nao tem permissao de Manage Printers:

```
[Enforcement][ERRO] Exception during PrintTicket update
```

Erro loggedo, NAO lanca excecao.

### Driver nao suporta

Se driver nao suporta propriedade:

```
Success=false, Details="Falha ao aplicar enforcement: Driver nao suporta..."
```

### Handle invalido

```
[Monitor] Falha ao abrir a impressora 'Nome'.
```

## 8) Eventos expostos

```csharp
public event EventHandler<string>? StatusChanged;    // Start/stop
public event EventHandler<string>? EnforcementLog;  // Operacoes
public event EventHandler<Exception>? Error;        // Excecoes
```

## 9) Limites e observacoes

- Requer permissao de gerenciamento na impressora (Manage Printers)
- Nao detecta mudanca de driver (apenas propriedades)
- Cada impressora precisa de instancia separada
- Debounce de 3s pode ser ajustefino
- Servico precisa de IPrintTicketService valido

## 10) Comparacao com abordagens tradicionais

| Abordagem | Reage a mudanca | Automatizacao | Complexidade |
|---|---|---|---|
| GPO | Parcial | Uma vez | Media |
| Logon script | Nao | Uma vez | Baixa |
| Monitoring manual | Nao | Nao | Baixa |
| Enforcement Ativo | Sim | Contínua | Media |

## 11) Checkpoint de implantacao

1. Verificar permissao de Manage Printers na fila
2. Testar IPrintTicketService isoladamente
3. Configurar Policy com valores desejados
4. Iniciar enforcer e verificar eventos
5. Ajustar debounce se necessario
6. Configurar log para producao
7. Planejar cleanup em shutdown

## 12) Estudo relacionado

Este estudo complementa o estudo geral de PrintTicketService (docs/estudo/PrintTicketService). O PrintTicketService fornece as capacidades de leitura/escrita que este enforcement consome.

## 13) POC

Link da POC: https://github.com/tiagobuchanellindd/WppQueuePoc