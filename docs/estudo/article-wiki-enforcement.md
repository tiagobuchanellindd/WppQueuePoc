# Enforcement ativo de PrintTicket via Win32 API: arquitetura de monitoramento reativo

Este estudo apresenta uma solução de enforcement ativo para configurações de impressão no Windows. Ele combina monitoramento de eventos via Win32 API com atualização automática de PrintTicket para manter configurações alinhadas com políticas organizacionais.

## 1) Contexto e problema

Em ambientes corporativos, há necessidade de garantir que impressoras mantenham configurações fixes, como por exemplo:

- duplex forçado (frente e verso)
- cor forçada (monochrome)
- orientação fixa (portrait ou landscape)

Usuários podem alterar essas configurações manualmente via driver ou diálogo de impressão, e a organização precisa reverter essas mudanças para garantir conformidade.

Abordagens tradicionais:

- GPO (Group Policy): funciona para políticas de driver, mas não para todas as propriedades
- Logon script: executa uma vez, não reage a mudanças subsequentes
- Monitoramento manual: dependência de operação constante

Solução proposta: **Enforcement Ativo** monitore eventos da impressora e restaure configurações automaticamente.

## 2) Visão geral da solução

A solução é composta por dois componentes principais:

- **PrinterPolicyEnforcer**: monitora eventos de mudança via Win32 API e coordena o fluxo de enforcement
- **PrintTicketEnforcementHelper**: helper estático que compara configurações atuais com a política e aplica correções

### Fluxo conceitual

```
[Impressora] --(evento de mudança)--> [MonitorLoop]
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

### Características principais

- Reage a eventos ativos (não apenas polling periódico)
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

Cada dimensão tem:

- flag (EnforceX): habilita monitoramento forçar
- valor (RequiredXValue): valor que deve ser restaurado

Valores null ou EnforceX=false significa "ignorar esta propriedade".

### 3.2 Loop duplo

A arquitetura usa dois loops assíncronos separados:

**MonitorLoop** (thread 1):

- Espera eventos via Win32 API (bloqueante em WaitForSingleObject)
- Quando detecta mudança, marca flag pendente
- Não pode ser bloqueado por operações lentas

**EnforcementWorkerLoop** (thread 2):

- Poll a cada 500ms
- verifica flag pendente
- executa enforcement (leitura, comparação, escrita)
- Atualiza timestamp para debounce

Benefício: monitoramento permanece responsivo mesmo durante processamento de enforcement.

### 3.3 Flag e lock thread-safe

Comunicação entre loops:

```csharp
private volatile bool _flagsEnforcementPending = false;
private readonly object _enforcementLock = new();
```

- `volatile`: garante visibilidade entre threads
- `lock`: protege acesso atômico

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
Mudança detecta -> Enforcement aplicado -> Driver reporta Nova mudança -> Enforcement aplicado -> ...
```

## 4) PrintTicketEnforcementHelper: lógica de comparação e aplicação

### 4.1 Helper estático para testabilidade

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

Benefícios:

- Sem estado, sem necessidade de instânciação
- Pode ser testado com mock de IPrintTicketService
- Reuso em diferentes contextos

### 4.2 Fluxo de comparação

```csharp
var info = service.GetDefaultTicketInfo(queueName);
if (!info.Available)
    return failure("Não foi possível ler: " + info.Details);

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

Comparação case-insensitive: drivers podem retornar valores em diferentes capitalizações.

### 4.3 Acúmulo de mudanças

```csharp
if (!requiresUpdate)
    return success("Já em conformidade", attemptedChange: false);

var request = new PrintTicketUpdateRequest(
    changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
    changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
    changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);

var update = service.UpdateDefaultTicket(queueName, request);
```

Única chamada de update para múltiplas propriedades (melhor performance).

### 4.4 Resultado estruturado

```csharp
public sealed record PrintTicketEnforcementResult(
    bool Success,        // Operação global
    string Details,    // Mensagem descritiva
    bool AttemptedChange);  // Tentou mudar?
```

Casos:

```
Success=true, AttemptedChange=false  -> Já estava conforme
Success=true, AttemptedChange=true  -> Mudou com sucesso
Success=false, AttemptedChange=true -> Tentou e falhou
```

## 5) Integração com Win32 API

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

Espera até timeout ou sinal.

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

Avança estado e retorna flags do evento.

### 5.4 Flags de monitoramento

```csharp
PrinterChangeNotificationNative.PRINTER_CHANGE_SET_PRINTER
```

Detecta qualquer alteração em configurações da impressora.

## 6) Fluxo de uso

### Exemplo: forçar duplex e monochrome

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
// [Monitor] Mudança capturada na impressora: 0x4.
// [Monitor] Enforcement pendente.
// [Enforcement] Enforcement realizado com sucesso.
//   Duplexing: OneSided -> TwoSidedLongEdge
//   OutputColor: Color -> Monochrome
```

## 7) Tratamento de erros

### Falta de permissão

Se o processo não tem permissão de Manage Printers:

```
[Enforcement][ERRO] Exception during PrintTicket update
```

Erro logado, NÃO lança exceção.

### Driver não suporta

Se driver não suporta propriedade:

```
Success=false, Details="Falha ao aplicar enforcement: Driver não suporta..."
```

### Handle inválido

```
[Monitor] Falha ao abrir a impressora 'Nome'.
```

## 8) Eventos expostos

```csharp
public event EventHandler<string>? StatusChanged;    // Start/stop
public event EventHandler<string>? EnforcementLog;  // Operações
public event EventHandler<Exception>? Error;        // Exceções
```

## 9) Limites e observações

- Requer permissão de gerenciamento na impressora (Manage Printers)
- Não detecta mudança de driver (apenas propriedades)
- Cada impressora precisa de instância separada
- Debounce de 3s pode ser ajuste fino
- Serviço precisa de IPrintTicketService válido

## 10) Comparação com abordagens tradicionais

| Abordagem | Reage a mudança | Automatização | Complexidade |
|---|---|---|---|
| GPO | Parcial | Uma vez | Média |
| Logon script | Não | Uma vez | Baixa |
| Monitoring manual | Não | Não | Baixa |
| Enforcement Ativo | Sim | Contínua | Média |

## 11) Checkpoint de implantação

1. Verificar permissão de Manage Printers na fila
2. Testar IPrintTicketService isoladamente
3. Configurar Policy com valores desejados
4. Iniciar enforcer e verificar eventos
5. Ajustar debounce se necessário
6. Configurar log para produção
7. Planejar cleanup em shutdown

## 12) Estudo relacionado

Este estudo complementa o estudo geral de PrintTicketService (docs/estudo/PrintTicketService).
O PrintTicketService fornece as capacidades de leitura/escrita que este enforcement consome.

## 13) POC

Link da POC: https://github.com/tiagobuchanellindd/WppQueuePoc

## 14) Testes de Monitoramento
![teste.gif](/.attachments/teste-a0ff081d-1f27-4176-9957-3d4de9678f9a.gif)