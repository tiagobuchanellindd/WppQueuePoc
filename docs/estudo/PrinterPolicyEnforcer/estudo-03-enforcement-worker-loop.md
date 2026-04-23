# Estudo 03 - EnforcementWorkerLoop

## Objetivo deste estudo

Entender como o EnforcementWorkerLoop processa a flag pendente e executa o enforcement.

## Separacao de responsabilidade

Por que separar MonitorLoop de EnforcementWorkerLoop?

### Motivos arquiteturais

1. **MonitorLoop** precisa ser o mais rapidopossivel para NAO perder eventos.
   - Espera bloqueante em I/O Win32
   - Precisa retornar rapidamente para esperar proximo evento

2. **EnforcementWorkerLoop** pode demorar (leitura de ticket, escrita, comparacao).
   - Operacoes potencialmente lentas
   - Pode falhar, precisa tratamento de erro

Se o mesmo loop fizesse ambos, um enforcement demorado faria o monitor acumular eventos ou perder timing.

### Solucao: Dois loops asynchronous separados

```
MonitorLoop (Thread 1)                    EnforcementWorkerLoop (Thread 2)
+------------------+                    +------------------------+
| Win32 wait       |                    | poll flag each 500ms    |
| Evento -> flag   |  ---> flag ------> | if pending -> process     |
+------------------+                    +------------------------+
```

## Implementacao do loop

```csharp
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
            // Executa enforcement
            var enforcement = PrintTicketEnforcementHelper.EnforceDefaultTicketPolicy(
                _printTicketService, _printerName, _policy);
            // ...
        }

        await Task.Delay(500, ct);  // Poll interval
    }
}
```

## Flag e lock

```csharp
private volatile bool _flagsEnforcementPending = false;
private readonly object _enforcementLock = new();
```

- `volatile`: Garante que a leitura/escrita nao seja reordenada pelo compilador
- `lock`: Protege acesso atomicto ao recurso compartilhado

O lock garante que:

- Se MonitorLoop marca flag, EnforcementWorkerLoop ve a marcação
- Nao ha race condition entre check e clear

## Poll interval: 500ms

O loop verifica a cada 500ms se ha trabalho pendente.

Consideracoes:

- 500ms e razoavelmente curto para reagir a mudanca
- Tambem nao sobrecarrega com verificacoes muito frequentes
- Intervalo maior (ex: 1s) aumentaria latencia
- Intervalo menor (ex: 100ms) gastaria mais CPU

Trade-off entre latencia e overhead.

## Pattern "consumer unico"

A flag so pode ser consumida uma vez:

```csharp
if (_flagsEnforcementPending)
{
    _flagsEnforcementPending = false;  // Limpa ANTES de processar
    shouldEnforce = true;
}
```

Isso evita duplicates se múltiplos eventos ocorrerem entre polls.

Mas espera: se varios eventos ocorrem entre polls, apenas uma execucao ocorre. Isso e intencional (debounce).

## Execucao do enforcement

```csharp
var enforcement = PrintTicketEnforcementHelper.EnforceDefaultTicketPolicy(
    _printTicketService,
    _printerName,
    _policy);
```

O metodo:

1. Le DefaultPrintTicket atual
2. Compara com politica (_policy)
3. Aplica atualizacao se necessario
4. Retorna resultado estruturado

Ver estudo 02-04 para detalhes do helper.

## Logging e erro

```csharp
if (shouldEnforce)
{
    try
    {
        // ...
        EnforcementLog?.Invoke(this, "[Enforcement] " + enforcement.Details);
    }
    catch (Exception ex)
    {
        // Excecao poco provavel (helper retorna resultado), mas tratar por via das dudas
        EnforcementLog?.Invoke(this, "[Enforcement][ERRO] " + ex.Message);
    }
}
```

Erros sao capturados e logados, NAO propagation como excecao.

## Atualizacao de timestamp

```csharp
_lastEnforcementUtc = DateTime.UtcNow;
```

Usado pelo MonitorLoop para debounce:

```csharp
if (_lastEnforcementUtc.HasValue && now - _lastEnforcementUtc.Value < _debounceInterval)
{
    // Ignora (debounce)
}
```

## Resumo

O EnforcementWorkerLoop implementa um pattern de consumidorperiodico:

1. Poll a cada 500ms
2. Se ha flag pendente, limpa e processa (so uma vez)
3. Executa helper de enforcement
4. Atualiza timestamp para debounce
5. Loga resultado

Separacao permite que monitoramento seja rapido enquanto enforcement processa em paralelo.