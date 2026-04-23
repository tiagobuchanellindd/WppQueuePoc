# Estudo 02 - Fluxo de comparacao

## Objetivo deste estudo

Entender como o helper compara os valores atuais do PrintTicket com os valores definidos na Policy.

## Passo a passo do fluxo

```csharp
public static PrintTicketEnforcementResult EnforceDefaultTicketPolicy(
    IPrintTicketService service,
    string queueName,
    PrinterPolicyEnforcer.Policy policy)
{
    // 1. Le ticket atual
    var info = service.GetDefaultTicketInfo(queueName);
    if (!info.Available)
        return failure("Nao foi possível ler PrintTicket: " + info.Details);

    // 2. Coleta mudancas necessarias
    var changes = new Dictionary<string, string?>();
    bool requiresUpdate = false;

    // 3. Compara duplex
    if (policy.EnforceDuplex && policy.RequiredDuplexValue != null)
    {
        info.Attributes.TryGetValue("Duplexing", out var curValue);
        if (!string.Equals(curValue, policy.RequiredDuplexValue, StringComparison.OrdinalIgnoreCase))
        {
            changes["Duplexing"] = policy.RequiredDuplexValue;
            requiresUpdate = true;
        }
    }

    // 4. Compara color (similar...)
    // 5. Compara orientation (similar...)

    // 6. Se什么都没有 alterado, retorna "ja conforme"
    if (!requiresUpdate)
        return success("Ja em conformidade", attemptedChange: false);

    // 7. Se diferenca, aplica update
    var update = service.UpdateDefaultTicket(queueName, request);

    // 8. Retorna resultado
    return ...
}
```

## Detalhes da comparacao

### Leitura do ticket

```csharp
var info = service.GetDefaultTicketInfo(queueName);
```

O `IPrintTicketService` retorna um objeto com:

- `bool Available`: Se o ticket foi lido com sucesso
- `Dictionary<string, string> Attributes`: Atributos capturados

Exemplo de Attributes:

```
{
    "Duplexing": "OneSided",
    "OutputColor": "Color",
    "PageOrientation": "Portrait",
    ...
}
```

### Comparacao com string.Equals

```csharp
if (!string.Equals(curValue, policy.RequiredDuplexValue, StringComparison.OrdinalIgnoreCase))
{
    // diferente!
}
```

Por que `OrdinalIgnoreCase`?

1. Drivers podem retornar `"TwoSidedLongEdge"`, `"twosidedlongedge"`, `"TWOSIDEDLONGEDGE"`
2. Todos sao semanticamente equivalentes
3. Comparacao case-insensitive evita "falso negativo"

### Comparacao "diferente de null"

Note a condicao:

```csharp
if (policy.EnforceDuplex && policy.RequiredDuplexValue != null)
```

Dois checks:

- `EnforceDuplex == true`: Esta dimensio esta ativa
- `RequiredDuplexValue != null`: Foi fornecido valor

Se qualquer um for false/null, a comparacao e ignorada.

## Acumulo de mudancas

O fluxo usa um Dictionary para acumular mudancas necessarias:

```csharp
var changes = new Dictionary<string, string?>();

if (diferente_em_duplex) changes["Duplexing"] = "TwoSidedLongEdge";
if (diferente_em_color)  changes["OutputColor"] = "Monochrome";
if (diferente_em_orientation) changes["PageOrientation"] = "Landscape";
```

Beneficio: **Unica chamada de update** para multiplas propriedades.

Isso e importante porque:

- Cada update pode falhar
- Update tem overhead (leitura, escrita, commit)
- melhor uma operacao que varias

## Criacao do request

```csharp
var request = new PrintTicketUpdateRequest(
    changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
    changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
    changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);
```

O request so inclui as propriedades que precisam mudanca (null se nao mudou).

## Por que "ja em conformidade"?

```csharp
if (!requiresUpdate)
    return new PrintTicketEnforcementResult(
        true,
        "Ja em conformidade com a política.",
        false);  // AttemptedChange = false
```

Isso permite ao chamador distinguir:

- Ja estava conforme -> Success=true, AttemptedChange=false (ok!)
- Forcar alteracao com sucesso -> Success=true, AttemptedChange=true (ok!)
- Tentou e falhou -> Success=false, AttemptedChange=true (problema!)

## Resumo do fluxo

```
Le ticket atual
    │
    ▼
Para cada dimensiao (duplex, color, orientation):
    │─ Se EnforceX && RequiredX != null
    │     │
    │     ▼
    │   Compara (case-insensitive)
    │     │
    │     ▼
    │   Se diferente: adiciona ao Dictionary de mudancas
    │
    ▼
Se zero mudancas: retorna "ja conforme"
    │
    ▼
Se mudancas: faz UNICA chamada de UpdateDefaultTicket
    │
    ▼
Retorna resultado estruturado
```

Design que minimiza operacoes e reporta claramente o que aconteceu.