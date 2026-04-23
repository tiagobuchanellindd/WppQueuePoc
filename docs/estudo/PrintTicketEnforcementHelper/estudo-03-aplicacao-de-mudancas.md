# Estudo 03 - Aplicacao de mudancas

## Objetivo deste estudo

Entender como o helper aplica as mudancas quando detecta divergencia entre valor atual e requerido.

## Transicao: comparacao -> aplicacao

Apos identificar mudancas necessarias (estudo 02), o fluxo aplica:

```csharp
if (!requiresUpdate)
    return success("Ja em conformidade", false);

// Ha diferenca! Criar request e aplicar
var request = new PrintTicketUpdateRequest(
    changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
    changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
    changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);

var update = service.UpdateDefaultTicket(queueName, request);

// Avaliar resultado
if (update.Applied)
    return success("Enforcement realizado...", true);
else
    return failure("Falha ao aplicar...", true);
```

## Criacao do request

```csharp
var request = new PrintTicketUpdateRequest(
    duplexValue,
    colorValue,
    orientationValue);
```

O `PrintTicketUpdateRequest` e um DTO com propriedades opcionais.

Valores que nao estao em changes vem como null:
- Null = "nao modifique esta propriedade"
- String = "aplique este valor"

Beneficio: O service de update so modifica o que foi pedido explicitamente.

## Chamada de update

```csharp
var update = service.UpdateDefaultTicket(queueName, request);
```

O IPrintTicketService:

1. Abre a fila (comoAdmin)
2. Le o DefaultPrintTicket atual
3. Aplica apenas as mudancas requests
4. Re-atribui o ticket
5. Faz Commit()
6. Retorna resultado

Detalhes em PrintTicketService estudio.

## Avaliacao do resultado

```csharp
if (update.Applied)
{
    // Sucesso! Montar mensagem detalhada
    var sb = new System.Text.StringBuilder();
    sb.AppendLine("Enforcement realizado com sucesso.");
    if (changes.ContainsKey("Duplexing"))
        sb.AppendLine($"  Duplexing: {info.Attributes["Duplexing"]} → {policy.RequiredDuplexValue}");
    // ...

    return new PrintTicketEnforcementResult(true, sb.ToString(), true);
}
else
{
    // Falha
    return new PrintTicketEnforcementResult(false, "Falha ao aplicar enforcement: " + update.Details, true);
}
```

## Por que duas avaliacoes?

`update.Applied` indica se a chamada de update funcionou.

Mas ha nuance:

1. **update.Applied == true**: Ticket foi escrito com sucesso
   - pode ser que driver tenha ignorado某些 valores
   - mas a operacao de set/commitNAO falhou

2. **update.Applied == false**: Operacao falhou
   - pode ser permissao insuficiente
   - pode ser driver nao suporta
   - pode ser queue nao existe

O helper propaga o detalhamento via Details.

## Mensagem detalhada

Quando bem-sucedido, a mensagem inclui diff detalhado:

```
Enforcement realizado com sucesso.
  Duplexing: OneSided → TwoSidedLongEdge
  OutputColor: Color → Monochrome
```

Isso ajuda diagnostico em producao.

Note: A messaging mostra valor ANTES (info.Attributes) e DEPOIS (policy.RequiredXValue).

Por que?

- O ticket lido pode ja estar diferente do que foi pedido
- diff mostra o que realmente mudou

## AttemptedChange

Sempre true quando chega nesta parte do fluxo:

```csharp
return new PrintTicketEnforcementResult(
    success, 
    message,
    true);  // AttemptedChange = true -至少有 Tentativa
```

Isso diferencia:

- Ja estava conforme (AttemptedChange=false)
- Tentou e conseguiu (AttemptedChange=true)
- Tentou e falhou (AttemptedChange=true)

Chamador pode usar isso para metricas.

## Tratamento de erro

Se `service.UpdateDefaultTicket` lancar excecao:

```csharp
try
{
    var update = service.UpdateDefaultTicket(queueName, request);
    // ...
}
catch (Exception ex)
{
    return failure("Exceção: " + ex.Message, attempted: true);
}
```

Mas na pratica isso nao acontece porque IPrintTicketService retorna resultado, nao lana excecao (design resiliente).

## Resumo da aplicacao

```
Ja identificou mudancas (requiresUpdate = true)
    │
    ▼
Criar PrintTicketUpdateRequest apenas com mudancas
    │
    ▼
Chamar IPrintTicketService.UpdateDefaultTicket()
    │
    ▼
Avaliar update.Applied
    │
    ├─ true: Montar mensagem com diff, return Success=true
    │
    └─ false: Propagar Details do erro, return Success=false
    │
    ▼
AttemptedChange = true (porque chegou aqui)
```

Fluxo que aplica apenas o necessario e reporta claramente.