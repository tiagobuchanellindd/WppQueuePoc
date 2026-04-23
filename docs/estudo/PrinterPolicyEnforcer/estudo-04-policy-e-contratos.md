# Estudo 04 - Policy e contratos de enforcement

## Objetivo deste estudo

Entender a estrutura `Policy` que define quais propriedades devem ser monitoradas e seus valores-alvo.

## O que e a Policy

Policy e uma classe interna (nested) de PrinterPolicyEnforcer que define o "contrato" de enforcement:

```csharp
public class Policy
{
    public bool EnforceDuplex { get; set; }
    public string? RequiredDuplexValue { get; set; }
    public bool EnforceColor { get; set; }
    public string? RequiredColorValue { get; set; }
    public bool EnforceOrientation { get; set; }
    public string? RequiredOrientationValue { get; set; }
}
```

## Semantica

Cada propriedade tem dois campos:

- **EnforceX** (bool): Define se essa dimensao esta ativa
- **RequiredXValue** (string?): Define o valor que deve ser forcado

Exemplo:

```csharp
var policy = new Policy
{
    EnforceDuplex = true,
    RequiredDuplexValue = "TwoSidedLongEdge",
    EnforceColor = true,
    RequiredColorValue = "Monochrome"
};
```

Traduzindo:

- "Duplex deve ser TwoSidedLongEdge" (frente e verso no eixo longo)
- "Cor deve ser Monochrome" (preto e branco)

## Por que dois campos?

Permite controle granular:

1. **Sem Enable booleano**: O valor poderia ser sempre string, mas ai como desativar?
   - Usar null como "desligado" conflita com null como valor invalido

2. **Separacao clara**: EnforceX=true significa "esta dimensio importa"
   - EnforceX=false significa "ignore esta propriedade"

3. **Extensibilidade**: Facil adicionar novas dimensoes sem mudar logica existente

## Dimensoes suportadas

### Duplexing (frente e verso)

Valores típicos (definidos pelo driver, mas convencionais):

- `"OneSided"` - Impressao em uma face
- `"TwoSidedLongEdge"` - Frente e verso, flip no eixo longo
- `"TwoSidedShortEdge"` - Frente e verso, flip no eixo curto

### OutputColor (cor)

- `"Color"` - Impressao colorida
- `"Monochrome"` - Apenas preto e branco

### PageOrientation (orientacao)

- `"Portrait"` - Retrato
- `"Landscape"` - Paisagem

## Como a Policy e usada

A Policy e passada no construtor e armazenada:

```csharp
private readonly Policy _policy;

public PrinterPolicyEnforcer(Policy policy, IPrintTicketService printTicketService)
{
    _policy = policy;
    _printTicketService = printTicketService;
}
```

Depois e usada no EnforcementWorkerLoop:

```csharp
var enforcement = PrintTicketEnforcementHelper.EnforceDefaultTicketPolicy(
    _printTicketService, _printerName, _policy);
```

Veja estudio PrintTicketEnforcementHelper para detalhes.

## Validacao

A Policy NAO faz validacao internamente. Validacaoes sao feitas pelo helper:

```csharp
if (policy.EnforceDuplex && policy.RequiredDuplexValue != null)
{
    // compara e atualiza
}
```

Isso permite que chamadas inválidas simplemente não façam nada (se Required for null ou se Enforce=false).

## Exemplo de uso completo

```csharp
// Criar politica: forcar duplex e monochrome
var policy = new PrinterPolicyEnforcer.Policy
{
    EnforceDuplex = true,
    RequiredDuplexValue = "TwoSidedLongEdge",
    EnforceColor = true,
    RequiredColorValue = "Monochrome"
};

var enforcer = new PrinterPolicyEnforcer(policy, printTicketService);

enforcer.EnforcementLog += (s, msg) => Console.WriteLine(msg);
enforcer.Start("MinhaImpressora");
```

Resultado: Quando usuario altera configuracao da impressora, componente detecta evento e restaura para duplex + monochrome.

## Extensibilidade futura

Para adicionar nova dimensao,只需要:

1. Adicionar campos na Policy:
   ```csharp
   public bool EnforceStapling { get; set; }
   public string? RequiredStaplingValue { get; set; }
   ```

2. Adicionar comparacao no helper:
   ```csharp
   if (policy.EnforceStapling && policy.RequiredStaplingValue != null)
   ```

3. Documentar valores aceitos

## Resumo

A estrutura Policy define contratos de enforcement de forma declarativa:

- Cada dimensiao tem flag de habilitacao + valor requerido
- Valores null ou flag=false significa "ignorar"
- Separação clara entre configuracao (Policy) e execucao (Enforcer)