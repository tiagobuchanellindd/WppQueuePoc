# Estudo 01 - Visao geral do PrinterPolicyEnforcer

## O que esta classe faz

`PrinterPolicyEnforcer` monitora eventos de mudanca em uma impressora e automaticamente restaura configuracoes (PrintTicket) para valores definidos em uma politica.

## Contexto do problema

Em ambientes corporativos, often ha necessidade de garantir que impressoras mantenham configuracoesfixas (ex: forcar duplex, forcar monochrome, forcar orientacao). Usuarios ou drivers podemalterar essas configuracoes manualmente, e a organizacao precisa reverter essas mudancas.

## Pattern empregados: "Enforcement Ativo"

O nome "Enforcement Ativo" indica que o componente NAO apenas le configuracoes, mas monitora eventos ativos e toma acao quando mudancas saodetectadas.

## Arquitetura de loop duplo

A classe usa dois loops asynconos separados:

1. **MonitorLoop**: Bloqueante em I/O (WaitForSingleObject). Espera eventos Wins32 de mudanca na impressora. Quando evento ocorre, marca uma flag (_flagsEnforcementPending).

2. **EnforcementWorkerLoop**: Executa em thread separada do pool. Periodicamente (500ms) verifica a flag pendente e executa o enforcement.

Beneficio: O monitoramento NAO fica bloqueado enquanto o enforcement processa.

## Flag e lock thread-safe

A comunicacao entre loops usa:

- `_flagsEnforcementPending` (bool volatile): marca se ha trabalho pendente
- `_enforcementLock` (object): protege a flag com lock

Vantagem: Permite que MonitorLoop marque pendencia enquanto EnforcementWorkerLoop consome (alternativamente).

## Debounce

Para evitar loops infinitos de re-enforcement (mudanca -> enforcement -> mudanca -> enforcement), ha um intervalode 3 segundos entre aplicacoes (_debounceInterval).

## Integracao com PrintTicketService

A classe NAO implementa logica de leitura/escrita de ticket diretamente. Ela usa `IPrintTicketService` (`PrintTicketEnforcementHelper`) para:

1. Ler DefaultPrintTicket atual
2. Comparar com politica
3. Aplicar atualizacao se necessario

Separacao que permite reuso e testabilidade.

## Eventos

- `StatusChanged`: Start/stop do componente
- `EnforcementLog`: Log operacional detalhado
- `Error`: Excecoes nao tratadas nos loops

## Riscos controlados

- Permissao insuficiente (Manage Printers) para update -> falha silenciosa com log
- Handle invalido ou impressora nao encontrada -> retorna com log descritivo
- Driver nao suporta propriedade -> log e continua

## Resumo da arquitetura

```
[Win32 Event] -> MonitorLoop -> [flag pendente] -> EnforcementWorkerLoop -> [PrintTicketService]
                                                                      ^
                                                                      |
                                                          [Policy: valores requeridos]
```

O pattern garante que:
1. O componente reage a mudancas ativas (evento)
2. Processamento NAO bloqueia monitoramento
3. Debounce evita loops infinitos
4. Falhas sao resilientes (log ao inves de excecao)