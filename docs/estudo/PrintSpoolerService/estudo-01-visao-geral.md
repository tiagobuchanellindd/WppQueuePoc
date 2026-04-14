# Estudo 01 - Visao geral do PrintSpoolerService

## O que esta classe faz

`PrintSpoolerService` e a camada que traduz casos de uso da POC para chamadas Winspool nativas.

Em alto nivel, ela cobre:

- criacao de porta WSD (`AddWsdPort`);
- criacao, atualizacao e exclusao de fila (`CreateQueue`, `UpdateQueue`, `DeleteQueue`);
- listagens de recursos (`ListQueues`, `ListPorts`, `ListDrivers`, `ListPrintProcessors`, `ListDataTypes`);
- inspecao heuristica de aderencia WPP (`InspectQueue`);
- helpers internos de leitura nativa (`GetQueueInfo`, `EnumeratePrinterInfo2`, `TryGetApPortInfo`, `ThrowLastWin32`).

## Modelo mental de funcionamento

1. Metodo de negocio valida entrada e decide estrategia.
2. Chama API nativa via `NativeMethods`.
3. Se houver retorno em buffer, aplica padrao de duas chamadas (query de tamanho + leitura real).
4. Faz marshal para structs gerenciadas.
5. Projeta para modelo de dominio quando necessario.
6. Libera recursos nativos em `finally`.

## Padroes tecnicos recorrentes

### 1) Duas chamadas em APIs de enumeracao

Varios metodos seguem o mesmo fluxo:

- primeira chamada com buffer nulo para descobrir `needed`;
- tratamento de `ERROR_INSUFFICIENT_BUFFER` como comportamento esperado;
- segunda chamada com buffer alocado;
- iteracao por `structSize` com `IntPtr.Add` + `PtrToStructure`.

### 2) Tratamento de erro Win32

- falhas nativas viram `Win32Exception` com contexto funcional;
- `ThrowLastWin32` padroniza esse comportamento;
- em alguns cenarios, ha tratamento especializado (ex.: `ERROR_ACCESS_DENIED`, `ERROR_NOT_SUPPORTED`).

### 3) Higiene de recursos nativos

- handles de impressora/xcv sempre fechados com `ClosePrinter` em `finally`;
- buffers alocados com `AllocHGlobal` sempre liberados com `FreeHGlobal`;
- estruturas criadas manualmente usam `DestroyStructure` antes de liberar ponteiro quando aplicavel.

## Onde esta o risco tecnico principal

- corrida entre primeira e segunda chamada de enumeracao (estado do spooler pode mudar);
- campos ponteiro sensiveis (`pDevMode`, `pSecurityDescriptor`) exigem cuidado em update;
- resultado da chamada Win32 e status funcional da operacao podem ser coisas diferentes (caso `XcvData`).

## Checklist de revisao

1. Entrada validada?
2. Nivel de acesso em `PRINTER_DEFAULTS` e o minimo necessario?
3. Erros esperados na fase de descoberta foram tratados?
4. Recursos nativos estao protegidos por `try/finally`?
5. Conversao de struct para modelo final preserva semantica?

## Aprendizados adicionais

- A classe ja esta bem orientada a observabilidade tecnica por meio de mensagens de erro com contexto.
- O desenho privilegia resiliencia em metodos de inspecao (retorno parcial em vez de quebra total).
