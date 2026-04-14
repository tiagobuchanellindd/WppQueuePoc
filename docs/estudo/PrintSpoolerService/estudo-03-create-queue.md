# Estudo 03 - CreateQueue

## Objetivo do metodo

`CreateQueue(...)` cria uma nova fila de impressao no spooler com base em `PRINTER_INFO_2`.

## Fluxo tecnico

1. Preenche `PRINTER_INFO_2` com nome, driver, porta, processador, datatype, comentario e localizacao.
2. Define campos operacionais iniciais (`Attributes`, prioridade, horarios, status).
3. Aloca memoria nativa para a struct.
4. Copia struct gerenciada para memoria nativa com `StructureToPtr`.
5. Chama `AddPrinter(null, 2, infoPtr)`.
6. Se handle vier nulo, levanta erro Win32 contextualizado.
7. Fecha handle de impressora criada.
8. Em `finally`, destroi struct e libera memoria.

## Decisoes tecnicas relevantes

- Usa nivel 2 porque `PRINTER_INFO_2` carrega os metadados necessarios para criacao completa.
- Define `PRINTER_ATTRIBUTE_QUEUED` como atributo inicial da fila.
- Mantem `pDevMode` e `pSecurityDescriptor` nulos na criacao da POC.

## Riscos e limites

- Driver/porta/processador/datatype invalidos serao rejeitados pelo spooler.
- Campos de string inconsistentes podem gerar erro de criacao dificil de diagnosticar sem contexto.

## Aprendizados adicionais

- O par `StructureToPtr` + `DestroyStructure` mostra um ciclo correto de vida para structs com campos marshalizados.
