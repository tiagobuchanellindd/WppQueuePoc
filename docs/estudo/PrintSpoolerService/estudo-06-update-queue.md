# Estudo 06 - UpdateQueue

## Objetivo do metodo

`UpdateQueue(...)` aplica alteracoes parciais em uma fila existente usando `SetPrinter` nivel 2.

## Fluxo tecnico

1. Le estado atual da fila via `GetQueueInfo(queueName)`.
2. Zera `pDevMode` e `pSecurityDescriptor` para evitar reuse indevido de ponteiros sensiveis.
3. Aplica apenas campos informados (nome, driver, porta, comentario, localizacao).
4. Verifica se houve ao menos uma mudanca (`hasChanges`).
5. Abre fila com `PRINTER_ACCESS_ADMINISTER`.
6. Aloca memoria, copia `PRINTER_INFO_2` atualizado e chama `SetPrinter`.
7. Libera memoria e fecha handle em `finally`.

## Ponto de seguranca importante

- O reset de ponteiros sensiveis protege contra envio acidental de dados nativos invalidos no update.

## Validacao de uso

- Se nenhum campo foi informado para mudanca, o metodo falha com mensagem clara.

## Aprendizados adicionais

- Update parcial com struct completa exige disciplina: alterar so o necessario sem comprometer campos estruturais da fila.
