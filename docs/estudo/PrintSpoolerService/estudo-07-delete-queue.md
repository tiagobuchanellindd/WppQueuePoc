# Estudo 07 - DeleteQueue

## Objetivo do metodo

`DeleteQueue(queueName)` remove uma fila existente com permissao elevada.

## Fluxo tecnico

1. Monta `PRINTER_DEFAULTS` com `PRINTER_ALL_ACCESS`.
2. Abre fila com `OpenPrinter`.
3. Executa `DeletePrinter(handle)`.
4. Em falha, le codigo Win32 e gera excecao contextualizada.
5. Fecha handle em `finally`.

## Tratamento especializado

- Para `ERROR_ACCESS_DENIED`, o metodo retorna mensagem orientativa (executar elevado e fechar consoles/props abertos).

## Aprendizados adicionais

- Esse metodo combina operacao destrutiva com diagnostico acionavel, o que ajuda bastante no suporte operacional.
