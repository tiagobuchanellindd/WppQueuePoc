# Estudo 10 - GetQueueInfo e ThrowLastWin32

## Objetivo dos metodos

- `GetQueueInfo(queueName)`: le configuracao completa da fila via `GetPrinter` nivel 2.
- `ThrowLastWin32(context)`: helper para padronizar excecoes Win32 com contexto funcional.

## Fluxo do GetQueueInfo

1. Abre fila com `PRINTER_ACCESS_USE`.
2. Chama `GetPrinter(..., IntPtr.Zero, 0, out needed)` para descobrir tamanho.
3. Espera `ERROR_INSUFFICIENT_BUFFER` como parte do fluxo.
4. Aloca buffer nativo com `needed`.
5. Chama `GetPrinter` novamente para leitura efetiva.
6. Converte buffer para `PRINTER_INFO_2`.
7. Libera buffer e fecha handle em `finally`.

## Por que esse metodo e central

- Serve como fonte de verdade para leitura antes de update.
- Tambem alimenta inspecao de fila com dados reais do spooler.

## Papel do ThrowLastWin32

- Encapsula `Marshal.GetLastWin32Error()` em um ponto unico.
- Mantem mensagens de erro consistentes e com contexto de operacao.

## Aprendizados adicionais

- Este par de metodos consolida o contrato de erro e leitura de metadados, reduzindo repeticao no restante da classe.
