# Estudo 05 - Assinaturas de CRUD de fila (AddPrinter, SetPrinter, DeletePrinter, GetPrinter)

## AddPrinter

Cria fila no nivel informado (neste projeto, nivel 2 com `PRINTER_INFO_2`).

Contrato:

- retorno e `IntPtr` de handle;
- `IntPtr.Zero` indica falha de criacao.

## SetPrinter

Atualiza propriedades da fila com dados de struct no nivel informado.

Ponto tecnico:

- `EntryPoint = "SetPrinterW"` garante chamada da variante Unicode.

## DeletePrinter

Exclui fila associada ao handle aberto.

Ponto tecnico:

- requer permissao adequada no handle (normalmente all access para cenarios de exclusao).

## GetPrinter

Le dados da fila no nivel informado (no projeto, nivel 2).

Padrao esperado:

- primeira chamada para tamanho (`pcbNeeded`);
- segunda chamada para leitura efetiva com buffer.

## Aprendizados adicionais

- Essas assinaturas formam o nucleo de ciclo de vida da fila; a seguranca operacional depende mais do uso correto do que da assinatura isolada.
