# Estudo 04 - ListQueues e EnumeratePrinterInfo2

## Objetivo dos metodos

- `ListQueues()` projeta filas para `QueueInfo` ordenado.
- `EnumeratePrinterInfo2()` executa a enumeracao nativa de filas e conexoes no nivel 2.

## Fluxo do EnumeratePrinterInfo2

1. Chama `EnumPrinters` com buffer nulo para descobrir `needed`.
2. Trata `ERROR_INSUFFICIENT_BUFFER` como esperado na fase de descoberta.
3. Se `needed == 0`, retorna lista vazia.
4. Aloca `AllocHGlobal(needed)`.
5. Chama `EnumPrinters` novamente para leitura efetiva.
6. Itera por deslocamento (`structSize`) e converte cada item com `PtrToStructure<PRINTER_INFO_2>`.
7. Libera buffer em `finally`.

## Fluxo do ListQueues

1. Recebe lista bruta de `EnumeratePrinterInfo2()`.
2. Projeta campos para `QueueInfo`.
3. Normaliza nulos para string vazia.
4. Calcula flag de compartilhamento com bitwise em `Attributes`.
5. Ordena por nome (`OrdinalIgnoreCase`).

## Pontos tecnicos importantes

- `PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS` traz visao consolidada de filas locais e conexoes.
- Marshal de strings acontece a partir de ponteiros presentes no buffer nativo.
- Estrategia de duas chamadas e obrigatoria para APIs de enumeracao Winspool.

## Risco conhecido

- Mudanca de estado do spooler entre primeira e segunda chamada pode invalidar o tamanho previamente calculado.

## Aprendizados adicionais

- A separacao entre "coleta nativa" e "projecao de dominio" deixa o metodo de listagem limpo e de facil manutencao.
