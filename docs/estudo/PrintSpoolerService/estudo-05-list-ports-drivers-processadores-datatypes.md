# Estudo 05 - ListPorts, ListDrivers, ListPrintProcessors e ListDataTypes

## Objetivo do bloco

Esses metodos formam a trilha de descoberta de recursos do host para apoiar criacao/atualizacao de filas.

## Padrao comum dos quatro metodos

1. Chamada inicial para descobrir tamanho de buffer (`needed`).
2. Tratamento de erro esperado de descoberta (`ERROR_INSUFFICIENT_BUFFER`).
3. Curto-circuito para lista vazia quando `needed == 0`.
4. Alocacao do buffer nativo.
5. Segunda chamada para leitura real.
6. Iteracao por struct em memoria e extracao de nome valido.
7. Ordenacao alfabetica (`OrdinalIgnoreCase`).
8. Liberacao de memoria em `finally`.

## Diferencas por metodo

- `ListPorts`: usa `EnumPorts` nivel 1 e `PORT_INFO_1`.
- `ListDrivers`: usa `EnumPrinterDrivers` nivel 2 e `DRIVER_INFO_2`.
- `ListPrintProcessors`: usa `EnumPrintProcessors` nivel 1 e `PRINTPROCESSOR_INFO_1`.
- `ListDataTypes`: usa `EnumPrintProcessorDatatypes` nivel 1 e `DATATYPES_INFO_1`, exigindo `printProcessor` valido.

## Validacoes de negocio relevantes

- `ListDataTypes` bloqueia entrada vazia com `InvalidOperationException`.
- Todos filtram nomes nulos/vazios antes de retornar para evitar ruido na CLI.

## Aprendizados adicionais

- O design reaproveita um esqueleto tecnico unico de enumeracao, reduzindo variacao acidental entre metodos semelhantes.
