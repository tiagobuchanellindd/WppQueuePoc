# Estudo 06 - Assinaturas de enumeracao

## APIs cobertas

- `EnumPrinters`
- `EnumPorts`
- `EnumPrinterDrivers`
- `EnumPrintProcessors`
- `EnumPrintProcessorDatatypes`

## Contrato comum

Todas seguem o mesmo desenho:

- entrada com nivel (`Level`) e buffer (`IntPtr` + tamanho);
- saida com `pcbNeeded` (bytes necessarios) e `pcReturned` (quantidade de itens);
- uso de `SetLastError = true` para diagnostico.

## Padrao de uso em runtime

1. Chamada com buffer nulo/tamanho 0 para descobrir `pcbNeeded`.
2. Alocacao de memoria com tamanho descoberto.
3. Segunda chamada para preencher dados.
4. Iteracao por deslocamento de struct no buffer.

## Diferencas de nivel/struct

- `EnumPrinters` nivel 2 => `PRINTER_INFO_2`.
- `EnumPorts` nivel 1 => `PORT_INFO_1`.
- `EnumPrinterDrivers` nivel 2 => `DRIVER_INFO_2`.
- `EnumPrintProcessors` nivel 1 => `PRINTPROCESSOR_INFO_1`.
- `EnumPrintProcessorDatatypes` nivel 1 => `DATATYPES_INFO_1`.

## Risco operacional conhecido

- Mudanca no estado do spooler entre as duas chamadas pode exigir retry com novo tamanho.

## Aprendizados adicionais

- O valor de `pcReturned` e o guia para iteracao logica; `pcbNeeded` guia alocacao de memoria.
