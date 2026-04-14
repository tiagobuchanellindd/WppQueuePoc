# Estudo 03 - Estruturas e layout de memoria

## Objetivo das structs

As structs traduzem layout Win32 para representacoes C# usadas no marshal.

## Structs principais

- `PRINTER_DEFAULTS`: parametros de abertura de handle (`OpenPrinter`).
- `PRINTER_INFO_2`: metadados completos da fila.
- `PORT_INFO_1`: nome de porta.
- `DRIVER_INFO_2`: metadados de driver.
- `PRINTPROCESSOR_INFO_1`: nome de processador.
- `DATATYPES_INFO_1`: nome de datatype.
- `AP_PORT_DATA_1`: retorno de `GetAPPortInfo` (versao, protocolo, URL).

## Decisoes de layout

- `StructLayout(LayoutKind.Sequential)`: preserva ordem de campos esperada no nativo.
- `CharSet.Unicode`: coerencia com API wide.
- `MarshalAs(UnmanagedType.LPWStr)`: ponteiro para string Unicode gerenciado.
- `ByValTStr` em `AP_PORT_DATA_1.DeviceOrServiceUrl`: string embutida de tamanho fixo.

## Cuidado com ponteiros sensiveis

Campos como `pDevMode` e `pSecurityDescriptor` sao ponteiros nativos e exigem cuidado em cenarios de update e copia de struct.

## Cuidado com tamanho fixo

`DeviceOrServiceUrl` usa `SizeConst = 260`, o que depende do contrato nativo esperado para esse payload.

## Aprendizados adicionais

- A maior fonte de bugs em interop costuma estar em layout e tipos; centralizar isso numa unica classe e uma protecao arquitetural forte.
