# Estudo 01 - Visao geral do NativeMethods

## O que esta classe representa

`NativeMethods` e a fronteira de interoperabilidade entre C# e Winspool (`winspool.drv`).

Ela centraliza:

- constantes Win32 usadas no dominio;
- structs com layout de memoria para marshal;
- assinaturas P/Invoke para operacoes de spooler.

## Papel arquitetural

- evitar espalhar `DllImport` por varios services;
- garantir consistencia de tipos e flags;
- reduzir risco de bug de interop por duplicacao de assinatura.

## Modelo mental de uso

1. Service prepara entrada gerenciada (strings, structs, flags).
2. Chama assinatura em `NativeMethods`.
3. Trata erros via `GetLastWin32Error`.
4. Faz marshal e projecao para modelos de dominio.

## Pilares tecnicos desta classe

- `SetLastError = true` para diagnostico correto de erro Win32;
- `CharSet.Unicode` para coerencia com APIs W;
- `StructLayout.Sequential` para layout previsivel em memoria;
- tipos numericos e ponteiros alinhados com contrato nativo.

## Aprendizados adicionais

- `NativeMethods` funciona como "contrato unico" da camada nativa: se ele estiver correto, os services ficam mais simples e seguros.
