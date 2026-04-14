# Estudo 06 - Cleanup e diagnostico

## DisposeIfPossible

Objetivo:

- liberar recursos quando objetos refletidos implementam `IDisposable`.

Uso:

- aplicado a `printQueue` e `localPrintServer` nos blocos `finally`.

Beneficio:

- evita vazamento de recurso nativo/gerenciado associado ao subsistema de impressao.

## GetInnermostMessage

Objetivo:

- extrair mensagem da excecao mais interna da cadeia (`InnerException`).

Uso:

- empregado nos retornos de falha para expor causa raiz sem stack trace completo.

Beneficio:

- melhora legibilidade do erro para operador/CLI sem perder contexto funcional.

## Estrategia global de erro na classe

- em vez de propagar excecao bruta, converte falhas para objetos de resultado com detalhes;
- facilita troubleshooting em ambientes onde variacoes de driver/permissao sao comuns.

## Aprendizados adicionais

- A combinacao de cleanup garantido com diagnostico amigavel e um dos pontos fortes do service para operacao real.
