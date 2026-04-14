# Estudo 02 - AddWsdPort

## Objetivo do metodo

`AddWsdPort(string portName)` tenta criar uma porta WSD via canal administrativo Xcv.

Ele protege dois cenarios:

- idempotencia (porta ja existe);
- diferenca entre sucesso da chamada Win32 e sucesso funcional do comando no monitor.

## Fluxo tecnico

1. Chama `ListPorts()` e compara nomes ignorando maiusculas/minusculas.
2. Se ja existir, loga e retorna sem erro.
3. Monta `PRINTER_DEFAULTS` com `SERVER_ACCESS_ADMINISTER`.
4. Abre `,XcvMonitor WSD Port` com `OpenPrinter`.
5. Monta payload Unicode com terminador nulo: `portName + '\0'`.
6. Executa `XcvData(..., "AddPort", ...)`.
7. Valida `status` retornado pelo monitor.
8. Fecha handle em `finally`.

## Pontos importantes

- `XcvData` pode retornar `true` e ainda assim o comando ser rejeitado via `status`.
- `ERROR_NOT_SUPPORTED` recebe mensagem orientando uso de descoberta/reuso de porta WSD.
- O payload Unicode com `\0` e necessario para comando textual no canal Xcv.

## Riscos e limites

- Dependencia de suporte do ambiente ao monitor WSD.
- Permissao administrativa insuficiente gera falha no `OpenPrinter`.

## Aprendizados adicionais

- Este metodo e um bom exemplo de integracao Win32 em duas camadas: transporte (chamada) e semantica (status).
