# Estudo 09 - TryGetApPortInfo e ProtocolToString

## Objetivo dos metodos

- `TryGetApPortInfo(portName)`: tenta obter metadados APMON de uma porta sem quebrar o fluxo principal.
- `ProtocolToString(protocol)`: traduz codigo numerico em rotulo legivel.

## Fluxo do TryGetApPortInfo

1. Valida `portName`.
2. Monta acesso administrativo e abre `,XcvPort {portName}`.
3. Aloca buffer para `AP_PORT_DATA_1`.
4. Executa `XcvData("GetAPPortInfo", ...)` com output struct.
5. Se chamada falhar ou status nao for sucesso, retorna `null`.
6. Em sucesso, converte struct nativa e retorna `ApPortData`.
7. Libera buffer e fecha handle sempre em `finally`.

## Estrategia de resiliencia

- Este metodo e "best effort": prefere retorno nulo a interromper a inspecao inteira.

## Fluxo do ProtocolToString

- `1 => WSD`
- `2 => IPP`
- demais valores => `Unknown`

## Aprendizados adicionais

- A combinacao desses metodos melhora legibilidade da auditoria sem acoplar classificacao a falhas de consulta auxiliar.
