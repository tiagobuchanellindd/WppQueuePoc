# Estudo 02 - Leitura de DefaultTicket e UserTicket

## Metodos cobertos

- `GetDefaultTicketInfo(queueName)`
- `GetUserTicketInfo(queueName)`

## Fluxo comum de leitura

1. Cria dicionario de atributos (case-insensitive).
2. Tenta resolver `System.Printing.LocalPrintServer` via `Type.GetType`.
3. Se montagem indisponivel, retorna `Available = false` com detalhe explicativo.
4. Cria `LocalPrintServer` (tentando acesso administrativo).
5. Resolve `PrintQueue` por nome.
6. Le a propriedade alvo do ticket (`DefaultPrintTicket` ou `UserPrintTicket`).
7. Extrai atributos por reflection (`OutputColor`, `PageMediaSize`, `PageOrientation`, `InputBin`, `Duplexing`, `CopyCount`, `Collation`, `Stapling`).
8. Retorna `PrintTicketInfoResult` com `Available = true` quando sucesso.
9. Em excecao, retorna resultado com detalhe da falha.
10. Sempre descarta `printQueue` e `localPrintServer` no `finally`.

## Diferenca entre os dois metodos

- apenas a propriedade lida do ticket muda (`DefaultPrintTicket` vs `UserPrintTicket`);
- o restante do fluxo e equivalente para permitir comparacao entre escopos.

## Valor de diagnostico

- o contrato de retorno facilita comparar configuracao padrao e configuracao do usuario sem depender de stack trace.

## Aprendizados adicionais

- O desenho favorece observabilidade de ambiente: distingue claramente "assembly ausente", "fila nao encontrada" e "ticket indisponivel".
