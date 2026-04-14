# Estudo 03 - UpdatePrintTicketInternal

## Metodos de entrada

- `UpdateDefaultTicket(...)` delega para escopo `Default`.
- `UpdateUserTicket(...)` delega para escopo `User`.

Ambos convergem para `UpdatePrintTicketInternal(...)`.

## Fluxo detalhado

1. Monta dicionario `requested` apenas com valores nao vazios (atualmente: `Duplexing`, `OutputColor`).
2. Valida disponibilidade de `System.Printing`; se indisponivel, retorna falha descritiva.
3. Cria servidor local e resolve fila alvo.
4. Localiza propriedade do ticket (`ticketTypeProperty`).
5. Le o ticket atual; se ausente, retorna falha orientativa.
6. Tenta aplicar diferencas com `WriteTicketAttribute`:
   - `Duplexing`
   - `OutputColor`
   - `PageOrientation`
7. Se houve mudanca, reatribui ticket na fila (`SetValue`) e tenta `Commit`.
8. Le valores aplicados para retorno (`applied`).
9. Retorna `PrintTicketUpdateResult` com mensagem final conforme sucesso parcial/total.
10. Em excecao, devolve mensagem com causa raiz (`GetInnermostMessage`).

## Observacao importante

- O metodo marca sucesso logico por `changed`; logo, pode retornar mudanca detectada mesmo com falha em `SetValue/Commit`, acompanhado de mensagem de erro em `resultMsg`.

## Riscos e comportamentos esperados

- Drivers podem ignorar propriedades suportadas parcialmente.
- Algumas filas exigem elevacao/permissao de "Manage Printers" para commit.
- Reflection pode localizar propriedade, mas tipo/conversao ainda falhar em runtime.

## Aprendizados adicionais

- O retorno com pares `requested` vs `applied` e excelente para auditoria de "pedido vs efetivo" em ambientes heterogeneos de driver.
