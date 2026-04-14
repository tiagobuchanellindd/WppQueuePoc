# Estudos tecnicos

Esta pasta organiza estudos por **arquivo-fonte**.

Padrao adotado:

- cada subpasta dentro de `docs/estudo` usa o nome do arquivo estudado;
- dentro da subpasta, os estudos sao separados em arquivos numerados;
- apenas este `README.md` e usado como indice mestre.

## Trilhas atuais

- `WppQueuePoc/Services/PrintSpoolerService.cs`
  - `docs/estudo/PrintSpoolerService/estudo-01-visao-geral.md`
  - `docs/estudo/PrintSpoolerService/estudo-02-add-wsd-port.md`
  - `docs/estudo/PrintSpoolerService/estudo-03-create-queue.md`
  - `docs/estudo/PrintSpoolerService/estudo-04-list-queues-e-enumerate-printer-info-2.md`
  - `docs/estudo/PrintSpoolerService/estudo-05-list-ports-drivers-processadores-datatypes.md`
  - `docs/estudo/PrintSpoolerService/estudo-06-update-queue.md`
  - `docs/estudo/PrintSpoolerService/estudo-07-delete-queue.md`
  - `docs/estudo/PrintSpoolerService/estudo-08-inspect-queue.md`
  - `docs/estudo/PrintSpoolerService/estudo-09-try-get-ap-port-info-e-protocol-to-string.md`
  - `docs/estudo/PrintSpoolerService/estudo-10-get-queue-info-e-throw-last-win32.md`

- `WppQueuePoc/Services/WppRegistryService.cs`
  - `docs/estudo/WppRegistryService/estudo-01-visao-geral.md`
  - `docs/estudo/WppRegistryService/estudo-02-get-wpp-status-fluxo.md`
  - `docs/estudo/WppRegistryService/estudo-03-get-wpp-status-regras-e-cenarios.md`
  - `docs/estudo/WppRegistryService/estudo-04-try-convert-to-int-e-registry-probe.md`

- `WppQueuePoc/Interop/NativeMethods.cs`
  - `docs/estudo/NativeMethods/estudo-01-visao-geral.md`
  - `docs/estudo/NativeMethods/estudo-02-constantes-e-semantica.md`
  - `docs/estudo/NativeMethods/estudo-03-estruturas-e-layout.md`
  - `docs/estudo/NativeMethods/estudo-04-assinaturas-open-close-xcv.md`
  - `docs/estudo/NativeMethods/estudo-05-assinaturas-crud-printer.md`
  - `docs/estudo/NativeMethods/estudo-06-assinaturas-enumeracao.md`

- `WppQueuePoc/Services/PrintTicketService.cs`
  - `docs/estudo/PrintTicketService/estudo-01-visao-geral.md`
  - `docs/estudo/PrintTicketService/estudo-02-leitura-default-e-user-ticket.md`
  - `docs/estudo/PrintTicketService/estudo-03-update-print-ticket-internal.md`
  - `docs/estudo/PrintTicketService/estudo-04-read-write-convert-attributes.md`
  - `docs/estudo/PrintTicketService/estudo-05-create-local-print-server-e-get-print-queue.md`
  - `docs/estudo/PrintTicketService/estudo-06-cleanup-e-diagnostico.md`

## Como usar este indice

- comece pelo `README.md` da trilha desejada;
- siga os estudos na ordem numerada (`estudo-01`, `estudo-02`, ...);
- ao fechar novo aprendizado, atualize o estudo existente ou crie uma continuacao.

Regra de manutencao:

- quando entrar novo tema, criar nova subpasta com nome do arquivo-fonte;
- criar estudos numerados diretamente na subpasta (sem `README.md` local);
- registrar aprendizados incrementais em cada estudo.

Regra de manutencao:

- quando entrar novo tema, criar nova subpasta com nome do arquivo-fonte;
- manter um `README.md` na subpasta com indice dos estudos;
- registrar aprendizados incrementais em cada estudo.
