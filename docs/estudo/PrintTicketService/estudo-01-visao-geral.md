# Estudo 01 - Visao geral do PrintTicketService

## O que esta classe faz

`PrintTicketService` le e atualiza atributos de `DefaultPrintTicket` e `UserPrintTicket` de filas de impressao.

## Decisao arquitetural principal

A implementacao usa reflection sobre `System.Printing` para evitar dependencia estatica forte.

Beneficios:

- roda em ambientes onde a montagem pode nao estar disponivel;
- degrada com mensagem descritiva em vez de quebrar fluxo da aplicacao;
- reduz acoplamento direto com tipos de impressao de desktop em tempo de compilacao.

## Capacidades do service

- leitura de atributos de ticket (`GetDefaultTicketInfo`, `GetUserTicketInfo`);
- atualizacao de atributos com deteccao de diferenca (`UpdateDefaultTicket`, `UpdateUserTicket`);
- utilitarios internos para conversao/reflection/dispose/diagnostico.

## Riscos controlados pelo design

- ausencia de `System.Printing` no runtime;
- divergencia de API por driver ou tipo de fila;
- excecoes de reflection/invocacao durante set/commit.

## Aprendizados adicionais

- A classe prioriza resiliencia operacional: quase tudo retorna resultado funcional com detalhe, evitando excecao vazando para cima.
