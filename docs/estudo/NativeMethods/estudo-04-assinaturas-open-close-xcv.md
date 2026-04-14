# Estudo 04 - Assinaturas OpenPrinter, ClosePrinter e XcvData

## OpenPrinter

Assinatura:

- recebe nome do alvo (`fila`, `,XcvMonitor ...`, `,XcvPort ...`);
- recebe `PRINTER_DEFAULTS` com acesso desejado;
- devolve handle em `phPrinter`.

Ponto-chave:

- `SetLastError = true` permite diagnostico confiavel em caso de falha.

## ClosePrinter

Responsavel por encerrar handle aberto por `OpenPrinter`.

Boas praticas:

- chamar sempre em `finally`;
- nao depender de GC para liberar recurso nativo.

## XcvData (duas sobrecargas)

### Sobrecarga com `byte[]`

Indicada para payload textual serializado (ex.: comando `AddPort` com string Unicode nula-terminada).

### Sobrecarga com `IntPtr`

Indicada para payload/retorno estruturado em memoria nativa (ex.: `GetAPPortInfo`).

## Dupla validacao em Xcv

- retorno bool da chamada indica sucesso de transporte da API;
- `pdwStatus` indica sucesso funcional do comando no monitor.

Ambos precisam ser validados para afirmar sucesso real.

## Aprendizados adicionais

- A existencia das duas sobrecargas reduz conversoes improvisadas e deixa explicito o formato de dados esperado em cada comando Xcv.
