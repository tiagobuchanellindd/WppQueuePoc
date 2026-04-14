# Estudo 04 - ReadTicketAttribute, WriteTicketAttribute e ConvertIfPossible

## ReadTicketAttribute

Objetivo:

- ler propriedade do ticket por reflection e salvar no dicionario de saida.

Comportamento:

- se ticket for nulo, retorna silenciosamente;
- se propriedade existir, salva `ToString()` (ou vazio se nulo);
- excecoes sao ignoradas (best effort) para nao bloquear leitura dos demais atributos.

## WriteTicketAttribute

Objetivo:

- escrever atributo somente quando houver mudanca real.

Fluxo:

1. valida ticket e valor de entrada;
2. busca propriedade e valida se e gravavel;
3. converte string para tipo da propriedade (`ConvertIfPossible`);
4. compara valor atual vs novo valor;
5. aplica `SetValue` apenas se diferente;
6. retorna `true` quando alterou de fato.

Regra de comparacao:

- para `string` e `enum`, compara por texto com `OrdinalIgnoreCase`;
- para outros tipos, usa `Equals`.

## ConvertIfPossible

Objetivo:

- converter string para tipo real da propriedade (incluindo `Nullable<T>`).

Suportes:

- enums via `Enum.Parse`;
- tipos comuns via `Convert.ChangeType`;
- `Nullable<T>` com uso de tipo subjacente.

Fallback:

- se destino for string e conversao falhar, preserva valor original;
- para demais tipos, relanca excecao.

## Aprendizados adicionais

- Esse trio implementa uma mini-camada de "binding dinamico" para ticket, com foco em nao aplicar mudancas desnecessarias e reduzir ruido operacional.
