# Estudo 02 - GetWppStatus (fluxo interno)

## Objetivo do metodo

`GetWppStatus()` resolve o estado final de WPP com base em probes de Registro conhecidos.

## Fluxo passo a passo

1. Inicializa `disabledCandidate` como nulo.
2. Itera pelos probes em ordem.
3. Tenta abrir `HKLM\{probe.Path}` com acesso somente leitura.
4. Se a chave nao existir, continua para o proximo probe.
5. Le `probe.ValueName`.
6. Se o valor nao existir, continua para o proximo probe.
7. Tenta converter o valor bruto para inteiro.
8. Se tipo nao suportado, retorna `Unknown` imediatamente.
9. Se valor bater com `EnabledValue`, retorna `Enabled` imediatamente.
10. Se valor bater com `DisabledValue`, guarda candidato e continua.
11. Se valor estiver fora do mapeamento, retorna `Unknown`.
12. Ao fim da iteracao, se houver candidato disabled, retorna ele.
13. Se nada for encontrado, retorna `Disabled` por padrao.

## Por que existe disabledCandidate

- A ordem de probes permite coexistencia de sinais.
- Se um probe indicar disabled e outro posterior indicar enabled, enabled deve vencer.
- Por isso, disabled nao retorna na hora; ele fica como candidato.

## Conteudo do retorno

Cada `WppStatusResult` inclui:

- status (`Enabled`, `Disabled`, `Unknown`);
- fonte (caminho completo no Registro ou mensagem de ausencia);
- detalhe textual da decisao;
- valor numerico quando disponivel.

## Aprendizados adicionais

- O metodo equilibra robustez (fallback) e rastreabilidade (motivo explicito no retorno).
