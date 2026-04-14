# Estudo 03 - GetWppStatus (regras e cenarios)

## Regras de decisao (resumo)

- qualquer probe com valor de habilitado conhecido => `Enabled` (retorno imediato);
- valor de desabilitado conhecido => candidato `Disabled` (segue varrendo);
- tipo nao suportado => `Unknown`;
- valor fora do mapeamento => `Unknown`;
- nenhum sinal conhecido encontrado => `Disabled` por padrao.

## Cenarios praticos

### Cenario A: chave ausente

- Probe 1 sem chave.
- Probe 2 sem chave.
- Resultado: `Disabled` por padrao, com fonte explicativa de ausencia.

### Cenario B: valor desabilitado no primeiro probe

- Probe 1 devolve `0` (disabled conhecido).
- Probe 2 inexistente.
- Resultado: `Disabled` pelo candidato salvo.

### Cenario C: desabilitado no primeiro e habilitado no segundo

- Probe 1 devolve disabled.
- Probe 2 devolve enabled.
- Resultado: `Enabled` (enabled prevalece).

### Cenario D: tipo invalido (ex.: binario nao conversivel)

- Primeiro valor encontrado nao converte para int.
- Resultado: `Unknown` imediato.

### Cenario E: valor numerico fora do mapeamento

- Valor encontrado e inteiro, porem diferente de enabled/disabled conhecidos.
- Resultado: `Unknown` imediato.

## Implicacao arquitetural

- A regra "enabled vence" reduz falso negativo quando existem multiplas politicas coexistindo.
- O retorno `Unknown` evita inferencia incorreta em cenarios nao catalogados.

## Aprendizados adicionais

- O metodo privilegia confiabilidade semantica sobre "forcar" um resultado binario em dados ambiguos.
