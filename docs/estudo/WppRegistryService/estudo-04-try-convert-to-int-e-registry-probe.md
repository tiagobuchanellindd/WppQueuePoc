# Estudo 04 - TryConvertToInt e RegistryProbe

## TryConvertToInt

### Objetivo

Normalizar o valor bruto do Registro para `int` com seguranca.

### Regras aceitas

- `int`: caminho principal (tipico de DWORD no Registry).
- `string` numerica: fallback aceito quando o valor foi persistido em texto.

### Saida

- retorna `true` e preenche `result` quando a conversao e valida;
- retorna `false` com `result` default quando tipo/valor nao sao suportados.

### Observacao

- O metodo e propositalmente restritivo: so aceita formatos que o dominio quer tratar como confiaveis.

## RegistryProbe

### O que representa

`RegistryProbe` encapsula:

- caminho da chave (`Path`);
- nome do valor (`ValueName`);
- valor que significa habilitado (`EnabledValue`);
- valor que significa desabilitado (`DisabledValue`).

### Beneficio de design

- separa "dados de mapeamento" da logica do metodo;
- facilita adicionar/remover probes sem alterar o algoritmo central;
- deixa mais claro o contrato semantico de cada fonte de configuracao.

## Aprendizados adicionais

- A combinacao `RegistryProbe + TryConvertToInt` cria uma fronteira limpa entre leitura tecnica (Registro) e interpretacao de negocio (status WPP).
