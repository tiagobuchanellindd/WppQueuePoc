# Estudo 01 - Visao geral do WppRegistryService

## O que esta classe faz

`WppRegistryService` implementa `IWppStatusProvider` e resolve o estado global de WPP lendo politicas no Registro (HKLM).

Ela converte sinais brutos (chave/valor/tipo) em um resultado de dominio `WppStatusResult`.

## Papel no sistema

- fornece status global para fluxos que precisam inferir aderencia a WPP;
- desacopla leitura de Registro do restante da aplicacao;
- concentra regras de interpretacao em um unico ponto.

## Estrategia central

1. Percorrer uma lista de probes conhecidos (`Path`, `ValueName`, `EnabledValue`, `DisabledValue`).
2. Tentar abrir a chave e ler o valor.
3. Normalizar o valor para inteiro (`TryConvertToInt`).
4. Aplicar regras semanticas para classificar em `Enabled`, `Disabled` ou `Unknown`.
5. Retornar `WppStatusResult` com fonte e justificativa.

## Caracteristicas de design

- fallback entre probes: um probe ausente nao derruba a resolucao;
- retorno explicativo: sempre traz contexto para auditoria;
- comportamento deterministico: ausencia de sinal conhecido cai para "disabled por padrao".

## Aprendizados adicionais

- A classe esta orientada a "diagnostico de operacao", nao apenas a um booleano final.
