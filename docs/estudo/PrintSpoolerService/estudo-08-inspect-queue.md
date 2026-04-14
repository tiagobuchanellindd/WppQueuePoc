# Estudo 08 - InspectQueue

## Objetivo do metodo

`InspectQueue(queueName)` classifica uma fila como `LikelyWpp`, `LikelyNotWpp` ou `Indeterminate` com base em sinais operacionais.

## Fontes de evidencia usadas

1. Estado global WPP vindo de `IWppStatusProvider`.
2. Padrao do nome de porta (porta iniciando com `WSD`).
3. Dados opcionais de APMON (`Protocol`, `DeviceOrServiceUrl`) via `TryGetApPortInfo`.

## Regras de classificacao

- WPP global habilitado + porta WSD ou protocolo APMON moderno (1/WSD ou 2/IPP) => `LikelyWpp`.
- WPP global desabilitado => `LikelyNotWpp`.
- Demais combinacoes => `Indeterminate` com detalhes para auditoria.

## Valor do retorno

Retorna `QueueInspectionResult` com:

- nome da fila;
- porta;
- estado global;
- classificacao;
- trilha textual de diagnostico.

## Limite conceitual

- O metodo e heuristico: orienta decisao, mas nao representa prova definitiva de conformidade.

## Aprendizados adicionais

- O design preserva rastreabilidade ao compor `details` com frases e evidencias tecnicas de apoio.
