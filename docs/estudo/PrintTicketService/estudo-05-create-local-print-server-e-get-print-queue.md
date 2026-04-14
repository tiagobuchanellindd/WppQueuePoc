# Estudo 05 - CreateLocalPrintServer e GetPrintQueue

## CreateLocalPrintServer

Objetivo:

- criar `LocalPrintServer` tentando acesso desejado e mantendo fallback de compatibilidade.

Fluxo:

1. tenta resolver enum `PrintSystemDesiredAccess`;
2. se nao existir, usa construtor padrao;
3. se existir, tenta construtor com acesso desejado;
4. se nao houver construtor compativel ou ocorrer erro, fallback para construtor padrao.

Resultado:

- maximiza chance de funcionamento em runtimes diferentes sem acoplamento forte.

## GetPrintQueue

Objetivo:

- obter fila por nome com estrategia de fallback entre formas de acesso.

Estrategia em ordem:

1. tenta construtor direto de `PrintQueue(PrintServer, string, desiredAccess)`;
2. fallback para `GetPrintQueue(string, string[])` pedindo propriedade do ticket;
3. fallback final para `GetPrintQueue(string)`.

## Valor pratico dessa estrategia

- evita quebrar quando uma sobrecarga especifica nao estiver disponivel;
- tenta acesso mais explicito primeiro (melhor para update/commit), degradando para alternativas.

## Aprendizados adicionais

- O metodo materializa uma politica de compatibilidade progressiva: "melhor caminho" -> "caminho compativel" -> "ultimo recurso".
