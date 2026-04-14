# Estudo 02 - Constantes e semantica

## Objetivo das constantes

As constantes encapsulam codigos Win32 e flags de spooler com nomes de dominio claros.

## Grupos principais

### Flags de enumeracao

- `PRINTER_ENUM_LOCAL`: inclui filas locais.
- `PRINTER_ENUM_CONNECTIONS`: inclui conexoes de impressora.

Uso tipico: `PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS` para visao consolidada.

### Atributos de fila

- `PRINTER_ATTRIBUTE_QUEUED`: fila baseada em spooling.
- `PRINTER_ATTRIBUTE_SHARED`: fila compartilhada em rede.

### Niveis de acesso

- `SERVER_ACCESS_ADMINISTER`: acesso administrativo em contexto de servidor/monitor (Xcv).
- `PRINTER_ACCESS_USE`: leitura/uso basico da fila.
- `PRINTER_ACCESS_ADMINISTER`: administrar configuracao da fila.
- `PRINTER_ALL_ACCESS`: mascara de acesso total (operacoes destrutivas).

### Codigos de erro Win32

- `ERROR_SUCCESS` (0): sucesso.
- `ERROR_ACCESS_DENIED` (5): permissao insuficiente.
- `ERROR_NOT_SUPPORTED` (50): operacao nao suportada.
- `ERROR_INSUFFICIENT_BUFFER` (122): buffer menor que o necessario.

## Regra pratica de leitura

- Em fase de descoberta de tamanho, `ERROR_INSUFFICIENT_BUFFER` e esperado.
- Fora desse contexto, o mesmo codigo geralmente indica necessidade de retry/realloc.

## Aprendizados adicionais

- Nomear constantes no interop reduz muito erro semantico em services (especialmente em acessos e tratamento de erro).
