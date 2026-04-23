# Estudo 02 - MonitorLoop e eventos Win32

## Objetivo deste estudo

Entender como o MonitorLoop captura eventos de mudanca na impressora usando APIs Win32.

## API Win32 utilizada

### FindFirstPrinterChangeNotification

```cpp
HANDLE FindFirstPrinterChangeNotification(
  HANDLE hPrinter,
  DWORD  fdwFlags,
  DWORD  fdwOptions,
  LPVOID pNotifyInfo
);
```

Esta funcao registra um handle de notificacao que sinaliza quando determinados eventosocorrem na impressora.

Parametros relevantes:

- `hPrinter`: Handle da impressora (aberta via OpenPrinter)
- `fdwFlags`: Quais eventos monitorar (ex: PRINTER_CHANGE_ADD_PRINTER, PRINTER_CHANGE_SET_PRINTER)
- `fdwOptions`: Opcoes adicionais (0 = sincrono)
- `pNotifyInfo`: Buffer para informacoes (pode ser IntPtr.Zero se apenas sinal)

Retorno: Handle de notificacao (ou INVALID_HANDLE_VALUE em caso de erro).

### WaitForSingleObject

```cpp
DWORD WaitForSingleObject(
  HANDLE hHandle,
  DWORD  dwMilliseconds
);
```

Espera que o handle seja sinalizado ou o timeout expirar.

Parametros:

- `hHandle`: Handle de notificacao
- `dwMilliseconds`: Timeout em milissegundos

Retorno:

- WAIT_OBJECT_0 (0): Handle sinalizado
- WAIT_TIMEOUT (258): Timeout expirou
- WAIT_FAILED: Erro

### FindNextPrinterChangeNotification

```cpp
BOOL FindNextPrinterChangeNotification(
  HANDLE hNotify,
  PDWORD pfdwChange,
  LPVOID pNotifyInfo,
  PVOID  pOptions
);
```

Avanca o estado do handle de notificacao e retorna qual evento ocorreu.

Parametros:

- `hNotify`: Handle de notificacao
- `pfdwChange`: Saida com flags do evento
- `pNotifyInfo`: Buffer de info (opcional)
- `pOptions`: Opcoes (opcional)

Retorna TRUE se bem-sucedido.

## Flags de mudanca

Definicos em PrinterChangeNotificationNative:

- PRINTER_CHANGE_ADD_PRINTER: Impressora foi adicionada
- PRINTER_CHANGE_DELETE_PRINTER: Impressora foi removida
- PRINTER_CHANGE_SET_PRINTER: Configuracao foi alterada
- PRINTER_CHANGE_PRINTER: Qualquer mudanca na impressora

No nosso codigo, usamos PRINTER_CHANGE_SET_PRINTER para detectar alteracoesmanuais nas propriedades da fila.

## Fluxo do monitoramento

```csharp
// 1. Abrir impressora
OpenPrinter(printerName, out hPrinter, ref defaults)

// 2. Registrar notificacao
hNotify = FindFirstPrinterChangeNotification(
    hPrinter,
    PRINTER_CHANGE_SET_PRINTER,  // mudou config
    0,
    IntPtr.Zero)

// 3. Loop de espera
while (!ct.IsCancellationRequested)
{
    waitResult = WaitForSingleObject(hNotify, 2000);  // 2s timeout

    if (waitResult == WAIT_OBJECT_0) {
        // Evento detectado
        FindNextPrinterChangeNotification(hNotify, out int change, ...);

        // Marcar enforcement pendente (se passou debounce)
    }
}
```

## Por que timeout de 2 segundos?

WaitForSingleObject com timeout permite:

1. Verificar cancellation token periodicamente (se nao bloqueasse, precisaria de CheckForCancellation separado)
2. Evitar deadlock se handle ficar sinalizado permanentemente
3. Manter responsividade do loop

O timeout curto (vs. timeout longo ou infinito) aceita reagir rapido a mudancas.

## Flag PRINTER_CHANGE_SET_PRINTER

Esta flag detecta qualquer alteracao nas propriedades da impressora/fila, incluindo:

- Mudanca em propriedades do driver
- Mudanca em DefaultPrintTicket
- Mudanca em configuracoes de porta

Nao detecta trabalhos de impressao sendo enviados.

## Tratamento de erro

Se OpenPrinter falha: loga erro e retorna.
Se FindFirstPrinterChangeNotification retorna INVALID_HANDLE: loga erro e retorna.
O loop quebra se WaitForSingleObject retorna erro (nao timeout).

## Recursos

O codigo garante cleanup em finally:

- FindClosePrinterChangeNotification(hNotify)
- ClosePrinter(hPrinter)

## Resumo

O MonitorLoop usa o pattern "espera bloqueante com timeout curto" via Win32 API parapermanecer responsivo a eventos de mudanca, sem consumir CPU intensiva em loops de polling.Toda vez que um evento é sinalizado, marca a flag pendente para oworker processar.