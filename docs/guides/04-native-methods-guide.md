# Native Methods Guide (`Interop/NativeMethods.cs`)

## Purpose
`NativeMethods.cs` is the Win32 interop boundary of this POC.  
It defines how managed C# code calls native Windows Print Spooler APIs (`winspool.drv`) and how native data is marshaled back into .NET types.

## How the file is structured
1. **Constants**  
   Access flags (`PRINTER_ACCESS_*`, `SERVER_ACCESS_ADMINISTER`), enumeration flags (`PRINTER_ENUM_*`), and common Win32 errors (`ERROR_*`).

2. **Native structs**  
   C# representations of Winspool structures (`PRINTER_INFO_2`, `PORT_INFO_1`, `DRIVER_INFO_2`, etc.) with explicit layout and Unicode string marshaling.

3. **P/Invoke signatures**  
   `[DllImport("winspool.drv", ...)]` declarations used by service classes (`PrintSpoolerService`).

## Why attributes matter
- **`[SupportedOSPlatform("windows")]`**  
  Declares this code is Windows-specific.

- **`[StructLayout(LayoutKind.Sequential)]`**  
  Keeps field order compatible with the native C structs.

- **`[MarshalAs(UnmanagedType.LPWStr)]`**  
  Marshals strings as UTF-16 pointers expected by Winspool.

- **`SetLastError = true`**  
  Enables retrieval of the native error via `Marshal.GetLastWin32Error()`.

## Main API roles
- **`OpenPrinter` / `ClosePrinter`**  
  Open and close native handles for queues, monitor endpoints (`XcvMonitor`), or ports (`XcvPort`).

- **`XcvData`**  
  Sends monitor/port administrative commands (for example `AddPort`, `GetAPPortInfo`).

- **`AddPrinter` / `SetPrinter` / `DeletePrinter`**  
  Create, update, and delete print queues.

- **`EnumPrinters` / `GetPrinter`**  
  Enumerate queues and read full queue metadata.

- **`EnumPorts` / `EnumPrinterDrivers` / `EnumPrintProcessors` / `EnumPrintProcessorDatatypes`**  
  Enumerate ports, drivers, print processors, and datatypes.

## Two-call buffer pattern (very important)
Most `Enum*` and `GetPrinter` operations use this pattern:

1. Call API with `IntPtr.Zero` buffer and size `0` to get `needed` bytes.
2. Allocate unmanaged memory (`Marshal.AllocHGlobal((int)needed)`).
3. Call API again to fill the buffer.
4. Walk the buffer entry-by-entry using:
   - `Marshal.SizeOf<T>()` (entry size),
   - `IntPtr.Add(basePtr, index * structSize)` (entry pointer),
   - `Marshal.PtrToStructure<T>(entryPtr)` (map bytes to C# struct).
5. Free unmanaged memory (`Marshal.FreeHGlobal`).

This is exactly how queue/port/driver/processor lists are materialized in this POC.

## Call flow examples in this project
- **Add WSD port**  
  `OpenPrinter(",XcvMonitor WSD Port")` -> `XcvData("AddPort")` -> check `dwStatus` -> `ClosePrinter`.

- **Inspect queue**  
  `GetPrinter` to read queue data -> optional `OpenPrinter(",XcvPort <port>")` + `XcvData("GetAPPortInfo")` for protocol evidence.

- **List print processors**  
  `EnumPrintProcessors` -> parse `PRINTPROCESSOR_INFO_1` entries from native buffer into `List<string>`.

## Error model
There are two layers of failure to check:

1. **API call return value (`bool`)**  
   If `false`, read `Marshal.GetLastWin32Error()`.

2. **Operation status from API output (`dwStatus` in `XcvData`)**  
   The function call can succeed while the monitor operation itself is rejected (for example `ERROR_NOT_SUPPORTED`).

## Safety practices used in the POC
- Handles are always closed in `finally`.
- Unmanaged buffers are always freed in `finally`.
- Marshaling is centralized in one file to reduce interop risk.
- Business rules stay outside `NativeMethods` (in services), keeping this layer purely technical.
