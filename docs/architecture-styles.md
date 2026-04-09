# Architecture Styles for Study

## Why this structure exists
This POC is intentionally organized to show complementary styles in the same project:

1. **Modern managed application style (WPF front-end)**
   - `WppQueuePoc.App/MainWindow.xaml(.cs)`: UI actions, validation, async execution, output rendering.
   - `Abstractions/*`: interfaces for service contracts.
   - `Models/*`: immutable records/enums for clean data flow.

2. **Native interop style (Windows Print Spooler)**
   - `Interop/NativeMethods.cs`: raw P/Invoke bindings, structs, constants.
   - `Services/PrintSpoolerService.cs`: safe wrapper around native calls (`OpenPrinter`, `XcvData`, `AddPrinter`, `SetPrinter`, `DeletePrinter`, `EnumPorts`, `EnumPrinters`, `GetPrinter`).
   - `Services/WppRegistryService.cs`: Registry-based WPP state detection.

3. **Managed print ticket style (Windows Desktop printing APIs)**
   - `Services/PrintTicketService.cs`: read-only diagnostics for default print ticket/capabilities.
   - Requires `Microsoft.WindowsDesktop.App` framework reference at build/runtime.

## How they work together
- The modern layer orchestrates behavior and keeps operation UX simple through the WPF interface.
- The interop layer isolates native complexity and Win32 details.
- This separation helps experimentation and troubleshooting without coupling UI logic to native memory/handle management code.

## Practical benefit in this POC
- Easier to compare and study modern .NET design versus low-level Windows API integration.
- Easier to extend with advanced scenarios such as:
  - `GetAPPortInfo` diagnostics for APMON/WSD/IPP ports.
  - **Print Ticket** diagnostics (`ticket-info`) for default ticket/capability analysis.

## Execution schema (Mermaid)
```mermaid
flowchart TD
    U["User Interface (WPF)"] --> A["MainWindow"]

    A -->|"wpp-status"| R["Services WppRegistryService"]
    A -->|"queue commands"| S["Services PrintSpoolerService"]
    A -->|"ticket-info"| T["Services PrintTicketService"]

    S --> N["Interop NativeMethods winspool.drv"]
    S --> R
    T --> SP["System.Printing Microsoft.WindowsDesktop.App"]

    R --> M1["Models WppStatusResult"]
    S --> M2["Models QueueInfo QueueInspectionResult ApPortData"]
    T --> M3["Models PrintTicketInfoResult"]

    M1 --> A
    M2 --> A
    M3 --> A
    A --> O["Output panel"]
```

## Command mapping: caller, executor, objective, result
| Command | Caller | Executor | Objective | Result model/output |
|---|---|---|---|---|
| `wpp-status` | `MainWindow` | `WppRegistryService` | Detect global WPP state from Registry | `WppStatusResult` (`Enabled/Disabled/Unknown`) |
| `add-wsd-port` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.XcvData` | Try creating WSD port in monitor | Success/error (`dwStatus`/Win32) |
| `create` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.AddPrinter` | Create queue with chosen driver/port | Queue creation confirmation or Win32 error |
| `list` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.EnumPrinters` | List installed queues | `QueueInfo[]` rendered to output panel |
| `list-ports` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.EnumPorts` | List available ports | Ordered port list rendered to output panel |
| `update` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.SetPrinterW` | Update queue metadata (comment/location) | Queue update confirmation or Win32 error |
| `delete` | `MainWindow` | `PrintSpoolerService` -> `NativeMethods.DeletePrinter` | Remove queue | Queue deletion confirmation or Win32 error |
| `inspect` | `MainWindow` | `PrintSpoolerService` (+ `WppRegistryService`) | Classify queue as likely WPP/not WPP | `QueueInspectionResult` with details (`APMON`, protocol, URL when available) |
| `ticket-info` | `MainWindow` | `PrintTicketService` -> `System.Printing` | Read default print ticket and capability snapshot | `PrintTicketInfoResult` (attributes/capability counts) |

## Business process flow (Mermaid)
```mermaid
flowchart TD
    A["Start"] --> B["Check global WPP status"]
    B --> C["List available ports"]
    C --> D{"WSD port available?"}
    D -->|No| E["Try add-wsd-port or onboard WSD device"]
    E --> C
    D -->|Yes| F["Create queue with driver + selected port"]
    F --> G["Inspect queue classification"]
    G --> H["Read ticket-info for default printing behavior"]
    H --> I{"Need metadata change?"}
    I -->|Yes| J["Update queue comment/location"]
    I -->|No| K["Keep current queue config"]
    J --> K
    K --> L{"Cleanup required?"}
    L -->|Yes| M["Delete queue"]
    L -->|No| N["End with queue kept"]
    M --> O["End with queue removed"]
```

## Technical sequence flow (Mermaid)
```mermaid
sequenceDiagram
    participant U as User
    participant A as MainWindow
    participant S as PrintSpoolerService
    participant R as WppRegistryService
    participant N as NativeMethods (winspool)
    participant T as PrintTicketService
    participant P as System.Printing

    U->>A: inspect --queue X
    A->>S: InspectQueue(X)
    S->>N: OpenPrinter/GetPrinter
    S->>R: GetWppStatus()
    R-->>S: WppStatusResult
    S->>N: OpenPrinter(",XcvPort ...") + XcvData(GetAPPortInfo)
    S-->>A: QueueInspectionResult
    A-->>U: Classification + details

    U->>A: ticket-info --queue X
    A->>T: GetDefaultTicketInfo(X)
    T->>P: LocalPrintServer.GetPrintQueue(X)
    T->>P: Read DefaultPrintTicket + capabilities
    T-->>A: PrintTicketInfoResult
    A-->>U: Ticket attributes/capabilities
```

## Classification decision flow (Mermaid)
```mermaid
flowchart TD
    A["Inspect queue"] --> B["Read global WPP status"]
    B --> C["Read port name and APMON protocol (if available)"]
    C --> D{"Global WPP = Disabled?"}
    D -->|Yes| E["Classification = LikelyNotWpp"]
    D -->|No| F{"Global WPP = Enabled?"}
    F -->|No| G["Classification = Indeterminate"]
    F -->|Yes| H{"Port is WSD OR APMON protocol is WSD/IPP?"}
    H -->|Yes| I["Classification = LikelyWpp"]
    H -->|No| J["Classification = Indeterminate"]
```
