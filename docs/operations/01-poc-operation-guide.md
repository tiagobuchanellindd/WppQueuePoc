# WPP POC Operation Guide

## Prerequisites
- Compatible Windows version (recommended: Windows 11).
- Administrative permissions for spooler and port monitor operations.
- A print driver available for queue creation.
- Build/runtime must include `Microsoft.WindowsDesktop.App` for `ticket-info` (`System.Printing` APIs).

## Target demonstration flow
1. Query global WPP status:
   - `wpp-status`
2. Create WSD port:
   - operation through `XcvData("AddPort")` on monitor `XcvMonitor WSD Port`
   - if `ERROR_NOT_SUPPORTED (50)` is returned, use `list-ports` and reuse an existing WSD port
3. Create queue using the WSD port:
   - `create --queue "<name>" --port "<wsd-port>" --driver "<driver>"`
4. List queues:
   - `list`
5. Update queue metadata:
   - `update --queue "<name>" [--new-queue "<new-name>"] [--driver "<driver>"] [--port "<port>"] [--comment "<text>"] [--location "<text>"]`
6. Inspect queue:
   - `inspect --queue "<name>"`
7. Read default print ticket data:
   - `ticket-info --queue "<name>"`
8. Update queue default print ticket:
   - `ticket-update-default --queue "<name>" --duplexing "<value>" --output-color "<value>" --orientation "<value>"`
9. Update queue user print ticket:
   - `ticket-update-user --queue "<name>" --duplexing "<value>" --output-color "<value>" --orientation "<value>"`
10. Delete queue:
   - `delete --queue "<name>"`

## Practical comparative scenario (A/B)
Use the same queue recipe in two contexts to demonstrate business behavior:

### Scenario A - without global WPP
1. Confirm `wpp-status` returns `Disabled` (or `Unknown` when policy mapping is unresolved).
2. Create queue with the target driver and a WSD/IPP-compatible port.
3. Run `inspect` and `ticket-info`.
4. Update metadata and delete queue to complete lifecycle validation.

Expected interpretation:
- lifecycle commands work (`create`, `update`, `delete`);
- `inspect` tends to `LikelyNotWpp` when global WPP is disabled.

### Scenario B - with global WPP
1. Enable WPP globally in Windows.
2. Recreate queue with the same driver/port strategy (prefer WSD/IPP).
3. Run `inspect` and `ticket-info`.
4. Update metadata and delete queue to confirm same lifecycle under policy.

Expected interpretation:
- lifecycle commands still work under the same operational flow;
- `inspect` tends to `LikelyWpp` when global WPP is enabled and queue signals are compatible.

## Expected result by command
- `wpp-status`: `Enabled`, `Disabled`, or `Unknown`.
- `create`/`list`/`update`/`delete`: explicit success or detailed Win32 error.
- `list-ports`: all available print ports (useful to identify existing `WSD...` ports).
- `list-drivers`: all installed print drivers (useful to copy exact `--driver` name for creation).
- `inspect`: `LikelyWpp`, `LikelyNotWpp`, or `Indeterminate`, with rationale.
- `ticket-info`: default print ticket/capability snapshot when available.
- `ticket-update-default`: updates queue default ticket (machine/queue scope), with validation and possible driver constraints.
- `ticket-update-user`: updates current user ticket for queue (user scope), with validation and possible driver constraints.

## Error handling
- Always log:
  - Win32 error code (`GetLastError`);
  - `dwStatus` returned by `XcvData`;
  - operation context (queue, port, command).
- Common deletion issue:
  - `Win32=5 (Access denied)`: run elevated and ensure no printer-management/property window is open for that queue.

## Example operation sequence (WPF)
1. Open `WppQueuePoc.App`.
2. Click `wpp-status`.
3. Click `add-wsd-port` (or use an existing `WSD-*` from `list-ports`).
4. Fill Queue/Driver/Port (and optional Processor/Datatype/Comment/Location) and click `create`.
5. Click `list`, then `inspect`, then `ticket-info`.
6. Fill ticket fields (`Ticket Duplexing`, `Ticket Color`, `Ticket Orientation`) only when testing ticket updates.
7. Click `ticket-update-default` and/or `ticket-update-user`.
8. Click `update` when metadata change is needed.
9. Click `delete` for cleanup.

## Validated E2E scenarios (RDP server)

| Scenario | Operation | Observed result |
|---|---|---|
| Baseline status | `wpp-status` | `Status: Disabled` and policy key not found (treated as disabled) |
| Baseline ports | `list-ports` | Multiple ports returned, including `WSD-*` ports |
| Brother E2E - create | `create` (`Queue=POC-BROTHER-E2E-02`, `Driver=Brother PCLXL(A3/Ledger) Generic Driver`, `Port=WSD-70b7c465-e297-44f9-8332-3e9432177f7b`) | `Queue created.` |
| Brother E2E - inspect | `inspect` (`Queue=POC-BROTHER-E2E-02`) | `LikelyNotWpp` with `APMON Protocol=WSD(1)` and URL evidence |
| Brother E2E - ticket info | `ticket-info` (`Queue=Brother MFC-J6935DW Printer`) | `Available: True`, with defaults such as `OutputColor=Color`, `Duplexing=OneSided`, `PageOrientation=Portrait` |
| Brother E2E - update | `update` (`Queue=POC-BROTHER-E2E-02`, `Comment=E2E Brother`) | `Queue updated.` |
| Brother E2E - delete | `delete` (`Queue=POC-BROTHER-E2E-02`) | `Queue deleted.` |
| Epson E2E - create | `create` (`Queue=POC-EPSON-E2E-01`, `Driver=EPSON Universal Print Driver`, `Port=WSD-b151cc40-3716-4889-addd-29e3a06ec25d`) | `Queue created.` |
| Epson E2E - inspect | `inspect` (`Queue=POC-EPSON-E2E-01`) | `LikelyNotWpp` (global WPP disabled) |
| Epson E2E - ticket info | `ticket-info` (`Queue=EPSONF4A10D (AM-C400 Series)`) | `Available: True`, with defaults such as `OutputColor=Color`, `Duplexing=TwoSidedLongEdge`, `Stapling=None` |
| Epson E2E - update | `update` (`Queue=POC-EPSON-E2E-01`, `Comment=E2E Epson`) | `Queue updated.` |
| Epson E2E - delete | `delete` (`Queue=POC-EPSON-E2E-01`) | `Queue deleted.` |

## WPF test scenarios (summary table)

| # | Scenario | Precondition | Action in WPF | Expected result |
|---|---|---|---|---|
| 1 | A/B baseline (without global WPP) | `wpp-status = Disabled` | Create queue with WSD + click `inspect` | Classification tends to `LikelyNotWpp` |
| 2 | A/B policy case (with global WPP) | `wpp-status = Enabled` | Recreate same queue style + click `inspect` | Classification tends to `LikelyWpp` (when compatible signals exist) |
| 3 | Global WPP status | App open | Click `wpp-status` | Returns `Enabled`, `Disabled`, or `Unknown` with source/details |
| 4 | List ports | App open | Click `list-ports` | Lists ports, including possible `WSD-*` |
| 5 | List drivers | App open | Click `list-drivers` | Lists installed drivers to copy exact driver name |
| 6 | List print processors | App open | Click `list-processors` | Lists print processors (for example `WinPrint`) |
| 7 | List datatypes | `Processor` filled (recommended: `WinPrint`) | Click `list-datatypes` | Lists available datatypes for selected processor |
| 8 | Create queue (valid) | Queue/Driver/Port valid | Fill inputs and click `create` | `Queue created.` |
| 9 | Validate required input | Required input missing (for example empty Queue on `ticket-info`) | Click command | Validation error in output; app remains open |
| 10 | Inspect queue | Existing queue in `Queue` field | Click `inspect` | Shows queue, port, global WPP, classification, details |
| 11 | Read ticket info | Existing queue in `Queue` field | Click `ticket-info` | Shows availability/details and attributes when available |
| 12 | Update default ticket | Existing queue + at least one ticket field filled | Click `ticket-update-default` | Shows `Applied=True` and effective values (or constrained result) |
| 13 | Update user ticket | Existing queue + at least one ticket field filled | Click `ticket-update-user` | Shows `Applied=True` and effective values (or constrained result) |
| 14 | Update and delete queue | Existing queue | Click `update`, then `delete` | `Queue updated.` and `Queue deleted.` |

## WPF test scenarios (12 detailed examples)

1. **A/B baseline (without global WPP)**
   - Suggested input:
     - `Queue = POC-WPF-AB-BASELINE`
     - `Driver = <exact value copied from list-drivers>`
     - `Port = <exact WSD value copied from list-ports>`
   - Action:
     1. Confirm `wpp-status` returns `Disabled` (or unresolved `Unknown`).
     2. Click `create`.
     3. Click `inspect`.
   - Expected result: classification tends to `LikelyNotWpp`, with global status evidence.

2. **A/B policy case (with global WPP)**
   - Suggested input:
     - `Queue = POC-WPF-AB-POLICY`
     - `Driver = <exact value copied from list-drivers>`
     - `Port = <exact WSD value copied from list-ports>`
   - Action:
     1. Enable WPP globally in Windows.
     2. Click `create`.
     3. Click `inspect`.
   - Expected result: classification tends to `LikelyWpp` when queue signals are compatible.

3. **Global WPP status**
   - Suggested input: none.
   - Action: click `wpp-status`.
   - Expected result: output with `Status`, `Source`, `Details` (and `RawValue` when present).

4. **List ports to choose WSD**
   - Suggested input: none.
   - Action: click `list-ports`.
   - Expected result: list of ports; choose one `WSD-*` value for queue creation.

5. **List installed drivers**
   - Suggested input: none.
   - Action: click `list-drivers`.
   - Expected result: list of installed print drivers; copy exact name to `Driver`.

6. **List print processors**
   - Suggested input: none.
   - Action: click `list-processors`.
   - Expected result: list containing available processors (usually includes `WinPrint`).

7. **List datatypes for processor**
   - Suggested input: `Processor = WinPrint`.
   - Action: click `list-datatypes`.
   - Expected result: list of datatypes for that processor (for example `RAW`, `NT EMF 1.008`).

8. **Create queue (happy path)**
   - Suggested inputs:
      - `Queue = POC-WPF-QUEUE-01`
      - `Driver = <exact value copied from list-drivers>`
      - `Port = <exact WSD value copied from list-ports>`
      - `Processor = WinPrint`
      - `Datatype = RAW`
      - `Comment = Created via WPF`
      - `Location = Lab TI`
   - Action: click `create`.
   - Expected result: `Queue created.`

9. **List queues and confirm creation**
   - Suggested input: none.
   - Action: click `list`.
   - Expected result: queue `POC-WPF-QUEUE-01` appears with selected port/driver.

10. **Inspect created queue**
   - Suggested input: `Queue = POC-WPF-QUEUE-01`.
   - Action: click `inspect`.
   - Expected result: shows `Queue`, `Port`, `GlobalWpp`, `Classification`, and `Details`.

11. **Read ticket info for created queue**
    - Suggested input: `Queue = POC-WPF-QUEUE-01`.
    - Action: click `ticket-info`.
    - Expected result: `Available` and `Details`; attributes listed when supported by runtime/queue.

12. **Update default ticket (queue scope)**
    - Suggested inputs:
      - `Queue = POC-WPF-QUEUE-01`
      - `Ticket Duplexing = TwoSidedLongEdge`
      - `Ticket Color = Color`
      - `Ticket Orientation = Landscape`
    - Action: click `ticket-update-default`.
    - Expected result:
      - output includes `Scope: default` and `Applied: True`;
      - output shows both `Requested` and `Effective` values;
      - if constrained by driver, `Details` explains requested vs applied values.

13. **Update user ticket (current user scope)**
    - Suggested inputs:
      - `Queue = POC-WPF-QUEUE-01`
      - `Ticket Duplexing = OneSided`
      - `Ticket Color = Monochrome`
      - `Ticket Orientation = Portrait`
    - Action: click `ticket-update-user`.
    - Expected result:
      - output includes `Scope: user` and `Applied: True`;
      - output shows both `Requested` and `Effective` values;
      - if constrained by driver, `Details` explains requested vs applied values.

14. **Update + delete + confirm removal**
    - Suggested inputs for update:
      - `Queue = POC-WPF-QUEUE-01`
      - `Comment = Updated via WPF`
      - `Location = Floor 2`
   - Actions:
     1. Click `update`.
     2. Click `delete` (same queue).
     3. Click `list`.
   - Expected result:
      - `Queue updated.`
      - `Queue deleted.`
      - final `list` does not include `POC-WPF-QUEUE-01`.

## Ticket update notes (WPF)
- Ticket fields are optional for queue creation (`create`) and metadata update (`update`).
- Ticket fields are used only by `ticket-update-default` and `ticket-update-user`.
- Accepted values follow `System.Printing` enum names (examples):
  - `Duplexing`: `OneSided`, `TwoSidedLongEdge`, `TwoSidedShortEdge`
  - `OutputColor`: `Color`, `Monochrome`
  - `PageOrientation`: `Portrait`, `Landscape`
