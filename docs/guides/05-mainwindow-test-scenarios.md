# MainWindow Test Scenarios (POC)

This guide provides a compact test checklist based on the actions implemented in `WppQueuePoc.App/MainWindow.xaml.cs`.

The app is used as an execution harness for service/API behavior validation.

## Test setup
- Run as administrator (recommended for spooler/port operations).
- Use a valid installed driver name from `list-drivers`.
- Use an existing WSD port from `list-ports` when `add-wsd-port` is not supported.
- Keep one queue name for the flow, for example: `POC-QUEUE-UI-01`.

## Core happy-path flow
1. Click `wpp-status`.
2. Click `list-ports` and pick a `WSD-*` port.
3. Click `list-drivers` and copy one exact driver name.
4. Fill queue fields and click `create`.
5. Click `list` and confirm queue appears.
6. Click `inspect`.
7. Click `ticket-info (default)` and `ticket-info (user)`.
8. Fill at least one ticket field and click `ticket-update-default`.
9. Fill at least one ticket field and click `ticket-update-user`.
10. Update queue metadata (`comment` or `location`) and click `update`.
11. Click `delete` and confirm removal with `list`.

## Scenario matrix by UI action

| UI Action | Example input | Expected result |
|---|---|---|
| `wpp-status` | none | Result shows `Enabled`, `Disabled`, or `Unknown` with source/details |
| `list` | none | Installed queues are listed |
| `list-ports` | none | Ports are listed (including possible `WSD-*`) |
| `list-drivers` | none | Installed drivers are listed |
| `list-processors` | none | Print processors are listed (usually includes `WinPrint`) |
| `list-data-types` | `Print Processor = WinPrint` | Datatypes are listed (for example `RAW`) |
| `add-wsd-port` | `Port Name = WSD-<id>` | Port created, or descriptive error (`ERROR_NOT_SUPPORTED`, permissions, etc.) |
| `create` | `Queue`, `Driver`, `Port`, `Processor=WinPrint`, `Datatype=RAW` | Queue created and set as current queue |
| `update` | current queue + one field (`Comment=Updated`) | Queue updated |
| `delete` | current queue | Queue deleted and current queue cleared |
| `inspect` | current queue | Classification + details (`GlobalWpp`, port evidence, APMON when available) |
| `ticket-info (default)` | current queue | Ticket availability/details + attribute snapshot |
| `ticket-info (user)` | current queue | Ticket availability/details + attribute snapshot |
| `ticket-update-default` | current queue + one ticket field (`Duplexing=TwoSidedLongEdge`) | Update result with applied/effective details |
| `ticket-update-user` | current queue + one ticket field (`OutputColor=Monochrome`) | Update result with applied/effective details |

## Validation and negative tests

| Case | Action | Expected validation/output |
|---|---|---|
| Missing required fields on create | Click `create` with empty required inputs | Warning: required fields list (`Queue Name`, `Driver Name`, etc.) |
| Missing processor for data types | Click `list-data-types` with empty processor | Warning asking for print processor name |
| Missing queue name | Click `inspect`/`delete`/`ticket-*` with no queue and no current queue | Warning asking to specify queue name or select current queue |
| Empty ticket update request | Click `ticket-update-default` or `ticket-update-user` with all ticket fields empty | Warning asking at least one ticket field |
| Update with no changes | Click `update` with no updatable fields | Warning asking for at least one update field |

## Example values for ticket updates
- `Duplexing`: `OneSided`, `TwoSidedLongEdge`, `TwoSidedShortEdge`
- `OutputColor`: `Color`, `Monochrome`
- `PageOrientation`: `Portrait`, `Landscape`

Note: supported values depend on driver/printer capabilities; requested and effective values may differ.
