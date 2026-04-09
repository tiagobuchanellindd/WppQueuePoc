# WPP POC Plan

## Objective
Implement a Windows POC (for example, on Windows 11) to:
1. create a queue in a WPP scenario;
2. update and delete a queue;
3. verify indicators that a queue is operating under WPP;
4. identify whether global WPP is enabled in Windows.

## Approach
- Build a .NET WPF app with explicit command actions:
  - `wpp-status`
  - `add-wsd-port`
  - `create`
  - `list`
  - `list-ports`
  - `update`
  - `delete`
  - `inspect`
  - `ticket-info`
- Use Winspool integration through P/Invoke for queue operations.
- Create WSD ports via `OpenPrinter(",XcvMonitor WSD Port")` + `XcvData("AddPort")`.
- Read a Registry key to determine global WPP status.
- Keep architecture separated between WPF orchestration and native interop for study.

## Implementation backlog
1. Prepare the base POC structure.
2. Map and implement Registry-based WPP status detection.
3. Implement WSD port creation with `XcvData`.
4. Implement queue CRUD.
5. Implement queue inspection classification (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`).
6. Add logs and demonstration flow.
7. Validate end-to-end on Windows 11.
8. Reorganize project into study-friendly layers/classes.
9. Explore Print Tickets as a final advanced phase (initial read-only diagnostics implemented).

## Important notes
- WPP is a global Windows configuration, not an isolated property of a single queue.
- Spooler and port-monitor operations typically require administrative privileges.
- If the Registry mapping varies by Windows build, status may be reported as `Unknown` until the exact key is confirmed.
