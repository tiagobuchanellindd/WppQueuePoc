# WPP Queue API Objectives Summary

## What this API currently delivers
1. Queue lifecycle operations in Windows spooler: create, list, update, and delete.
2. Discovery endpoints for supporting data: ports, drivers, print processors, and datatypes.
3. WSD port administration via Xcv monitor command (`AddPort`) with status-aware validation.
4. Global WPP status detection from Windows Registry (`Enabled`, `Disabled`, `Unknown`).
5. Queue inspection heuristics (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`) using global status + port/APMON evidence.
6. Print ticket diagnostics and updates for both scopes:
   - default queue ticket,
   - current user ticket.

## Why this matters
- WPP is a global Windows policy and directly changes expected queue behavior.
- Printing operations require a reliable operational workflow with clear diagnostics.
- Native spooler calls are complex; this API centralizes them behind safer service contracts.

## Current project intent
- Keep a service/API-first architecture with explicit operation results.
- Preserve interoperability fidelity with `winspool.drv` while returning actionable messages.
- Provide study-ready documentation in two tracks:
  - concise technical guides in `docs/guides/`,
  - deep method-level notes in `docs/estudo/` (PT-BR).
