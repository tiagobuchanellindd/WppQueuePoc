# WPP Business Overview

## Purpose
This POC exists to validate how queue management behaves when **Windows Protected Print (WPP)** is enabled, and what that means for operational workflows.

## Business view of how it works
- **WPP is a Windows-level setting**: it is enabled/disabled globally, not per queue.
- When WPP is enabled, queue behavior follows WPP policy automatically.
- In this model, traditional TCP/IP port flows are replaced by WSD-based behavior in supported scenarios.

## Operational impact
- Queue lifecycle operations still matter:
  - create queue,
  - update queue properties,
  - delete queue.
- The platform must also expose:
  - global WPP status (`Enabled`/`Disabled`/`Unknown`),
  - queue inspection signals indicating whether a queue is likely operating under WPP.

## Practical business example (without WPP vs with WPP)
Consider the same branch office printer and the same operational flow (create, inspect, update, delete):

| Context | What operations do | Expected POC interpretation |
|---|---|---|
| Global WPP **Disabled** | Create queue normally (even with WSD/IPP), then inspect | Queue can be created and used, but `inspect` should tend to `LikelyNotWpp` because global policy is not active |
| Global WPP **Enabled** | Recreate queue in same style (prefer WSD/IPP), then inspect | `inspect` should tend to `LikelyWpp` when queue signals are compatible (WSD/IPP/APMON evidence) |

This is the business value of the POC: same administration workflow, but with explicit diagnostics showing when the environment is likely under WPP policy.

## Why this POC is needed
- To reduce ambiguity for support and operations teams.
- To provide a repeatable way to test queue administration in WPP-enabled environments.
- To document expected outcomes and error patterns before moving to a production-grade implementation.
