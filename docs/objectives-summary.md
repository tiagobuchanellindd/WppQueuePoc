# WPP POC Objectives Summary

## What we are aiming to deliver
1. Create print queues in a WPP-compatible Windows environment (for example, Windows 11).
2. Provide queue lifecycle operations: create, list, update, and delete.
3. Provide a practical way to evaluate whether a queue is likely operating under WPP.
4. Expose the global WPP status from Windows Registry (`Enabled`, `Disabled`, or `Unknown`).

## Why this matters
- WPP is configured at OS level and directly affects print queue behavior.
- Operations/support teams need a predictable, testable flow to manage queues in WPP scenarios.
- This POC reduces risk before a production-grade implementation.

## Initial implementation scope
- WPF interface as the primary operation surface, using the same service contracts of the POC.
- First technical milestone: `wpp-status` action based on Registry lookup.
