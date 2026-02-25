# PDC WinForms GUI

Windows GUI version of the AVX PDC (Production Data Collection) system.
Replaces the Clipper/Harbour console UI with C# WinForms while preserving the same business logic and ADS/DBF data.

## Status

**Phase 0 — Scaffold complete.**

Working:
- Login dialog
- Main window with full PDC menu structure
- DBF/DAT file viewer (Tools > Open DBF Table, Ctrl+O)
- Status bar with user info and clock
- Keyboard shortcuts (F1-F6)

Stubs (pending implementation):
- Batch Operations (Arrive, Enter, Leave, Cancel, Reprint)
- Queries (WIP by Process, WIP by Workstation, Batch Path)
- Shipments (Outside Prod, Proforma Invoice)
- Labels (Packing, Production)
- Reports (Yield Analysis, Production Area Stages)

## Build

### Requirements
- .NET 8.0 SDK
- Windows 10/11

### Compile and Run
```
cd PdcGui
dotnet build
dotnet run
```

Or with username:
```
dotnet run -- VITALY
```

## Project Structure

```
PDC-harbour-WINGUI/
  PdcGui/
    PdcGui.csproj      .NET 8 WinForms project
    Program.cs          Entry point + login dialog
    MainForm.cs         Main window, menu, DBF viewer
```

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| DbfDataReader | 1.0.0 | Read DBF/DAT files (Clipper/dBASE format) |

When ADS client (`ace32.dll`) becomes available:
- `Advantage.Data.Provider` — native ADS access with indexes, locks, transactions

## Reference

The original Harbour/Clipper source is in `PDC-clean/` (not in this repo).
See that project for business logic that will be ported screen by screen.

## Migration Plan

| Phase | Scope | Risk |
|-------|-------|------|
| 0 | Scaffold + DBF viewer | Done |
| 1 | Read-only queries (WIP, Batch Path, Shipments) | Low |
| 2 | Input forms (Batch Arrive/Enter/Leave) | Medium |
| 3 | Shipments + Label printing | Medium |
| 4 | Reports (Yield, Production Area) | Medium |
| 5 | Full transition, Harbour retired | - |
