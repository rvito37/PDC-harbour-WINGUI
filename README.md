# PDC WinForms GUI

Windows GUI version of the AVX PDC (Production Data Collection) system.
Replaces the Clipper/Harbour console UI with C# WinForms while preserving the same business logic and ADS/DBF data.

## Status

**Phase 1 — Read-only queries in progress.**

Working:
- Login dialog
- Main window with full PDC menu structure
- ADS Local Server integration (ace32.dll + adsloc32.dll bundled)
- DBF/DAT file viewer via ADS (Tools > Open DBF Table, Ctrl+O)
- Fallback to DbfDataReader if ADS unavailable
- Status bar with ADS status, user info, and clock
- Keyboard shortcuts (F1-F6)
- **WIP by Process (StockInProc)** — fully working query screen:
  - Process selector from C_PROC.DBF
  - Batch grid with all original Clipper columns (Batch, Pu, St, P_Line, Val, Strips, Pcs, Days, Slack, Priority, Cust, PO, Started)
  - Alt+S sort toggle: iSlack / nwislack / imslack / qcislack (matching original CDX index expressions)
  - **Real Slack calculation**: `Slack = B_dprom - (Today + LTime_NoRoute)` using m_linemv + c_leadt lead times (matching Clipper DELIVSCH.PRG)
  - Graceful fallback if lead time tables not available
  - Green row highlighting for started batches (matching Clipper B/Bg colorBlock)

Stubs (pending implementation):
- Batch Operations (Arrive, Enter, Leave, Cancel, Reprint)
- Queries (WIP by Workstation, Batch Path)
- Shipments (Outside Prod, Proforma Invoice)
- Labels (Packing, Production)
- Reports (Yield Analysis, Production Area Stages)

## Build

### Requirements
- .NET 8.0 SDK
- Windows 10/11 (x86 or x64)

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
    PdcGui.csproj          .NET 8 WinForms project (x86 target)
    Program.cs              Entry point + login dialog
    MainForm.cs             Main window, menu, ADS/DBF viewer
    StockInProcForm.cs      WIP by Process query (Clipper STOKPROC.PRG)
    ADS/
      ace32.dll             ADS client library (32-bit)
      adsloc32.dll          ADS Local Server engine
      axcws32.dll           ADS communication layer
      adslocal.cfg          Local server configuration
      ansi.chr              ANSI character set
      extend.chr            Extended character set
  TestQuery/
    Program.cs              Standalone ADS query test utility
```

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| Advantage.Data.Provider | 8.10.1.2 | Native ADS access — SELECT, indexes, locks, transactions |
| DbfDataReader | 1.0.0 | Fallback DBF reader (no ADS dependency) |

## ADS Connection

The app uses ADS Local Server mode (no server license required):
```
Data Source={directory};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;
```

This reads the same DBF/CDX files as the Clipper/Harbour PDC — both can run simultaneously against the same data.

### Required Data Tables

| Table | Records | Purpose |
|-------|---------|---------|
| D_LINE.DBF | ~134K | WIP batch data |
| C_PROC.DBF | ~50 | Process definitions |
| m_linemv.DBF | ~3.3M | Batch movement steps (for Slack/LTime) |
| c_leadt.DBF | ~17K | Lead times per process type (for Slack/LTime) |

## Reference

The original Clipper source is in `C:\Users\AVXUser\PDC\` (not in this repo).
This is the authoritative reference for all business logic ported screen by screen.

## Migration Plan

| Phase | Scope | Status |
|-------|-------|--------|
| 0 | Scaffold + ADS + DBF viewer | Done |
| 1 | Read-only queries (WIP, Batch Path, Shipments) | In Progress |
| 2 | Input forms (Batch Arrive/Enter/Leave) | Planned |
| 3 | Shipments + Label printing | Planned |
| 4 | Reports (Yield, Production Area) | Planned |
| 5 | Full transition, Harbour retired | - |
