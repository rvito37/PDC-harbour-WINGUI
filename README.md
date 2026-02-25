# PDC WinForms GUI

Windows GUI version of the AVX PDC (Production Data Collection) system.
Replaces the Clipper/Harbour console UI with C# WinForms while preserving the same business logic and ADS/DBF data.

## Status

**Phase 0 — Scaffold + ADS integration complete.**

Working:
- Login dialog
- Main window with full PDC menu structure
- ADS Local Server integration (ace32.dll + adsloc32.dll bundled)
- DBF/DAT file viewer via ADS (Tools > Open DBF Table, Ctrl+O)
- Fallback to DbfDataReader if ADS unavailable
- Status bar with ADS status, user info, and clock
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
    PdcGui.csproj      .NET 8 WinForms project (x86 target)
    Program.cs          Entry point + login dialog
    MainForm.cs         Main window, menu, ADS/DBF viewer
    ADS/
      ace32.dll         ADS client library (32-bit)
      adsloc32.dll      ADS Local Server engine
      axcws32.dll       ADS communication layer
      adslocal.cfg      Local server configuration
      ansi.chr          ANSI character set
      extend.chr        Extended character set
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

This reads the same DBF/CDX files as the Harbour PDC — both can run simultaneously against the same data.

## Reference

The original Harbour/Clipper source is in `PDC-clean/` (not in this repo).
See that project for business logic that will be ported screen by screen.

## Migration Plan

| Phase | Scope | Status |
|-------|-------|--------|
| 0 | Scaffold + ADS + DBF viewer | Done |
| 1 | Read-only queries (WIP, Batch Path, Shipments) | Next |
| 2 | Input forms (Batch Arrive/Enter/Leave) | Planned |
| 3 | Shipments + Label printing | Planned |
| 4 | Reports (Yield, Production Area) | Planned |
| 5 | Full transition, Harbour retired | - |
