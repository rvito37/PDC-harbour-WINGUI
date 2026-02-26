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
  - Process selector with descriptions from C_PROC.DBF ("ID - Name" format)
  - Batch grid with all original Clipper columns (Batch, Pu, St, P_Line, Val, Strips, Pcs, Days, Slack, Priority, Cust, PO, Started)
  - Alt+S sort toggle: iSlack / nwislack / imslack / qcislack (matching original CDX index expressions)
  - Sort modes per process: 190.0 (QC) → QCISLACK/ISLACK/NWISLACK, 001.0 (Metal) → IMSLACK only, other → ISLACK/NWISLACK
  - **Real Slack calculation**: `Slack = B_dprom - (Today + LTime_NoRoute)` using m_linemv + c_leadt lead times (matching Clipper DELIVSCH.PRG)
  - Graceful fallback if lead time tables not available
  - Green row highlighting for started batches (matching Clipper B/Bg colorBlock)
- **WIP by Workstation (StockInWork)** — fully working query screen:
  - Workstation selector with descriptions ("ID - Name" format) + GRU0-GRU7 process groups
  - 12 columns: Pu, Esn, Batch, Wfr, Pcs, Proc, Proc Name, St, Pr, Days, Slack, Comments
  - Styled column headers (dark blue background, white bold text) matching StockInProc
  - GRU groups filter by CDX process sets (Substrate Cleaning, Lito, Polymid, Chemical, Electroplating, Cure, Measurements, ET-6)
  - Real Slack calculation (shared SlackCalculator)
  - Green row highlighting for started batches
- **Batch Path Query (BatchPath)** — fully working query screen:
  - 6-digit batch number input (digits only, zero-padded)
  - Info panel: Purpose (c_bpurp), Type (c_btype), B/N, Location (m_loc)
  - 9-column grid: Pr, Stage, Process & Name, Wfr, Str, Pcs, Arrive/Start/Finish DateTimes
  - Green row highlighting for in-progress stages (arrived but not finished)
  - Pack memo popup for CP_PCCODE ≥ 4550 (read-only)
  - Two-level ESC: grid → input → close (matching Clipper behavior)
  - c_btype lookup via DbfDataReader (bypasses CDX Error 3010)
  - Location defaults to "Prod. floor" when m_loc has no data

Stubs (pending implementation):
- Batch Operations (Arrive, Enter, Leave, Cancel, Reprint)
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
    StockInWorkForm.cs      WIP by Workstation query (Clipper STOKWORK.PRG)
    BatchPathForm.cs        Batch Path Query (Clipper BNPATH.PRG)
    SlackCalculator.cs      Shared Slack/LTime_NoRoute calculation
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
Data Source={directory};ServerType=LOCAL;TableType=CDX;LockMode=COMPATIBLE;CharType=OEM;TrimTrailingSpaces=TRUE;
```

This reads the same DBF/CDX files as the Clipper/Harbour PDC — both can run simultaneously against the same data.

### ADS SQL Limitations

ADS Local Server uses a limited SQL dialect, **not** full SQL-92. Known limitations encountered during migration:

| Feature | Status | Workaround |
|---------|--------|------------|
| `SELECT`, `WHERE`, `ORDER BY`, `JOIN` | Supported | — |
| `IN (...)` clause | Supported | Used for GRU group process filters |
| `UPPER()`, `TRIM()` | Supported | Used for status filtering |
| `STR()` function | **Not supported** in SQL | Use plain column ordering instead |
| `IF()` / `IIF()` function | **Not supported** in SQL | Drop or use plain columns |
| `CHR()` function | **Not supported** in SQL | Use literal chars: `CHR(55)` → `'7'` |
| `CASE WHEN ... END` | **Not supported** in ORDER BY | Use plain columns |
| ISAM `Seek`/`SetRange` on conditional CDX | **Error 7038** | Use SQL `WHERE` instead |
| CDX with alias-qualified keys (`table->field`) | **Error 3010** | Use `DbfDataReader` NuGet to read DBF directly, bypassing CDX |
| Computed expressions in ORDER BY | **Very limited** | Use stored index fields (islack, qcislack, imslack) or plain columns |

**Key insight:** CDX indexes like ISLACK, QCISLACK, IMSLACK store precomputed sort keys as fields in the DBF.
These stored fields can be used directly in `ORDER BY islack` — much simpler than replicating the CDX key expression in SQL.
For computed indexes like NWISLACK (no stored field), use the individual columns in ORDER BY.

### Required Data Tables

| Table | Records | Purpose |
|-------|---------|---------|
| D_LINE.DBF | ~177K | WIP batch data (main table) |
| C_PROC.DBF | ~50 | Process definitions (proc_id, proc_nme/proc_nmh) |
| m_linemv.DBF | ~3.3M | Batch movement steps (for Slack/LTime_NoRoute) |
| c_leadt.DBF | ~17K | Lead times per process type (for Slack/LTime_NoRoute) |
| C_WKSTN.DBF | ~15 | Workstation definitions (wkstn_id, wkstn_nme) |
| C_BPURP.DBF | ~10 | Batch purpose codes (b_purp → bpurp_nme) |
| C_BTYPE.DBF | ~50 | Batch type codes (b_type+esnxx_id+pline_id → btype_nme) |
| M_LOC.DBF | ~1 | Batch location tracking (b_id → loc_nme) |

## Reference

The original Clipper source is in `C:\Users\AVXUser\PDC\` (not in this repo).
This is the authoritative reference for all business logic ported screen by screen.

Key Clipper source files:
| File | WinForms equivalent | Purpose |
|------|---------------------|---------|
| STOKPROC.PRG | StockInProcForm.cs | WIP by Process query |
| STOKWORK.PRG | StockInWorkForm.cs | WIP by Workstation query |
| DELIVSCH.PRG | SlackCalculator.cs | Slack / LTime_NoRoute calculation |
| BNPATH.PRG | BatchPathForm.cs | Batch Path Query |
| AVXFUNCS.PRG | (inline) | Shared utility functions (Slack, NotCommas) |

## Migration Plan

| Phase | Scope | Status |
|-------|-------|--------|
| 0 | Scaffold + ADS + DBF viewer | Done |
| 1 | Read-only queries (WIP, Batch Path, Shipments) | In Progress |
| 2 | Input forms (Batch Arrive/Enter/Leave) | Planned |
| 3 | Shipments + Label printing | Planned |
| 4 | Reports (Yield, Production Area) | Planned |
| 5 | Full transition, Harbour retired | - |
