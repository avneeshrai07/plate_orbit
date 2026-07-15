# PlateNova / plate_orbit

Consolidated project for the PlateNova Excel add-in: an Excel VBA front end that
drives a compiled VB.NET COM DLL (`RaghavStaadExtractor`) to pull beam/plate data
out of STAAD models, sort/finalize sections, and generate DXF drawings.

## Why this repo exists

The real logic of the PlateNova extractor lived only inside the compiled
`RaghavStaadExtractor.dll`. The former developer who wrote it (the VBA/add-in
carries the author tags "Peeyush" / "Raghav") left the company without handing
over the source. What's here was reconstructed by **decompiling the shipped
DLL** and hand-porting the lost classes back to VB.NET (see "Recovered via
decompilation" below). This repo therefore serves two purposes:

1. preserve a buildable source copy of the add-in, and
2. act as the **reference specification for OptiPEB** — the from-scratch Python
   reimplementation of exactly this tool (see "Relationship to OptiPEB").

## Layout

```
vba/            VBA modules exported from the host workbook's VBA project
com/
  RaghavStaadExtractor/            VB.NET source for the main COM add-in (RaghavStaadExtractor.dll)
                                    — complete: all files the .vbproj references are present,
                                    including StaadDataExtractor.vb / DxfExporterRaghav.vb /
                                    ExecutionValidation.vb, ported back from decompiled C# (see below)
  RaghavStaadExtractor_decompiled/ Full C# decompile of bin/RaghavStaadExtractor.dll (via ilspycmd) —
                                    kept as the source of truth these 3 files were ported from
  RaghavStaadExtractor_decompiled_fresh/  A second, independent decompile of the same DLL (member-by-
                                    member, via tools/decompile_dll/ — see below), produced later to spot-
                                    check the hand-ported VB.NET against the compiled binary again. Not a
                                    replacement for RaghavStaadExtractor_decompiled/, just corroboration.
  RaghavTekNova/                   VB.NET source for a sibling COM add-in (bolt/assembly/drawing lists)
                                    — not currently referenced by anything in vba/, kept for reference
tools/
  decompile_dll/                   Re-runnable script to regenerate RaghavStaadExtractor_decompiled_fresh/
                                    from bin/RaghavStaadExtractor.dll without needing the .NET SDK or
                                    ilspycmd (uses ICSharpCode.Decompiler via pythonnet instead) — see its
                                    own README for setup/usage and why it decompiles member-by-member.
bin/            Latest compiled RaghavStaadExtractor.dll/.tlb + netDxf dependency
workbooks/      Host Excel workbook (PlateNova) and the section database workbook
samples/staad/  Sample .STD STAAD models for testing extraction
```

## How it fits together

Most `vba/` modules are thin shims — the button calls
`CreateObject("RaghavStaadExtractor.<Class>")` and invokes one method on it.
`com/RaghavStaadExtractor/` is the VB.NET source that compiles to that DLL
(registered for COM interop, exposed via `<ProgId(...)>` attributes).

| VBA macro (`vba/`) | ProgID / class | Source file |
|---|---|---|
| `ExtractStaadDataR1` (A1_StaadExtraction.vba) | `StaadDataExtractor` | `com/RaghavStaadExtractor/StaadDataExtractor.vb` (ported from decompile — see below) |
| `ByPassBeamDataFromStaad`, `ByPassForAnalysis` (BYPASS_BOQMODE.vba) | `StaadDataExtractor` | same as above |
| `SectionSortingS2` (A2_create_Mps.vba) | `SectionSortingS2` | `com/RaghavStaadExtractor/SectionSortingS2.vb` |
| `SectionFinalization` (A2_create_Mps.vba) | `SectionFinalization` | `com/RaghavStaadExtractor/SectionFinalization.vb` |
| `ExportToDXF` (A3_DxfCreation.vba) | `DxfExporterRaghav` | `com/RaghavStaadExtractor/DxfExporterRaghav.vb` (ported from decompile — see below) |
| `AssignSectionDatabase` (ASSIGN_SAVEAS_CLEAR.vba) | `AssignSectionDatabase` | `com/RaghavStaadExtractor/AssignSectionDatabase.vb` |
| `SaveBeamOutput`, `SavePlateSummary` (ASSIGN_SAVEAS_CLEAR.vba) | `SaveSheetsManager` | `com/RaghavStaadExtractor/SaveAsPlate.vb` |
| `Clear_Sheet1/2/AllSheets` (ASSIGN_SAVEAS_CLEAR.vba) | `ClearSheet` | `com/RaghavStaadExtractor/ClearSheet.vb` |
| `BoqMode`, `DrawingMode`, `NormalMode` (DIFFRENT_SHEET_MODES.vba) | `ColumnHider` | `com/RaghavStaadExtractor/ColumnHider.vb` |
| `UnlockRange` (DIFFRENT_SHEET_MODES.vba) | `Unlock` | `com/RaghavStaadExtractor/Unlock.vb` |
| `QuickToggleFilter`, `QuicksaveFilter` (Module2.vba) | `Unlock` | `com/RaghavStaadExtractor/Unlock.vb` |
| `CallGridSystem` (GRID_SYSTEM.vba) | `GridSystem` | `com/RaghavStaadExtractor/My Project/GridSystem.vb` |
| `ExportGrid` (Module1.vba) | `RaghavStaadExporter.Grid` | **not found** — confirmed not in `RaghavStaadExtractor.dll`; see "Still missing" below |
| — (not called by any current VBA) | `RaghavDatabase` | `com/RaghavStaadExtractor/RaghavDatabase.vb` |
| — (not called by any current VBA) | `SheetSwitch` | `com/RaghavStaadExtractor/My Project/SheetSwitch.vb` |
| — (not called by any current VBA) | `SupportReaction` | `com/RaghavStaadExtractor/SupportReaction.vb` |

A handful of `vba/` modules **do** carry their own logic (they aren't COM shims):

- **`KgToTon.vba`** (`ConvertWeightUnits`) — toggles the Sheet2 material summary
  between kg↔ton and mm↔m, stashing the original kg in column Z.
- **`DRAWING_MODE_VIEWS.vba`** — `FrontAndRun` / `BackAndRun` / `LeftAndRun` /
  `RightAndRun` / `TopAndRun`: each writes the view code into `AR2`/`AS2`, then
  chains `ByPassBeamDataFromStaad` + `ExportToDXF` (the elevation-drawing buttons).
- **`GRID_SYSTEM.vba`** — `WriteX` / `WriteZ` / `WriteG` set the `AW2`
  frame-analysis prompt (alongside `CallGridSystem`).
- **`sheet2.sheet`** — a `Worksheet_Change` handler that live-sums column E into
  `F4`.
- **`OPEN_LINKEDIN.vba`** — opens the author's LinkedIn.

`com/RaghavTekNova/` is a **separate sibling COM project** — a Tekla-Structures
BOQ/fabrication tool, unrelated to STAAD. Nothing in this workbook's `vba/` calls
it; it lives in its own workbook. Full breakdown in
"[The RaghavTekNova sibling add-in](#the-raghavteknova-sibling-add-in)" below.

## The data pipeline (what the tool actually does)

Everything runs against a **live STAAD.Pro session** (`StaadPro.OpenSTAAD`) and a
**live Excel** instance — nothing parses `.STD` files directly. The PlateNova
workbook has four working sheets, all protected with password **`2022`**:

| Sheet | Role |
|---|---|
| **Sheet1** | Master beam list — raw extraction + per-member section "farming" and weight |
| **Sheet2** | Per-plate summary → aggregated **Material Summary List** (the BOM) |
| **Sheet3** | STAAD section database, copied from `RAGHAV DATABASE.xlsm` (a 32-sheet catalog of built-up profiles like `I80012B50012`, with `Area(cm²)` → `kg/m`) |
| **Sheet4** | "Nova" grid system — X/Z grid names + coordinate spacing |

### 1. Extract → Sheet1 (`StaadDataExtractor.ExtractBeamDataFromStaad`)

Iterates the **currently selected beams** in STAAD (`GetNoOfSelectedBeams` /
`GetSelectedBeams` — *selected*, not the whole model) and per member reads
`GetBeamLength`, `GetBeamSectionName`, `GetBeamSectionPropertyRefNo`,
`GetBeamSectionPropertyTypeNo`, `GetBetaAngle`, `GetMemberIncidence` (+
`GetNodeCoordinates` for both ends), and the raw section-parameter array
`GetSectionPropertyValuesEx` (F1..F7 for tapers, up to 11 slots otherwise). The
job header block comes from `GetFullJobInfo`.

### 2. Section farming + weight (`ProcessSectionFarmingAndWeight`)

Decomposes each member into plates (WEB / TOP FLANGE / BOTTOM FLANGE) and
computes weights, dispatching on the STAAD **section type code**:

| Type code | Handling |
|---|---|
| **680** TAPER | general I-section: web depth minus flange thicknesses, two flanges from F4..F7 |
| **697** UPT general | double section ("2x…"); split into two equal plates when flange ≠ web |
| **690–699** UPT | area-based total weight only, tagged `UPT` |
| **613 / 614 / 615** | prismatic plate composites (web-only / flange-only / both) |
| **671 / 672** | passed through with **zero** computed weight |
| name = `TUBE` / `PIPE` / `ROD` | closed-section formulas from outer/inner dimensions |
| default | rolled I-section: average web + top/bottom flange |

Steel density is hard-coded **7850 kg/m³** (written as `7850`, or `7.85` alongside
mm dimensions). Weights round to 3 decimals.

### 3. Summarize → Sheet2 (`SectionSortingS2` → `SectionFinalization`)

`SummerizedToSheet2` explodes Sheet1 into one row per plate
(`THK./SECTION | LENGTH(mm) | WIDTH(mm) | WEIGHT(kg)`); `FinalSummarySheet2`
then groups identical sections into the **Material Summary List** (BOM), summing
length and weight.

### 4. Grids + frame analysis (optional, Sheet4)

The grid definitions live on Sheet4: `GridSystem.ShowGridInputForm` pops a
borderless WinForms dialog ("DEFINE GRIDS AS PER STAAD'S AXIS SYSTEM") with three
column groups — **grid naming** (`B2:B4` = X/Y/Z name lists), **grid
co-ordinates** (`C2:C4` = offset strings like `0 5*6000`), and **line
extensions** (`D2:D4` margins, `E2` bubble text height) — plus FLIP X / FLIP Z
(reverse a name list) and AUTO (fill default A–Z / 1–80 names or swap X↔Z) helpers.

The `X` / `Z` / `G` prompt in cell `AW2` then drives frame analysis
(`StaadDataExtractor`): beams are bucketed by their constant X or Z coordinate
into named grids, and frames are compared for geometric similarity (same
sections, spans, beta angle) to find repeated frames. `G` regenerates grid
definitions; `X` / `Z` restrict analysis to predefined grids and write a
similarity report to a `Frame Analysis/` folder.

### 5. Drawings (`DxfExporterRaghav`, `netDxf`)

`ExportBeamsToDxfWithPrompt` reads each member's end coordinates (cols P–U) from
Sheet1, scales m → mm (×1000), applies a **view rotation** (`AS2` = 1 FRONT /
2 BACK / 3 LEFT / 4 RIGHT / 5 TOP) and the member's beta angle, and draws a true
**3D wireframe cross-section** per STAAD section-type code — pipe/rod
(655/660), square/rect hollow tube (651/654/650/672), tapered I (680/610/697),
tapered channel (630), and I-sections with extra cover plates (613/614/615) —
built from `netDxf` `Line`/`Circle`/`Text`/`MText` entities on `Beams` /
`LeaderLines` / `BeamText` / `GRID_*` layers. View-mode flags in `AS5` switch to
per-section colour layers (100), colour layers without length text (101), or
single-line mode (200). Each member gets a leader line + `MText` size label
(web/TF/BF from cols 37–39, length from col 3). It also renders the Sheet4 grid
as named bubbles (circle + text) around the model extents. Output goes to
`<workbook>\PlateNovaDrawings\<name>.dxf` (AutoCad2010), with a success dialog
offering "Show in Folder" / "Open in CAD".

### Supporting classes

**SupportReaction** — per-node Fx/Fy/Fz/Mx/My/Mz over a load-case range via
`Output.GetSupportReactions`, with global min/max (present but not wired to any
current button). **ColumnHider** — BOQ / Drawing / Normal column views (F:AJ are
intermediate calc columns, always hidden). **ClearSheet** / **SaveSheetsManager**
(Save-As beam output / material summary). **AssignSectionDatabase** +
**RaghavDatabase** — build the Sheet3 unit-weight DB (`kg/m = Ax(cm²) × 0.785`)
from the 32-sheet `RAGHAV DATABASE.xlsm` catalog. **Unlock**, **GridSystem**
(WinForms input form). **SheetSwitch** — `SwitchToSheet1/2` nav plus
`NormalizeView`, which restores the full Excel UI (ribbon/headings/gridlines)
but is gated behind a **separate owner password `AVPR`** (not the `2022` sheet
password). Every class starts with an `ExecutionValidation` license check.

## Relationship to OptiPEB (the Python rewrite)

**OptiPEB** (`C:\Users\avnee\Desktop\OptiPEB`) is a ground-up reimplementation of
this add-in in Python — the same OpenSTAAD extraction and section/weight logic,
but delivered as a FastAPI service (+ Next.js frontend) instead of Excel/VBA
driving a COM DLL. This repo is its reference spec. Rough mapping:

| PlateNova (VB.NET here) | OptiPEB (Python) |
|---|---|
| `StaadDataExtractor.ProcessBeamData` | `routes/beam_data` + `services/{geometry,property}_service` |
| `ProcessSectionFarmingAndWeight` (F1..F7 → plates) | `core/section_geometry.taper_geometry` + `services/sections_service` |
| `SectionSortingS2.SummerizedToSheet2` (per-plate rows) | `SectionPlate` / `build_member_plates` |
| `SectionFinalization.FinalSummarySheet2` (BOM) | `build_sections_bom` (`/sections` endpoint) |
| Nova grid system (Sheet4) | OptiPEB grid systems |
| 7850 kg/m³, mm output, 345 MPa yield default | `services/sections_service` (same constants) |

Key behavioural differences to keep in mind when porting: PlateNova extracts the
**selected** beams into Excel, whereas OptiPEB serves **all** members over HTTP;
and OptiPEB reads per-member design yield (FYLD) from the steel design brief,
which PlateNova does not. The 7850 kg/m³ density, mm rounding, and section-type
decomposition carry over verbatim.

## The RaghavTekNova sibling add-in

`com/RaghavTekNova/` is a **completely separate** COM add-in (assembly name
`RaghavTekNova`) with **no STAAD involvement** and **no connection to OptiPEB's
scope**. Where PlateNova extracts a *design* model from STAAD, RaghavTekNova is a
downstream **fabrication BOQ** tool: it imports **Tekla Structures** report
exports and reformats them into billing/material lists. It ships in its own Excel
workbook (not the PlateNova one — nothing in this repo's `vba/` calls it) and
reuses the same licensing, `2022` sheet password, 2027-04-09 expiry, and a
"Raghav" print footer.

### Inputs (external Tekla `.xls`/`.xlsm` exports, placed next to the workbook)

`PR_Assembly_list_Vba.xls` (+ `_Net_`/`_Gross_` variants), `PR_Bolt_list_Vba.xls`,
`Assigned Bolt.xlsm`, `PR_Material_Part_list_vba.xls`,
`PR_Drawing_Tracking_list_Vba.xls`.

### Workbook sheet map

| Sheet | Role | Sheet | Role |
|---|---|---|---|
| Sheet1 | Assembly List (formatted BOM) | Sheet7 | bolt staging (imported) |
| Sheet2 | Bolt List | Sheet8 | assigned-bolt / bolt standards |
| Sheet3 | BOQ Summary | Sheet9 | part staging (imported) |
| Sheet4 | Part List | Sheet10 | Drawing Tracking (formatted) |
| Sheet5 | Material Summary | Sheet11 | drawing staging (imported) |
| Sheet6 | assembly staging (imported) | Sheet12+ | per-category split sheets |

### Classes (each ProgId is `RaghavTekNova.<Class>`)

| Class / file | Key methods | What it does |
|---|---|---|
| `GetAssemblyList` | `A1_GET…ASSEMBLY_LIST`, `B1_Get…Bolt_List`, `B2_Copy_Assigned_BoltData`, `P1_GET_MATERIAL_PART_LIST`, `V1_GET…Drawing_Tracking`, `SUMMARY_C1` | Import each Tekla `.xls` export into its staging sheet (6/7/8/9/11); build the Sheet3 **BOQ summary** by totaling weight(kg)/area(sqm) across the category sheets (12+) |
| `RunAssemblyList` | `A2_RUN…ASSEMBLY_LIST` | Format staging Sheet6 → Sheet1 as "BILL OF MATERIAL (ASSEMBLY LIST)" (SR/CATEGORY/ASSEMBLY/PROFILE/MARK/LEN/QTY/WEIGHT/AREA/PARTS/DRAWINGS/GUID), summing weight & area |
| `RunDrawingList` | `V2_RUN…Drawing_Tracking` | Format Sheet11 → Sheet10 as "BILL OF MATERIAL (DRAWING TRACKING LIST)" |
| `AssemblyFinal` | `A4_Finalized_Assembly_list` | Clear the PROFILE column for assemblies containing >1 part |
| `DuplicateAndUnlockedCheck` | `A3_Remove_Duplicate_Drawing_Check_Combine` (+ split `A3_Drawing_Check` / `A3_Remove_Duplicate` / toggle-filter) | Highlight assemblies whose drawing is **missing or unlocked in the Tekla model** (cols M/N); merge duplicate assembly marks (sum qty/weight/area, max unit values) |
| `BoltList` | `Run_PR_Bolt_List`, `Calculate_Total_Bolt_weight`, `Organize_Boltdata` | Match each bolt dia/length against a standards sheet for unit weight; build `M-`/`H-` grade names; consolidate duplicate bolts |
| `PartList` | `P2_RUN_MATERIAL_PART_LIST`, `P3_MATERIAL_LIST_SUMMARY`, `P4_SUMMERIZED_DATA` | Build the part list (Sheet9→Sheet4); summarize by profile+material (Sheet4→Sheet5); normalize plate profiles (`PL`/`CHEQ_PLT` → smallest dimension) and consolidate |
| `SplitAntiSplit` | `GenerateSheetsByCategory`, `DeleteAndRenumberSheets` | Split the assembly list into one sheet per category (sheets 12+); or delete the generated sheets and renumber |
| `GuidManager` (`Guid.vb`; also `DisplayGuid`/`RemoveGuid`) | `DisplayGuid` / `RemoveGuidDisplay` | Show/hide the assembly-GUID column O so a row can be searched back in the Tekla model |
| `ClearSheets` | `ClearAllSheets`, `A8/B8/S3/P7/P8/V4_Clear_Sheet…` | Reset each list sheet to a blank template with an on-sheet button-legend panel |
| `SaveAs` | `SaveSheet1..5/10As…`, `SaveCurrentSheet` | Export a list sheet to a clean `.xlsx` in `BOQ\` or `SaveAsBreakup\`, stripping controls/shapes, unlocking, deleting cols L+, and inserting a company logo into A2:B5 |

**It cannot be built as-is:** its `.vbproj` references `ExecutionValidation.vb`,
but that file is **absent from disk** (only `RaghavStaadExtractor` has a copy),
and no compiled `RaghavTekNova.dll` was found to decompile it from.

## Recovered via decompilation, ported back to VB.NET

`StaadDataExtractor.vb`, `DxfExporterRaghav.vb`, and `ExecutionValidation.vb`
were missing from every source location on disk (VS2022 working copy, all
`BACKUP/` snapshots, the one `.zip` installer bundle that could be inspected).
They only existed compiled into `bin/RaghavStaadExtractor.dll`.

Recovery was two steps:
1. Decompiled the DLL with `ilspycmd` (ICSharpCode.Decompiler) — decompilers
   only emit C#, regardless of the original language, so this produced a full
   C# project at `com/RaghavStaadExtractor_decompiled/` (kept as reference —
   it's the source of truth the VB.NET below was ported from).
2. Hand-ported those 3 classes from the decompiled C# back to VB.NET, cleaning
   up decompiler artifacts (`goto`-based control flow reconstructed from IL,
   `NewLateBinding`/`RuntimeHelpers.GetObjectValue`/`Conversions.` wrapper
   noise, compiler-generated closure classes for lambdas) into idiomatic,
   hand-written-quality VB.NET, while preserving exact behavior. Now living
   directly in `com/RaghavStaadExtractor/` alongside the files that were
   always there — **the project is complete**: every file
   `RaghavStaadExtractor.vbproj` references now exists.

The classes that matter:

- `StaadDataExtractor.vb` (1712 lines, from a 4448-line decompile) — the STAAD
  extraction engine every extraction macro depends on. Talks to a *live*
  STAAD.Pro session via the `StaadPro.OpenSTAAD` COM API (`GetSelectedBeams`,
  `GetBeamLength`, `GetMemberIncidence`, etc.) — it does not parse `.STD`
  files directly.
- `DxfExporterRaghav.vb` (1892 lines, from a 3187-line decompile) — the DXF
  export engine used by `ExportToDXF`, builds 3D beam geometry via the
  `netDxf` library.
- `ExecutionValidation.vb` (276 lines, from a 383-line decompile) — confirmed
  licensing logic (AES key/IV `AshuVedantRaghav`, PKCS7): during the free-trial
  window (through **2025-12-31**) no license file is needed; after that it
  requires `C:\PlateNova\ExecutionFile.lic`, whose lines are per-machine
  `<BIOS-serial>=<expiry|forever>` entries (BIOS serial via WMI `Win32_BIOS`).
  Anti-tamper: an AES-encrypted last-run timestamp in `C:\PlateNova\PlateNova.tst`
  rejects clock rollback; on any failure it shows one of **10 deliberately fake
  runtime-error strings** (decoys) and closes the workbook. Explains why it was
  kept out of the working source tree and behind `ConfuserEx` obfuscation.

  A **second, independent** expiry gate lives in each COM class's own
  `CheckExpirationDateAndBackdate` (e.g. in `StaadDataExtractor`): a hard expiry
  of **2027-04-09** plus a separate clock-rollback check against
  `%AppData%\RaghavStaad.last`.

`Guid`/`ProgId` attributes were checked against what `vba/` actually calls
(e.g. `[ProgId("RaghavStaadExtractor.StaadDataExtractor")]`) to confirm these
are the right classes, not just similarly-named ones. No compiler was
available in this environment (no Visual Studio/MSBuild) to do a real build
check — verification was manual (method-by-method comparison against the
decompiled source, checked for leftover decompiler artifacts, balanced
`Class`/`End Class` and `Function`/`End Function` blocks).

## Still missing

**`RaghavStaadExporter.Grid`** — the ProgID `Module1.vba`'s `ExportGrid` calls.
Confirmed (via `ilspycmd -l c` type listing) that it is **not** inside
`RaghavStaadExtractor.dll` at all — note the assembly name difference,
`RaghavStaadExporter` vs `RaghavStaadExtractor`. It's a genuinely separate DLL
that doesn't appear anywhere searched in `PN/`.
`com/RaghavStaadExtractor/My Project/DxfExporterGrid.vb` (the VB.NET source
tree's stub for it) is an empty 3-byte placeholder.

The two `.rar` archives in the original `PN/` folder
(`PlateNova v1.0.9.4.25_installer+files.rar`,
`PlateNova v1.00.09.04.25_installer+files.rar`) have since been checked (7-Zip
installed for this). Same story as the `.zip`: compiled distributable bundles
only (dll + xlsm + license + installer script), no `.vb`/`.cs` source, and no
`RaghavStaadExporter.Grid` DLL either. One interesting side note: the older
archive's `RaghavStaadExtractor.dll` (365KB, dated 2025-09-13) is an
obfuscated release build — `ConfusedByAttribute` and garbled Unicode class
names confirm ConfuserEx — of the same codebase, not a different/richer
version, so it adds nothing over the clean build already decompiled above.

`RaghavStaadExporter.Grid` remains genuinely unrecovered — it isn't in any
file searched across `PN/`, including both archives.

## Building the COM DLL

`com/RaghavStaadExtractor/` and `com/RaghavTekNova/` are VS2022 VB.NET Class
Library projects (COM-visible, registered via `regasm`/COM interop).
`com/RaghavStaadExtractor/` now has every file its `.vbproj` references and
should build in Visual Studio (needs `netDxf.netstandard` restored via
`packages.config`, and the `Microsoft.Office.Interop.Excel`/`Core` COM
references available). This has not been build-verified in this environment
(no Visual Studio/MSBuild installed here) — worth a real build the first time
it's opened in Visual Studio to catch anything the manual port missed.
`com/RaghavTekNova/` still can't build — it's missing its own copy of
`ExecutionValidation.vb` (not yet recovered; no compiled `RaghavTekNova.dll`
has been located anywhere to decompile). `bin/` holds the last known good
compiled `RaghavStaadExtractor.dll` so the workbook keeps working regardless
of whether a rebuild is attempted.
