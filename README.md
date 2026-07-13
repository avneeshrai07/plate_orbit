# PlateNova / plate_orbit

Consolidated project for the PlateNova Excel add-in: an Excel VBA front end that
drives a compiled VB.NET COM DLL (`RaghavStaadExtractor`) to pull beam/plate data
out of STAAD models, sort/finalize sections, and generate DXF drawings.

## Layout

```
vba/            VBA modules exported from the host workbook's VBA project
com/
  RaghavStaadExtractor/            VB.NET source for the main COM add-in (RaghavStaadExtractor.dll)
                                    — missing StaadDataExtractor.vb, DxfExporterRaghav.vb,
                                    ExecutionValidation.vb (recovered below, not ported back yet)
  RaghavStaadExtractor_decompiled/ Full C# decompile of bin/RaghavStaadExtractor.dll (via ilspycmd),
                                    recovers the 3 files missing from the VB.NET source tree
  RaghavTekNova/                   VB.NET source for a sibling COM add-in (bolt/assembly/drawing lists)
                                    — not currently referenced by anything in vba/, kept for reference
bin/            Latest compiled RaghavStaadExtractor.dll/.tlb + netDxf dependency
workbooks/      Host Excel workbook (PlateNova) and the section database workbook
samples/staad/  Sample .STD STAAD models for testing extraction
```

## How it fits together

The `vba/` modules never contain real logic — every button in the workbook
calls `CreateObject("RaghavStaadExtractor.<Class>")` and invokes a method on it.
`com/RaghavStaadExtractor/` is the VB.NET source that compiles to that DLL
(registered for COM interop, exposed via `<ProgId(...)>` attributes).

| VBA macro (`vba/`) | ProgID / class | Source file |
|---|---|---|
| `ExtractStaadDataR1` (A1_StaadExtraction.vba) | `StaadDataExtractor` | `com/RaghavStaadExtractor_decompiled/RaghavStaadExtractor/StaadDataExtractor.cs` (decompiled — see below) |
| `ByPassBeamDataFromStaad`, `ByPassForAnalysis` (BYPASS_BOQMODE.vba) | `StaadDataExtractor` | same as above |
| `SectionSortingS2` (A2_create_Mps.vba) | `SectionSortingS2` | `com/RaghavStaadExtractor/SectionSortingS2.vb` |
| `SectionFinalization` (A2_create_Mps.vba) | `SectionFinalization` | `com/RaghavStaadExtractor/SectionFinalization.vb` |
| `ExportToDXF` (A3_DxfCreation.vba) | `DxfExporterRaghav` | `com/RaghavStaadExtractor_decompiled/RaghavStaadExtractor/DxfExporterRaghav.cs` (decompiled — see below) |
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

`com/RaghavTekNova/` is a separate, related COM project (bolt lists, assembly
lists, drawing tracking) that lives alongside `RaghavStaadExtractor` on disk.
Nothing in `vba/` currently calls it; it's included here because it's part of
the same author's toolset, not because it's wired into this workbook.

## Recovered via decompilation

`StaadDataExtractor.vb`, `DxfExporterRaghav.vb`, and `ExecutionValidation.vb`
were missing from every source location on disk (VS2022 working copy, all
`BACKUP/` snapshots, the one `.zip` installer bundle that could be inspected).
They only existed compiled into `bin/RaghavStaadExtractor.dll`. Recovered by
decompiling that DLL with `ilspycmd` (ICSharpCode.Decompiler) — output is in
**`com/RaghavStaadExtractor_decompiled/`** as a full C# project (decompilation
only produces C#, not VB.NET, but it's functionally equivalent and readable).
The classes that matter:

- `RaghavStaadExtractor/StaadDataExtractor.cs` (4448 lines) — the STAAD
  parsing/extraction engine every extraction macro depends on.
- `RaghavStaadExtractor/DxfExporterRaghav.cs` (3187 lines) — the DXF export
  engine used by `ExportToDXF`.
- `RaghavStaadExtractor/ExecutionValidation.cs` (383 lines) — confirmed to be
  licensing logic: AES-encrypted trial-date check against
  `C:\PlateNova\PlateNova.tst`, plus a bank of fake runtime-error strings shown
  on failed validation (anti-tamper decoys). Explains why it was kept out of
  the working source tree and behind `ConfuserEx` obfuscation.

`Guid`/`ProgId` attributes on the decompiled classes were checked against what
`vba/` actually calls (e.g. `[ProgId("RaghavStaadExtractor.StaadDataExtractor")]`)
to confirm these are the right classes, not just similarly-named ones.

## Still missing

**`RaghavStaadExporter.Grid`** — the ProgID `Module1.vba`'s `ExportGrid` calls.
Confirmed (via `ilspycmd -l c` type listing) that it is **not** inside
`RaghavStaadExtractor.dll` at all — note the assembly name difference,
`RaghavStaadExporter` vs `RaghavStaadExtractor`. It's a genuinely separate DLL
that doesn't appear anywhere searched in `PN/`.
`com/RaghavStaadExtractor/My Project/DxfExporterGrid.vb` (the VB.NET source
tree's stub for it) is an empty 3-byte placeholder.

Two `.rar` archives in the original `PN/` folder
(`PlateNova v1.0.9.4.25_installer+files.rar`,
`PlateNova v1.00.09.04.25_installer+files.rar`) were never inspected — no
`7z`/`unrar` in this environment — but the one `.zip` archive that *could* be
checked, with identical "installer+files" naming, turned out to be a compiled
distributable bundle (dll + xlsm + license + installer script) with no source
at all. Worth a manual check only if `RaghavStaadExporter.Grid` still needs
recovering.

## Building the COM DLL

`com/RaghavStaadExtractor/` and `com/RaghavTekNova/` are VS2022 VB.NET Class
Library projects (COM-visible, registered via `regasm`/COM interop). Neither
will currently build as-is: the three files recovered by decompilation live in
`com/RaghavStaadExtractor_decompiled/` as C#, not as `.vb` files back in
`com/RaghavStaadExtractor/` — porting them back to VB.NET (or converting the
project to C#) is needed before the original project compiles again. The
`RaghavStaadExporter.Grid` gap is unresolved regardless. `bin/` holds the last
known good compiled output so the workbook keeps working without a rebuild.
