# PlateNova / plate_orbit

Consolidated project for the PlateNova Excel add-in: an Excel VBA front end that
drives a compiled VB.NET COM DLL (`RaghavStaadExtractor`) to pull beam/plate data
out of STAAD models, sort/finalize sections, and generate DXF drawings.

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

The `vba/` modules never contain real logic — every button in the workbook
calls `CreateObject("RaghavStaadExtractor.<Class>")` and invokes a method on it.
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

`com/RaghavTekNova/` is a separate, related COM project (bolt lists, assembly
lists, drawing tracking) that lives alongside `RaghavStaadExtractor` on disk.
Nothing in `vba/` currently calls it; it's included here because it's part of
the same author's toolset, not because it's wired into this workbook.

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
  licensing logic: AES-encrypted trial-date check against
  `C:\PlateNova\PlateNova.tst`, a clock-rollback tamper check against
  `%AppData%\RaghavStaad.last`, a hard expiry date, and a bank of 10
  deliberately fake runtime-error strings shown on failure instead of an
  honest license error (anti-crack decoys). Explains why it was kept out of
  the working source tree and behind `ConfuserEx` obfuscation.

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
