# PlateNova / plate_orbit

Consolidated project for the PlateNova Excel add-in: an Excel VBA front end that
drives a compiled VB.NET COM DLL (`RaghavStaadExtractor`) to pull beam/plate data
out of STAAD models, sort/finalize sections, and generate DXF drawings.

## Layout

```
vba/            VBA modules exported from the host workbook's VBA project
com/
  RaghavStaadExtractor/   VB.NET source for the main COM add-in (RaghavStaadExtractor.dll)
  RaghavTekNova/          VB.NET source for a sibling COM add-in (bolt/assembly/drawing lists)
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
| `ExtractStaadDataR1` (A1_StaadExtraction.vba) | `StaadDataExtractor` | **missing** — see below |
| `ByPassBeamDataFromStaad`, `ByPassForAnalysis` (BYPASS_BOQMODE.vba) | `StaadDataExtractor` | **missing** |
| `SectionSortingS2` (A2_create_Mps.vba) | `SectionSortingS2` | `com/RaghavStaadExtractor/SectionSortingS2.vb` |
| `SectionFinalization` (A2_create_Mps.vba) | `SectionFinalization` | `com/RaghavStaadExtractor/SectionFinalization.vb` |
| `ExportToDXF` (A3_DxfCreation.vba) | `DxfExporterRaghav` | **missing** |
| `AssignSectionDatabase` (ASSIGN_SAVEAS_CLEAR.vba) | `AssignSectionDatabase` | `com/RaghavStaadExtractor/AssignSectionDatabase.vb` |
| `SaveBeamOutput`, `SavePlateSummary` (ASSIGN_SAVEAS_CLEAR.vba) | `SaveSheetsManager` | `com/RaghavStaadExtractor/SaveAsPlate.vb` |
| `Clear_Sheet1/2/AllSheets` (ASSIGN_SAVEAS_CLEAR.vba) | `ClearSheet` | `com/RaghavStaadExtractor/ClearSheet.vb` |
| `BoqMode`, `DrawingMode`, `NormalMode` (DIFFRENT_SHEET_MODES.vba) | `ColumnHider` | `com/RaghavStaadExtractor/ColumnHider.vb` |
| `UnlockRange` (DIFFRENT_SHEET_MODES.vba) | `Unlock` | `com/RaghavStaadExtractor/Unlock.vb` |
| `QuickToggleFilter`, `QuicksaveFilter` (Module2.vba) | `Unlock` | `com/RaghavStaadExtractor/Unlock.vb` |
| `CallGridSystem` (GRID_SYSTEM.vba) | `GridSystem` | `com/RaghavStaadExtractor/My Project/GridSystem.vb` |
| `ExportGrid` (Module1.vba) | `RaghavStaadExporter.Grid` | `com/RaghavStaadExtractor/My Project/DxfExporterGrid.vb` (stub, see below) |
| — (not called by any current VBA) | `RaghavDatabase` | `com/RaghavStaadExtractor/RaghavDatabase.vb` |
| — (not called by any current VBA) | `SheetSwitch` | `com/RaghavStaadExtractor/My Project/SheetSwitch.vb` |
| — (not called by any current VBA) | `SupportReaction` | `com/RaghavStaadExtractor/SupportReaction.vb` |

`com/RaghavTekNova/` is a separate, related COM project (bolt lists, assembly
lists, drawing tracking) that lives alongside `RaghavStaadExtractor` on disk.
Nothing in `vba/` currently calls it; it's included here because it's part of
the same author's toolset, not because it's wired into this workbook.

## Known gaps

`RaghavStaadExtractor.vbproj` and `RaghavTekNova.vbproj` both reference source
files that do not exist anywhere on disk (checked the VS2022 working copy, all
`BACKUP/` snapshots, and the one `.zip` installer bundle that could be inspected
without a `.rar`-capable tool):

- **`StaadDataExtractor.vb`** — the actual STAAD parsing/extraction engine.
  This is the most important missing piece; every extraction macro depends on it.
- **`DxfExporterRaghav.vb`** — the DXF export engine used by `ExportToDXF`.
- **`ExecutionValidation.vb`** — referenced by *both* projects, missing from
  both. Almost certainly licensing/anti-piracy logic, deliberately kept out of
  the working copies (there's a `ConfuserEx` obfuscator and an
  `ExecutionFile.lic` file sitting next to the compiled DLLs, which fits).
- **`com/RaghavStaadExtractor/My Project/DxfExporterGrid.vb`** is present but
  is an empty 3-byte stub, not real source, for the same `RaghavStaadExporter.Grid`
  class `Module1.vba`'s `ExportGrid` calls.

These three are only present compiled into `bin/RaghavStaadExtractor.dll`
(decompilation, not source recovery, would be the only way to recover them from
here). Two `.rar` archives in the original `PN/` folder
(`PlateNova v1.0.9.4.25_installer+files.rar`,
`PlateNova v1.00.09.04.25_installer+files.rar`) were **not** inspected — no
`7z`/`unrar` was available in this environment — but they follow the same
"installer+files" naming as the one `.zip` that *was* checked, which turned out
to be a compiled distributable bundle (dll + xlsm + license + installer
script) with no `.vb` source. Worth a manual check if the missing files still
need recovering.

## Building the COM DLL

`com/RaghavStaadExtractor/` and `com/RaghavTekNova/` are VS2022 VB.NET Class
Library projects (COM-visible, registered via `regasm`/COM interop). Neither
will currently build as-is because of the missing files listed above — they'd
need to be re-added (or stubbed) before compiling. `bin/` holds the last known
good compiled output so the workbook keeps working without a rebuild.
