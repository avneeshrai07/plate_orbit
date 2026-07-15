# decompile_dll

Re-extracts `com/RaghavStaadExtractor_decompiled_fresh/` from `bin/RaghavStaadExtractor.dll`
without needing Visual Studio, the .NET SDK, or `ilspycmd` (a .NET SDK is not
assumed to be installed — only the shared runtime).

## Setup (one-time)

```
pip install pythonnet
python fetch_decompiler.py
```

`fetch_decompiler.py` downloads `ICSharpCode.Decompiler` straight from
NuGet's flat-container API as a raw `.nupkg` (just a zip) and extracts the
`netstandard2.0` build into `decompiler_pkg/` — no `dotnet tool install`
required (that needs an SDK, not just the runtime).

## Run

```
python dump_all.py
```

Writes one `.cs` file per `RaghavStaadExtractor.*` type into
`../../com/RaghavStaadExtractor_decompiled_fresh/`.

## Why member-by-member

Decompiling a whole type in one call crashes on some of the larger methods in
this environment (`StaadDataExtractor.ProcessSectionFarmingAndWeight`, etc.) —
`System.TypeLoadException: Could not load type
'System.Reflection.Metadata.TypeName'`, a version mismatch between what the
decompiler package expects and the `System.Reflection.Metadata` copy bundled
with the installed .NET runtime (this reproduced identically across decompiler
versions 7.2.1 / 8.2.0 / 10.1.1, so it's a runtime-side limit, not a decompiler
bug to chase). Decompiling field-by-field / property-by-property /
method-by-method sidesteps it entirely and got 100% (or near-100%, a couple of
struct field initializers aside) coverage — see `dump_all.py` for the loop.

## What this is for

`com/RaghavStaadExtractor_decompiled/` (no `_fresh` suffix) is the original
ilspycmd-based decompile the hand-ported `com/RaghavStaadExtractor/*.vb` files
were produced from (see the main README's "Recovered via decompilation"
section). `RaghavStaadExtractor_decompiled_fresh/` is a second, independent
decompile pulled straight from the DLL by this tool — useful for verifying the
hand-ported VB.NET against the compiled binary again later without repeating
the environment archaeology (no SDK, `ilspycmd` unavailable, decompiler
version hunting) that went into building this script.
