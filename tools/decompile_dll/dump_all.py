"""
Decompile every RaghavStaadExtractor.* type in bin/RaghavStaadExtractor.dll into
one .cs file per type, member-by-member.

Why member-by-member: no .NET SDK is assumed to be present (only the shared
runtime), so `dotnet tool install ilspycmd` isn't available. This uses the
ICSharpCode.Decompiler library directly via pythonnet instead. Whole-type
decompile crashes this decompiler build on a handful of methods (an internal
ConvertConstantValue/ReflectionHelper issue tied to the runtime's
System.Reflection.Metadata version), but decompiling one member at a time and
concatenating is reliable and gets 100% (or near-100%) coverage.

Setup (one-time):
    pip install pythonnet
    python fetch_decompiler.py      # downloads ICSharpCode.Decompiler from NuGet

Run:
    python dump_all.py

Output: ../../com/RaghavStaadExtractor_decompiled_fresh/*.cs (one file per type)
"""
import os
import clr_loader
from pythonnet import set_runtime

set_runtime(clr_loader.get_coreclr())

import clr
HERE = os.path.dirname(os.path.abspath(__file__))
DECOMPILER_DLL = os.path.join(HERE, "decompiler_pkg", "lib", "netstandard2.0", "ICSharpCode.Decompiler.dll")
clr.AddReference(DECOMPILER_DLL)

from ICSharpCode.Decompiler.CSharp import CSharpDecompiler
from ICSharpCode.Decompiler import DecompilerSettings

TARGET_DLL = os.path.join(HERE, "..", "..", "bin", "RaghavStaadExtractor.dll")
OUT_DIR = os.path.join(HERE, "..", "..", "com", "RaghavStaadExtractor_decompiled_fresh")

# Skip generated/interop noise; keep the actual business-logic namespace(s).
NAMESPACE_PREFIXES = ("RaghavStaadExtractor.",)
SKIP_SUBSTRINGS = ("My.Resources", "My.MyApplication", "My.MyComputer", "My.MyProject",
                   "My.MySettings", "AnonymousDelegate")

settings = DecompilerSettings()
decompiler = CSharpDecompiler(TARGET_DLL, settings)
module = decompiler.TypeSystem.MainModule

os.makedirs(OUT_DIR, exist_ok=True)

summary = []

for t in module.TypeDefinitions:
    full_name = str(t.FullName)
    if not full_name.startswith(NAMESPACE_PREFIXES):
        continue
    if any(s in full_name for s in SKIP_SUBSTRINGS):
        continue

    chunks = [f"// ===== {full_name} =====", "// Decompiled member-by-member from RaghavStaadExtractor.dll", ""]
    ok_count = 0
    fail_count = 0

    for member_group in (t.Fields, t.Properties, t.Methods):
        for m in member_group:
            try:
                chunks.append(decompiler.DecompileAsString([m.MetadataToken]))
                ok_count += 1
            except Exception as e:
                chunks.append(f"// FAILED to decompile {m.Name}: {e}")
                fail_count += 1

    out_name = full_name.replace("RaghavStaadExtractor.", "").replace(".", "_") + ".cs"
    out_path = os.path.join(OUT_DIR, out_name)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write("\n\n".join(chunks))

    summary.append((full_name, ok_count, fail_count, out_name))
    print(f"{full_name}: {ok_count} ok, {fail_count} failed -> {out_name}")

print("\n=== DONE ===")
total_ok = sum(s[1] for s in summary)
total_fail = sum(s[2] for s in summary)
print(f"{len(summary)} types, {total_ok} members ok, {total_fail} members failed")
