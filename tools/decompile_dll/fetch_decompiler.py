"""
One-time setup for dump_all.py: downloads ICSharpCode.Decompiler from NuGet
(as a raw .nupkg, extracted directly - no .NET SDK required) into
./decompiler_pkg/, since `dotnet tool install ilspycmd` needs an SDK that may
not be installed.
"""
import os
import urllib.request
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = "7.2.1.6856"  # any recent version works; this one was verified working
URL = f"https://api.nuget.org/v3-flatcontainer/icsharpcode.decompiler/{VERSION}/icsharpcode.decompiler.{VERSION}.nupkg"
NUPKG_PATH = os.path.join(HERE, "icsharpcode.decompiler.nupkg")
OUT_DIR = os.path.join(HERE, "decompiler_pkg")

print(f"Downloading {URL} ...")
urllib.request.urlretrieve(URL, NUPKG_PATH)

print(f"Extracting to {OUT_DIR} ...")
with zipfile.ZipFile(NUPKG_PATH) as zf:
    zf.extractall(OUT_DIR)

dll_path = os.path.join(OUT_DIR, "lib", "netstandard2.0", "ICSharpCode.Decompiler.dll")
assert os.path.exists(dll_path), f"Expected DLL not found at {dll_path}"
print(f"OK: {dll_path}")
