"""Quick dependency check for project runtime environment."""

from importlib import import_module
import sys

MODULES = ["pandas", "openpyxl", "tkinter"]

print(f"Python executable: {sys.executable}")
print(f"Python version: {sys.version.split()[0]}")
print()

failed = []
for module_name in MODULES:
    try:
        import_module(module_name)
        print(f"[OK] {module_name}")
    except Exception as exc:
        print(f"[FAIL] {module_name}: {exc}")
        failed.append(module_name)

print()
if failed:
    print("Missing/failed modules:", ", ".join(failed))
    print("Install with: python -m pip install pandas openpyxl")
    raise SystemExit(1)

print("All required modules imported successfully.")
