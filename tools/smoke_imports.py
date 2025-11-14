import importlib
import sys
import traceback

mods = [
    "src.reports.excel_report",
    "src.analysis.noise_chart_generator",
    "src.reports.builders.excel_sheet_helper",
    "src.reports.builders.sap_sheet_builder",
]

ok = True
for m in mods:
    try:
        importlib.import_module(m)
        print(m + " OK")
    except Exception as e:
        print("ERROR importing {}: {}".format(m, e))
        traceback.print_exc()
        ok = False

sys.exit(0 if ok else 1)
