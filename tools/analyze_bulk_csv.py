import csv
import sys
from pathlib import Path
from collections import Counter, defaultdict


def year_from_path(p: str) -> str:
    if not p:
        return ""
    parts = Path(p).parts
    # Try to find a 4-digit year component
    for part in parts:
        if part.isdigit() and len(part) == 4:
            return part
    return ""


def ext_from_path(p: str) -> str:
    if not p:
        return ""
    return Path(p).suffix.lower()


def main(csv_path: str) -> int:
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            # Normalize fields to ints where applicable
            r["has_scheda"] = int(r.get("has_scheda", 0) or 0)
            r["has_collaudo"] = int(r.get("has_collaudo", 0) or 0)
            r["scheda_headers_detected"] = int(r.get("scheda_headers_detected", 0) or 0)
            r["scheda_labels_found"] = int(r.get("scheda_labels_found", 0) or 0)
            r["collaudo_media_candidates"] = int(r.get("collaudo_media_candidates", 0) or 0)
            r["collaudo_best_numeric_count"] = int(r.get("collaudo_best_numeric_count", 0) or 0)
            rows.append(r)

    total = len(rows)
    not_loaded = [r for r in rows if not r.get("source_path")]
    loaded = [r for r in rows if r.get("source_path")]
    none_after_load = [r for r in loaded if r["has_scheda"] == 0 and r["has_collaudo"] == 0]
    collaudo_only = [r for r in loaded if r["has_scheda"] == 0 and r["has_collaudo"] == 1]

    print("Summary")
    print(f"- Total rows: {total}")
    print(f"- Not loaded (no workbook): {len(not_loaded)}")
    print(f"- Loaded but None: {len(none_after_load)}")
    print(f"- Collaudo only: {len(collaudo_only)}")

    def group_stats(items, key_fn):
        c = Counter(key_fn(r) for r in items)
        for k, v in c.most_common():
            if k:
                print(f"  {k}: {v}")

    print("\nNot loaded by year:")
    group_stats(not_loaded, lambda r: year_from_path(r.get("source_path", "")))

    print("\nNot loaded examples (up to 10):")
    for r in not_loaded[:10]:
        print(f"  {r['test_number']}")

    print("\nNone-after-load by year:")
    group_stats(none_after_load, lambda r: year_from_path(r.get("source_path", "")))
    print("None-after-load by ext:")
    group_stats(none_after_load, lambda r: ext_from_path(r.get("source_path", "")))
    print("\nNone-after-load examples (up to 10):")
    for r in none_after_load[:10]:
        print(f"  {r['test_number']} -> {r['source_path']}")

    print("\nCollaudo-only by year:")
    group_stats(collaudo_only, lambda r: year_from_path(r.get("source_path", "")))
    print("Collaudo-only by ext:")
    group_stats(collaudo_only, lambda r: ext_from_path(r.get("source_path", "")))

    print("\nCollaudo-only diagnostics (up to 16):")
    for r in collaudo_only[:16]:
        print(
            "  {tn} -> {p} | headers={h} labels={l} scheda_sheet={sf}".format(
                tn=r["test_number"],
                p=r.get("source_path", ""),
                h=r["scheda_headers_detected"],
                l=r["scheda_labels_found"],
                sf=r.get("scheda_sheet_found", ""),
            )
        )

    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python tools/analyze_bulk_csv.py <path-to-csv>")
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
