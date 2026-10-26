#!/usr/bin/env python3
"""Diff two template versions (contract.json or .pptx/.potx on either side).

    python3 diff.py OLD NEW [--json] [--tolerance-pct 0.25]

breaking  = a build spec written against OLD may fail or misplace content
            (layout/placeholder/pattern/slot removed or renamed, placeholder moved
            or resized beyond tolerance, slide size changed)
changed   = still builds, but output looks different (theme colours/fonts, chrome)
additive  = new layouts, placeholders or patterns
Exit codes: 0 no breaking changes, 1 breaking changes, 2 usage/read error.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_core as core


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    ap.add_argument(
        "--tolerance-pct", type=float, default=0.25, help="geometry tolerance, %% of slide width"
    )
    a = ap.parse_args(argv)
    try:
        old, new = core.contract_or_profile(a.old), core.contract_or_profile(a.new)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    d = core.diff_contracts(old, new, a.tolerance_pct / 100)
    same_file = old["source"]["sha256"] == new["source"]["sha256"]
    if a.json:
        print(json.dumps({"identical_file": same_file, **d}, indent=1))
    else:
        print(
            f"{old['source']['file']} -> {new['source']['file']}"
            + (" (byte-identical)" if same_file else "")
        )
        for k in ("breaking", "changed", "additive"):
            print(f"{k.upper()} ({len(d[k])})")
            for line in d[k]:
                print(f"  - {line}")
    return 1 if d["breaking"] else 0


if __name__ == "__main__":
    sys.exit(main())
