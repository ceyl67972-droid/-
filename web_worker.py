from __future__ import annotations

import json
import sys
from pathlib import Path

from reconcile_bank_statement import reconcile


def main() -> int:
    manifest_path = Path(sys.argv[1])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    output_path = Path(manifest["output_path"])
    stats = reconcile(
        [Path(item) for item in manifest["pdf_paths"]],
        Path(manifest["excel_path"]),
        output_path,
        None,
        int(manifest["date_tolerance"]),
    )
    payload = {
        "result_path": str(output_path),
        "stats": {
            "statement_count": stats[0],
            "exact_count": stats[1],
            "tolerance_count": stats[2],
            "unmatched_count": stats[3],
            "unrecognized_pdf_count": stats[4],
            "unknown_bank_count": stats[5],
        },
    }
    print("__RESULT__" + json.dumps(payload, ensure_ascii=True), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
