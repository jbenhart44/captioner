#!/usr/bin/env python3
"""
Smoke test — deterministic, no vision call.

Runs the full extract -> apply --dry-run chain against the bundled synthetic
deck (examples/sample.pptx, 3 slides / 2 pictures) and asserts:

  1. extract_images.py produces a manifest with n_pictures == 2
  2. apply_captions.py --dry-run (fed the committed examples/sample_captions.json)
     produces an audit CSV with exactly one data row per picture
  3. every action is the dry-run "would-add" form (no .pptx was modified)

The --dry-run path skips the Claude-Code vision step, so this test is fully
deterministic and runnable in CI or on a fresh clone with no API access.

Run:  python3 tests/test_smoke.py     (exit 0 = pass)
"""

import csv
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
SAMPLE = REPO / "examples" / "sample.pptx"
CAPTIONS_FIXTURE = REPO / "examples" / "sample_captions.json"


def main() -> int:
    assert SAMPLE.exists(), f"missing fixture: {SAMPLE}"
    assert CAPTIONS_FIXTURE.exists(), f"missing fixture: {CAPTIONS_FIXTURE}"

    work = Path(tempfile.mkdtemp(prefix="capsmoke_"))
    try:
        # 1. extract
        subprocess.run(
            [sys.executable, str(REPO / "scripts" / "extract_images.py"),
             str(SAMPLE), str(work)],
            check=True, capture_output=True, text=True,
        )
        manifest = json.loads((work / "manifest.json").read_text())
        deck = manifest["sample"]
        assert deck["n_pictures"] == 2, f"expected 2 pictures, got {deck['n_pictures']}"
        print(f"  [1/3] extract OK — {deck['n_pictures']} pictures, {deck['n_slides']} slides")

        # 2. apply --dry-run with the committed captions fixture
        shutil.copy(CAPTIONS_FIXTURE, work / "captions.json")
        subprocess.run(
            [sys.executable, str(REPO / "scripts" / "apply_captions.py"),
             str(work), "--dry-run"],
            check=True, capture_output=True, text=True,
        )
        audit = work / "audit" / "sample_audit.csv"
        assert audit.exists(), f"missing audit CSV: {audit}"
        rows = list(csv.DictReader(audit.read_text().splitlines()))
        assert len(rows) == 2, f"expected 2 audit rows, got {len(rows)}"
        print(f"  [2/3] apply --dry-run OK — {len(rows)} audit rows == 2 pictures")

        # 3. every action is a dry-run would-add (nothing modified)
        for r in rows:
            assert r["action"].startswith("dry-run-would-"), \
                f"unexpected action: {r['action']}"
        print(f"  [3/3] all actions are dry-run (no .pptx modified)")

        print("SMOKE TEST PASSED")
        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
