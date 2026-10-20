#!/usr/bin/env python3
"""Prove the test suites would notice the defects they exist to prevent.

    python3 tests/mutation_check.py      # exit 0 = every mutant caught

A green suite only claims the tests would notice a defect. This makes the claim true: each mutant
injects one real defect into scripts/, re-runs the suite aimed at it, and requires that suite to go
red AND to name the check that caught it. Every mutant must leave working code that behaves wrongly;
a mutant that fails to import tests nothing and is reported BROKEN.

Mutants are written in place (the scripts import their siblings by directory), and each original is
restored in a finally block, then re-hashed, so an interrupted run cannot leave a mutated file behind.
"""
from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
SUITES = {
    "smoke": REPO / "tests" / "test_smoke.py",
    "descr": REPO / "tests" / "test_descr_writer.py",
    "placement": REPO / "tests" / "test_placement.py",
    "rerun": REPO / "tests" / "test_rerun.py",
}
# (why it matters, file, exact source line, the broken version, suite, strings the red run must show)
MUTANTS = [
    ("a caption that overlaps an obstacle in both axes is treated as clear",
     "scripts/_geometry.py", "        if _rects_block(cap_rect, r, vfrac):", "        if False:",
     "placement", ["both-axis overlap blocks", "text-overlap rejected"]),
    ("re-running deletes a caption a person edited by hand",
     "scripts/apply_captions.py", "        if oracle and not force_overwrite_edits:", "        if False:",
     "placement", ["unedited removed, edited preserved"]),
    ("alt-text a person edited is overwritten instead of preserved",
     "scripts/_descr_writer.py", "        return 'descr-preserved-hand-edit', None", "        return 'descr-authored', alt_text",
     "descr", ["test_decide_descr_matrix"]),
    ("alt-text the deck already had is overwritten (preserve-by-default broken)",
     "scripts/_descr_writer.py", "    return 'descr-preserved-pre-existing', None", "    return 'descr-authored', alt_text",
     "descr", ["test_decide_descr_matrix"]),
    ("a re-run on an already-captioned deck duplicates its captions (no refusal)",
     "scripts/apply_captions.py", "        if n_prior:", "        if False:",
     "rerun", ["a second run without --update-existing refuses", "a refused run writes no deck"]),
    ("--update-existing leaves in-picture band captions behind, so they duplicate",
     "scripts/apply_captions.py", "for p in (CAPTION_SHAPE_NAME_PREFIX, CAPTION_BAND_NAME_PREFIX)",
     "for p in (CAPTION_SHAPE_NAME_PREFIX,)",
     "rerun", ["--update-existing strips band captions"]),
    ("a deck holding only band captions is not seen as already captioned",
     "scripts/apply_captions.py", "CAPTIONER_SHAPE_PREFIXES = (CAPTION_SHAPE_NAME_PREFIX, CAPTION_BAND_NAME_PREFIX,",
     "CAPTIONER_SHAPE_PREFIXES = (CAPTION_SHAPE_NAME_PREFIX,",
     "rerun", ["a band caption counts as a captioner shape"]),
]


def run(suite: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SUITES[suite])], capture_output=True, text=True, cwd=REPO)


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main() -> int:
    for name, path in SUITES.items():
        base = run(name)
        if base.returncode != 0:
            print(f"SETUP ERROR: the unmutated {name} suite is red:\n{base.stdout[-800:]}{base.stderr[-800:]}")
            return 2
    holes = []
    for why, rel, old, new, suite, must in MUTANTS:
        target = REPO / rel
        src = target.read_text(encoding="utf-8")
        if src.count(old) != 1:
            print(f"STALE   {why}: anchor found {src.count(old)} times in {rel} (the code moved; re-point it)")
            holes.append(why)
            continue
        before = sha(target)
        try:
            target.write_text(src.replace(old, new), encoding="utf-8")
            p = run(suite)
        finally:
            target.write_text(src, encoding="utf-8")
        if sha(target) != before:
            print(f"SETUP ERROR: {rel} was not restored byte-for-byte")
            return 2
        out = p.stdout + p.stderr
        if "ModuleNotFoundError" in out or "SyntaxError" in out or "ImportError" in out:
            print(f"BROKEN  {why}: the mutant did not import (a broken mutant tests nothing)")
            holes.append(why)
            continue
        # A name counts only where it is reported as failing: a FAIL line, or a traceback line
        # ("File ..., line N, in <test>"). Check names also appear on PASS lines, so a plain
        # substring test would call a mutant caught on the exit code alone.
        missing = [m for m in must
                   if not any(m in ln and ("FAIL" in ln or ", line " in ln) for ln in out.splitlines())]
        caught = p.returncode != 0 and not missing
        print(f"{'CAUGHT' if caught else 'MISSED'}  {why}" + (f"  (red run did not name: {missing})" if missing else ""))
        if not caught:
            holes.append(why)
    print(f"\n{len(MUTANTS) - len(holes)}/{len(MUTANTS)} mutants caught" + (" (PROVEN)" if not holes else f"; HOLES: {holes}"))
    return 1 if holes else 0


if __name__ == "__main__":
    sys.exit(main())
