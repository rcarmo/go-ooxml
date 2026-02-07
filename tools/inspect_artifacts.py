#!/usr/bin/env python3
import argparse
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run(cmd):
    print(f"$ {' '.join(cmd)}")
    result = subprocess.run(cmd, text=True)
    return result.returncode


def find_artifacts(artifacts_dir):
    artifacts = {}
    for path in Path(artifacts_dir).glob("*"):
        if path.suffix.lower() in (".pptx", ".xlsx", ".docx"):
            artifacts[path.name] = path
    return artifacts


def main(argv):
    parser = argparse.ArgumentParser(description="Inspect generated artifacts and compare with repaired files.")
    parser.add_argument("--artifacts", default=str(ROOT / "artifacts"), help="Artifacts directory")
    parser.add_argument("--repaired-suffix", default=" - Repaired", help="Suffix before extension for repaired files")
    parser.add_argument("--diff", action="store_true", help="Show diffs for changed XML files")
    parser.add_argument("--diff-limit", type=int, default=200, help="Max diff lines per file")
    args = parser.parse_args(argv)

    artifacts = find_artifacts(args.artifacts)
    if not artifacts:
        print(f"No artifacts found in {args.artifacts}")
        return 1

    status = 0
    for name, path in sorted(artifacts.items()):
        stem = path.stem
        repaired = Path(args.artifacts) / f"{stem}{args.repaired_suffix}{path.suffix}"
        if path.suffix.lower() == ".pptx":
            cmd = [sys.executable, str(ROOT / "tools" / "inspect_pptx.py"), str(path)]
            if repaired.exists():
                cmd += ["--compare", str(repaired)]
        elif path.suffix.lower() == ".xlsx":
            cmd = [sys.executable, str(ROOT / "tools" / "inspect_xlsx.py"), str(path)]
            if repaired.exists():
                cmd += ["--compare", str(repaired)]
        else:
            print(f"Skipping {path} (no inspector for {path.suffix})")
            continue
        if args.diff:
            cmd.append("--diff")
            cmd.append(f"--diff-limit={args.diff_limit}")
        rc = run(cmd)
        status = status or rc
    return status


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
