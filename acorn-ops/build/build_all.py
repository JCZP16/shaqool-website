"""
Rebuilds everything, then verifies it.

    python3 build_all.py

Order matters: the workbooks come from schema.py, the docs come from schema.py,
and the verification checks that the hand-written VBA still agrees with schema.py.
A non-zero exit means something is inconsistent - do not ship it.
"""

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent

STEPS = [
    ("Data workbooks", "build_data.py"),
    ("Console", "build_console.py"),
    ("Word templates", "build_templates.py"),
    ("Documentation", "build_docs.py"),
]


def run(script, *args):
    result = subprocess.run([sys.executable, str(HERE / script), *args],
                            capture_output=True, text=True)
    sys.stdout.write(result.stdout)
    if result.returncode != 0:
        sys.stderr.write(result.stderr)
    return result.returncode


def main():
    print("Acorn Recyclers operations platform - full build")
    print("=" * 55)

    for title, script in STEPS:
        print(f"\n{title}")
        if run(script) != 0:
            print(f"\n{title} failed. Stopping.")
            return 1

    print("\nVBA static checks")
    lint = run("lint_vba.py")

    print("\nWorkbook and consistency checks")
    verify = run("verify.py")

    print()
    if lint == 0 and verify == 0:
        print("Build complete and verified.")
        return 0
    print("Build produced output, but the checks failed. Fix the problems above before shipping.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
