"""
Static checks on the VBA modules.

There is no VBA compiler in this environment, so this catches the classes of
mistake that would otherwise only surface as a compile error on the user's PC:

  * unbalanced Sub/Function/If/With/For/Do blocks
  * a line continuation with nothing after it, or a stray one at end of file
  * calls to modXxx.Something where Something is not Public in that module
  * button actions on the Start sheet pointing at routines that do not exist
  * duplicate procedure names across modules (VBA resolves these silently and
    unpredictably)
  * ReDim Preserve growing anything other than the last dimension
  * missing Option Explicit or Attribute VB_Name
  * lines over VBA's 1023-character limit

    python3 lint_vba.py [vba_dir]
"""

import re
import sys
from pathlib import Path

PROC_RE = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+(\w+)",
                     re.IGNORECASE)
END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
QUALIFIED_CALL_RE = re.compile(r"\b(mod[A-Z]\w*)\.(\w+)")
DECL_RE = re.compile(r"^\s*(?:Public|Private|Global)?\s*(?:Const|Declare|Enum|Type)\b", re.IGNORECASE)


def strip_code(line: str) -> str:
    """Remove string literals and comments so keywords inside them are ignored."""
    out, in_str = [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
            i += 1
            continue
        if not in_str and ch == "'":
            break
        if not in_str and line[i:i + 4].upper() == "REM " and (i == 0 or line[i - 1] in " \t:"):
            break
        if not in_str:
            out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(raw_lines):
    """Join VBA line continuations into single logical lines, keeping the
    physical line number of where each one started."""
    joined, buf, start = [], "", None
    for n, raw in enumerate(raw_lines, 1):
        line = raw.rstrip("\n")
        if start is None:
            start = n
        stripped = line.rstrip()
        if stripped.endswith(" _") or stripped.endswith("\t_"):
            buf += stripped[:-1]
        else:
            joined.append((start, buf + line))
            buf, start = "", None
    if buf:
        joined.append((start, buf))
    return joined


BLOCK_OPENERS = [
    (re.compile(r"^\s*(?:Public |Private |Friend |Static )*(?:Sub|Function)\s+\w+", re.I), "proc"),
    (re.compile(r"\bWith\b", re.I), "with"),
    (re.compile(r"^\s*(?:Do)\s*(?:While|Until)?\b", re.I), "do"),
    (re.compile(r"^\s*Select\s+Case\b", re.I), "select"),
]


def check_blocks(path, lines):
    """Balance of the block constructs, tracked as a stack so the error message
    can name what was left open."""
    problems = []
    stack = []

    for lineno, logical in lines:
        code = strip_code(logical).strip()
        if not code:
            continue
        low = code.lower()

        if re.match(r"^(public |private |friend |static )*(sub|function)\s+\w+", low):
            if any(k == "proc" for k, _ in stack):
                problems.append(f"{path.name}:{lineno}: a procedure starts before the previous one ends")
            stack.append(("proc", lineno))
            continue
        if re.match(r"^end\s+(sub|function)\b", low):
            if not stack or stack[-1][0] != "proc":
                problems.append(f"{path.name}:{lineno}: End Sub/Function with no matching opener")
            else:
                stack.pop()
            continue

        # If ... Then with nothing after it opens a block; a one-liner does not.
        m = re.match(r"^if\b.*\bthen\b(.*)$", low)
        if m and not m.group(1).strip():
            stack.append(("if", lineno))
            continue
        if re.match(r"^end\s+if\b", low):
            if not stack or stack[-1][0] != "if":
                problems.append(f"{path.name}:{lineno}: End If with no matching If")
            else:
                stack.pop()
            continue

        if re.match(r"^with\b", low):
            stack.append(("with", lineno))
            continue
        if re.match(r"^end\s+with\b", low):
            if not stack or stack[-1][0] != "with":
                problems.append(f"{path.name}:{lineno}: End With with no matching With")
            else:
                stack.pop()
            continue

        if re.match(r"^select\s+case\b", low):
            stack.append(("select", lineno))
            continue
        if re.match(r"^end\s+select\b", low):
            if not stack or stack[-1][0] != "select":
                problems.append(f"{path.name}:{lineno}: End Select with no matching Select Case")
            else:
                stack.pop()
            continue

        if re.match(r"^for\b", low):
            stack.append(("for", lineno))
            continue
        if re.match(r"^next\b", low):
            if not stack or stack[-1][0] != "for":
                problems.append(f"{path.name}:{lineno}: Next with no matching For")
            else:
                stack.pop()
            continue

        if re.match(r"^do\b", low):
            stack.append(("do", lineno))
            continue
        if re.match(r"^loop\b", low):
            if not stack or stack[-1][0] != "do":
                problems.append(f"{path.name}:{lineno}: Loop with no matching Do")
            else:
                stack.pop()
            continue

    for kind, lineno in stack:
        problems.append(f"{path.name}:{lineno}: {kind} block is never closed")
    return problems


def collect_procs(path, lines):
    """{name: 'Public'|'Private'} for every procedure in a module."""
    procs = {}
    for lineno, logical in lines:
        code = strip_code(logical)
        m = PROC_RE.match(code)
        if m and not DECL_RE.match(code):
            scope = (m.group(1) or "Public").title()
            procs[m.group(3)] = scope
    return procs


def main():
    vba_dir = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent.parent / "src" / "vba")
    files = sorted(vba_dir.glob("*.bas"))
    if not files:
        print(f"no .bas files in {vba_dir}")
        return 1

    all_procs = {}       # module -> {proc: scope}
    problems = []
    parsed = {}

    for path in files:
        raw = path.read_text(encoding="utf-8").splitlines(keepends=True)
        lines = logical_lines(raw)
        parsed[path] = lines

        text = "".join(raw)
        if not text.startswith("Attribute VB_Name"):
            problems.append(f"{path.name}: missing the Attribute VB_Name header, so it will not import")
        if not re.search(r"^Option Explicit\s*$", text, re.MULTILINE):
            problems.append(f"{path.name}: no Option Explicit")

        for n, raw_line in enumerate(raw, 1):
            if len(raw_line.rstrip("\n")) > 1023:
                problems.append(f"{path.name}:{n}: line is over VBA's 1023-character limit")
            stripped = raw_line.rstrip("\n").rstrip()
            if stripped.endswith(" _") and n == len(raw):
                problems.append(f"{path.name}:{n}: file ends on a line continuation")

        for lineno, logical in lines:
            code = strip_code(logical)
            m = re.search(r"ReDim\s+Preserve\s+(\w+)\s*\((.*)\)\s*$", code, re.IGNORECASE)
            if m and m.group(2).count(",") >= 1:
                dims = [d.strip() for d in m.group(2).split(",")]
                if any("To" in d and not d.startswith("1 To " + dims[-1].split()[-1]) for d in dims[:-1]):
                    problems.append(
                        f"{path.name}:{lineno}: ReDim Preserve can only resize the last dimension")

        all_procs[path.stem] = collect_procs(path, lines)
        problems += check_blocks(path, lines)

    # Duplicate public names across modules: VBA picks one silently.
    seen = {}
    for module, procs in all_procs.items():
        for name, scope in procs.items():
            if scope != "Public":
                continue
            if name in seen:
                problems.append(
                    f"{module}.{name} duplicates {seen[name]}.{name}; VBA will resolve the "
                    f"unqualified name to whichever it feels like")
            else:
                seen[name] = module

    # Button OnAction targets are string literals, so strip_code hides them from
    # the call check above - and a button wired to a routine that no longer
    # exists fails only when somebody presses it.
    # Only the button wiring, not the "modX.Proc" strings that Err.Raise uses to
    # name where an error came from.
    literal_re = re.compile(r'"(mod[A-Z]\w*\.\w+)"')
    button_line_re = re.compile(r'OnAction|Array\(\s*"[^"]*"\s*,\s*"mod[A-Z]')
    for path in files:
        for n, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if not button_line_re.search(raw_line):
                continue
            for target in literal_re.findall(raw_line):
                module, name = target.split(".")
                if module not in all_procs:
                    problems.append(f"{path.name}:{n}: button points at unknown module {module}")
                elif name not in all_procs[module]:
                    problems.append(f"{path.name}:{n}: button points at missing {target}")
                elif all_procs[module][name] != "Public":
                    problems.append(f"{path.name}:{n}: button points at Private {target}")

    # Qualified cross-module calls must resolve to a Public procedure.
    for path, lines in parsed.items():
        for lineno, logical in lines:
            code = strip_code(logical)
            for module, name in QUALIFIED_CALL_RE.findall(code):
                if module not in all_procs:
                    problems.append(f"{path.name}:{lineno}: call into unknown module {module}")
                elif name not in all_procs[module]:
                    problems.append(f"{path.name}:{lineno}: {module}.{name} does not exist")
                elif all_procs[module][name] != "Public":
                    problems.append(f"{path.name}:{lineno}: {module}.{name} is Private")

    total_procs = sum(len(p) for p in all_procs.values())
    for path in files:
        print(f"{path.name:<24} {len(parsed[path]):>4} logical lines  "
              f"{len(all_procs[path.stem]):>2} procedures")
    print()
    if problems:
        for p in problems:
            print(f"  - {p}")
        print(f"\n{len(problems)} problem(s) across {len(files)} modules, {total_procs} procedures.")
        return 1
    print(f"All checks passed: {len(files)} modules, {total_procs} procedures.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
