"""On-demand FEHD check for the Mac, with popups.

Driven by 'Check FEHD.command'. Wraps update_list.py:
  pull -> check -> update -> commit -> push -> report
Every exit path ends in a dialog, so double-clicking always says something.
"""
import json
import os
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAX_LISTED = 12


def dialog(message, title="Paws & Plates", icon="note"):
    """Show a Mac popup. argv avoids any quote-escaping problems."""
    if os.environ.get("PAWS_NO_DIALOG"):  # set by tests; prints instead
        print(f"[{icon}] {message}")
        return
    subprocess.run([
        "osascript",
        "-e", "on run argv",
        "-e", f'display dialog (item 1 of argv) with title "{title}" '
              f'buttons {{"OK"}} default button "OK" with icon {icon}',
        "-e", "end run",
        "--", message,
    ], capture_output=True)


def die(message):
    dialog(message, icon="stop")
    sys.exit(1)


def git(*args, check=True):
    r = subprocess.run(["git", "-C", REPO, *args], capture_output=True, text=True)
    if check and r.returncode != 0:
        die(f"Git step failed:\n\n{' '.join(args)}\n\n{r.stderr.strip()[:600]}")
    return r


def names(items, arrow):
    lines = []
    for x in items[:MAX_LISTED]:
        en, tc = x["en"], x["tc"]
        label = en if en == tc else f"{en}  /  {tc}"
        lines.append(f"{arrow} {label}")
    if len(items) > MAX_LISTED:
        lines.append(f"   …and {len(items) - MAX_LISTED} more (see summary.txt)")
    return lines


for stale in ("summary.txt", "summary.json"):
    path = os.path.join(REPO, stale)
    if os.path.exists(path):
        os.remove(path)

# Get in step with the scheduled cloud job before touching the sheet.
if git("status", "--porcelain", "--", "seating.xlsx").stdout.strip():
    die("seating.xlsx has changes you haven't committed yet.\n\n"
        "Open GitHub Desktop, commit them (or discard them), then run this again.\n\n"
        "Stopping now so your edits aren't overwritten.")

pull = git("pull", "--ff-only", "origin", "main", check=False)
if pull.returncode != 0:
    die("Could not sync with GitHub before checking.\n\n"
        f"{pull.stderr.strip()[:400]}\n\n"
        "Usually this means you have commits that haven't been pushed. "
        "Open GitHub Desktop and click 'Push origin', then run this again.")

venv_python = sys.executable
run = subprocess.run([venv_python, os.path.join(REPO, "scripts", "update_list.py")],
                     cwd=REPO, capture_output=True, text=True)
output = (run.stdout + run.stderr).strip()

if run.returncode != 0:
    die(f"The FEHD check could not finish.\n\n{output[-600:]}")

def publish(message_file=None, fallback="FEHD checked, no changes"):
    """Commit and push whatever changed. Returns False if the push failed."""
    git("add", "--", "seating.xlsx", "docs/status.json")
    if not git("diff", "--cached", "--quiet", check=False).returncode:
        return True  # nothing staged; already in step with GitHub
    if message_file and os.path.exists(message_file):
        git("commit", "-F", message_file)
    else:
        git("commit", "-m", fallback)
    return git("push", "origin", "main", check=False).returncode == 0


summary_path = os.path.join(REPO, "summary.json")
changed = os.path.exists(summary_path) and \
    bool(git("status", "--porcelain", "--", "seating.xlsx").stdout.strip())

if not changed:
    # Still publish the refreshed timestamp so the phone page stays honest.
    pushed = publish()
    note = "" if pushed else "\n\n(Could not reach GitHub to update the phone page.)"
    dialog("No update required.\n\n" + run.stdout.strip() + note)
    sys.exit(0)

with open(summary_path, encoding="utf-8") as f:
    summary = json.load(f)

if not publish(os.path.join(REPO, "summary.txt")):
    die("The spreadsheet was updated and committed, but the push to GitHub failed.\n\n"
        "Open GitHub Desktop and click 'Push origin' to finish.")

added, removed = summary["added"], summary["removed"]
parts = ["✅ Updated and pushed to GitHub.",
         f"FEHD list dated {summary['fehd_date']}.", ""]
parts.append(f"NEW RESTAURANTS ({len(added)}):")
parts += names(added, "+") if added else ["   none"]
parts.append("")
parts.append(f"REMOVED — moved to the 'Delisted' tab ({len(removed)}):")
parts += names(removed, "−") if removed else ["   none"]
if added:
    parts += ["", "Remember to set indoor/outdoor for the new ones."]

dialog("\n".join(parts))
