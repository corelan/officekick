#!/usr/bin/env python3
"""Strip the broken pane reference stamp from Office decks (macOS 16.114).

PowerPoint stamps a deck's reopened task-pane reference with
store="developer" storeType="Registry"; the Mac host's launch-hydration
cannot resolve a wef-folder-sideloaded add-in from that store, so opening
the deck throws the "Add-in Error: New Office Add-in" toast. Removing the
two attributes makes the host resolve purely by add-in id (the path the
manual Ribbon open uses and that always works). Verified end-to-end (2026-11-02, PowerPoint 16.114.1027.3): after the
repair, decks open via Finder double-click WITH the pane auto-loading and
no toast, and pane-open saves keep working (the repaired reference
survives re-saves). Note: opening via PowerPoint's in-app open dialog
skips pane restore silently (no error) - Finder double-click is the
verified-good launch path.

Usage:
    python3 repair_deck_paneref.py deck.pptx [more.pptx ...]
    python3 repair_deck_paneref.py /folder/with/decks        (non-recursive)
    python3 repair_deck_paneref.py -r /folder/with/decks     (subfolders too)
    python3 repair_deck_paneref.py file1.pptx /folder mixed  (any mix)
    python3 repair_deck_paneref.py --dry-run deck.pptx /folder

Before the first change the original is copied to an ``orig/`` subfolder next
to the deck(s) as ``<origfilename>.pptx.<YYYYMMDD_HHMMSS>.bak``; the deck is
then patched in place, keeping the same filename. Decks are always patched
atomically (write-to-temp + os.replace).
"""
import argparse
import glob
import os
import re
import shutil
import sys
import time
import zipfile


def repair(path, dry_run=False):
    if not path.lower().endswith(".pptx"):
        return (path, "skip (not a .pptx)")
    try:
        z = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        return (path, "skip (unreadable: %s)" % str(e).replace(chr(10), " ")[:80])
    changed = False
    entries = []
    for it in z.infolist():
        data = z.read(it.filename)
        if "webextension" in it.filename.lower() and it.filename.endswith(".xml"):
            x = data.decode("utf-8", "replace")
            x2 = x.replace(' store="developer" storeType="Registry"', "")
            if x2 != x:
                data = x2.encode("utf-8")
                changed = True
        entries.append((it.filename, data))
    z.close()
    if not changed:
        return (path, "ok (no broken pane reference)")
    if dry_run:
        return (path, "would repair")
    folder = os.path.dirname(os.path.abspath(path))
    base = os.path.basename(path)
    orig_dir = os.path.join(folder, "orig")
    os.makedirs(orig_dir, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    bak = os.path.join(orig_dir, base + "." + ts + ".bak")
    shutil.copyfile(path, bak)
    tmp = path + ".tmp_repair"
    out = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for name, data in entries:
        out.writestr(name, data)
    out.close()
    os.replace(tmp, path)  # patch in place, same filename, atomic
    return (path, "repaired (original kept: %s)" % bak)


def collect(args):
    """Returns (decks, notes): decks = real .pptx to process (-r includes subfolders),
    notes = lines printed for anything in the scan that is not processed."""
    decks, notes = [], []
    seen = set()
    for item in args.files:
        if os.path.isdir(item):
            if args.recursive:
                for root, _dirs, files in os.walk(item):
                    for f in sorted(files):
                        p = os.path.join(root, f)
                        classify(p, f, decks, notes, seen)
            else:
                for f in sorted(os.listdir(item)):
                    classify(os.path.join(item, f), f, decks, notes, seen)
        elif os.path.basename(item).startswith("~$"):
            notes.append(("locked", item))
        elif item.lower().endswith(".pptx"):
            classify(item, os.path.basename(item), decks, notes, seen)
        else:
            notes.append(("skip", item))
    return decks, notes


def classify(path, name, decks, notes, seen):
    if path in seen:
        return
    seen.add(path)
    if name.startswith("~$"):
        notes.append(("locked", path))
    elif name.lower().endswith(".pptx"):
        decks.append(path)
    elif name.lower().endswith(".bak"):
        notes.append(("bak", path))
    else:
        notes.append(("not-deck", path))
def main():
    ap = argparse.ArgumentParser(description="Remove the unresolvable pane-reference stamp from decks")
    ap.add_argument("files", nargs="*", help="deck.pptx file(s) and/or folder(s)")
    ap.add_argument("-r", "--recursive", action="store_true",
                    help="when a folder is given: include .pptx in subfolders too")
    ap.add_argument("--dry-run", action="store_true", help="report without writing")
    a = ap.parse_args()
    if not a.files:
        ap.error("give at least one .pptx file or folder (-- recursive with -r)")
    rc = 0
    decks, notes = collect(a)
    for kind, path in notes:
        print("%-10s %s" % (kind, path))
    if not decks:
        print("no .pptx found in the given path(s)")
        sys.exit(1)
    for deck in decks:
        _p, msg = repair(deck, a.dry_run)
        print("%-10s %s" % (msg.split()[0], deck), "-", msg)
        if msg.startswith("skip"):
            rc = 1
    print("%d deck(s) processed" % len(decks))
    sys.exit(rc)


if __name__ == "__main__":
    main()