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
    python3 repair_deck_paneref.py --dry-run deck.pptx

A .bak backup of each deck is written next to it before the first change.
"""
import argparse
import re
import shutil
import sys
import zipfile


def repair(path, dry_run=False):
    if not path.lower().endswith(".pptx"):
        return (path, "skip (not a .pptx)")
    try:
        z = zipfile.ZipFile(path)
    except OSError as e:
        return (path, "skip (unreadable: %s)" % e)
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
    bak = path + ".bak"
    shutil.copyfile(path, bak)
    out = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    for name, data in entries:
        out.writestr(name, data)
    out.close()
    return (path, "repaired (backup: %s)" % bak)


def main():
    ap = argparse.ArgumentParser(description="Remove the unresolvable pane-reference stamp from decks")
    ap.add_argument("decks", nargs="+", help="deck.pptx file(s) to repair")
    ap.add_argument("--dry-run", action="store_true", help="report without writing")
    a = ap.parse_args()
    rc = 0
    for deck in a.decks:
        _p, msg = repair(deck, a.dry_run)
        print("%-10s %s" % (msg.split()[0], deck), "-", msg)
        if msg.startswith("skip"):
            rc = 1
    sys.exit(rc)


if __name__ == "__main__":
    main()