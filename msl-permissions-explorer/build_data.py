"""
Build data.js for the MSL Permissions Explorer from an MSL permissions export.

Usage:
    pip install openpyxl
    python build_data.py "MSL Permissions 1.xlsx"

The export has one sheet per permission (Admin permissions, Edit news, ...),
all with the same 15 columns. Each row says a role (Grouping) has that
permission on a target organisation, either "direct" or "inherited" from a
parent organisation (the "Source target org").

This script:
  1. reads every sheet,
  2. keeps the direct grants (the real decisions),
  3. works out which organisation sits inside which from the inherited rows,
  4. records which direct grants carry down to child organisations,
  5. writes everything to data.js as window.MSL_DATA.
"""
import json
import sys
from collections import defaultdict

import openpyxl

COLS = ["Grouping Id", "Grouping name", "Organisation Id", "Organisation name",
        "Grouping type", "Category name", "Permission name", "Assignment",
        "Target organisation Id", "Target organisation name",
        "Target organisation type", "Target organisation category",
        "Source target org Id", "Source target org name", "Permission Id"]


def load(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    rows, sheets = [], []
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        header = next(it, None)
        if not header or header[0] is None:
            continue  # empty sheet
        col = {name: k for k, name in enumerate(header) if name}
        missing = [c for c in COLS[:14] if c not in col]
        if missing:
            print(f"  skipping '{ws.title}': missing columns {missing}")
            continue
        n = 0
        for r in it:
            if r is None or r[col["Grouping Id"]] is None:
                continue
            rows.append({c: r[col[c]] for c in COLS[:14]} | {"sheet": ws.title})
            n += 1
        if n:
            sheets.append(ws.title)
        print(f"  {ws.title}: {n:,} rows")
    return rows, sheets, len(wb.worksheets)


def build(rows, sheets, sheet_count):
    # ---- organisations ----
    orgs = {}
    for r in rows:
        orgs.setdefault(r["Target organisation Id"], (
            r["Target organisation name"], r["Target organisation type"],
            r["Target organisation category"]))
    for r in rows:  # teams that own roles but are never a target
        orgs.setdefault(r["Organisation Id"], (r["Organisation name"], "Organisation", ""))

    # ---- hierarchy: an inherited row means target sits inside source ----
    cover = defaultdict(set)
    for r in rows:
        if r["Assignment"] == "inherited":
            cover[r["Source target org Id"]].add(r["Target organisation Id"])
    for s in cover:
        cover[s].add(s)
    sources = sorted(cover, key=lambda s: len(cover[s]))

    def parent(o):
        # smallest source that contains o and isn't o itself (or equal to it)
        for s in sources:
            if s != o and o in cover[s] and not (o in cover and cover[o] >= cover[s]):
                return s
        return None

    ids = sorted(orgs)
    idx = {o: i for i, o in enumerate(ids)}
    org_list = [[o, *orgs[o], (idx[parent(o)] if parent(o) is not None else None)] for o in ids]

    # ---- permissions (one per sheet) ----
    perm_idx = {s: k for k, s in enumerate(sheets)}
    perm_list = [[s, family(s)] for s in sheets]

    # ---- roles ----
    roles = {}
    for r in rows:
        roles.setdefault(r["Grouping Id"], (r["Grouping name"], r["Organisation Id"],
                                            r["Grouping type"], r["Category name"]))
    gids = sorted(roles)
    gidx = {g: i for i, g in enumerate(gids)}
    role_list = [[n, idx[team], gt, cat] for n, team, gt, cat in (roles[g] for g in gids)]

    # ---- grants: (role, org) -> permissions given directly, and which carry down ----
    carries = {(r["Grouping Id"], r["sheet"], r["Source target org Id"])
               for r in rows if r["Assignment"] == "inherited"}
    grants = defaultdict(lambda: [set(), set()])
    direct = 0
    for r in rows:
        if r["Assignment"] != "direct":
            continue
        direct += 1
        key = (gidx[r["Grouping Id"]], idx[r["Target organisation Id"]])
        p = perm_idx[r["sheet"]]
        grants[key][0].add(p)
        if (r["Grouping Id"], r["sheet"], r["Target organisation Id"]) in carries:
            grants[key][1].add(p)

    grant_list = [[role, org, enc(a), enc(c)] for (role, org), (a, c) in sorted(grants.items())]
    return dict(orgs=org_list, roles=role_list, perms=perm_list, grants=grant_list,
                total=len(rows), direct=direct, sheets=sheet_count)


CLUB_TOOLS = {"Edit website pages", "Edit news", "Submit news", "Approve news", "Send messages",
              "Edit event", "Edit calendar", "Edit signups", "Edit groups", "Edit members",
              "View members", "Edit organisation details", "Upload resources", "Edit resources",
              "Edit products", "View sales reports", "Configure Membership API",
              "Manage contact requests"}


def family(sheet):
    if sheet.startswith("SGF"):
        return "Finance (SGF)"
    return "Club & society tools" if sheet in CLUB_TOOLS else "SU-wide & staff tools"


def enc(perms):
    # compact: one character per permission index (skips the backslash)
    return "".join(chr(48 + p) if 48 + p < 92 else chr(48 + p + 1) for p in sorted(perms))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit('Usage: python build_data.py "MSL Permissions 1.xlsx"')
    print(f"Reading {sys.argv[1]} ...")
    rows, sheets, n = load(sys.argv[1])
    data = build(rows, sheets, n)
    with open("data.js", "w", encoding="utf-8") as f:
        f.write("window.MSL_DATA = ")
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)
        f.write(";\n")
    print(f"Wrote data.js: {len(data['roles']):,} roles, {len(data['orgs']):,} organisations, "
          f"{len(data['perms'])} permissions, {data['total']:,} rows.")
