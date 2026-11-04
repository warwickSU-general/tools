# MSL Permissions Explorer

A two-tab page for looking up MSL permissions:

- **Roles**: pick a role and see every permission it has, and where.
- **Organisations**: pick an organisation and filter which roles can do what there.

## Files

| File | What it is |
|---|---|
| `index.html` | The page. Open it in a browser. |
| `data.js` | The permissions data, built from the MSL export. |
| `build_data.py` | Rebuilds `data.js` from a new MSL export. |

## Run it

**Quickest:** double-click `index.html`. It opens in your browser and works offline. The font loads from Google Fonts when you're online and falls back to a system font when you're not.

**In VS Code:**
1. Open this folder (File → Open Folder).
2. Install the **Live Server** extension (by Ritwick Dey).
3. Right-click `index.html` → **Open with Live Server**.

Live Server reloads the page whenever you save a change.

## Update the data from a new export

You need Python 3 installed.

```bash
pip install openpyxl
python build_data.py "MSL Permissions 1.xlsx"
```

This overwrites `data.js`. Refresh the page to see the new data.

It takes about a minute for a full export (~370,000 rows across 44 sheets).
