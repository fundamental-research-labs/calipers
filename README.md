# Calipers

CLI for **Excel-golden generation** and **Excel-vs-engine verification**.

Open an `.xlsx` in desktop Microsoft Excel, save it again, and record sidecar metadata. The Excel rewrite is the golden: later engine output is compared to that golden, not to the original file.

v1 is **Windows-only**. Goldens are produced via Excel COM (`Excel.Application`). Do **not** generate goldens on macOS — Excel for Mac serializes OOXML differently.

This repository is the verification tool and its init corpus. It is not a spreadsheet engine.

## Requirements

- Go 1.22+
- For `excel-save` and `excel-run`: Windows + Microsoft Excel (desktop)

Off Windows, those commands exit with:

```
calipers: excel-save requires Windows + Excel (COM)
```

`version` and `-h` / `--help` work on any OS.

## Build

From the repo root:

```bash
go build -o calipers ./cmd/calipers
go test ./...
```

## Commands

```bash
# Open input in Excel, Save As xlsx (load+save, no Office.js)
calipers excel-save <input.xlsx> <output.xlsx>

# Open input, run Office.js inside Excel (sideloaded add-in), Save As
calipers excel-run <input.xlsx> <script.js> <output.xlsx>

# Tool version (on Windows, also prints Excel version when COM works)
calipers version

# Usage
calipers -h
```

Examples:

```bash
calipers excel-save verification/cases/tier_a_simple/init.xlsx verification/cases/tier_a_simple/golden.xlsx
calipers excel-run verification/cases/tier_a_simple_set_a1/init.xlsx \
  verification/cases/tier_a_simple_set_a1/script.js \
  verification/cases/tier_a_simple_set_a1/golden.xlsx
```

A successful run writes `golden.xlsx.meta.json` beside the xlsx (`host`, Excel version/build when available, OS, tool name, input basename, optional `script` identity, UTC `generatedAt`). Load+save goldens omit `script`. `excel-run` records the script basename.

CI never runs Excel. A Windows machine with Excel generates goldens; those files are committed and compared later.

## Cases

Each case lives in [`verification/cases/<id>/`](verification/cases/) as `tier_{a|b|c}_<feature>/` with required `init.xlsx`, optional `script.js`, and a dedicated `golden.xlsx` (not mixed into the inits). Missing or empty `script.js` means load+save only — skip script execution. **92 cases** (`tier_a_` 58, `tier_b_` 26, `tier_c_` 8). First Office.js case: [`tier_a_simple_set_a1`](verification/cases/tier_a_simple_set_a1/) (sets A1; `tier_a_simple` stays load+save). See [`verification/README.md`](verification/README.md) for tiers, sources, and compare rules.

| Prefix | Role |
|--------|------|
| `tier_a_` | Engine claims support (values, formulas, styles, sheets, merges, names, freeze, unicode, themes, number formats). First goldens. |
| `tier_b_` | Remaining work (charts, pivot, tables/autofilter, CF, validation, comments, drawings, hyperlinks, protection, print). Goldens still useful: Excel keeps these, an engine may drop them. |
| `tier_c_` | Later / hostile (strict OOXML, password, XML bomb, huge stress, ATP, OLE embed). **Do not** run in the default golden pass. |

A default golden pass is `tier_a` and `tier_b`; it skips `tier_c` hostiles (`tier_c_xmlbomb`, `tier_c_password`, …).

## Host

`excel-save` drives Excel through COM:

- `Excel.Application` on an STA thread
- `DisplayAlerts` / `Visible` off
- timeout plus started-PID cleanup
- `Workbooks.Open` then `SaveAs` xlsx format 51

`excel-run` uses the same COM lifecycle, then a **sideloaded Office.js add-in** (not AppSource, not Office Scripts):

- Local HTTP server serves `internal/excel/web/` (task pane + Office.js from CDN)
- GET `/job` hands the corpus script to the add-in; the add-in `Excel.run`s it; POST `/done` signals completion
- WEF trusted catalog + workbook web-extension stamp so the task pane can auto-open
- Excel is shown (`Visible=true`) so WebView2 can run

First-time Windows box: if the task pane does not appear, Insert → Add-ins → Shared Folder → **Calipers Office.js runner** (catalog is registered under `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`). After that, `excel-run` should auto-open it.

Office.js cannot be eval’d through COM. Corpus scripts stay `Excel.run` (not Automate / Office Scripts).

Goldens must be generated on **Windows Excel**. macOS Excel is not used.
