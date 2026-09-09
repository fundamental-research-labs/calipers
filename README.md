# Calipers

CLI for **Excel-golden generation** and **Excel-vs-engine verification**.

Open an `.xlsx` in desktop Microsoft Excel, save it again, and record sidecar metadata. The Excel rewrite is the golden: later engine output is compared to that golden, not to the original file.

v1 is **Windows-only**. Goldens are produced via Excel COM (`Excel.Application`). Do **not** generate goldens on macOS — Excel for Mac serializes OOXML differently.

This repository is the verification tool and its init corpus. It is not a spreadsheet engine.

`verify` currently compares ZIP parts after selected metadata exclusions and XML line-ending normalization. PASS means a package match under those rules; FAIL counts differing parts, not cells or proven defects. Text-node whitespace is preserved. A resolved workbook semantic comparator is not yet implemented; see [comparison scope](verification/README.md#what-to-compare).

## Requirements

- Go 1.22+
- For `excel-save` and `excel-run`: Windows + Microsoft Excel (desktop)

Off Windows, those commands exit with:

```
calipers: excel-save requires Windows + Excel (COM)
```

`version`, `verify`, and `-h` / `--help` work on any OS. `verify --engine excel` needs Windows + Excel. `verify --engine <bin>` execs a caller-supplied binary (`save` / `run`).

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

# Default-pass goldens (tier_a and tier_b; skip Office.js and tier_c)
calipers excel-save-pass [cases-dir]

# Open input, run Office.js inside Excel (sideloaded add-in), Save As
calipers excel-run <input.xlsx> <script.js> <output.xlsx>

# Engine vs committed Excel goldens (binary host or Excel)
calipers verify --engine /path/to/engine
calipers verify --engine /path/to/engine --suite roundtrip
calipers verify --engine /path/to/engine --suite scratch
calipers verify --engine /path/to/engine --case roundtrip/tier_a_simple
calipers verify --engine excel --case roundtrip/tier_a_simple

# Opt in to recalculation with an external engine supporting --recalculate
calipers verify --engine /path/to/mog --recalculate --case roundtrip/tier_a_formulas

# Tool version (on Windows, also prints Excel version when COM works)
calipers version

# Usage
calipers -h
```

Examples:

```bash
calipers excel-save verification/cases/roundtrip/tier_a_simple/init.xlsx verification/cases/roundtrip/tier_a_simple/golden.xlsx
calipers excel-save-pass
calipers excel-run verification/cases/default/tier_a_simple_set_a1/init.xlsx \
  verification/cases/default/tier_a_simple_set_a1/script.js \
  verification/cases/default/tier_a_simple_set_a1/golden.xlsx

# External engine binary (argv: save <in> <out> / run <in> <script.js> <out>)
calipers verify --engine ./vendor/mog/target-native/debug/mog
calipers verify --engine ./vendor/mog/target-native/debug/mog --suite roundtrip
calipers verify --engine ./vendor/mog/target-native/debug/mog --suite scratch
calipers verify --engine ./vendor/mog/target-native/debug/mog --case roundtrip/tier_a_simple
```

A successful run writes `golden.xlsx.meta.json` beside the xlsx (`host`, Excel version/build when available, OS, tool name, input basename, optional `script` identity, UTC `generatedAt`). Load+save goldens omit `script`. `excel-run` records the script basename.

CI never runs Excel. A Windows machine with Excel generates goldens; those files are committed and compared later.

`verify` prints the selected calculation policy. By default it leaves each host's policy unchanged; Mog preserves imported caches. `--recalculate` passes `save --recalculate <in> <out>` or `run --recalculate <in> <script> <out>` to an external engine that supports this contract. Mog recalculates before export, after the script for `run`. The flag is rejected for the Excel host, whose calculation is not explicitly controlled. Recalculation does not make random, time-dependent, or environment-dependent results and their dependents equal previously captured goldens; those require controlled assertions.

## Cases

Each case lives in [`verification/cases/<suite>/<id>/`](verification/cases/). Roundtrip and default cases are `tier_{a|b|c}_<feature>/`; scratch cases are unprefixed feature names. Required `init.xlsx`, optional `script.js`, and a dedicated `golden.xlsx` (not mixed into the inits). Missing or empty `script.js` means load+save only — skip script execution. Suites: [`roundtrip/`](verification/cases/roundtrip/) (load+save package comparison), [`default/`](verification/cases/default/) (untriaged until categorized), and [`scratch/`](verification/cases/scratch/) (Office.js from an empty init; committed `excel-run` goldens). **105 cases** (`tier_a_` 57, `tier_b_` 25, `tier_c_` 8, plus 15 unprefixed scratch; 81 roundtrip, 9 default, 15 scratch). Office.js goldens: [`tier_a_simple_set_a1`](verification/cases/default/tier_a_simple_set_a1/) (sets A1; `tier_a_simple` stays load+save in roundtrip) and the 15 `scratch/` cases. See [`verification/README.md`](verification/README.md) for tiers, sources, and compare rules.

`calipers verify` walks committed-golden cases (roundtrip, default, and scratch). `--suite scratch` runs only the empty-init Office.js cases; `--suite roundtrip` runs one suite; `--case roundtrip/tier_a_simple` runs one test. Case ids are `suite/name`.

| Prefix | Role |
|--------|------|
| `tier_a_` | Engine claims support (values, formulas, styles, sheets, merges, names, freeze, unicode, themes, number formats). First goldens. |
| `tier_b_` | Remaining work (charts, pivot, tables/autofilter, CF, validation, comments, drawings, hyperlinks, protection, print). Goldens still useful: Excel keeps these, an engine may drop them. |
| `tier_c_` | Later / hostile (strict OOXML, password, XML bomb, huge stress, ATP, OLE embed). **Do not** run in the default golden pass. |

A default golden pass (`excel-save-pass`) is `tier_a` and `tier_b`; it skips `tier_c` hostiles (`tier_c_xmlbomb`, `tier_c_password`, …) and Office.js cases (`tier_a_simple_set_a1` and `scratch/`). Those Office.js cases have committed `excel-run` goldens (`script=script.js`).

## Host

`excel-save` drives Excel through COM:

- `Excel.Application` on an STA thread
- `DisplayAlerts` / `Visible` off
- timeout plus started-PID cleanup
- `Workbooks.Open` then `SaveAs` xlsx format 51

`excel-run` starts **excel.exe** (COM `CreateObject` does not load Office.js / WebView2), then a **sideloaded Office.js add-in** (not AppSource, not Office Scripts):

- Local HTTP server on `http://localhost:<port>` serves `internal/excel/web/` (task pane + Office.js from CDN)
- GET `/job` hands the corpus script to the add-in; the add-in `Excel.run`s it; POST `/done` signals completion
- WEF **Developer** sideload (`HKCU\...\WEF\Developer\<add-in GUID>` = path to the manifest) so the workbook web-extension stamp (`store=developer`, `storeType=Registry`) can auto-open the task pane
- WEF **TrustedCatalogs** `{GUID}` key with `Id` and a UNC `Url` (`\\localhost\<drive>$\...`) so Insert → Add-ins → Shared Folder can list it
- Catalog lives under `%LOCALAPPDATA%\calipers\wef` so Shared Folder trust survives across runs

First-time Windows box: if the task pane does not appear, Insert → Add-ins → Shared Folder → **Calipers Office.js runner**. After that, `excel-run` should auto-open it.

Office.js cannot be eval’d through COM. Corpus scripts stay `Excel.run` (not Automate / Office Scripts).

Goldens must be generated on **Windows Excel**. macOS Excel is not used.
