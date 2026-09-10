# Calipers

CLI for **Excel-golden generation** and **Excel-vs-engine verification**.

Open an `.xlsx` in desktop Microsoft Excel, save it again, and record sidecar metadata. The Excel rewrite is the golden: later engine output is compared to that golden, not to the original file.

v1 is **Windows-only**. Goldens are produced via Excel COM (`Excel.Application`). Do **not** generate goldens on macOS — Excel for Mac serializes OOXML differently.

This repository is the verification tool and its init corpus. It is not a spreadsheet engine.

`verify` PASS/FAIL is resolved workbook semantics (values, types, formulas, styles, sheets, names, merges, freeze, `date1904`), not ZIP-part or byte equality. Optional `--package` prints ZIP-part diffs as a diagnostic and does not change the gate. See [comparison scope](verification/README.md#what-to-compare).

## Requirements

- Go 1.22+
- For `excel-save` and `excel-run`: Windows + Microsoft Excel (desktop)

Off Windows, those commands exit with:

```
calipers: excel-save requires Windows + Excel (COM)
```

`version`, `verify`, and `-h` / `--help` work on any OS. `verify` requires `--engine` (a binary path or `excel`). `verify --engine excel` needs Windows + Excel. `verify --engine <bin>` execs a caller-supplied binary (`save` / `run`).

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

# Default-pass goldens (load+save; skip Office.js)
calipers excel-save-pass [cases-dir]

# Open input, run Office.js inside Excel (sideloaded add-in), Save As
calipers excel-run <input.xlsx> <script.js> <output.xlsx>

# Pending scripted goldens (Windows; skips cases that already have a golden)
calipers excel-run-pass [cases-dir]

# Optional peak-memory / duration budgets (Windows)
calipers measure-budgets [--engine excel|PATH] [--margin 1.5] [--suite NAME]

# Engine vs committed Excel goldens (binary host or Excel)
calipers verify --engine /path/to/engine
calipers verify --engine /path/to/engine --suite roundtrip
calipers verify --engine /path/to/engine --suite scratch
calipers verify --engine /path/to/engine --case roundtrip/simple
calipers verify --engine excel --case roundtrip/simple
calipers verify --engine /path/to/engine --package --case roundtrip/simple

# Opt in to recalculation with an external engine supporting --recalculate
calipers verify --engine /path/to/engine --recalculate --case roundtrip/formulas

# Tool version (on Windows, also prints Excel version when COM works)
calipers version

# Usage
calipers -h
```

Examples:

```bash
calipers excel-save verification/cases/roundtrip/simple/init.xlsx verification/cases/roundtrip/simple/golden.xlsx
calipers excel-save-pass
calipers excel-run verification/cases/default/simple_set_a1/init.xlsx \
  verification/cases/default/simple_set_a1/script.js \
  verification/cases/default/simple_set_a1/golden.xlsx

# Regenerate pending Office.js goldens (Windows)
calipers excel-run-pass
calipers excel-run-pass verification/cases/officejs

# Optional peak-memory / duration budgets (Windows)
calipers measure-budgets --engine excel --suite officejs

# External engine binary (argv: save <in> <out> / run <in> <script.js> <out>)
calipers verify --engine /path/to/engine
calipers verify --engine /path/to/engine --suite roundtrip
calipers verify --engine /path/to/engine --suite scratch
calipers verify --engine /path/to/engine --case roundtrip/simple
# example: a Mog CLI that implements the same argv
calipers verify --engine ./vendor/mog/target-native/debug/mog --case roundtrip/simple
```

A successful run writes `golden.xlsx.meta.json` beside the xlsx (`host`, Excel version/build when available, OS, tool name, input basename, optional `script` identity, UTC `generatedAt`). Load+save goldens omit `script`. `excel-run` records the script basename.

CI never runs Excel. A Windows machine with Excel generates goldens; those files are committed and compared later.

`verify` prints the selected calculation policy. By default it leaves each host's policy unchanged. `--recalculate` passes `save --recalculate <in> <out>` or `run --recalculate <in> <script> <out>` to an external engine that supports this contract. The flag is rejected for the Excel host, whose calculation is not explicitly controlled. Recalculation does not make random, time-dependent, or environment-dependent results and their dependents equal previously captured goldens; those require controlled assertions.

## Cases

Each case lives in [`verification/cases/<suite>/<id>/`](verification/cases/) as a feature-named directory. Required `init.xlsx`, optional `script.js`, optional `config.json` (peak-memory / duration budgets), and a dedicated `golden.xlsx` (not mixed into the inits). Missing or empty `script.js` means load+save only — skip script execution. Missing `config.json` is valid (no budget). Suites: [`roundtrip/`](verification/cases/roundtrip/) (load+save semantic comparison), [`default/`](verification/cases/default/) (untriaged until categorized), [`scratch/`](verification/cases/scratch/) (Office.js from an empty init; committed `excel-run` goldens), and [`officejs/`](verification/cases/officejs/) (~100 blank-init Office.js cases with committed `excel-run` goldens). **203 cases** (83 roundtrip, 5 default, 15 scratch, 100 officejs). Six XLSX roundtrip/lost-info cases have committed Excel-win goldens; see [`verification/GOLDENS.md`](verification/GOLDENS.md). Eight hostiles live in [`_disabled/`](verification/cases/_disabled/) and are not loaded. Office.js goldens: [`simple_set_a1`](verification/cases/default/simple_set_a1/) (sets A1; `simple` stays load+save in roundtrip), the 15 `scratch/` cases, and `officejs/`. Budget numbers are a Windows follow-up (`measure-budgets`). See [`verification/README.md`](verification/README.md) for sources and compare rules.

`calipers verify` walks committed-golden cases (roundtrip, default, scratch, and officejs). `--suite scratch` runs only the empty-init Office.js cases with goldens; `--suite officejs` runs the larger blank-init set; `--suite roundtrip` runs one suite; `--case roundtrip/simple` runs one test. Case ids are `suite/name`.

A default golden pass (`excel-save-pass`) is unscripted cases; it skips Office.js cases (`simple_set_a1` and `scratch/`). Those Office.js cases have committed `excel-run` goldens (`script=script.js`). Hostiles in `_disabled/` are not loaded.

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
