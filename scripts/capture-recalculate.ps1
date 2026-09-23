# Windows + desktop Excel. Engine must implement Calipers save --recalculate.
param([Parameter(Mandatory = $true)][string]$Engine)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$enginePath = (Resolve-Path $Engine).Path
$root = Split-Path $PSScriptRoot -Parent
Push-Location $root
try {
    $out = Join-Path $root 'tmp/recalculate'
    New-Item -ItemType Directory -Force $out | Out-Null
    $calipers = Join-Path $out 'calipers.exe'
    & go build -o $calipers ./cmd/calipers
    if ($LASTEXITCODE -ne 0) { throw 'Calipers build failed' }

    & $calipers excel-save-pass verification/cases/recalculate | Tee-Object "$out/generate.log"
    if ($LASTEXITCODE -ne 0) { throw 'Excel golden generation failed' }
    & go run ./scripts/gen-recalculate-cases -check-goldens | Tee-Object "$out/expected-results.log"
    if ($LASTEXITCODE -ne 0) { throw 'Excel results do not match the synthetic fixture expectations' }

    & $calipers measure-budgets --engine excel --suite recalculate --force | Tee-Object "$out/budgets.log"
    if ($LASTEXITCODE -ne 0) { throw 'Excel budget measurement failed' }
    foreach ($case in Get-ChildItem verification/cases/recalculate -Directory) {
        $budget = Get-Content (Join-Path $case.FullName 'config.json') -Raw | ConvertFrom-Json
        if ($budget.maxPeakMemoryBytes -le 0 -or $budget.maxDurationMs -le 0) {
            throw "Missing measured budget for $($case.Name)"
        }
    }

    & $calipers verify --engine excel --suite recalculate --out-dir "$out/excel" | Tee-Object "$out/excel-verify.log"
    if ($LASTEXITCODE -ne 0) { throw 'Excel did not match its goldens' }
    if (!(Select-String -Path "$out/excel-verify.log" -Pattern '^verify: 5 pass, 0 fail, 0 error$' -Quiet)) {
        throw 'Excel verification did not run all five cases'
    }

    # A failing verifier is the expected baseline. Preserve stdout and stderr,
    # then distinguish semantic failures from launch/parse errors and skips.
    $ErrorActionPreference = 'Continue'
    & $calipers verify --engine $enginePath --suite recalculate --out-dir "$out/mog" 2>&1 | Tee-Object "$out/mog-verify.log"
    $mogStatus = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    if ($mogStatus -ne 1 -or !(Select-String -Path "$out/mog-verify.log" -Pattern '^verify: 1 pass, 4 fail, 0 error$' -Quiet)) {
        throw "Unexpected Mog baseline (exit $mogStatus); inspect $out/mog-verify.log"
    }
    if (!(Select-String -Path "$out/mog-verify.log" -Pattern 'recalculate/offset_controls PASS' -Quiet)) {
        throw 'The literal/reference controls did not pass'
    }
    Write-Host "Captured Excel goldens, budgets, and expected Mog failures. Evidence: $out"
    Write-Host 'Commit verification/cases/recalculate (golden.xlsx, sidecars, and case config.json files).'
    Write-Host 'Keep tmp/recalculate as the evidence bundle; it is ignored by Git.'
} finally {
    Pop-Location
}
