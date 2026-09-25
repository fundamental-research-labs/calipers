# Windows + desktop Excel only. Preserve the imported goldens; capture budgets
# and compare recalculated Excel exports against the supplied references.
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$root = Split-Path $PSScriptRoot -Parent
Push-Location $root
try {
    $out = Join-Path $root 'tmp/pv_goldens'
    New-Item -ItemType Directory -Force $out | Out-Null
    $calipers = Join-Path $out 'calipers.exe'
    & go build -o $calipers ./cmd/calipers
    if ($LASTEXITCODE -ne 0) { throw 'Calipers build failed' }

    & go run ./scripts/gen-pv-golden-inits -check
    if ($LASTEXITCODE -ne 0) { throw 'Golden/input integrity check failed' }
    $cases = @(Get-ChildItem verification/cases/pv_goldens -Directory | Where-Object {
        Test-Path (Join-Path $_.FullName 'init.xlsx')
    })
    if ($cases.Count -ne 107) { throw "Expected 107 cases, found $($cases.Count)" }

    # Continue through native stderr so failures are retained in the logs.
    $ErrorActionPreference = 'Continue'
    & $calipers measure-budgets --engine excel --suite pv_goldens --force 2>&1 | Tee-Object "$out/budgets.log"
    $budgetStatus = $LASTEXITCODE
    & $calipers verify --engine excel --suite pv_goldens --out-dir "$out/excel" 2>&1 | Tee-Object "$out/excel-verify.log"
    $verifyStatus = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'

    & go run ./scripts/gen-pv-golden-inits -check
    if ($LASTEXITCODE -ne 0) { throw 'Golden/input files changed during capture' }
    if ($budgetStatus -ne 0 -or !(Select-String -Path "$out/budgets.log" -Pattern '^measure-budgets: 107 wrote, 0 failed, 0 skipped$' -Quiet)) {
        throw "Incomplete Excel budget capture. Inspect $out/budgets.log and rerun after resolving errors."
    }
    foreach ($case in $cases) {
        $budget = Get-Content (Join-Path $case.FullName 'config.json') -Raw | ConvertFrom-Json
        if ($budget.maxPeakMemoryBytes -le 0 -or $budget.maxDurationMs -le 0) {
            throw "Missing measured budget for $($case.Name)"
        }
    }
    Write-Host 'Captured 107 measured case configs. Commit only verification/cases/pv_goldens/*/config.json.'
    Write-Host "Keep $out as the evidence bundle (ignored by Git)."
    if ($verifyStatus -ne 0 -or !(Select-String -Path "$out/excel-verify.log" -Pattern '^verify: 107 pass, 0 fail, 0 error$' -Quiet)) {
        throw "Excel does not fully reproduce the supplied goldens. Share $out/excel-verify.log; preserve the goldens for investigation. Budgets are ready to commit."
    }
    Write-Host 'Excel matched all 107 supplied goldens. Mog verification is the next, separate step.'
} finally {
    Pop-Location
}
