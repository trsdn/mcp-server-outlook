#Requires -Version 7.0
<#
.SYNOPSIS
    Fails the build if every RequiresOutlook=true test was skipped in a given TRX result file.

.DESCRIPTION
    Regression guard for issue #22: Outlook integration tests use Xunit.SkippableFact to report
    a real "Skipped" outcome (never a false "Passed") when no classic Outlook desktop instance is
    available. That is correct behavior on a developer machine or a runner without Outlook, but a
    CI job that is SUPPOSED to have Outlook available must fail loudly if it skipped every single
    Outlook test - a wall of skips there means the runner's Outlook profile is broken/missing, not
    that "everything passed".

    This script inspects a .trx file for tests whose fully qualified name matches the Outlook smoke
    test class and asserts that at least one of them actually ran (Passed or Failed), not just
    Skipped/NotExecuted.

.PARAMETER TrxPath
    Path to the .trx result file produced by `dotnet test --logger "trx;LogFileName=..."`.

.PARAMETER TestNamePattern
    Regex applied to each test's fully qualified name to select the Outlook test population.
    Defaults to the OutlookSeedSmokeTests class.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$TrxPath,

    [string]$TestNamePattern = 'OutlookSeedSmokeTests'
)

if (-not (Test-Path $TrxPath)) {
    Write-Error "TRX file not found: $TrxPath"
    exit 1
}

[xml]$trx = Get-Content $TrxPath -Raw

# TRX namespaces are versioned; use a namespace-agnostic XPath via local-name().
$results = $trx.SelectNodes("//*[local-name()='UnitTestResult']")

if ($results.Count -eq 0) {
    Write-Error "No <UnitTestResult> entries found in $TrxPath - cannot verify Outlook test execution."
    exit 1
}

$outlookResults = $results | Where-Object { $_.testName -match $TestNamePattern }

if ($outlookResults.Count -eq 0) {
    Write-Error "No tests matching pattern '$TestNamePattern' were found in $TrxPath. Expected the Outlook smoke test suite to be present."
    exit 1
}

$executed = $outlookResults | Where-Object { $_.outcome -in @('Passed', 'Failed') }
$skipped = $outlookResults | Where-Object { $_.outcome -notin @('Passed', 'Failed') }

Write-Output "Outlook test results: $($outlookResults.Count) total, $($executed.Count) executed, $($skipped.Count) skipped."

if ($executed.Count -eq 0) {
    Write-Error "All $($outlookResults.Count) Outlook test(s) were skipped. This CI job is expected to run against a live classic Outlook desktop instance - a 100% skip rate means the runner's Outlook profile is missing or broken, not that the suite passed. See docs/AZURE_SELFHOSTED_RUNNER_SETUP.md and issue #22/#31."
    exit 1
}

Write-Output "OK: at least one Outlook test executed (not skipped)."
exit 0
