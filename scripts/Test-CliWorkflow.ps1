#!/usr/bin/env pwsh
<#
.SYNOPSIS
    End-to-end smoke test for the outlookcli surface.

.DESCRIPTION
    Exercises the CLI exactly as a user would, without requiring Outlook to be
    running. Every assertion holds in both states:

    1. `diag ping`            - the daemon starts and answers.
    2. `diag echo`            - a parameter round-trips through the pipe.
    3. `diag outlook`         - flavour detection returns a well-formed payload.
    4. `service status`       - the daemon reports itself running.
    5. `application get-status` - a *generated* command reaches Core. This is the
       important one: it proves the generated dispatch surface is wired. It
       succeeds when classic Outlook is running and fails cleanly when it is
       not, so the assertion is that the JSON is well formed and the process
       exit code AGREES with the payload's `success` field (issue #63).
    6. `--output` on a failing command must not leave a file behind.

    Deliberately does NOT assert that Outlook is present. Only the self-hosted
    integration runner (#31) can do that.

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1 -Verbose
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Find CLI executable (prefer Release build)
$candidateCliPaths = @(
    "..\src\OutlookMcp.CLI\bin\Release\net9.0-windows\outlookcli.exe",
    "..\src\OutlookMcp.CLI\bin\Debug\net9.0-windows\outlookcli.exe",
    "..\src\OutlookMcp.CLI\bin\Release\net10.0-windows\outlookcli.exe",
    "..\src\OutlookMcp.CLI\bin\Debug\net10.0-windows\outlookcli.exe"
) | ForEach-Object { Join-Path $PSScriptRoot $_ }

$cliPath = $candidateCliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $cliPath) {
    Write-Error "CLI not found. Build first: dotnet build src/OutlookMcp.CLI"
    exit 1
}

$cli = (Resolve-Path $cliPath).Path
Write-Host "Using CLI: $cli" -ForegroundColor Cyan

$passed = 0
$failed = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host "`n[$Name]" -ForegroundColor Yellow
    try {
        $ok = & $Action
        if (-not $ok) {
            Write-Host "  FAIL: assertion returned false" -ForegroundColor Red
            $script:failed++
            return
        }
        Write-Host "  PASS" -ForegroundColor Green
        $script:passed++
    }
    catch {
        Write-Host "  FAIL: $_" -ForegroundColor Red
        $script:failed++
    }
}

# Runs the CLI and returns the parsed JSON plus the real exit code.
function Invoke-Cli {
    param([Parameter(ValueFromRemainingArguments)] [string[]]$CliArgs)

    $raw = & $cli -q @CliArgs 2>&1 | Out-String
    $code = $LASTEXITCODE
    $json = $null
    try { $json = $raw | ConvertFrom-Json } catch { }
    Write-Verbose "args=[$($CliArgs -join ' ')] exit=$code raw=$raw"
    [pscustomobject]@{ Json = $json; ExitCode = $code; Raw = $raw }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Outlook CLI Workflow Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Test-Step "diag ping - daemon starts and answers" {
    $r = Invoke-Cli diag ping
    if ($r.ExitCode -ne 0) { throw "expected exit 0, got $($r.ExitCode)" }
    if (-not $r.Json) { throw "response was not JSON: $($r.Raw)" }
    $r.Json.success -eq $true
}

Test-Step "diag echo - parameter round-trips through the pipe" {
    $marker = "smoke-$(Get-Random)"
    $r = Invoke-Cli diag echo --message $marker
    if ($r.ExitCode -ne 0) { throw "expected exit 0, got $($r.ExitCode)" }
    if (-not $r.Json) { throw "response was not JSON: $($r.Raw)" }
    if ($r.Raw -notmatch [regex]::Escape($marker)) { throw "echo did not return '$marker'" }
    $r.Json.success -eq $true
}

Test-Step "diag outlook - flavour detection returns a well-formed payload" {
    $r = Invoke-Cli diag outlook
    if ($r.ExitCode -ne 0) { throw "expected exit 0, got $($r.ExitCode)" }
    if (-not $r.Json) { throw "response was not JSON: $($r.Raw)" }
    $r.Json.success -eq $true
}

Test-Step "service status - daemon reports itself running" {
    $r = Invoke-Cli service status
    if ($r.ExitCode -ne 0) { throw "expected exit 0, got $($r.ExitCode)" }
    if (-not $r.Json) { throw "response was not JSON: $($r.Raw)" }
    $r.Json.running -eq $true
}

# The generated dispatch surface. Whether this succeeds depends on Outlook being
# present, so assert only on well-formedness and on exit code / payload agreement.
Test-Step "application get-status - generated command reaches Core, exit code agrees with payload" {
    $r = Invoke-Cli application get-status
    if (-not $r.Json) { throw "response was not JSON: $($r.Raw)" }
    if ($null -eq $r.Json.success) { throw "payload has no 'success' property: $($r.Raw)" }

    $expected = if ($r.Json.success -eq $true) { 0 } else { 1 }
    if ($r.ExitCode -ne $expected) {
        throw "success=$($r.Json.success) but exit code was $($r.ExitCode), expected $expected (issue #63)"
    }
    Write-Host "  (success=$($r.Json.success), exit=$($r.ExitCode))" -ForegroundColor DarkGray
    $true
}

Test-Step "unknown action is rejected with a non-zero exit code" {
    $r = Invoke-Cli application definitely-not-an-action
    if ($r.ExitCode -eq 0) { throw "unknown action returned exit 0" }
    $true
}

Test-Step "--output writes no file when the operation fails" {
    $outFile = Join-Path $env:TEMP "cli-smoke-out-$(Get-Random).json"
    $r = Invoke-Cli application definitely-not-an-action --output $outFile
    try {
        if ($r.ExitCode -eq 0) { throw "expected a non-zero exit code" }
        if (Test-Path $outFile) { throw "output file was written despite failure: $outFile" }
        $true
    }
    finally {
        Remove-Item $outFile -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Passed: $passed  Failed: $failed" -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Red' })
Write-Host "========================================" -ForegroundColor Cyan

exit $(if ($failed -eq 0) { 0 } else { 1 })
