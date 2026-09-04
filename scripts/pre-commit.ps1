#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks, Core Commands coverage, CLI Settings usage, Success flag violations, CLI workflow, and MCP Server functionality

.DESCRIPTION
    Runs checks before allowing commits:
    0. Process cleanup - kills stale outlookcli and MCP server processes to prevent file locks
    1. COM leak checker - ensures no Outlook COM objects are leaked
    2. Coverage check - ensures Core Commands are exposed via MCP Server (CoreCommandsCoverageTests)
    3. CLI Settings usage check - ensures every CLI Settings property is actually passed to the daemon
    4. Success flag validation - ensures Success=true never paired with ErrorMessage (Rule 0)
    5. CLI workflow smoke test - validates end-to-end CLI functionality
    6. MCP Server smoke test - validates all MCP tools work correctly

    Ensures code quality and prevents regression.

.EXAMPLE
    .\pre-commit.ps1

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy .git/hooks/pre-commit (bash) or configure Git to use this PowerShell version.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

# CRITICAL: Check branch FIRST - never commit directly to master (Rule 6)
Write-Host "Checking current branch..." -ForegroundColor Cyan
$currentBranch = git branch --show-current

if ($currentBranch -eq "master") {
    Write-Host ""
    Write-Host "BLOCKED: Cannot commit directly to 'master' branch!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Rule 6: All Changes Via Pull Requests" -ForegroundColor Yellow
    Write-Host "   'Never commit to master. Create feature branch -> PR -> CI/CD + review -> merge.'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   To fix:" -ForegroundColor Cyan
    Write-Host "   1. git stash                                    # Save your changes" -ForegroundColor White
    Write-Host "   2. git checkout -b feature/your-feature-name    # Create feature branch" -ForegroundColor White
    Write-Host "   3. git stash pop                                # Restore changes" -ForegroundColor White
    Write-Host "   4. git add <files>                              # Stage changes" -ForegroundColor White
    Write-Host "   5. git commit -m 'your message'                 # Commit to feature branch" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "Branch check passed - on '$currentBranch' (not master)" -ForegroundColor Green
Write-Host ""

# Kill stale CLI and MCP server processes to avoid file locks on Release binaries
Write-Host "Killing stale CLI and server processes..." -ForegroundColor Cyan

$killedProcesses = @()
foreach ($procName in @("outlookcli", "OutlookMcp.McpServer", "OutlookMcp.Service")) {
    $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        $killedProcesses += "$procName ($($procs.Count))"
    }
}

if ($killedProcesses.Count -gt 0) {
    Write-Host "   Killed: $($killedProcesses -join ', ')" -ForegroundColor Yellow
    # Brief pause to let file handles release
    Start-Sleep -Milliseconds 500
}
else {
    Write-Host "   No stale processes found" -ForegroundColor Gray
}

Write-Host "Process cleanup done" -ForegroundColor Green
Write-Host ""

Write-Host "Checking for COM object leaks..." -ForegroundColor Cyan

try {
    $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
    & $leakCheckScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "COM object leaks detected! Fix them before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "COM leak check passed" -ForegroundColor Green
}
catch {
    # A gate that cannot run is a failed gate, not a warning. Downgrading this to
    # "continuing" is how check-dynamic-casts.ps1 silently stopped running for every
    # commit made through Windows PowerShell 5.1 (see #82).
    Write-Host "Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   This check is mandatory and could not be completed." -ForegroundColor Red
    exit 1
}

Write-Host ""

Write-Host "Checking nothing final-releases the shared Outlook.Application..." -ForegroundColor Cyan

try {
    $sharedApplicationScript = Join-Path $rootDir "scripts\check-shared-application-release.ps1"
    & $sharedApplicationScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "The shared Outlook.Application is being final-released. Fix it before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "Shared Application release check passed" -ForegroundColor Green
}
catch {
    Write-Host "Error running shared Application release check: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   This check is mandatory and could not be completed." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking Core Commands coverage (Outlook surface)..." -ForegroundColor Cyan

try {
    # audit-core-coverage.ps1 (regex-scraping a hand-authored ToolActions.cs) predates the move
    # to Roslyn source generators (#5/#11) -- actions are now generated directly from
    # [ServiceAction] attributes on Core interfaces, so ToolActions.cs no longer exists and the
    # old script silently reported "0/0 = 100% coverage", a false-green gate (#25). This runs the
    # reflection-based CoreCommandsCoverageTests instead, which enumerates the live Outlook Core
    # interfaces/generated enums and genuinely fails if any [ServiceAction] method lacks a
    # matching enum value.
    $coverageOutput = dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests" --verbosity minimal 2>&1 | Out-String
    $coverageExitCode = $LASTEXITCODE

    if (-not ($coverageOutput -match "Passed!.*Passed:\s*[1-9]")) {
        Write-Host ""
        Write-Host "CRITICAL: No coverage tests passed! Filter may have matched zero tests." -ForegroundColor Red
        Write-Host $coverageOutput -ForegroundColor Gray
        exit 1
    }

    if ($coverageExitCode -ne 0) {
        Write-Host ""
        Write-Host "Coverage issues detected!" -ForegroundColor Red
        Write-Host "   All Core methods must be exposed via MCP Server with a matching enum action." -ForegroundColor Red
        Write-Host $coverageOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "Coverage check passed - Core methods have matching enum actions" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running coverage check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking CLI Settings property usage..." -ForegroundColor Cyan

try {
    $cliSettingsScript = Join-Path $rootDir "scripts\check-cli-settings-usage.ps1"
    & $cliSettingsScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "CLI Settings property usage issues detected!" -ForegroundColor Red
        Write-Host "   A Settings property is defined but not passed to the daemon -- user values would be silently dropped." -ForegroundColor Red
        exit 1
    }

    Write-Host "CLI Settings usage check passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running CLI Settings usage check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking Success flag violations (Rule 0)..." -ForegroundColor Cyan

try {
    $successFlagScript = Join-Path $rootDir "scripts\check-success-flag.ps1"
    & $successFlagScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Success flag violations detected!" -ForegroundColor Red
        Write-Host "   CRITICAL: Success=true with ErrorMessage confuses LLMs and causes data corruption." -ForegroundColor Red
        Write-Host "   Fix the violations before committing (add Success=false in catch blocks)." -ForegroundColor Red
        exit 1
    }

    Write-Host "Success flag check passed - all flags match reality" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running success flag check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# NOTE: CLI coverage checks removed - commands are now auto-generated by Roslyn source generators
# The CLI generator produces all command classes and registration from Core interfaces
# Validation is handled by:
# - Build-time generator errors if interfaces are malformed
# - CLI workflow smoke test below (end-to-end validation)

Write-Host ""
Write-Host "Regenerating skill files (Release build)..." -ForegroundColor Cyan

# This block used to assert that "the Release build already ran". It had not: the only Release
# build in this script is the one inside the CLI workflow smoke test *below*. So an edit to
# skills/shared/*.md was staged while the copies under skills/outlook-*/references/ were still
# stale on disk, and the build that refreshed them happened after the staging - leaving the
# regenerated files out of the commit, which is the exact failure the block exists to prevent.
# Build first, then stage, so the claim is true by construction.
$buildOutput = dotnet build -c Release --nologo 2>&1 | Out-String
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Release build failed - skill files cannot be regenerated." -ForegroundColor Red
    Write-Host $buildOutput -ForegroundColor Gray
    exit 1
}
Write-Host "Release build succeeded" -ForegroundColor Green

Write-Host ""
Write-Host "Auto-staging generated SKILL.md files..." -ForegroundColor Cyan

try {
    # SKILL.md files and their references/ copies are generated during the Release build above.
    # Auto-stage them so developers never have to think about it.
    $skillPaths = @(
        "skills/outlook-mcp/SKILL.md",
        "skills/outlook-cli/SKILL.md",
        "skills/outlook-mcp/references/",
        "skills/outlook-cli/references/"
    )

    # git writes advisory notices such as "LF will be replaced by CRLF" to stderr. With
    # $ErrorActionPreference = 'Stop', Windows PowerShell turns those into terminating errors, so
    # this whole block used to abort on a *warning*: git add never ran, the catch printed
    # "Continuing", and the script still reported "All pre-commit checks passed!" while the
    # regenerated skill files sat unstaged. Redirection alone does not help - the error record is
    # created before it is discarded - so relax the preference around the git calls and check
    # $LASTEXITCODE explicitly instead.
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $skillDiff = git diff --name-only -- @skillPaths 2>$null
        $untrackedSkills = git ls-files --others --exclude-standard -- @skillPaths 2>$null

        $allChanges = @()
        if ($skillDiff) { $allChanges += $skillDiff }
        if ($untrackedSkills) { $allChanges += $untrackedSkills }

        if ($allChanges.Count -gt 0) {
            git add -- @skillPaths 2>$null
            if ($LASTEXITCODE -ne 0) {
                throw "git add returned exit code $LASTEXITCODE"
            }

            # Prove the staging actually happened. Reporting success here without checking is the
            # failure this block already shipped once.
            $stillUnstaged = git diff --name-only -- @skillPaths 2>$null
            if ($stillUnstaged) {
                throw "still unstaged after git add: $($stillUnstaged -join ', ')"
            }

            Write-Host "Skill files were regenerated and auto-staged ($($allChanges.Count) files)" -ForegroundColor Green
            $allChanges | ForEach-Object { Write-Host "   + $_" -ForegroundColor DarkGray }
        } else {
            Write-Host "Skill files are already up to date" -ForegroundColor Green
        }
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }
}
catch {
    # Not recoverable: committing without the regenerated skill files ships a tool surface whose
    # documentation disagrees with it. Fail the commit rather than continue.
    Write-Host ""
    Write-Host "Error auto-staging SKILL.md files: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Generated skill files would be left out of this commit." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Running CLI workflow smoke test..." -ForegroundColor Cyan

try {
    $cliWorkflowScript = Join-Path $rootDir "scripts\Test-CliWorkflow.ps1"
    $cliWorkflowOutput = & $cliWorkflowScript 2>&1 | Out-String
    $cliWorkflowExitCode = $LASTEXITCODE

    if ($cliWorkflowExitCode -ne 0) {
        Write-Host ""
        Write-Host "CLI workflow smoke test failed!" -ForegroundColor Red
        Write-Host "   This test validates the end-to-end CLI workflow." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $cliWorkflowOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "CLI workflow smoke test passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running CLI workflow smoke test: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Running MCP Server smoke test..." -ForegroundColor Cyan

# Stop OutlookMcp Service before smoke test to prevent DLL locking
& "$PSScriptRoot\Stop-OutlookMcpProcesses.ps1"

try {
    # Run the smoke test - validates all MCP tools work correctly
    $smokeTestFilter = "FullyQualifiedName~McpServerIntegrationTests.SmokeTest_AllTools_E2EWorkflow"

    Write-Host "   dotnet test --filter `"$smokeTestFilter`"" -ForegroundColor Gray

    # Capture output to verify tests actually ran (dotnet test returns 0 even if no tests match!)
    $testOutput = dotnet test --filter $smokeTestFilter --verbosity minimal 2>&1 | Out-String
    $testExitCode = $LASTEXITCODE

    # Check if any tests actually passed (critical - filter typos cause silent failures!)
    # Note: "No test matches" appears for projects without the test, so we check for "Passed"
    if (-not ($testOutput -match "Passed!.*Passed:\s*[1-9]")) {
        Write-Host ""
        Write-Host "CRITICAL: No smoke tests passed! Filter may have matched zero tests." -ForegroundColor Red
        Write-Host "   Filter: $smokeTestFilter" -ForegroundColor Yellow
        Write-Host "   This likely means the test was renamed or deleted." -ForegroundColor Yellow
        Write-Host "   Verify the test exists: McpServerIntegrationTests.SmokeTest_AllTools_E2EWorkflow" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    if ($testExitCode -ne 0) {
        Write-Host ""
        Write-Host "MCP Server smoke test failed! Core functionality is broken." -ForegroundColor Red
        Write-Host "   This test validates all MCP tools work correctly." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "MCP Server smoke test passed - all tools functional" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running smoke test: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking for undocumented ((dynamic)) casts..." -ForegroundColor Cyan

try {
    $dynamicCastScript = Join-Path $rootDir "scripts\check-dynamic-casts.ps1"
    & $dynamicCastScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Undocumented ((dynamic)) casts detected!" -ForegroundColor Red
        Write-Host "   Add a justification comment (// PIA gap:, // TODO:, or // Reason:) before each cast." -ForegroundColor Red
        Write-Host "   See docs/PIA-COVERAGE.md for guidance." -ForegroundColor Red
        exit 1
    }

    Write-Host "Dynamic cast check passed - all casts are documented" -ForegroundColor Green
}
catch {
    # See the COM leak check above: a gate that errors out has not passed.
    Write-Host "Error running dynamic cast check: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   This check is mandatory and could not be completed." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "All pre-commit checks passed!" -ForegroundColor Green
exit 0
