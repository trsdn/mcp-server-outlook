# Fails the build when anything final-releases the shared Outlook.Application.
#
# The Application returned by OutlookInteropRunner.TryGetRunningApplication is the user's
# already-running Outlook instance, cached per-process by the CLR's RCW table. Calling
# Marshal.FinalReleaseComObject on it - which is what the generic
# OutlookInteropRunner.ReleaseComObject overload does - zeroes the refcount for *every* holder in
# the process, not just the caller. Later code is then handed a wrapper that has been separated
# from its RCW.
#
# In production that is #19. In the test suite it does not surface as a test failure at all: the
# test host dies with STATUS_STACK_BUFFER_OVERRUN (0xc0000409), which reads as infrastructure
# flakiness. See #116.
#
# Use OutlookInteropRunner.ReleaseSharedComObject (a plain ref-count decrement) instead, or better,
# go through OutlookInteropRunner.Execute and never hold the Application at all.

$ErrorActionPreference = "Stop"

$rootDir = Split-Path -Parent $PSScriptRoot
$searchRoots = @(
    (Join-Path $rootDir "src"),
    (Join-Path $rootDir "tests")
)

# Variable names that denote the shared Outlook.Application in this codebase.
$applicationNames = "application|app|outlookApplication|outlookApp"

$patterns = @(
    # OutlookInteropRunner.ReleaseComObject(ref application) - the generic overload is FinalRelease.
    "ReleaseComObject\(\s*ref\s+($applicationNames)\b",
    # Marshal.FinalReleaseComObject(application)
    "FinalReleaseComObject\(\s*($applicationNames)\b"
)

$violations = @()
$filesChecked = 0

foreach ($searchRoot in $searchRoots) {
    if (-not (Test-Path $searchRoot)) { continue }

    foreach ($file in Get-ChildItem -Path $searchRoot -Filter *.cs -Recurse -File) {
        if ($file.FullName -match "\\(obj|bin)\\") { continue }

        $filesChecked++
        $lineNumber = 0

        foreach ($line in Get-Content -Path $file.FullName) {
            $lineNumber++

            # ReleaseSharedComObject is the correct call and contains "ReleaseComObject" as a
            # substring, so exclude it explicitly rather than by pattern subtlety.
            if ($line -match "ReleaseSharedComObject") { continue }

            foreach ($pattern in $patterns) {
                if ($line -match $pattern) {
                    $violations += [PSCustomObject]@{
                        File = $file.FullName.Substring($rootDir.Length + 1)
                        Line = $lineNumber
                        Text = $line.Trim()
                    }
                    break
                }
            }
        }
    }
}

Write-Host "Checked $filesChecked C# files for final-release of the shared Outlook.Application"

if ($violations.Count -gt 0) {
    Write-Host ""
    Write-Host "Found $($violations.Count) site(s) final-releasing the shared Outlook.Application:" -ForegroundColor Red
    foreach ($violation in $violations) {
        Write-Host ("  {0}:{1}" -f $violation.File, $violation.Line) -ForegroundColor Yellow
        Write-Host ("      {0}" -f $violation.Text) -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "Use OutlookInteropRunner.ReleaseSharedComObject instead, or go through" -ForegroundColor Red
    Write-Host "OutlookInteropRunner.Execute and do not hold the Application at all. See #19 and #116." -ForegroundColor Red
    exit 1
}

Write-Host "No site final-releases the shared Outlook.Application" -ForegroundColor Green
exit 0
