<#
.SYNOPSIS
    Stops the OutlookMcp Service gracefully before build.
.DESCRIPTION
    Pre-build cleanup script that:
    1. Gracefully stops the OutlookMcp Service via named pipe (service.shutdown)

    This prevents file locking issues during build when the service
    holds handles to assemblies or presentations.
.NOTES
    Called from Directory.Build.props as a BeforeBuild target.
    Safe to run when no processes are running (silently succeeds).
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = 'SilentlyContinue'

function Write-Status($message) {
    if ($Verbose) {
        Write-Host "  [pre-build] $message" -ForegroundColor DarkGray
    }
}

# ----------------------------------------------
# 1. Gracefully stop OutlookMcp Service via CLI
# ----------------------------------------------
function Stop-OutlookMcpService {
    # Look for outlookcli in build output directories (Debug/Release)
    $scriptDir = Split-Path -Parent $PSScriptRoot  # repo root
    $cliPaths = @(
        "$scriptDir\src\OutlookMcp.CLI\bin\Debug\net10.0-windows\outlookcli.exe",
        "$scriptDir\src\OutlookMcp.CLI\bin\Release\net10.0-windows\outlookcli.exe"
    )
    $outlookcli = $cliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($outlookcli) {
        Write-Status "Using CLI: $outlookcli"
        $output = & $outlookcli service stop --quiet 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            # Parse JSON to check if service was running
            try {
                $result = $output | ConvertFrom-Json
                if ($result.message -eq 'Service is not running.') {
                    Write-Status "OutlookMcp Service was not running"
                } else {
                    Write-Host "  OutlookMcp Service stopped gracefully" -ForegroundColor Green
                }
            } catch {
                Write-Status "Service stop completed (exit code 0)"
            }
        } else {
            Write-Status "CLI service stop returned exit code $exitCode, falling back to process kill"
            Stop-OutlookMcpServiceFallback
        }
    } else {
        Write-Status "outlookcli not found (first build?), using fallback"
        Stop-OutlookMcpServiceFallback
    }
}

function Stop-OutlookMcpServiceFallback {
    # Fallback: direct named pipe shutdown (works without CLI binary)
    $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    $pipeName = "OutlookMcp-$sid"

    $pipeExists = Test-Path "\\.\pipe\$pipeName"
    if (-not $pipeExists) {
        Write-Status "OutlookMcp Service not running (no pipe found)"
        return
    }

    Write-Status "OutlookMcp Service detected, sending shutdown via pipe..."
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect(3000)

        $writer = New-Object System.IO.StreamWriter($pipe, [System.Text.Encoding]::UTF8, 4096)
        $writer.AutoFlush = $true
        $reader = New-Object System.IO.StreamReader($pipe, [System.Text.Encoding]::UTF8)

        $writer.WriteLine('{"Command":"service.shutdown"}')
        $response = $reader.ReadLine()
        Write-Status "Service response: $response"

        $reader.Dispose()
        $writer.Dispose()
        $pipe.Dispose()

        Start-Sleep -Milliseconds 500
        Write-Host "  OutlookMcp Service stopped gracefully" -ForegroundColor Green
    }
    catch {
        Write-Status "Could not connect to pipe: $($_.Exception.Message)"
        $serviceProcs = Get-Process -Name 'OutlookMcp.McpServer', 'OutlookMcp.Service' -ErrorAction SilentlyContinue
        if ($serviceProcs) {
            $serviceProcs | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host "  OutlookMcp Service processes killed (pipe unavailable)" -ForegroundColor Yellow
        }
    }
}

# ----------------------------------------------
# Run cleanup
# ----------------------------------------------
Stop-OutlookMcpService
