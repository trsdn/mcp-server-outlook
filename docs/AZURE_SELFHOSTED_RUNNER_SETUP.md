# Azure Self-Hosted Runner Setup

This document describes the minimum setup needed to activate `.github/workflows/integration-tests.yml`.

## Why a Self-Hosted Runner Is Required

Real Outlook COM automation requires a Windows machine with Microsoft Outlook (classic) installed, a configured mail profile, and an interactive desktop session. GitHub-hosted runners do not provide that environment.

The integration workflow is therefore present in the repository but only becomes active when:

- repository variable `ENABLE_OUTLOOK_INTEGRATION_CI` is set to `true`
- a self-hosted Windows runner with the label `outlook` is available

Until then, the workflow exits with a status message instead of pretending that Outlook integration is covered in CI.

## Recommended Host Requirements

- Windows 11 or Windows Server with desktop experience
- Microsoft 365 Apps / Outlook (classic desktop) installed, licensed, and with a working mail profile
- .NET SDK `9.0.x`
- `uv` available on PATH for `llm-tests/`
- Stable disk space for build outputs and test artifacts
- Runner labels: `self-hosted`, `windows`, `outlook`

## Desktop Session Requirement

Outlook COM automation is not reliably headless. Use a runner host that keeps an interactive desktop session available for the runner user.

Recommended practice:

- dedicate the machine to Outlook integration workloads
- use a dedicated local/service account for the runner
- verify that Outlook can open and close normally under that account, without a first-run or profile prompt, before enabling CI

## Basic Setup Steps

1. Provision the Windows host or Azure VM.
2. Install Outlook (classic) and confirm it opens successfully for the runner account with a default mail profile.
3. Install the .NET 9 SDK.
4. Install `uv`.
5. Register the GitHub Actions runner for this repository.
6. Add the `outlook` label to that runner.
7. Set repository variable `ENABLE_OUTLOOK_INTEGRATION_CI=true`.
8. Optionally add secret `AZURE_OPENAI_ENDPOINT` if you want workflow-dispatch LLM gate runs.

## Validation Checklist

Before enabling the repository variable, validate on the runner host:

```powershell
dotnet build src\OutlookMcp.CLI\OutlookMcp.CLI.csproj -c Release
dotnet build src\OutlookMcp.McpServer\OutlookMcp.McpServer.csproj -c Release
.\scripts\Test-CliWorkflow.ps1
dotnet test tests\OutlookMcp.McpServer.Tests\OutlookMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~McpServerIntegrationTests.SmokeTest_AllTools_E2EWorkflow"
```

If those pass locally on the runner host, enable `ENABLE_OUTLOOK_INTEGRATION_CI` and trigger `integration-tests.yml` with `workflow_dispatch`.

## Optional LLM Regression Gate

The workflow can also run the canonical LLM regression gate when dispatched manually.

Prerequisites:

- `AZURE_OPENAI_ENDPOINT` secret configured
- runner host already passes the regular Outlook smoke/integration steps

Manual local command:

```powershell
.\scripts\Test-LlmRegressionGate.ps1
```
