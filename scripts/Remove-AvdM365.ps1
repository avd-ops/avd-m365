# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()][string]$RunManifestPath = './m365-run-manifest.json',
    [Parameter()][switch]$DiscardLastCopy,
    [Parameter()][switch]$DryRun
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-GraphDelete {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()][bool]$IsDryRun = $false
    )

    $mockPath = $env:AVD_OPS_MOCK_GRAPH_PATH
    if (-not [string]::IsNullOrWhiteSpace($mockPath)) {
        if (-not (Test-Path -LiteralPath $mockPath)) {
            throw "Mock Graph state file not found at '$mockPath'."
        }
        $mockState = Get-Content -LiteralPath $mockPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
        if (-not $mockState.ContainsKey('requests')) {
            $mockState['requests'] = @()
        }
        $mockState['requests'] += [ordered]@{
            method = 'DELETE'
            uri    = $Uri
            body   = $null
        }
        $policyId = ($Uri -split '/')[-1]
        if ($mockState.ContainsKey('configurationPolicies')) {
            $remaining = @()
            foreach ($pol in @($mockState['configurationPolicies'])) {
                if ([string]$pol['id'] -ne $policyId) {
                    $remaining += $pol
                }
            }
            $mockState['configurationPolicies'] = $remaining
        }
        $mockState | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $mockPath -Encoding utf8
        return
    }

    if ($IsDryRun) {
        return
    }

    $rawOut = & az rest --method DELETE --uri $Uri
    if ($LASTEXITCODE -ne 0) {
        throw "az rest DELETE $Uri failed with exit code $LASTEXITCODE ($($rawOut -join ' '))."
    }
}

if (-not $DiscardLastCopy.IsPresent) {
    throw "Refusing to remove M365 baseline policies or delete the last copy of '$RunManifestPath' without -DiscardLastCopy."
}

if (-not (Test-Path -LiteralPath $RunManifestPath)) {
    throw "M365 run manifest not found at '$RunManifestPath'."
}

$isDryRun = $DryRun.IsPresent -or ($env:AVD_OPS_DRY_RUN -eq '1')
$runManifest = Get-Content -LiteralPath $RunManifestPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable

$environment = [string]$runManifest['environment']
$prefix = [string]$runManifest['prefix']
$stamp = if ($runManifest.ContainsKey('stamp') -and -not [string]::IsNullOrWhiteSpace([string]$runManifest['stamp'])) {
    [string]$runManifest['stamp']
}
else {
    "[avd-ops-$prefix-$environment]"
}

foreach ($resource in @($runManifest['resources'])) {
    if ($null -eq $resource) {
        continue
    }
    $wasCreated = [bool]$resource['created']
    $displayName = [string]$resource['displayName']
    $resourceId = [string]$resource['id']
    $uri = if ($resource.ContainsKey('uri') -and -not [string]::IsNullOrWhiteSpace([string]$resource['uri'])) {
        [string]$resource['uri']
    }
    else {
        "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$resourceId"
    }

    if (-not $wasCreated) {
        Write-Verbose "Skipping pre-existing policy '$displayName' ($resourceId) because created=false."
        continue
    }

    if (-not $displayName.StartsWith($stamp, [System.StringComparison]::Ordinal)) {
        Write-Verbose "Skipping policy '$displayName' ($resourceId) because it does not match stamp '$stamp'."
        continue
    }

    if ($PSCmdlet.ShouldProcess("$displayName ($resourceId)", 'Delete M365 baseline configuration policy')) {
        Invoke-GraphDelete -Uri $uri -IsDryRun $isDryRun
    }
}

if ($PSCmdlet.ShouldProcess($RunManifestPath, 'Delete last copy of m365-run-manifest.json')) {
    Remove-Item -LiteralPath $RunManifestPath -Force
}
