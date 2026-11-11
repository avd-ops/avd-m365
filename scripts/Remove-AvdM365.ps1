# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()][string]$RunManifestPath = './m365-run-manifest.json',
    [Parameter()][ValidateSet('commercial', 'usgov')][string]$Cloud = 'commercial',
    [Parameter()][switch]$DiscardLastCopy,
    [Parameter()][switch]$DryRun
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-AvdCloudEnvironment {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable], [System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter()][ValidateSet('commercial', 'usgov')][string]$Cloud = 'commercial',
        [Parameter()][string]$ContractPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'contracts/cloud-environments.json')
    )
    if (-not (Test-Path -LiteralPath $ContractPath)) {
        throw "Cloud environments contract (C26) not found at '$ContractPath'."
    }
    $contract = Get-Content -LiteralPath $ContractPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
    if (-not $contract.ContainsKey('clouds') -or -not $contract['clouds'].ContainsKey($Cloud)) {
        throw "Cloud '$Cloud' not defined in '$ContractPath'."
    }
    return $contract['clouds'][$Cloud]
}

function Test-AvdM365CloudBoundary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('commercial', 'usgov')][string]$TargetCloud,
        [Parameter(Mandatory)][hashtable]$CloudConfig,
        [Parameter()][hashtable]$RunManifest
    )

    if ($null -ne $RunManifest -and $RunManifest.ContainsKey('cloud') -and -not [string]::IsNullOrWhiteSpace([string]$RunManifest['cloud'])) {
        $manifestCloud = [string]$RunManifest['cloud']
        if ($manifestCloud -ne $TargetCloud) {
            throw "Cloud boundary mismatch: m365-run-manifest.json specifies cloud '$manifestCloud', but target cloud is '$TargetCloud'. Refusing cross-cloud operation."
        }
    }

    $expectedAzEnv = [string]$CloudConfig['psEnvironmentName']
    $expectedEntraEnv = [string]$CloudConfig['entraEnvironment']

    $mockPath = $env:AVD_OPS_MOCK_GRAPH_PATH
    if (-not [string]::IsNullOrWhiteSpace($mockPath) -and (Test-Path -LiteralPath $mockPath)) {
        $mockState = Get-Content -LiteralPath $mockPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
        if ($mockState.ContainsKey('azEnvironment') -and -not [string]::IsNullOrWhiteSpace([string]$mockState['azEnvironment'])) {
            $mockAzEnv = [string]$mockState['azEnvironment']
            if ($mockAzEnv -ne $expectedAzEnv) {
                throw "Cloud boundary mismatch: active Az environment '$mockAzEnv' does not match target cloud '$TargetCloud' ('$expectedAzEnv')."
            }
        }
        if ($mockState.ContainsKey('mgEnvironment') -and -not [string]::IsNullOrWhiteSpace([string]$mockState['mgEnvironment'])) {
            $mockMgEnv = [string]$mockState['mgEnvironment']
            if ($mockMgEnv -ne $expectedEntraEnv) {
                throw "Cloud boundary mismatch: active MgGraph environment '$mockMgEnv' does not match target cloud '$TargetCloud' ('$expectedEntraEnv')."
            }
        }
    }

    $skipLiveSessionCheck = -not [string]::IsNullOrWhiteSpace($mockPath)
    $azCtxCmd = Get-Item -LiteralPath 'Function:Get-AzContext' -ErrorAction SilentlyContinue
    if ($null -eq $azCtxCmd -and -not $skipLiveSessionCheck -and $null -ne (Get-Module -Name 'Az.Accounts')) {
        $azCtxCmd = Get-Command 'Get-AzContext' -ErrorAction SilentlyContinue
    }
    if ($null -ne $azCtxCmd) {
        $azCtx = & $azCtxCmd -ErrorAction SilentlyContinue
        if ($null -ne $azCtx -and $null -ne $azCtx.Environment -and -not [string]::IsNullOrWhiteSpace([string]$azCtx.Environment.Name)) {
            $activeAzEnv = [string]$azCtx.Environment.Name
            if ($activeAzEnv -ne $expectedAzEnv) {
                throw "Cloud boundary mismatch: active Az context environment '$activeAzEnv' does not match target cloud '$TargetCloud' ('$expectedAzEnv')."
            }
        }
    }

    $mgCtxCmd = Get-Item -LiteralPath 'Function:Get-MgContext' -ErrorAction SilentlyContinue
    if ($null -eq $mgCtxCmd -and -not $skipLiveSessionCheck -and $null -ne (Get-Module -Name 'Microsoft.Graph.Authentication')) {
        $mgCtxCmd = Get-Command 'Get-MgContext' -ErrorAction SilentlyContinue
    }
    if ($null -ne $mgCtxCmd) {
        $mgCtx = & $mgCtxCmd -ErrorAction SilentlyContinue
        if ($null -ne $mgCtx -and -not [string]::IsNullOrWhiteSpace([string]$mgCtx.Environment)) {
            $activeMgEnv = [string]$mgCtx.Environment
            if ($activeMgEnv -ne $expectedEntraEnv) {
                throw "Cloud boundary mismatch: active MgGraph environment '$activeMgEnv' does not match target cloud '$TargetCloud' ('$expectedEntraEnv'). Use Connect-MgGraph -Environment $expectedEntraEnv."
            }
        }
    }
    if (Test-Path -LiteralPath 'Alias:Invoke-GraphRequest') {
        Remove-Item -LiteralPath 'Alias:Invoke-GraphRequest' -Force -ErrorAction SilentlyContinue
    }
}

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

$resolvedCloud = if ($PSBoundParameters.ContainsKey('Cloud')) {
    $Cloud
}
elseif ($runManifest.ContainsKey('cloud') -and -not [string]::IsNullOrWhiteSpace([string]$runManifest['cloud'])) {
    [string]$runManifest['cloud']
}
else {
    $Cloud
}

if ($resolvedCloud -notin @('commercial', 'usgov')) {
    throw "Invalid cloud profile '$resolvedCloud' in '$RunManifestPath'. Expected 'commercial' or 'usgov'."
}

$cloudConfig = Resolve-AvdCloudEnvironment -Cloud $resolvedCloud
Test-AvdM365CloudBoundary -TargetCloud $resolvedCloud -CloudConfig $cloudConfig -RunManifest $runManifest
$graphEndpoint = ([string]$cloudConfig['graphEndpoint']).TrimEnd('/')

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
        "$graphEndpoint/beta/deviceManagement/configurationPolicies/$resourceId"
    }

    if (-not $wasCreated) {
        Write-Verbose "Skipping pre-existing policy '$displayName' ($resourceId) because created=false."
        continue
    }

    if (-not $displayName.StartsWith($stamp, [System.StringComparison]::Ordinal)) {
        Write-Verbose "Skipping policy '$displayName' ($resourceId) because it does not match stamp '$stamp'."
        continue
    }

    if (-not $uri.StartsWith("$graphEndpoint/", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Cloud boundary mismatch on delete URI '$uri': expected prefix '$graphEndpoint/' for cloud '$resolvedCloud'."
    }

    if ($PSCmdlet.ShouldProcess("$displayName ($resourceId)", 'Delete M365 baseline configuration policy')) {
        Invoke-GraphDelete -Uri $uri -IsDryRun $isDryRun
    }
}

if ($PSCmdlet.ShouldProcess($RunManifestPath, 'Delete last copy of m365-run-manifest.json')) {
    Remove-Item -LiteralPath $RunManifestPath -Force
}
