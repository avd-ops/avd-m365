# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
# Offline tests for Azure US Government & Microsoft 365 GCC High (usgov) support across
# Deploy-AvdM365.ps1, Remove-AvdM365.ps1, deploy-m365.sh, and C26 cloud-environments.json.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Command 'Describe' -ErrorAction SilentlyContinue)) {
    function global:Describe {
        param([string]$Name, [scriptblock]$Fixture)
        Write-Host "Describe: $Name"
        $script:CurrentBeforeEach = $null
        $script:CurrentAfterEach = $null
        & $Fixture
    }
    function global:BeforeAll {
        param([scriptblock]$Action)
        & $Action
    }
    function global:BeforeEach {
        param([scriptblock]$Action)
        $script:CurrentBeforeEach = $Action
    }
    function global:AfterEach {
        param([scriptblock]$Action)
        $script:CurrentAfterEach = $Action
    }
    function global:It {
        param([string]$Name, [scriptblock]$Test)
        $global:TestDrive = Join-Path ([System.IO.Path]::GetTempPath()) ('pester-shim-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $global:TestDrive -Force
        try {
            if ($null -ne $script:CurrentBeforeEach) {
                & $script:CurrentBeforeEach
            }
            & $Test
            Write-Host "  [+] $Name"
        } finally {
            if ($null -ne $script:CurrentAfterEach) {
                & $script:CurrentAfterEach
            }
            if (Test-Path -LiteralPath $global:TestDrive) {
                Remove-Item -LiteralPath $global:TestDrive -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    function global:Should {
        param(
            [Parameter(ValueFromPipeline)][object]$ActualValue,
            [switch]$BeTrue,
            [switch]$BeFalse,
            [switch]$BeNullOrEmpty,
            [object]$Be = $null,
            [object]$BeGreaterThan = $null,
            [string]$Match = '',
            [string]$Throw = '',
            [switch]$Not,
            [string]$Because = ''
        )
        begin {
            $items = [System.Collections.Generic.List[object]]::new()
        }
        process {
            $items.Add($ActualValue)
        }
        end {
            $val = if ($items.Count -eq 1) { $items[0] } else { $items.ToArray() }

            if ($PSBoundParameters.ContainsKey('Throw')) {
                $threw = $false
                $msg = ''
                try {
                    & $val
                } catch {
                    $threw = $true
                    $msg = $_.Exception.Message
                }
                if ($Not) {
                    if ($threw) { throw "Expected scriptblock not to throw, but it threw '$msg'. $Because" }
                    return
                }
                if (-not $threw) {
                    throw "Expected scriptblock to throw '$Throw', but it did not throw. $Because"
                }
                if (-not [string]::IsNullOrEmpty($Throw) -and $msg -notlike $Throw) {
                    throw "Expected exception message matching '$Throw', got '$msg'. $Because"
                }
                return
            }
            if ($BeNullOrEmpty) {
                $isEmpty = ($null -eq $val) -or ($val -is [string] -and [string]::IsNullOrEmpty($val)) -or ($val -is [System.Collections.IEnumerable] -and @($val).Count -eq 0)
                if ($Not -and $isEmpty) { throw "Expected value not to be null or empty. $Because" }
                if (-not $Not -and -not $isEmpty) { throw "Expected value to be null or empty, got '$val'. $Because" }
                return
            }
            if ($BeTrue) {
                if ([bool]$val -ne (-not $Not)) { throw "Expected $(if ($Not) {'False'} else {'True'}), got '$val'. $Because" }
                return
            }
            if ($BeFalse) {
                if ([bool]$val -ne [bool]$Not) { throw "Expected $(if ($Not) {'True'} else {'False'}), got '$val'. $Because" }
                return
            }
            if ($PSBoundParameters.ContainsKey('BeGreaterThan')) {
                if (-not ($val -gt $BeGreaterThan)) { throw "Expected '$val' to be greater than '$BeGreaterThan'. $Because" }
                return
            }
            if ($PSBoundParameters.ContainsKey('Be')) {
                if ($val -is [System.Array] -and $Be -is [System.Array]) {
                    $same = ($val.Count -eq $Be.Count)
                    if ($same) {
                        for ($i = 0; $i -lt $val.Count; $i++) {
                            if ($val[$i] -ne $Be[$i]) { $same = $false; break }
                        }
                    }
                    if ($Not -and $same) { throw "Expected arrays not to be equal. $Because" }
                    if (-not $Not -and -not $same) { throw "Expected array ($($Be -join ', ')), got ($($val -join ', ')). $Because" }
                    return
                }
                if ($Not) {
                    if ($val -eq $Be) { throw "Expected not '$Be', got '$val'. $Because" }
                } else {
                    if ($val -ne $Be) { throw "Expected '$Be', got '$val'. $Because" }
                }
                return
            }
            if ($PSBoundParameters.ContainsKey('Match')) {
                $matched = ([string]$val -match $Match)
                if ($Not -and $matched) {
                    throw "Expected '$val' not to match '$Match'. $Because"
                }
                if (-not $Not -and -not $matched) {
                    throw "Expected '$val' to match '$Match'. $Because"
                }
            }
        }
    }
}

BeforeAll {
    $script:root = Split-Path -Parent $PSScriptRoot
    $script:deployScript = Join-Path $script:root 'scripts/Deploy-AvdM365.ps1'
    $script:removeScript = Join-Path $script:root 'scripts/Remove-AvdM365.ps1'
    $script:c26Path = Join-Path $script:root 'contracts/cloud-environments.json'
    $script:assessC26Path = Join-Path (Split-Path -Parent $script:root) 'avd-assess/contracts/cloud-environments.json'
    $script:schemaPath = Join-Path $script:root 'contracts/m365-manifest.schema.json'
    $script:baselinePath = Join-Path $script:root 'baselines/vdi-workplace-baseline.json'

    $deployAst = [System.Management.Automation.Language.Parser]::ParseFile($script:deployScript, [ref]$null, [ref]$null)
    foreach ($fnName in @('Resolve-AvdCloudEnvironment', 'Test-AvdM365CloudBoundary')) {
        $fnAst = $deployAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $fnName
        }, $true) | Select-Object -First 1
        if ($null -ne $fnAst) {
            $globalFnText = $fnAst.Extent.Text -replace "^function\s+$([regex]::Escape($fnName))", "function global:$fnName"
            . ([scriptblock]::Create($globalFnText))
        }
    }
}

Describe 'Contract C26 (contracts/cloud-environments.json) & Parameter Declarations' {
    It 'has a valid contracts/cloud-environments.json pinned from avd-assess' {
        Test-Path -LiteralPath $script:c26Path | Should -BeTrue
        $c26 = Get-Content -LiteralPath $script:c26Path -Raw -Encoding utf8 | ConvertFrom-Json
        $c26.schema | Should -Be 'avd-cloud-environments/v1'
        $c26.clouds.commercial.graphEndpoint | Should -Be 'https://graph.microsoft.com'
        $c26.clouds.usgov.graphEndpoint | Should -Be 'https://graph.microsoft.us'
        $c26.clouds.usgov.entraEnvironment | Should -Be 'USGov'
        $c26.clouds.usgov.psEnvironmentName | Should -Be 'AzureUSGovernment'

        if (Test-Path -LiteralPath $script:assessC26Path) {
            $localHash = (Get-FileHash -LiteralPath $script:c26Path -Algorithm SHA256).Hash
            $assessHash = (Get-FileHash -LiteralPath $script:assessC26Path -Algorithm SHA256).Hash
            $localHash | Should -Be $assessHash
        }
    }

    It 'resolves commercial and usgov M365 suffixes and Teams VDI default modes via Resolve-AvdCloudEnvironment' {
        $comm = Resolve-AvdCloudEnvironment -Cloud 'commercial' -ContractPath $script:c26Path
        $usgov = Resolve-AvdCloudEnvironment -Cloud 'usgov' -ContractPath $script:c26Path

        $comm['graphEndpoint'] | Should -Be 'https://graph.microsoft.com'
        $comm['m365']['sharePointSuffix'] | Should -Be '.sharepoint.com'
        $comm['capabilities']['teamsVdiDefaultMode'] | Should -Be 'SlimCore'

        $usgov['graphEndpoint'] | Should -Be 'https://graph.microsoft.us'
        $usgov['m365']['sharePointSuffix'] | Should -Be '.sharepoint.us'
        $usgov['capabilities']['teamsVdiDefaultMode'] | Should -Be 'WebRtc'
    }

    It 'declares [ValidateSet("commercial", "usgov")] -Cloud on Deploy-AvdM365.ps1 and Remove-AvdM365.ps1' {
        foreach ($scriptPath in @($script:deployScript, $script:removeScript)) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$null)
            $cloudParam = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Cloud' }
            $cloudParam | Should -Not -BeNullOrEmpty -Because $scriptPath
            $validateSet = $cloudParam.Attributes | Where-Object { $_.TypeName.Name -eq 'ValidateSet' }
            @($validateSet.PositionalArguments.Value) | Should -Be @('commercial', 'usgov')
        }
    }
}

Describe 'Static AST Zero-Commercial-URI Guard' {
    It 'does not hardcode commercial Graph, ARM, Login, BAP, or .windows.net URIs in scripts/*.ps1' {
        $forbiddenPatterns = @(
            'management\.azure\.com',
            'graph\.microsoft\.com',
            'login\.microsoftonline\.com',
            'api\.bap\.microsoft\.com',
            '\.windows\.net'
        )
        $psFiles = Get-ChildItem -Path (Join-Path $script:root 'scripts') -Filter '*.ps1' -Recurse
        $psFiles.Count | Should -BeGreaterThan 0

        foreach ($file in $psFiles) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $stringLiterals = $ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
                $node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
            }, $true)

            foreach ($lit in $stringLiterals) {
                foreach ($pattern in $forbiddenPatterns) {
                    $lit.Extent.Text | Should -Not -Match $pattern -Because "File $($file.Name) line $($lit.Extent.StartLineNumber) must resolve cloud endpoints from C26 instead of hardcoding '$pattern'"
                }
            }
        }
    }
}

Describe 'Mocked AzureUSGovernment / GCC High Execution & Fail-Closed Cloud Boundary Guard' {
    BeforeEach {
        $script:testDir = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $script:testDir -Force | Out-Null

        $script:entraManifestPath = Join-Path $script:testDir 'entra-manifest.json'
        $script:m365ManifestPath = Join-Path $script:testDir 'm365-manifest.json'
        $script:runManifestPath = Join-Path $script:testDir 'm365-run-manifest.json'
        $script:mockGraphPath = Join-Path $script:testDir 'mock-graph.json'

        $entraPayload = [ordered]@{
            schema             = 'avd-ops-entra-manifest/v1'
            environment        = 'dev'
            cloud              = 'usgov'
            prefix             = 'avd'
            groups             = [ordered]@{
                avdUsersGroupId            = '00000000-0000-0000-0000-000000000011'
                avdAdminsGroupId           = '00000000-0000-0000-0000-000000000012'
                sessionHostsDynamicGroupId = '00000000-0000-0000-0000-000000000013'
                breakGlassExcludeGroupId   = '00000000-0000-0000-0000-000000000014'
            }
            conditionalAccess  = @()
            entraKerberosReady = $true
        }
        $entraPayload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:entraManifestPath -Encoding utf8

        $mockState = [ordered]@{
            azEnvironment  = 'AzureUSGovernment'
            mgEnvironment  = 'USGov'
            subscribedSkus = @(
                [ordered]@{
                    skuPartNumber    = 'SPE_E5_USGOV_GCCHIGH'
                    capabilityStatus = 'Enabled'
                    servicePlans     = @(
                        @{ servicePlanName = 'SHAREPOINTENTERPRISE'; provisioningStatus = 'Success' },
                        @{ servicePlanName = 'TEAMS1'; provisioningStatus = 'Success' },
                        @{ servicePlanName = 'OFFICESUBSCRIPTION'; provisioningStatus = 'Success' }
                    )
                }
            )
            organization   = @(
                [ordered]@{
                    id              = '00000000-0000-0000-0000-000000000099'
                    displayName     = 'Contoso USGov Sandbox'
                    verifiedDomains = @(
                        @{ name = 'contoso.onmicrosoft.us'; isInitial = $true }
                    )
                }
            )
            siteRoot       = [ordered]@{
                id     = 'contoso,00000000-0000-0000-0000-000000000099,00000000-0000-0000-0000-000000000001'
                webUrl = 'https://contoso.sharepoint.us'
            }
            configurationPolicies = @()
            requests              = @()
        }
        $mockState | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:mockGraphPath -Encoding utf8
        $env:AVD_OPS_MOCK_GRAPH_PATH = $script:mockGraphPath
    }

    AfterEach {
        Remove-Item Env:\AVD_OPS_MOCK_GRAPH_PATH -ErrorAction SilentlyContinue
    }

    It 'routes all Graph calls to https://graph.microsoft.us, validates .sharepoint.us KFM, and sets WebRtc Teams mode in usgov' {
        & $script:deployScript `
            -EntraManifestPath $script:entraManifestPath `
            -BaselinePath $script:baselinePath `
            -OutputPath $script:m365ManifestPath `
            -RunManifestPath $script:runManifestPath `
            -Cloud 'usgov' `
            -Confirm:$false

        Test-Path -LiteralPath $script:m365ManifestPath | Should -BeTrue
        $manifest = Get-Content -LiteralPath $script:m365ManifestPath -Raw -Encoding utf8 | ConvertFrom-Json
        $manifest.cloud | Should -Be 'usgov'
        $manifest.sharePointSuffix | Should -Be '.sharepoint.us'
        $manifest.teamsVdiMode | Should -Be 'WebRtc'
        $manifest.oneDriveKfmEnabled | Should -BeTrue
        $manifest.teamsVdiOptimizationEnabled | Should -BeTrue
        $manifest.sharedComputerActivationEnabled | Should -BeTrue

        $stateAfter = Get-Content -LiteralPath $script:mockGraphPath -Raw -Encoding utf8 | ConvertFrom-Json
        $stateAfter.requests.Count | Should -BeGreaterThan 0
        foreach ($req in $stateAfter.requests) {
            $req.uri | Should -Match '^https://graph\.microsoft\.us/'
            $req.uri | Should -Not -Match 'graph\.microsoft\.com'
        }

        $oneDrivePol = $stateAfter.configurationPolicies | Where-Object { $_.name -like '*OneDrive*' } | Select-Object -First 1
        $oneDrivePol | Should -Not -BeNullOrEmpty
        ($oneDrivePol.settings | ConvertTo-Json -Depth 10) | Should -Match '\.sharepoint\.us'

        $teamsPol = $stateAfter.configurationPolicies | Where-Object { $_.name -like '*Teams VDI*' } | Select-Object -First 1
        $teamsPol | Should -Not -BeNullOrEmpty
        ($teamsPol.settings | ConvertTo-Json -Depth 10) | Should -Match 'WebRtc'
        ($teamsPol.settings | ConvertTo-Json -Depth 10) | Should -Not -Match 'SlimCoreAndWebRtc'
    }

    It 'fails closed before issuing any Graph request when entra-manifest.json cloud mismatches -Cloud' {
        $commercialEntra = Get-Content -LiteralPath $script:entraManifestPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
        $commercialEntra['cloud'] = 'commercial'
        $commercialEntra | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:entraManifestPath -Encoding utf8

        {
            & $script:deployScript `
                -EntraManifestPath $script:entraManifestPath `
                -BaselinePath $script:baselinePath `
                -OutputPath $script:m365ManifestPath `
                -RunManifestPath $script:runManifestPath `
                -Cloud 'usgov' `
                -Confirm:$false
        } | Should -Throw '*Cloud boundary mismatch*'

        $stateAfter = Get-Content -LiteralPath $script:mockGraphPath -Raw -Encoding utf8 | ConvertFrom-Json
        @($stateAfter.requests).Count | Should -Be 0
    }

    It 'fails closed before issuing any Graph request when MgGraph environment is Global while -Cloud is usgov' {
        $badMock = Get-Content -LiteralPath $script:mockGraphPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
        $badMock['mgEnvironment'] = 'Global'
        $badMock | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:mockGraphPath -Encoding utf8

        {
            & $script:deployScript `
                -EntraManifestPath $script:entraManifestPath `
                -BaselinePath $script:baselinePath `
                -OutputPath $script:m365ManifestPath `
                -RunManifestPath $script:runManifestPath `
                -Cloud 'usgov' `
                -Confirm:$false
        } | Should -Throw '*Cloud boundary mismatch*MgGraph*'

        $stateAfter = Get-Content -LiteralPath $script:mockGraphPath -Raw -Encoding utf8 | ConvertFrom-Json
        @($stateAfter.requests).Count | Should -Be 0
    }

    It 'deletes usgov policies via https://graph.microsoft.us in Remove-AvdM365.ps1 and fails closed on cloud mismatch' {
        & $script:deployScript `
            -EntraManifestPath $script:entraManifestPath `
            -BaselinePath $script:baselinePath `
            -OutputPath $script:m365ManifestPath `
            -RunManifestPath $script:runManifestPath `
            -Cloud 'usgov' `
            -Confirm:$false

        {
            & $script:removeScript `
                -RunManifestPath $script:runManifestPath `
                -Cloud 'commercial' `
                -DiscardLastCopy `
                -Confirm:$false
        } | Should -Throw '*Cloud boundary mismatch*'

        Test-Path -LiteralPath $script:runManifestPath | Should -BeTrue

        & $script:removeScript `
            -RunManifestPath $script:runManifestPath `
            -Cloud 'usgov' `
            -DiscardLastCopy `
            -Confirm:$false

        Test-Path -LiteralPath $script:runManifestPath | Should -BeFalse
        $stateAfter = Get-Content -LiteralPath $script:mockGraphPath -Raw -Encoding utf8 | ConvertFrom-Json
        $deleteCalls = @($stateAfter.requests | Where-Object { $_.method -eq 'DELETE' })
        $deleteCalls.Count | Should -Be 3
        foreach ($del in $deleteCalls) {
            $del.uri | Should -Match '^https://graph\.microsoft\.us/'
        }
    }
}
