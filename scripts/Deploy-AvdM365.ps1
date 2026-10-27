# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()][string]$EntraManifestPath = './entra-manifest.json',
    [Parameter()][string]$BaselinePath = './baselines/vdi-workplace-baseline.json',
    [Parameter()][string]$OutputPath = './m365-manifest.json',
    [Parameter()][string]$RunManifestPath = './m365-run-manifest.json',
    [Parameter()][switch]$DryRun
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PythonCommand {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    foreach ($candidate in @('python3', 'python')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -ne $cmd -and $cmd.Source -notlike '*WindowsApps*') {
            return $cmd.Source
        }
    }
    $pyLauncher = Get-Command 'py' -ErrorAction SilentlyContinue
    if ($null -ne $pyLauncher) {
        return $pyLauncher.Source
    }
    $localPy = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
    if (Test-Path -LiteralPath $localPy) {
        return $localPy
    }
    throw 'Python 3 interpreter not found on PATH.'
}

function Invoke-GraphRequest {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable], [System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()][hashtable]$Body,
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
            method = $Method
            uri    = $Uri
            body   = $Body
        }

        $response = $null
        if ($Uri -like '*/v1.0/subscribedSkus*') {
            $skus = if ($mockState.ContainsKey('subscribedSkus')) { $mockState['subscribedSkus'] } else { @() }
            $response = @{ value = $skus }
        }
        elseif ($Uri -like '*/v1.0/organization*') {
            $orgs = if ($mockState.ContainsKey('organization')) {
                $mockState['organization']
            }
            else {
                @(@{ id = '00000000-0000-0000-0000-000000000000'; displayName = 'Contoso Sandbox' })
            }
            $response = @{ value = $orgs }
        }
        elseif ($Uri -like '*/beta/admin/microsoft365Apps/installationOptions*') {
            $opts = if ($mockState.ContainsKey('installationOptions')) {
                $mockState['installationOptions']
            }
            else {
                @{ updateChannel = 'monthlyEnterprise'; appsForWindows = @{ isMicrosoft365AppsEnabled = $true } }
            }
            $response = $opts
        }
        elseif ($Uri -like '*/beta/deviceManagement/configurationPolicies/*/assign') {
            $response = @{ status = 'assigned' }
        }
        elseif ($Uri -like '*/beta/deviceManagement/configurationPolicies*') {
            if (-not $mockState.ContainsKey('configurationPolicies')) {
                $mockState['configurationPolicies'] = @()
            }
            if ($Method -eq 'GET') {
                $response = @{ value = @($mockState['configurationPolicies']) }
            }
            elseif ($Method -eq 'POST') {
                $idx = @($mockState['configurationPolicies']).Count + 1
                $newId = ('00000000-0000-0000-0000-{0:d12}' -f $idx)
                $createdPolicy = [ordered]@{
                    id           = $newId
                    name         = $Body['name']
                    description  = $Body['description']
                    platforms    = $Body['platforms']
                    technologies = $Body['technologies']
                    settings     = $Body['settings']
                }
                $mockState['configurationPolicies'] = @($mockState['configurationPolicies']) + @($createdPolicy)
                $response = $createdPolicy
            }
            elseif ($Method -eq 'PATCH' -or $Method -eq 'PUT') {
                $policyId = ($Uri -split '/')[-1]
                $updatedList = @()
                $matchedPolicy = $null
                foreach ($pol in @($mockState['configurationPolicies'])) {
                    if ($pol['id'] -eq $policyId) {
                        $pol['name'] = $Body['name']
                        $pol['description'] = $Body['description']
                        $pol['settings'] = $Body['settings']
                        $matchedPolicy = $pol
                    }
                    $updatedList += $pol
                }
                $mockState['configurationPolicies'] = $updatedList
                $response = if ($null -ne $matchedPolicy) { $matchedPolicy } else { @{ id = $policyId } }
            }
        }

        $mockState | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $mockPath -Encoding utf8
        return $response
    }

    if ($IsDryRun) {
        if ($Uri -like '*/v1.0/subscribedSkus*') {
            return @{
                value = @(
                    @{
                        skuPartNumber    = 'SPE_E5'
                        capabilityStatus = 'Enabled'
                        servicePlans     = @(
                            @{ servicePlanName = 'SHAREPOINTENTERPRISE'; provisioningStatus = 'Success' },
                            @{ servicePlanName = 'TEAMS1'; provisioningStatus = 'Success' },
                            @{ servicePlanName = 'OFFICESUBSCRIPTION'; provisioningStatus = 'Success' }
                        )
                    }
                )
            }
        }
        if ($Uri -like '*/v1.0/organization*') {
            return @{
                value = @(
                    @{
                        id          = '00000000-0000-0000-0000-000000000000'
                        displayName = 'Contoso Lab Tenant'
                    }
                )
            }
        }
        if ($Uri -like '*/beta/admin/microsoft365Apps/installationOptions*') {
            return @{
                updateChannel  = 'monthlyEnterprise'
                appsForWindows = @{ isMicrosoft365AppsEnabled = $true }
            }
        }
        if ($Uri -like '*/beta/deviceManagement/configurationPolicies/*/assign') {
            return @{ status = 'assigned' }
        }
        if ($Uri -like '*/beta/deviceManagement/configurationPolicies*') {
            if ($Method -eq 'GET') {
                return @{ value = @() }
            }
            if ($Method -eq 'POST') {
                $hashBytes = [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes([string]$Body['name']))
                $suffix = ([System.BitConverter]::ToString($hashBytes[0..5])).Replace('-', '').ToLowerInvariant()
                return @{
                    id          = "00000000-0000-0000-0000-$suffix"
                    name        = $Body['name']
                    description = $Body['description']
                }
            }
            return @{
                id   = ($Uri -split '/')[-1]
                name = $Body['name']
            }
        }
        return @{}
    }

    $azArgs = @('rest', '--method', $Method, '--uri', $Uri, '--output', 'json')
    $tempFile = $null
    try {
        if ($null -ne $Body) {
            $tempFile = [System.IO.Path]::GetTempFileName()
            $jsonPayload = $Body | ConvertTo-Json -Depth 20 -Compress
            Set-Content -LiteralPath $tempFile -Value $jsonPayload -Encoding utf8 -NoNewline
            $azArgs += @('--headers', 'Content-Type=application/json', '--body', "@$tempFile")
        }
        $rawOut = & az @azArgs
        if ($LASTEXITCODE -ne 0) {
            throw "az rest $Method $Uri failed with exit code $LASTEXITCODE."
        }
        if ([string]::IsNullOrWhiteSpace(($rawOut -join ''))) {
            return @{}
        }
        return (($rawOut -join "`n") | ConvertFrom-Json -AsHashtable)
    }
    finally {
        if ($null -ne $tempFile -and (Test-Path -LiteralPath $tempFile)) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-M365SubscribedSku {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$IsDryRun
    )

    $skuResponse = Invoke-GraphRequest -Method 'GET' -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -IsDryRun $IsDryRun
    $enabledPlans = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($sku in @($skuResponse['value'])) {
        if ($null -eq $sku) {
            continue
        }
        $capStatus = [string]$sku['capabilityStatus']
        if ($capStatus -and $capStatus -notin @('Enabled', 'Warning')) {
            continue
        }
        foreach ($plan in @($sku['servicePlans'])) {
            if ($null -ne $plan -and [string]$plan['provisioningStatus'] -ne 'Disabled') {
                $planName = [string]$plan['servicePlanName']
                if (-not [string]::IsNullOrWhiteSpace($planName)) {
                    [void]$enabledPlans.Add($planName)
                }
            }
        }
    }

    $missingCapabilities = @()
    if (-not ($enabledPlans.Contains('SHAREPOINTENTERPRISE') -or $enabledPlans.Contains('SHAREPOINTWAC'))) {
        $missingCapabilities += 'OneDrive/SharePoint (SHAREPOINTENTERPRISE or SHAREPOINTWAC)'
    }
    if (-not ($enabledPlans.Contains('TEAMS1') -or $enabledPlans.Contains('MCOEV'))) {
        $missingCapabilities += 'Microsoft Teams (TEAMS1 or MCOEV)'
    }
    if (-not ($enabledPlans.Contains('OFFICESUBSCRIPTION') -or $enabledPlans.Contains('O365_BUSINESS'))) {
        $missingCapabilities += 'Microsoft 365 Apps (OFFICESUBSCRIPTION)'
    }

    if ($missingCapabilities.Count -gt 0) {
        throw "Preflight SKU check failed. Missing required M365 service plans in subscribedSkus (never auto-purchasing licenses): $($missingCapabilities -join '; ')."
    }
}

function Get-TenantOrganizationId {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][bool]$IsDryRun
    )

    $orgResponse = Invoke-GraphRequest -Method 'GET' -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName' -IsDryRun $IsDryRun
    $orgs = @($orgResponse['value'])
    if ($orgs.Count -eq 0 -or [string]::IsNullOrWhiteSpace([string]$orgs[0]['id'])) {
        throw 'Unable to resolve dynamic tenant ID from Microsoft Graph organization endpoint.'
    }
    return [string]$orgs[0]['id']
}

function ConvertTo-ConfigurationPolicyBody {
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][hashtable]$ConfigValues
    )

    $settingsList = @()
    foreach ($key in $ConfigValues.Keys) {
        $val = $ConfigValues[$key]
        $settingsList += [ordered]@{
            '@odata.type'   = '#microsoft.graph.deviceManagementConfigurationSetting'
            settingInstance = [ordered]@{
                '@odata.type'        = '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance'
                settingDefinitionId  = "device_vendor_msft_policy_config_$($key.ToLowerInvariant())"
                simpleSettingValue   = [ordered]@{
                    '@odata.type' = '#microsoft.graph.deviceManagementConfigurationStringSettingValue'
                    value         = ($val | ConvertTo-Json -Compress -Depth 5)
                }
            }
        }
    }

    return [ordered]@{
        name         = $Name
        description  = $Description
        platforms    = 'windows10'
        technologies = 'mdm'
        settings     = $settingsList
    }
}

$isDryRun = $DryRun.IsPresent -or ($env:AVD_OPS_DRY_RUN -eq '1')

if (-not (Test-Path -LiteralPath $EntraManifestPath)) {
    throw "Entra handoff manifest not found at '$EntraManifestPath'."
}
if (-not (Test-Path -LiteralPath $BaselinePath)) {
    throw "M365 workplace baseline file not found at '$BaselinePath'."
}

$entraManifest = Get-Content -LiteralPath $EntraManifestPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
$baseline = Get-Content -LiteralPath $BaselinePath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable

$environment = [string]$entraManifest['environment']
if ($environment -notin @('dev', 'test', 'prod')) {
    throw "Invalid or missing 'environment' in '$EntraManifestPath': '$environment'."
}

$prefix = if ($entraManifest.ContainsKey('prefix') -and -not [string]::IsNullOrWhiteSpace([string]$entraManifest['prefix'])) {
    [string]$entraManifest['prefix']
}
else {
    'avd'
}

$groups = $entraManifest['groups']
if ($null -eq $groups -or [string]::IsNullOrWhiteSpace([string]$groups['avdUsersGroupId']) -or [string]::IsNullOrWhiteSpace([string]$groups['sessionHostsDynamicGroupId'])) {
    throw "Entra manifest '$EntraManifestPath' must include groups.avdUsersGroupId and groups.sessionHostsDynamicGroupId."
}

$avdUsersGroupId = [string]$groups['avdUsersGroupId']
$sessionHostsGroupId = [string]$groups['sessionHostsDynamicGroupId']
$stamp = "[avd-ops-$prefix-$environment]"

if ($PSCmdlet.ShouldProcess("Microsoft 365 tenant ($stamp)", 'Verify SKUs and deploy OneDrive KFM, Teams VDI, and M365 Apps baselines')) {
    Test-M365SubscribedSku -IsDryRun $isDryRun
    $tenantId = Get-TenantOrganizationId -IsDryRun $isDryRun

    [void](Invoke-GraphRequest -Method 'GET' -Uri 'https://graph.microsoft.com/beta/admin/microsoft365Apps/installationOptions' -IsDryRun $isDryRun)

    $existingPoliciesResponse = Invoke-GraphRequest -Method 'GET' -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -IsDryRun $isDryRun
    $existingPolicies = @($existingPoliciesResponse['value'])

    $policySpecs = @(
        [ordered]@{
            pillar      = 'oneDrive'
            displayName = "$stamp M365 OneDrive VDI & KFM Baseline"
            description = "Managed by avd-m365 ($stamp). Configures OneDrive Silent Account Configuration, Known Folder Move (KFM), and Files On-Demand."
            values      = [ordered]@{
                SilentAccountConfig  = [bool]$baseline['oneDrive']['silentAccountConfig']
                KFMSilentOptIn       = [ordered]@{
                    TenantId = $tenantId
                    Folders  = @($baseline['oneDrive']['knownFolderMove'])
                }
                FilesOnDemandEnabled = [bool]$baseline['oneDrive']['filesOnDemand']
            }
            targets     = @($sessionHostsGroupId, $avdUsersGroupId)
        },
        [ordered]@{
            pillar      = 'teamsVdi'
            displayName = "$stamp M365 Teams VDI Media Optimization"
            description = "Managed by avd-m365 ($stamp). Configures Microsoft Teams VDI media optimization ($($baseline['teamsVdi']['vdiMediaOptimization'])) and GPU hardware acceleration."
            values      = [ordered]@{
                VdiMediaOptimization                    = [string]$baseline['teamsVdi']['vdiMediaOptimization']
                DisableGpuHardwareAccelerationInSession = [bool]$baseline['teamsVdi']['disableGpuHardwareAccelerationInSession']
            }
            targets     = @($avdUsersGroupId, $sessionHostsGroupId)
        },
        [ordered]@{
            pillar      = 'm365Apps'
            displayName = "$stamp M365 Apps Shared Computer Activation & Outlook VDI"
            description = "Managed by avd-m365 ($stamp). Configures M365 Apps Shared Computer Licensing, update channel, and Outlook Cached Exchange Mode sync window."
            values      = [ordered]@{
                SharedComputerLicensing         = if ([bool]$baseline['m365Apps']['sharedComputerLicensing']) { 1 } else { 0 }
                UpdateChannel                   = [string]$baseline['m365Apps']['updateChannel']
                OutlookCachedExchangeSyncMonths = [int]$baseline['m365Apps']['outlookCachedExchangeSyncMonths']
            }
            targets     = @($sessionHostsGroupId, $avdUsersGroupId)
        }
    )

    $runResources = @()
    foreach ($spec in $policySpecs) {
        $policyBody = ConvertTo-ConfigurationPolicyBody -Name $spec.displayName -Description $spec.description -ConfigValues $spec.values
        $matched = $null
        foreach ($item in $existingPolicies) {
            if ($null -ne $item -and [string]$item['name'] -eq $spec.displayName) {
                $matched = $item
                break
            }
        }

        if ($null -ne $matched) {
            $policyId = [string]$matched['id']
            $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId"
            [void](Invoke-GraphRequest -Method 'PATCH' -Uri $uri -Body $policyBody -IsDryRun $isDryRun)
            $wasCreated = $false
            $operation = 'updated'
        }
        else {
            $createUri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies'
            $createdPolicy = Invoke-GraphRequest -Method 'POST' -Uri $createUri -Body $policyBody -IsDryRun $isDryRun
            $policyId = [string]$createdPolicy['id']
            $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId"
            $wasCreated = $true
            $operation = 'created'
        }

        $uniqueGroups = @($spec.targets | Select-Object -Unique)
        $assignments = @()
        foreach ($gid in $uniqueGroups) {
            $assignments += [ordered]@{
                target = [ordered]@{
                    '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                    groupId       = $gid
                }
            }
        }
        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId/assign"
        [void](Invoke-GraphRequest -Method 'POST' -Uri $assignUri -Body @{ assignments = $assignments } -IsDryRun $isDryRun)

        $runResources += [ordered]@{
            id           = $policyId
            displayName  = $spec.displayName
            resourceType = 'deviceManagementConfigurationPolicy'
            pillar       = $spec.pillar
            uri          = $uri
            assignedTo   = $uniqueGroups
            created      = $wasCreated
            operation    = $operation
        }
    }

    $runManifest = [ordered]@{
        schema      = 'avd-ops-m365-run-manifest/v1'
        environment = $environment
        prefix      = $prefix
        stamp       = $stamp
        tenantId    = $tenantId
        resources   = $runResources
    }
    $runManifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $RunManifestPath -Encoding utf8

    $m365Manifest = [ordered]@{
        schema                          = 'avd-ops-m365-manifest/v1'
        environment                     = $environment
        oneDriveKfmEnabled              = [bool]($baseline['oneDrive']['silentAccountConfig'] -and (@($baseline['oneDrive']['knownFolderMove']).Count -gt 0))
        teamsVdiOptimizationEnabled     = [bool]($baseline['teamsVdi']['vdiMediaOptimization'] -eq 'SlimCoreAndWebRtc')
        sharedComputerActivationEnabled = [bool]$baseline['m365Apps']['sharedComputerLicensing']
    }
    $m365Manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutputPath -Encoding utf8

    $pythonExe = Get-PythonCommand
    $validateScript = Join-Path $PSScriptRoot 'validate.py'
    & $pythonExe $validateScript --manifest $OutputPath --baseline $BaselinePath
    if ($LASTEXITCODE -ne 0) {
        throw "Contract validation failed for emitted manifest '$OutputPath'."
    }
}
