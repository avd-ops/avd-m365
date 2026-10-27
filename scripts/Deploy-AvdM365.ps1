# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()][string]$EntraManifestPath = './entra-manifest.json',
    [Parameter()][string]$OutputPath = './m365-manifest.json'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($PSCmdlet.ShouldProcess('Microsoft 365', 'Deploy greenfield OneDrive, Teams VDI and M365 Apps baselines')) {
    & bash ./scripts/deploy-m365.sh $EntraManifestPath $OutputPath
}
