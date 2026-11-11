# Contracts for avd-m365

Consumes `entra-manifest.json` user group targets and `contracts/cloud-environments.json` (`C26`, `avd-cloud-environments/v1`) and produces `m365-manifest.json` (cloud profile `"commercial"` | `"usgov"`, OneDrive, Teams VDI, and M365 Apps readiness facts).

See [avd-ops decision 0009](https://avd-ops.dev/docs/decisions/0009-split-landing-zone-into-five-control-planes.html) and [avd-ops decision 0013](https://avd-ops.dev/docs/decisions/0013-azure-us-government-and-gcc-high-support.html).

