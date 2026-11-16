# avd-m365

Part of [avd-ops](https://avd-ops.dev/), the open toolkit for Azure Virtual Desktop.

Greenfield Microsoft 365 workplace and VDI productivity control plane (`avd-m365`), orchestrated by [avd-landing-zone](https://github.com/avd-ops/avd-landing-zone#readme) ([avd-ops docs](https://avd-ops.dev/docs/repos/avd-m365.html), [decision 0009](https://avd-ops.dev/docs/decisions/0009-split-landing-zone-into-five-control-planes.html)).

> [!WARNING]
> **Community research project. Use at your own risk. No warranty.**
>
> - This is my own personal project, built on my own time **for research and learning**. It is **not affiliated with, endorsed by or supported by my employer**, Liquidware or Microsoft, and it does not represent their views, products or guidance.
> - Test it first in a dedicated lab subscription and tenant before using it in production; you are responsible for reviewing the code and validating it for your environment.
> - It is provided **"as is", without warranty of any kind**, under the [MIT License](LICENSE). You use it at your own risk.
> - **Contributions and forks are welcome and encouraged.** See [Contributing](#contributing).

## How it connects

| Direction | Repo | Relationship |
|---|---|---|
| Front door | [avd-ops](https://github.com/avd-ops/avd-ops#readme) | Overview, [repo page](https://avd-ops.dev/docs/repos/avd-m365.html), and platform decisions |
| Orchestrator | [avd-landing-zone](https://github.com/avd-ops/avd-landing-zone#readme) | Invokes `avd-m365` in the multi-plane deployment pipeline |
| Upstream | [avd-entra](https://github.com/avd-ops/avd-entra#readme) | Supplies `entra-manifest.json` user group targets |
| Assess & remediate | [avd-assess](https://github.com/avd-ops/avd-assess#readme) / [avd-remediate](https://github.com/avd-ops/avd-remediate#readme) | M365 VDI standards (`methods/m365.md`) and `C24` remediation plane |

## What it does

- Configures OneDrive for Business Silent Account Configuration, Known Folder Move (KFM), and VDI sync baselines.
- Configures SharePoint tenant settings (external sharing guardrails, idle session sign-out, and VDI sync file exclusions).
- Applies Microsoft Teams VDI media optimization policies and meeting baselines.
- Configures Microsoft 365 Apps Shared Computer Activation and Exchange Online / Outlook caching defaults for FSLogix containers.
- Verifies required M365 tenant SKUs in preflight without purchasing recurring commitments.
- Driven from Azure Cloud Shell (PowerShell 7, Bash, Python and Microsoft Graph REST).

## Contracts

Consumes `entra-manifest.json` user group targets and produces `m365-manifest.json` (OneDrive, Teams VDI, and M365 Apps readiness facts).

## Docs

- Agent instructions & guardrails: [AGENTS.md](AGENTS.md)
- Repo page on avd-ops: [avd-ops.dev/docs/repos/avd-m365](https://avd-ops.dev/docs/repos/avd-m365.html)

## Contributing

Fork it, open an issue or a pull request. Follow the family guardrails in [AGENTS.md](AGENTS.md): no tenant IDs, secrets, domains or customer names, and propose changes via pull request only. Contributions are accepted under the [MIT License](LICENSE).

## License and disclaimer

[MIT](LICENSE). Community research project, provided as is, use at your own risk, not affiliated with the author's employer, Liquidware or Microsoft.
