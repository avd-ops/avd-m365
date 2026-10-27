# avd-m365

Part of [avd-ops](https://avd-ops.dev/), the open toolkit for Azure Virtual Desktop.

> [!WARNING]
> **Personal project. Research and learning only. Not for production. No warranty.**
>
> - This is my own personal project, built on my own time **for research and learning**. It is **not affiliated with, endorsed by or supported by my employer**, Liquidware or Microsoft, and it does not represent their views, products or guidance.
> - **Do not run any of it in production** or against subscriptions and tenants that matter. Test it first in a dedicated lab subscription and tenant.
> - It is provided **"as is", without warranty of any kind**, under the [MIT License](LICENSE). You use it at your own risk.
> - **Contributions and forks are welcome and encouraged.** See [Contributing](#contributing).

## Role in avd-ops

Greenfield Microsoft 365 workplace and VDI productivity control plane (`avd-m365`), orchestrated by [avd-landing-zone](https://github.com/avd-ops/avd-landing-zone) ([decision 0009](https://avd-ops.dev/docs/decisions/0009-split-landing-zone-into-five-control-planes.html)).

## What it does

- Configures OneDrive for Business Silent Account Configuration, Known Folder Move (KFM), and VDI sync baselines.
- Applies Microsoft Teams VDI media optimization policies and meeting baselines.
- Configures Microsoft 365 Apps Shared Computer Activation and Exchange Online / Outlook caching defaults for FSLogix containers.
- Verifies required M365 tenant SKUs in preflight without purchasing recurring commitments.
- Driven from Azure Cloud Shell (PowerShell 7, Bash, Python and Microsoft Graph REST).

## Contracts

Consumes `entra-manifest.json` user group targets and produces `m365-manifest.json` (OneDrive, Teams VDI, and M365 Apps readiness facts).

## Contributing

Fork it, open an issue or a pull request. Follow the family guardrails in [AGENTS.md](AGENTS.md): no tenant IDs, secrets, domains or customer names, and propose changes via pull request only. Contributions are accepted under the [MIT License](LICENSE).

## License and disclaimer

[MIT](LICENSE). Personal project, not for production, provided as is, not affiliated with the author's employer, Liquidware or Microsoft.
