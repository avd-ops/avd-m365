#!/usr/bin/env bash
# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENTRA_MANIFEST="${1:-./entra-manifest.json}"
OUT_FILE="${2:-./m365-manifest.json}"
BASELINE_FILE="${3:-${SCRIPT_DIR}/../baselines/vdi-workplace-baseline.json}"
RUN_MANIFEST="${4:-./m365-run-manifest.json}"

PWSH_ARGS=(
  -NoProfile
  -File "${SCRIPT_DIR}/Deploy-AvdM365.ps1"
  -EntraManifestPath "${ENTRA_MANIFEST}"
  -BaselinePath "${BASELINE_FILE}"
  -OutputPath "${OUT_FILE}"
  -RunManifestPath "${RUN_MANIFEST}"
)

if [[ "${AVD_OPS_DRY_RUN:-0}" == "1" ]]; then
  PWSH_ARGS+=(-DryRun)
fi

pwsh "${PWSH_ARGS[@]}"

if command -v python3 >/dev/null 2>&1 && python3 --version >/dev/null 2>&1; then
  python3 "${SCRIPT_DIR}/validate.py" --manifest "${OUT_FILE}" --baseline "${BASELINE_FILE}"
elif command -v py >/dev/null 2>&1; then
  py -3 "${SCRIPT_DIR}/validate.py" --manifest "${OUT_FILE}" --baseline "${BASELINE_FILE}"
else
  python "${SCRIPT_DIR}/validate.py" --manifest "${OUT_FILE}" --baseline "${BASELINE_FILE}"
fi

echo "Configured and validated greenfield M365 OneDrive, Teams VDI and M365 Apps baselines (input: ${ENTRA_MANIFEST}, output: ${OUT_FILE})."
