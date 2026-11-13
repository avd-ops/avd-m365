#!/usr/bin/env bash
# Personal project, not for production, as is; not affiliated with the author's employer, Liquidware or Microsoft.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CLOUD="${AVD_OPS_CLOUD:-commercial}"
POSITIONAL=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --cloud)
      if [[ $# -lt 2 ]]; then
        echo "Error: --cloud requires an argument ('commercial' or 'usgov')." >&2
        exit 1
      fi
      CLOUD="$2"
      shift 2
      ;;
    --cloud=*)
      CLOUD="${1#*=}"
      shift
      ;;
    *)
      POSITIONAL+=("$1")
      shift
      ;;
  esac
done

ENTRA_MANIFEST="${POSITIONAL[0]:-./entra-manifest.json}"
OUT_FILE="${POSITIONAL[1]:-./m365-manifest.json}"
BASELINE_FILE="${POSITIONAL[2]:-${SCRIPT_DIR}/../baselines/vdi-workplace-baseline.json}"
RUN_MANIFEST="${POSITIONAL[3]:-./m365-run-manifest.json}"
if [[ -n "${POSITIONAL[4]:-}" ]]; then
  CLOUD="${POSITIONAL[4]}"
fi

if [[ "${CLOUD}" != "commercial" && "${CLOUD}" != "usgov" ]]; then
  echo "Error: Invalid cloud '${CLOUD}'. Expected 'commercial' or 'usgov'." >&2
  exit 1
fi

PWSH_ARGS=(
  -NoProfile
  -File "${SCRIPT_DIR}/Deploy-AvdM365.ps1"
  -EntraManifestPath "${ENTRA_MANIFEST}"
  -BaselinePath "${BASELINE_FILE}"
  -OutputPath "${OUT_FILE}"
  -RunManifestPath "${RUN_MANIFEST}"
  -Cloud "${CLOUD}"
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

echo "Configured and validated greenfield M365 OneDrive, SharePoint, Teams VDI and M365 Apps baselines (cloud: ${CLOUD}, input: ${ENTRA_MANIFEST}, output: ${OUT_FILE})."
