#!/usr/bin/env python3
"""Validate all JSON templates, contract schemas, manifests, and no-leak guardrails in this repository."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
SCHEMA_PATH = ROOT / "contracts" / "m365-manifest.schema.json"
DEFAULT_BASELINE_PATH = ROOT / "baselines" / "vdi-workplace-baseline.json"

EMAIL = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
ALLOWED_DOMAINS = ("example.com", "contoso.com", "users.noreply.github.com")
JWT_PATTERN = re.compile(r"\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b")


def check_text_leaks(text: str, label: str, errors: list[str]) -> None:
    """Ensure text does not contain disallowed email domains or raw bearer tokens."""
    for match in EMAIL.finditer(text):
        addr = match.group(0)
        if not any(addr.lower().endswith("@" + d) for d in ALLOWED_DOMAINS):
            errors.append(f"{label}: disallowed email address {addr}")
    if JWT_PATTERN.search(text):
        errors.append(f"{label}: potential raw JWT token detected")


def validate_baseline_data(data: Any, label: str, errors: list[str]) -> None:
    """Validate the structure and values of baselines/vdi-workplace-baseline.json."""
    if not isinstance(data, dict):
        errors.append(f"{label}: baseline root must be a JSON object")
        return

    name = data.get("name")
    if not isinstance(name, str) or not name.strip():
        errors.append(f"{label}: 'name' must be a non-empty string")

    onedrive = data.get("oneDrive")
    if not isinstance(onedrive, dict):
        errors.append(f"{label}: 'oneDrive' must be an object")
    else:
        if not isinstance(onedrive.get("silentAccountConfig"), bool):
            errors.append(f"{label}: 'oneDrive.silentAccountConfig' must be a boolean")
        kfm = onedrive.get("knownFolderMove")
        valid_folders = {"Desktop", "Documents", "Pictures"}
        if not isinstance(kfm, list) or not kfm or not all(isinstance(x, str) and x in valid_folders for x in kfm):
            errors.append(f"{label}: 'oneDrive.knownFolderMove' must be a non-empty list of {sorted(valid_folders)}")
        if not isinstance(onedrive.get("filesOnDemand"), bool):
            errors.append(f"{label}: 'oneDrive.filesOnDemand' must be a boolean")

    teams = data.get("teamsVdi")
    if not isinstance(teams, dict):
        errors.append(f"{label}: 'teamsVdi' must be an object")
    else:
        if teams.get("vdiMediaOptimization") not in ("SlimCoreAndWebRtc", "WebRtc"):
            errors.append(f"{label}: 'teamsVdi.vdiMediaOptimization' must be 'SlimCoreAndWebRtc' or 'WebRtc'")
        if not isinstance(teams.get("disableGpuHardwareAccelerationInSession"), bool):
            errors.append(f"{label}: 'teamsVdi.disableGpuHardwareAccelerationInSession' must be a boolean")

    apps = data.get("m365Apps")
    if not isinstance(apps, dict):
        errors.append(f"{label}: 'm365Apps' must be an object")
    else:
        if not isinstance(apps.get("sharedComputerLicensing"), bool):
            errors.append(f"{label}: 'm365Apps.sharedComputerLicensing' must be a boolean")
        if apps.get("updateChannel") not in ("MonthlyEnterprise", "Current", "SemiAnnual", "SemiAnnualPreview"):
            errors.append(f"{label}: 'm365Apps.updateChannel' must be a supported M365 Apps channel")
        months = apps.get("outlookCachedExchangeSyncMonths")
        if not isinstance(months, int) or isinstance(months, bool) or not (1 <= months <= 24):
            errors.append(f"{label}: 'm365Apps.outlookCachedExchangeSyncMonths' must be an integer between 1 and 24")


def validate_m365_manifest(manifest_path: Path, schema: dict[str, Any], errors: list[str]) -> None:
    """Validate an emitted m365-manifest.json against contracts/m365-manifest.schema.json."""
    label = str(manifest_path)
    if not manifest_path.is_file():
        errors.append(f"{label}: manifest file does not exist")
        return

    try:
        raw = manifest_path.read_text(encoding="utf-8")
        check_text_leaks(raw, label, errors)
        data = json.loads(raw)
    except Exception as exc:
        errors.append(f"{label}: invalid JSON: {exc}")
        return

    if not isinstance(data, dict):
        errors.append(f"{label}: manifest root must be an object")
        return

    required = schema.get("required", [])
    properties = schema.get("properties", {})
    additional_allowed = schema.get("additionalProperties", True)

    for req_key in required:
        if req_key not in data:
            errors.append(f"{label}: missing required property '{req_key}'")

    if additional_allowed is False:
        for key in data:
            if key not in properties:
                errors.append(f"{label}: additional property '{key}' is not allowed by schema")

    for key, prop_schema in properties.items():
        if key not in data:
            continue
        val = data[key]
        if "const" in prop_schema and val != prop_schema["const"]:
            errors.append(f"{label}: property '{key}' must equal '{prop_schema['const']}' (got {val!r})")
        expected_type = prop_schema.get("type")
        if expected_type == "string":
            if not isinstance(val, str):
                errors.append(f"{label}: property '{key}' must be a string")
            elif "enum" in prop_schema and val not in prop_schema["enum"]:
                errors.append(f"{label}: property '{key}' must be one of {prop_schema['enum']} (got {val!r})")
        elif expected_type == "boolean":
            if not isinstance(val, bool):
                errors.append(f"{label}: property '{key}' must be a boolean")


def run_validation(manifest_paths: list[Path] | None = None, baseline_path: Path | None = None) -> list[str]:
    """Run repository and manifest validations and return a list of error strings."""
    errors: list[str] = []

    for jf in ROOT.rglob("*.json"):
        if ".git" in jf.parts:
            continue
        try:
            data = json.loads(jf.read_text(encoding="utf-8"))
        except Exception as exc:
            errors.append(f"{jf.relative_to(ROOT)}: invalid JSON: {exc}")
            continue
        if not isinstance(data, (dict, list)):
            errors.append(f"{jf.relative_to(ROOT)}: root must be object or array")

    for tf in ROOT.rglob("*"):
        if ".git" in tf.parts or "__pycache__" in tf.parts or not tf.is_file():
            continue
        try:
            text = tf.read_text(encoding="utf-8")
        except Exception:
            continue
        check_text_leaks(text, str(tf.relative_to(ROOT)), errors)

    schema_data: dict[str, Any] = {}
    if not SCHEMA_PATH.is_file():
        errors.append("contracts/m365-manifest.schema.json: missing schema file")
    else:
        try:
            schema_data = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        except Exception as exc:
            errors.append(f"contracts/m365-manifest.schema.json: invalid JSON: {exc}")

    target_baseline = baseline_path or DEFAULT_BASELINE_PATH
    if not target_baseline.is_file():
        errors.append(f"{target_baseline}: missing baseline file")
    else:
        try:
            b_data = json.loads(target_baseline.read_text(encoding="utf-8"))
            validate_baseline_data(b_data, str(target_baseline), errors)
        except Exception as exc:
            errors.append(f"{target_baseline}: invalid JSON: {exc}")

    if manifest_paths and schema_data:
        for mp in manifest_paths:
            validate_m365_manifest(mp, schema_data, errors)

    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--manifest",
        action="append",
         default=[],
        help="Path to an m365-manifest.json file to validate against contracts/m365-manifest.schema.json",
    )
    parser.add_argument(
        "--baseline",
        help="Optional path to a workplace baseline JSON file to validate",
    )
    args = parser.parse_args(argv)

    manifest_paths = [Path(p) for p in args.manifest]
    baseline_path = Path(args.baseline) if args.baseline else None
    errors = run_validation(manifest_paths=manifest_paths, baseline_path=baseline_path)

    if errors:
        for err in errors:
            print(f"::error::{err}")
        return 1

    print("All JSON templates, schemas, and leak checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
