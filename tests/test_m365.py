#!/usr/bin/env python3
"""Offline unit and integration tests for avd-m365 scripts and contracts."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import validate  # noqa: E402


def find_pwsh() -> str:
    """Resolve PowerShell 7 executable."""
    for candidate in ("pwsh", "pwsh.exe"):
        path = shutil.which(candidate)
        if path:
            return path
    raise RuntimeError("pwsh not found on PATH")


def find_bash() -> str | None:
    """Resolve Git Bash on Windows or bash on Linux/macOS."""
    win_git_bash = Path(r"C:\Program Files\Git\bin\bash.exe")
    if win_git_bash.is_file():
        return str(win_git_bash)
    return shutil.which("bash")


class TestM365ControlPlane(unittest.TestCase):
    """Test suite for avd-m365 validation, idempotent CRUD, SKU checks, and teardown."""

    def setUp(self) -> None:
        self.pwsh = find_pwsh()
        self.temp_dir = Path(tempfile.mkdtemp(prefix="avd-m365-test-"))
        self.entra_manifest_path = self.temp_dir / "entra-manifest.json"
        self.m365_manifest_path = self.temp_dir / "m365-manifest.json"
        self.run_manifest_path = self.temp_dir / "m365-run-manifest.json"
        self.mock_graph_path = self.temp_dir / "mock-graph.json"

        entra_payload = {
            "schema": "avd-ops-entra-manifest/v1",
            "environment": "dev",
            "prefix": "avd",
            "groups": {
                "avdUsersGroupId": "00000000-0000-0000-0000-000000000011",
                "avdAdminsGroupId": "00000000-0000-0000-0000-000000000012",
                "sessionHostsDynamicGroupId": "00000000-0000-0000-0000-000000000013",
                "breakGlassExcludeGroupId": "00000000-0000-0000-0000-000000000014",
            },
            "conditionalAccess": [],
            "entraKerberosReady": True,
        }
        self.entra_manifest_path.write_text(json.dumps(entra_payload, indent=2), encoding="utf-8")

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_repo_and_schema_validation(self) -> None:
        """Verify repository JSON files, baseline rules, and contract validator."""
        errors = validate.run_validation()
        self.assertEqual(errors, [])

        bad_manifest = self.temp_dir / "bad-manifest.json"
        bad_manifest.write_text(
            json.dumps(
                {
                    "schema": "wrong-schema",
                    "environment": "staging",
                    "oneDriveKfmEnabled": "yes",
                    "extraField": True,
                }
            ),
            encoding="utf-8",
        )
        schema = json.loads((ROOT / "contracts" / "m365-manifest.schema.json").read_text(encoding="utf-8"))
        manifest_errors: list[str] = []
        validate.validate_m365_manifest(bad_manifest, schema, manifest_errors)
        self.assertTrue(any("missing required property" in e for e in manifest_errors))
        self.assertTrue(any("additional property 'extraField'" in e for e in manifest_errors))
        self.assertTrue(any("must equal 'avd-ops-m365-manifest/v1'" in e for e in manifest_errors))

    def test_preflight_sku_check_rejects_missing_service_plans(self) -> None:
        """Verify Deploy-AvdM365.ps1 fails in preflight when required M365 SKUs are absent."""
        mock_state = {
            "subscribedSkus": [
                {
                    "skuPartNumber": "FLOW_FREE",
                    "capabilityStatus": "Enabled",
                    "servicePlans": [
                        {"servicePlanName": "FLOW_P2_VIRAL", "provisioningStatus": "Success"}
                    ],
                }
            ],
            "organization": [{"id": "00000000-0000-0000-0000-000000000099", "displayName": "Contoso"}],
            "configurationPolicies": [],
        }
        self.mock_graph_path.write_text(json.dumps(mock_state, indent=2), encoding="utf-8")

        env = os.environ.copy()
        env["AVD_OPS_MOCK_GRAPH_PATH"] = str(self.mock_graph_path)

        proc = subprocess.run(
            [
                self.pwsh,
                "-NoProfile",
                "-File",
                str(ROOT / "scripts" / "Deploy-AvdM365.ps1"),
                "-EntraManifestPath",
                str(self.entra_manifest_path),
                "-OutputPath",
                str(self.m365_manifest_path),
                "-RunManifestPath",
                str(self.run_manifest_path),
                "-Confirm:$false",
            ],
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("Preflight SKU check failed", proc.stderr + proc.stdout)
        self.assertFalse(self.m365_manifest_path.exists())

    def test_idempotent_create_and_update_workflow(self) -> None:
        """Verify first run creates stamped policies and second run updates them idempotently."""
        tenant_id = "00000000-0000-0000-0000-000000000099"
        mock_state = {
            "subscribedSkus": [
                {
                    "skuPartNumber": "SPE_E5",
                    "capabilityStatus": "Enabled",
                    "servicePlans": [
                        {"servicePlanName": "SHAREPOINTENTERPRISE", "provisioningStatus": "Success"},
                        {"servicePlanName": "TEAMS1", "provisioningStatus": "Success"},
                        {"servicePlanName": "OFFICESUBSCRIPTION", "provisioningStatus": "Success"},
                    ],
                }
            ],
            "organization": [{"id": tenant_id, "displayName": "Contoso Sandbox"}],
            "configurationPolicies": [],
        }
        self.mock_graph_path.write_text(json.dumps(mock_state, indent=2), encoding="utf-8")

        env = os.environ.copy()
        env["AVD_OPS_MOCK_GRAPH_PATH"] = str(self.mock_graph_path)

        cmd = [
            self.pwsh,
            "-NoProfile",
            "-File",
            str(ROOT / "scripts" / "Deploy-AvdM365.ps1"),
            "-EntraManifestPath",
            str(self.entra_manifest_path),
            "-OutputPath",
            str(self.m365_manifest_path),
            "-RunManifestPath",
            str(self.run_manifest_path),
            "-Confirm:$false",
        ]

        first = subprocess.run(cmd, env=env, capture_output=True, text=True, check=False)
        self.assertEqual(first.returncode, 0, msg=first.stderr + first.stdout)

        run_manifest_1 = json.loads(self.run_manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(run_manifest_1["stamp"], "[avd-ops-avd-dev]")
        self.assertEqual(run_manifest_1["tenantId"], tenant_id)
        self.assertEqual(len(run_manifest_1["resources"]), 3)
        for res in run_manifest_1["resources"]:
            self.assertTrue(res["created"])
            self.assertEqual(res["operation"], "created")
            self.assertTrue(res["displayName"].startswith("[avd-ops-avd-dev]"))

        m365_manifest = json.loads(self.m365_manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(
            m365_manifest,
            {
                "schema": "avd-ops-m365-manifest/v1",
                "environment": "dev",
                "oneDriveKfmEnabled": True,
                "teamsVdiOptimizationEnabled": True,
                "sharedComputerActivationEnabled": True,
            },
        )
        self.assertEqual(validate.run_validation(manifest_paths=[self.m365_manifest_path]), [])

        # Verify dynamic tenant ID was embedded in the OneDrive KFM policy settings
        state_after_first = json.loads(self.mock_graph_path.read_text(encoding="utf-8"))
        self.assertEqual(len(state_after_first["configurationPolicies"]), 3)
        onedrive_pol = next(
            p for p in state_after_first["configurationPolicies"] if "OneDrive" in p["name"]
        )
        self.assertIn(tenant_id, json.dumps(onedrive_pol["settings"]))

        # Second run against existing Graph policies should PATCH (update) and mark created=False
        second = subprocess.run(cmd, env=env, capture_output=True, text=True, check=False)
        self.assertEqual(second.returncode, 0, msg=second.stderr + second.stdout)

        run_manifest_2 = json.loads(self.run_manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(len(run_manifest_2["resources"]), 3)
        for res in run_manifest_2["resources"]:
            self.assertFalse(res["created"])
            self.assertEqual(res["operation"], "updated")

        state_after_second = json.loads(self.mock_graph_path.read_text(encoding="utf-8"))
        self.assertEqual(len(state_after_second["configurationPolicies"]), 3)
        patch_calls = [
            r for r in state_after_second["requests"] if r["method"] == "PATCH"
        ]
        self.assertEqual(len(patch_calls), 3)

    def test_remove_enforces_discard_last_copy_and_selective_deletion(self) -> None:
        """Verify Remove-AvdM365.ps1 requires -DiscardLastCopy and only deletes created stamped policies."""
        mock_state = {
            "configurationPolicies": [
                {"id": "00000000-0000-0000-0000-000000000001", "name": "[avd-ops-avd-dev] M365 OneDrive VDI & KFM Baseline"},
                {"id": "00000000-0000-0000-0000-000000000002", "name": "[avd-ops-avd-dev] M365 Teams VDI Media Optimization"},
                {"id": "00000000-0000-0000-0000-000000000003", "name": "[preexisting] Corporate Policy"},
            ],
            "requests": [],
        }
        self.mock_graph_path.write_text(json.dumps(mock_state, indent=2), encoding="utf-8")

        run_manifest = {
            "schema": "avd-ops-m365-run-manifest/v1",
            "environment": "dev",
            "prefix": "avd",
            "stamp": "[avd-ops-avd-dev]",
            "resources": [
                {
                    "id": "00000000-0000-0000-0000-000000000001",
                    "displayName": "[avd-ops-avd-dev] M365 OneDrive VDI & KFM Baseline",
                    "resourceType": "deviceManagementConfigurationPolicy",
                    "uri": "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/00000000-0000-0000-0000-000000000001",
                    "created": True,
                },
                {
                    "id": "00000000-0000-0000-0000-000000000002",
                    "displayName": "[avd-ops-avd-dev] M365 Teams VDI Media Optimization",
                    "resourceType": "deviceManagementConfigurationPolicy",
                    "uri": "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/00000000-0000-0000-0000-000000000002",
                    "created": False,
                },
                {
                    "id": "00000000-0000-0000-0000-000000000003",
                    "displayName": "[preexisting] Corporate Policy",
                    "resourceType": "deviceManagementConfigurationPolicy",
                    "uri": "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/00000000-0000-0000-0000-000000000003",
                    "created": True,
                },
            ],
        }
        self.run_manifest_path.write_text(json.dumps(run_manifest, indent=2), encoding="utf-8")

        env = os.environ.copy()
        env["AVD_OPS_MOCK_GRAPH_PATH"] = str(self.mock_graph_path)

        # Without -DiscardLastCopy -> must fail and preserve run manifest
        without_flag = subprocess.run(
            [
                self.pwsh,
                "-NoProfile",
                "-File",
                str(ROOT / "scripts" / "Remove-AvdM365.ps1"),
                "-RunManifestPath",
                str(self.run_manifest_path),
                "-Confirm:$false",
            ],
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(without_flag.returncode, 0)
        self.assertIn("-DiscardLastCopy", without_flag.stderr + without_flag.stdout)
        self.assertTrue(self.run_manifest_path.exists())

        # With -DiscardLastCopy -> must delete only resource 1 and remove run manifest
        with_flag = subprocess.run(
            [
                self.pwsh,
                "-NoProfile",
                "-File",
                str(ROOT / "scripts" / "Remove-AvdM365.ps1"),
                "-RunManifestPath",
                str(self.run_manifest_path),
                "-DiscardLastCopy",
                "-Confirm:$false",
            ],
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(with_flag.returncode, 0, msg=with_flag.stderr + with_flag.stdout)
        self.assertFalse(self.run_manifest_path.exists())

        final_state = json.loads(self.mock_graph_path.read_text(encoding="utf-8"))
        remaining_ids = [p["id"] for p in final_state["configurationPolicies"]]
        self.assertEqual(
            remaining_ids,
            [
                "00000000-0000-0000-0000-000000000002",
                "00000000-0000-0000-0000-000000000003",
            ],
        )
        delete_calls = [r for r in final_state["requests"] if r["method"] == "DELETE"]
        self.assertEqual(len(delete_calls), 1)
        self.assertTrue(delete_calls[0]["uri"].endswith("00000000-0000-0000-0000-000000000001"))

    def test_bash_wrapper_dry_run(self) -> None:
        """Verify deploy-m365.sh executes Deploy-AvdM365.ps1 in dry-run mode and validates output."""
        bash_exe = find_bash()
        if not bash_exe:
            self.skipTest("bash not available")

        env = os.environ.copy()
        env["AVD_OPS_DRY_RUN"] = "1"

        proc = subprocess.run(
            [
                bash_exe,
                str(ROOT / "scripts" / "deploy-m365.sh"),
                str(self.entra_manifest_path),
                str(self.m365_manifest_path),
                str(ROOT / "baselines" / "vdi-workplace-baseline.json"),
                str(self.run_manifest_path),
            ],
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
        self.assertTrue(self.m365_manifest_path.is_file())
        self.assertTrue(self.run_manifest_path.is_file())


if __name__ == "__main__":
    unittest.main()
