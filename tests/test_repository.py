from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "ci.yml"
CHECKOUT_PIN = "34e114876b0b11c390a56381ad16ebd13914f8d5"
SETUP_PYTHON_PIN = "a26af69be951a213d495a4c3e4e4022e16d87065"


def assert_ci_workflow_contract(workflow: str) -> None:
    required = (
        "name: CI",
        "permissions:\n  contents: read",
        "timeout-minutes: 15",
        (
            "concurrency:\n"
            "  group: cafe-ci-${{ github.workflow }}-${{ github.ref }}\n"
            "  cancel-in-progress: true"
        ),
        "persist-credentials: false",
        "python-version: '3.12'",
        "python -m pip install -r requirements-dev.txt pip-audit",
        "python -m pip check",
        "pip-audit -r requirements-dev.txt",
        "python -m unittest discover -v",
    )
    for snippet in required:
        if snippet not in workflow:
            raise AssertionError(f"CI workflow is missing required contract: {snippet}")

    if workflow.count("\non:\n") != 1 or workflow.count("\npermissions:") != 1:
        raise AssertionError("CI workflow must declare one explicit trigger block")
    trigger_block = workflow.split("\non:\n", 1)[1].split("\npermissions:", 1)[0]
    if trigger_block.splitlines() != ["  push:", "  pull_request:"]:
        raise AssertionError("CI workflow must run only on push and pull_request")

    if workflow.count("permissions:") != 1 or workflow.count("concurrency:") != 1:
        raise AssertionError("CI workflow must declare one permission and concurrency block")
    permission_block = workflow.split("\npermissions:\n", 1)[1].split(
        "\nconcurrency:", 1
    )[0]
    if permission_block.splitlines() != ["  contents: read"]:
        raise AssertionError("CI workflow permissions must remain read-only")
    if workflow.count("timeout-minutes:") != 1:
        raise AssertionError("Every CI job must retain its explicit timeout")
    if workflow.count("persist-credentials: false") != 1:
        raise AssertionError("Checkout credentials must not persist")

    lowered = workflow.casefold()
    forbidden = (
        "env:",
        "services:",
        "continue-on-error",
        "|| true",
        "set +e",
        "if: always()",
        "write-all",
        "microsoft excel",
        "excel.exe",
        "libreoffice",
        "soffice",
        "powershell",
        "pwsh",
        "osascript",
        "run macro",
        "run vba",
    )
    for snippet in forbidden:
        if snippet in lowered:
            raise AssertionError(f"CI workflow contains forbidden behavior: {snippet}")
    if re.search(r"\bsecrets\s*(?:\.|\[)", lowered):
        raise AssertionError("CI workflow must not reference GitHub secrets")

    expected_actions = {
        "actions/checkout": (CHECKOUT_PIN, "v4.3.1"),
        "actions/setup-python": (SETUP_PYTHON_PIN, "v5.6.0"),
    }
    action_refs = re.findall(
        r"uses:\s*([^@\s]+)@([0-9a-f]{40})\s+#\s+([^\s]+)", workflow
    )
    if workflow.count("uses:") != len(action_refs):
        raise AssertionError("CI workflow actions must use immutable SHA references")
    if len(action_refs) != len(expected_actions):
        raise AssertionError("CI workflow must use exactly the approved official actions")
    for action, ref, tag in action_refs:
        if expected_actions.get(action) != (ref, tag):
            raise AssertionError(f"CI action must use its approved immutable SHA: {action}")


class RepositoryPresentationTests(unittest.TestCase):
    def test_ci_checks_workbook_and_vba_without_office_automation(self):
        workflow = WORKFLOW_PATH.read_text(encoding="utf-8")
        assert_ci_workflow_contract(workflow)

        hostile_variants = (
            workflow.replace(CHECKOUT_PIN, "v4", 1),
            workflow.replace(SETUP_PYTHON_PIN, "v5.6.0", 1),
            workflow.replace("cancel-in-progress: true", "cancel-in-progress: false", 1),
            workflow.replace(
                "group: cafe-ci-${{ github.workflow }}-${{ github.ref }}",
                "group:",
                1,
            ),
            workflow.replace("  pull_request:\n", "  workflow_dispatch:\n", 1),
            workflow.replace("  push:\n", "", 1),
            workflow.replace("  pull_request:\n", "", 1),
            workflow.replace(
                "      - name: Test workbook contracts\n",
                "      - name: Read a secret\n"
                "        run: echo \"${{ secrets['CHECKOUT_TOKEN'] }}\"\n"
                "      - name: Test workbook contracts\n",
                1,
            ),
            workflow
            + "\nservices:\n  office:\n    image: libreoffice:latest\n",
            workflow.replace(
                "python -m unittest discover -v",
                "python -m unittest discover -v || true",
                1,
            ),
        )
        for hostile_workflow in hostile_variants:
            with self.subTest(hostile_workflow=hostile_workflow):
                with self.assertRaises(AssertionError):
                    assert_ci_workflow_contract(hostile_workflow)

    def test_readme_explains_the_portfolio_project(self):
        readme_path = ROOT / "README.md"
        self.assertTrue(readme_path.exists(), "README.md should exist")
        readme = readme_path.read_text(encoding="utf-8")

        required_copy = (
            "新店・文山・大安咖啡廳推薦系統",
            "Taipei Cafe Recommender",
            "問題與目標",
            "核心功能",
            "技術架構",
            "使用方式",
            "安全性",
            "限制與資料說明",
            "2026-04",
            "通勤時間",
            "Hunter20041004",
        )
        for text in required_copy:
            with self.subTest(text=text):
                self.assertIn(text, readme)

    def test_data_notice_limits_public_snapshot_claims(self):
        notice_path = ROOT / "DATA_NOTICE.md"
        self.assertTrue(notice_path.exists(), "DATA_NOTICE.md should exist")
        notice = notice_path.read_text(encoding="utf-8")

        self.assertIn("2026-04", notice)
        self.assertIn("Google Maps", notice)
        self.assertIn("即時", notice)
        self.assertIn("不隸屬", notice)

    def test_repository_declares_code_license_and_build_dependency(self):
        license_path = ROOT / "LICENSE"
        requirements_path = ROOT / "requirements-dev.txt"
        self.assertTrue(license_path.exists(), "LICENSE should exist")
        self.assertTrue(requirements_path.exists(), "requirements-dev.txt should exist")

        self.assertIn("MIT License", license_path.read_text(encoding="utf-8"))
        self.assertEqual(
            "pyOpenVBA==3.0.1",
            requirements_path.read_text(encoding="utf-8").strip(),
        )

    def test_readme_screenshots_are_real_png_artifacts(self):
        for name in ("dashboard.png", "analysis.png"):
            with self.subTest(name=name):
                path = ROOT / "screenshots" / name
                self.assertTrue(path.exists(), f"{name} should exist")
                self.assertGreater(path.stat().st_size, 10_000)
                self.assertEqual(b"\x89PNG\r\n\x1a\n", path.read_bytes()[:8])


if __name__ == "__main__":
    unittest.main()
