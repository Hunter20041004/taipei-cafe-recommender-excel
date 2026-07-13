from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class RepositoryPresentationTests(unittest.TestCase):
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
