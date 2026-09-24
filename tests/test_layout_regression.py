"""레이아웃 회귀 검사를 기본 검사 묶음에 연결한다.

검사 본체는 배포 폴더(deck/scripts/layout_regression.py)에 있다. 설치본에서 덱을
고치는 에이전트도 같은 스크립트를 직접 돌리게 하려고 tests/ 가 아니라 그쪽에 둔다.
"""
import subprocess
import sys
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "deck" / "scripts" / "layout_regression.py"


class LayoutRegressionTests(unittest.TestCase):
    def test_layout_regression_script_passes(self) -> None:
        result = subprocess.run(
            [sys.executable, str(SCRIPT)], capture_output=True, text=True, timeout=240
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("PASS: deck layout regression", result.stdout)


if __name__ == "__main__":
    unittest.main()
