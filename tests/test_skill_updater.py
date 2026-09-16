from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "deck" / "scripts"))
import update_skill  # noqa: E402


def git(path: Path, *args: str) -> str:
    return subprocess.run(("git", *args), cwd=path, check=True, text=True,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE).stdout.strip()


class SkillUpdaterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.origin = self.root / "origin.git"
        self.seed = self.root / "seed"
        self.clone = self.root / "clone"
        git(self.root, "init", "--bare", str(self.origin))
        git(self.root, "init", "-b", "main", str(self.seed))
        git(self.seed, "config", "user.email", "test@example.com")
        git(self.seed, "config", "user.name", "Test")
        (self.seed / "deck" / "scripts").mkdir(parents=True)
        (self.seed / "deck" / "SKILL.md").write_text("old", encoding="utf-8")
        (self.seed / "deck" / "scripts" / "update_skill.py").write_text("placeholder", encoding="utf-8")
        (self.seed / ".gitignore").write_text("deck/.deck-update.json\n", encoding="utf-8")
        git(self.seed, "add", ".")
        git(self.seed, "commit", "-m", "initial")
        git(self.seed, "remote", "add", "origin", str(self.origin))
        git(self.seed, "push", "-u", "origin", "main")
        git(self.root, "clone", "--branch", "main", str(self.origin), str(self.clone))
        self.skill = self.clone / "deck"
        self.trust = mock.patch.object(update_skill, "is_trusted_origin", return_value=True)
        self.trust.start()

    def tearDown(self) -> None:
        self.trust.stop()
        self.temp.cleanup()

    def _enable(self) -> None:
        self.assertIn("ENABLED", update_skill.enable(self.skill))

    def test_enable_requires_main_clean_linked_checkout(self) -> None:
        self._enable()
        config = update_skill.load_config(self.skill)
        self.assertTrue(config["enabled"])
        self.assertEqual("main", config["branch"])
        self.assertEqual(str(self.clone.resolve()), config["repo_root"])

    def test_auto_fast_forwards_and_reports_new_skill(self) -> None:
        self._enable()
        (self.seed / "deck" / "SKILL.md").write_text("new", encoding="utf-8")
        git(self.seed, "add", ".")
        git(self.seed, "commit", "-m", "update skill")
        git(self.seed, "push")
        message = update_skill.auto_update(self.skill, now=lambda: 10_000)
        self.assertIn("UPDATED", message)
        self.assertEqual("new", (self.skill / "SKILL.md").read_text(encoding="utf-8"))

    def test_auto_skips_dirty_and_disabled_without_fetching(self) -> None:
        self.assertIn("꺼져", update_skill.auto_update(self.skill))
        self._enable()
        (self.clone / "local-note.txt").write_text("dirty", encoding="utf-8")
        self.assertIn("작업 트리", update_skill.auto_update(self.skill, now=lambda: 10_000))

    def test_auto_skips_diverged_checkout(self) -> None:
        self._enable()
        (self.seed / "remote.txt").write_text("remote", encoding="utf-8")
        git(self.seed, "add", ".")
        git(self.seed, "commit", "-m", "remote")
        git(self.seed, "push")
        (self.clone / "local.txt").write_text("local", encoding="utf-8")
        git(self.clone, "add", ".")
        git(self.clone, "commit", "-m", "local")
        self.assertIn("갈라", update_skill.auto_update(self.skill, now=lambda: 10_000))

    def test_check_is_read_only_and_reports_origin_state(self) -> None:
        self.assertIn("STATUS disabled", update_skill.check(self.skill))


if __name__ == "__main__":
    unittest.main()
