#!/usr/bin/env python3
"""Safely update a linked deck skill checkout after an explicit opt-in.

This script intentionally has no import-time side effects and never installs a
global hook.  It only updates the checkout that contains this script.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Sequence


BRANCH = "main"
CHECK_INTERVAL_SECONDS = 60 * 60
TRUSTED_ORIGINS = {
    "https://github.com/So-Yul-e/deck-builder-skill.git",
    "git@github.com:So-Yul-e/deck-builder-skill.git",
}
CONFIG_NAME = ".deck-update.json"
LOCK_NAME = "deck-update.lock"


@dataclass(frozen=True)
class Checkout:
    skill_dir: Path
    repo_root: Path
    git_dir: Path
    origin_url: str
    branch: str


class UpdateError(RuntimeError):
    """A checkout is not eligible for an automated update."""


def _run(command: Sequence[str], cwd: Path, timeout: int = 20) -> str:
    completed = subprocess.run(
        list(command), cwd=str(cwd), text=True, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, timeout=timeout, check=False,
    )
    if completed.returncode:
        detail = completed.stderr.strip() or completed.stdout.strip() or "git command failed"
        raise UpdateError(detail)
    return completed.stdout.strip()


def _git(repo_root: Path, *args: str, timeout: int = 20) -> str:
    return _run(("git", *args), repo_root, timeout)


def _default_skill_dir() -> Path:
    return Path(__file__).resolve().parents[1]


def is_trusted_origin(url: str) -> bool:
    """Accept only the two documented official clone URLs."""
    return url.strip() in TRUSTED_ORIGINS


def locate_checkout(skill_dir: Path | None = None) -> Checkout:
    """Return a linked checkout, rejecting copied skill installations."""
    resolved_skill = (skill_dir or _default_skill_dir()).resolve()
    try:
        repo_root = Path(_git(resolved_skill, "rev-parse", "--show-toplevel")).resolve()
        git_dir_text = _git(resolved_skill, "rev-parse", "--git-dir")
        git_dir = Path(git_dir_text)
        if not git_dir.is_absolute():
            git_dir = (resolved_skill / git_dir).resolve()
        origin_url = _git(repo_root, "remote", "get-url", "origin")
        branch = _git(repo_root, "symbolic-ref", "--quiet", "--short", "HEAD")
    except UpdateError as error:
        raise UpdateError("연결된 Git 원본을 찾지 못했습니다: " + str(error)) from error

    # A real linked install lives at <checkout>/deck.  A copied deck directory
    # can have a Git parent elsewhere, but it must never update itself.
    if resolved_skill.parent != repo_root:
        raise UpdateError("복사된 설치본입니다. 원본 checkout에 연결된 deck 스킬에서만 업데이트할 수 있습니다.")
    return Checkout(resolved_skill, repo_root, git_dir, origin_url, branch)


def _config_path(skill_dir: Path) -> Path:
    return skill_dir / CONFIG_NAME


def load_config(skill_dir: Path) -> dict[str, object]:
    path = _config_path(skill_dir)
    if not path.exists():
        return {}
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise UpdateError(f"업데이트 설정을 읽을 수 없습니다: {error}") from error
    if not isinstance(loaded, dict):
        raise UpdateError("업데이트 설정 형식이 올바르지 않습니다.")
    return loaded


def save_config(skill_dir: Path, config: dict[str, object]) -> None:
    path = _config_path(skill_dir)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def _config_for(checkout: Checkout, enabled: bool, last_checked: float = 0) -> dict[str, object]:
    return {
        "repo_root": str(checkout.repo_root),
        "origin_url": checkout.origin_url,
        "branch": BRANCH,
        "enabled": enabled,
        "last_checked": last_checked,
    }


def enable(skill_dir: Path | None = None) -> str:
    checkout = locate_checkout(skill_dir)
    if not is_trusted_origin(checkout.origin_url):
        raise UpdateError("공식 deck-builder-skill origin이 아닙니다.")
    if checkout.branch != BRANCH:
        raise UpdateError("main checkout에서만 자동 업데이트를 켤 수 있습니다.")
    if _git(checkout.repo_root, "status", "--porcelain"):
        raise UpdateError("작업 트리가 변경되어 있습니다. 깨끗한 main checkout에서만 켤 수 있습니다.")
    save_config(checkout.skill_dir, _config_for(checkout, True))
    return "ENABLED 자동 업데이트가 켜졌습니다 (main, 공식 origin)."


def disable(skill_dir: Path | None = None) -> str:
    directory = (skill_dir or _default_skill_dir()).resolve()
    config = load_config(directory)
    config["enabled"] = False
    config.setdefault("branch", BRANCH)
    config.setdefault("last_checked", 0)
    save_config(directory, config)
    return "DISABLED 자동 업데이트가 꺼졌습니다."


def check(skill_dir: Path | None = None) -> str:
    """Read local state only; --check never fetches or changes files."""
    directory = (skill_dir or _default_skill_dir()).resolve()
    config = load_config(directory)
    try:
        checkout = locate_checkout(directory)
    except UpdateError as error:
        return f"SKIPPED {error}"
    enabled = bool(config.get("enabled", False))
    state = "enabled" if enabled else "disabled"
    origin = "trusted" if is_trusted_origin(checkout.origin_url) else "untrusted"
    return (
        f"STATUS {state}; branch={checkout.branch}; origin={origin}; "
        f"last_checked={config.get('last_checked', 0)}"
    )


def _eligible(checkout: Checkout, config: dict[str, object]) -> str | None:
    if not is_trusted_origin(checkout.origin_url):
        return "origin이 공식 URL과 다릅니다"
    if checkout.branch != BRANCH:
        return "main checkout이 아닙니다"
    if config.get("repo_root") != str(checkout.repo_root):
        return "설정의 checkout 경로가 달라졌습니다"
    if config.get("origin_url") != checkout.origin_url or config.get("branch") != BRANCH:
        return "설정의 origin 또는 branch가 변경되었습니다"
    if _git(checkout.repo_root, "status", "--porcelain"):
        return "작업 트리가 변경되어 있습니다"
    return None


def _acquire_lock(git_dir: Path) -> int | None:
    path = git_dir / LOCK_NAME
    try:
        return os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError:
        return None


def auto_update(
    skill_dir: Path | None = None,
    now: Callable[[], float] = time.time,
) -> str:
    """Fetch and fast-forward an opted-in, clean official main checkout only."""
    directory = (skill_dir or _default_skill_dir()).resolve()
    config = load_config(directory)
    if not bool(config.get("enabled", False)):
        return "SKIPPED 자동 업데이트가 꺼져 있습니다."
    try:
        checkout = locate_checkout(directory)
        reason = _eligible(checkout, config)
    except UpdateError as error:
        return f"WARNING 업데이트하지 않았습니다: {error}"
    if reason:
        return f"SKIPPED 업데이트하지 않았습니다: {reason}"

    current_time = now()
    last_checked = config.get("last_checked", 0)
    if isinstance(last_checked, (int, float)) and current_time - last_checked < CHECK_INTERVAL_SECONDS:
        return "SKIPPED 최근 1시간 안에 확인했습니다."

    lock_fd = _acquire_lock(checkout.git_dir)
    if lock_fd is None:
        return "SKIPPED 다른 업데이트 확인이 진행 중입니다."
    lock_path = checkout.git_dir / LOCK_NAME
    try:
        before = _git(checkout.repo_root, "rev-parse", "HEAD")
        try:
            _git(checkout.repo_root, "fetch", "origin", BRANCH, timeout=20)
        except (UpdateError, subprocess.TimeoutExpired) as error:
            return f"WARNING 네트워크 확인 실패: {error}"
        config["last_checked"] = current_time
        save_config(checkout.skill_dir, config)

        # Fetch can reveal a divergence.  Do not merge in that case.
        counts = _git(checkout.repo_root, "rev-list", "--left-right", "--count", f"HEAD...origin/{BRANCH}")
        try:
            ahead, behind = (int(value) for value in counts.split())
        except ValueError:
            return "WARNING 원격 비교 결과를 해석하지 못했습니다."
        if ahead:
            return "SKIPPED 로컬 checkout이 원격과 갈라져 있습니다."
        if not behind:
            return "OK 이미 최신 버전입니다."
        try:
            _git(checkout.repo_root, "merge", "--ff-only", f"origin/{BRANCH}")
        except UpdateError as error:
            return f"WARNING fast-forward 업데이트 실패: {error}"
        after = _git(checkout.repo_root, "rev-parse", "HEAD")
        if before != after:
            return "UPDATED 새 deck SKILL을 반영했습니다. 다음 요청부터 새 규칙을 읽습니다."
        return "OK 이미 최신 버전입니다."
    finally:
        os.close(lock_fd)
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="연결된 deck 스킬 checkout의 안전한 업데이트")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--enable", action="store_true", help="공식 main checkout의 자동 업데이트를 켭니다")
    group.add_argument("--disable", action="store_true", help="자동 업데이트를 끕니다")
    group.add_argument("--auto", action="store_true", help="opt-in된 경우에만 안전하게 업데이트합니다")
    args = parser.parse_args(argv)
    try:
        if args.enable:
            message = enable()
        elif args.disable:
            message = disable()
        elif args.auto:
            message = auto_update()
        else:
            message = check()
    except UpdateError as error:
        print(f"ERROR {error}")
        return 2
    print(message)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
