#!/usr/bin/env python3
"""Release helper

Usage:
  python release.py [--dry-run]

Requirements: set environment variable GITHUB_TOKEN with a token that can create releases.
"""
import os
import re
import sys
import subprocess
import json
from datetime import datetime
from typing import List, Tuple, Optional

import requests


def run(cmd: List[str]) -> str:
    return subprocess.check_output(cmd, stderr=subprocess.STDOUT).decode().strip()


def get_repo() -> Tuple[str, str]:
    """Return (owner, repo) from git remote url"""
    try:
        url = run(["git", "remote", "get-url", "origin"])
    except Exception:
        # fallback to environment
        repo_env = os.environ.get("GITHUB_REPOSITORY")
        if repo_env and "/" in repo_env:
            owner, repo = repo_env.split("/", 1)
            return owner, repo
        raise RuntimeError("Cannot determine git remote 'origin' URL and GITHUB_REPOSITORY not set")

    # parse
    if url.startswith("git@"):
        # git@github.com:owner/repo.git
        m = re.match(r"git@[^:]+:([^/]+)/([^.]+)(\.git)?", url)
    else:
        # https://github.com/owner/repo.git
        m = re.match(r"https?://[^/]+/([^/]+)/([^.]+)(\.git)?", url)

    if not m:
        raise RuntimeError(f"Unrecognized remote URL: {url}")

    owner, repo = m.group(1), m.group(2)
    return owner, repo


def get_latest_release(owner: str, repo: str, token: str) -> Optional[dict]:
    url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        return r.json()
    if r.status_code == 404:
        return None
    r.raise_for_status()


def parse_version(s: str) -> Tuple[int, int, int]:
    s = s.strip()
    if s.startswith("v"):
        s = s[1:]
    parts = s.split(".")
    parts += ["0"] * (3 - len(parts))
    return tuple(int(p) for p in parts[:3])


def bump_version(current: Tuple[int, int, int], level: str) -> Tuple[int, int, int]:
    major, minor, patch = current
    if level == "major":
        return (major + 1, 0, 0)
    if level == "minor":
        return (major, minor + 1, 0)
    return (major, minor, patch + 1)


def get_commits_since(tag: Optional[str]) -> List[Tuple[str, str]]:
    # return list of (sha, message)
    cmd = ["git", "log", "--pretty=format:%H%x1f%B%x1e"]
    if tag:
        cmd.insert(2, f"{tag}..HEAD")
    out = run(cmd)
    commits = []
    if not out:
        return commits
    for item in out.split("\x1e"):
        item = item.strip()
        if not item:
            continue
        sha, body = item.split("\x1f", 1)
        message = body.strip().splitlines()[0]
        commits.append((sha, message))
    return commits


def determine_bump(commits: List[Tuple[str, str]]) -> str:
    # simple rules: breaking -> major, feat -> minor, else patch
    level = "patch"
    for _, msg in commits:
        if "BREAKING CHANGE" in msg or "!" in msg.split(":")[0]:
            return "major"
        if msg.lower().startswith("feat"):
            level = "minor"
    return level


def write_version_md(version: str):
    with open("version.md", "w", encoding="utf-8") as f:
        f.write(version + "\n")


def read_version_md() -> Optional[str]:
    if not os.path.exists("version.md"):
        return None
    with open("version.md", "r", encoding="utf-8") as f:
        return f.read().strip()


def prepend_changelog(version: str, commits: List[Tuple[str, str]]):
    changelog = """# Changelog\n\n"""
    date = datetime.utcnow().strftime("%Y-%m-%d")
    header = f"## {version} - {date}\n\n"
    body_lines = []
    for sha, msg in commits:
        body_lines.append(f"- {msg} ({sha[:7]})")
    body = "\n".join(body_lines) + "\n\n"

    existing = ""
    if os.path.exists("CHANGELOG.md"):
        with open("CHANGELOG.md", "r", encoding="utf-8") as f:
            existing = f.read()

    with open("CHANGELOG.md", "w", encoding="utf-8") as f:
        f.write(changelog + header + body + existing)


def create_release_on_github(owner: str, repo: str, tag: str, name: str, body: str, token: str, draft=False):
    url = f"https://api.github.com/repos/{owner}/{repo}/releases"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    payload = {"tag_name": tag, "name": name, "body": body, "draft": draft}
    r = requests.post(url, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    return r.json()


def main(dry_run=False):
    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        print("Error: GITHUB_TOKEN not set in environment")
        sys.exit(1)

    owner, repo = get_repo()
    latest = get_latest_release(owner, repo, token)
    latest_tag = latest["tag_name"] if latest else None
    print(f"Latest release tag: {latest_tag}")

    current_version = read_version_md()
    if current_version:
        current_parsed = parse_version(current_version)
    elif latest_tag:
        current_parsed = parse_version(latest_tag)
    else:
        current_parsed = (0, 0, 0)

    commits = get_commits_since(latest_tag)
    if not commits:
        print("No commits since last release; nothing to do.")
        return

    bump = determine_bump(commits)
    new_version_tuple = bump_version(current_parsed, bump)
    new_version = f"v{new_version_tuple[0]}.{new_version_tuple[1]}.{new_version_tuple[2]}"

    print(f"Bump level: {bump}, new version: {new_version}")

    # create changelog
    prepend_changelog(new_version, commits)

    # update version.md
    write_version_md(new_version)

    # git add, commit, tag, push
    run(["git", "add", "version.md", "CHANGELOG.md"]) if not dry_run else print("git add version.md CHANGELOG.md")
    run(["git", "commit", "-m", f"chore(release): {new_version}"]) if not dry_run else print(f"git commit -m chore(release): {new_version}")
    run(["git", "tag", new_version]) if not dry_run else print(f"git tag {new_version}")
    run(["git", "push"]) if not dry_run else print("git push")
    run(["git", "push", "origin", new_version]) if not dry_run else print(f"git push origin {new_version}")

    # create release on GitHub
    with open("CHANGELOG.md", "r", encoding="utf-8") as f:
        changelog_body = f.read().splitlines()
    # only include top section
    body = "\n".join(changelog_body[:2000])
    if dry_run:
        print("Would create GitHub release with body:\n", body[:1000])
    else:
        rel = create_release_on_github(owner, repo, new_version, new_version, body, token)
        print(f"Created release: {rel.get('html_url')}")


if __name__ == "__main__":
    dry = "--dry-run" in sys.argv
    main(dry_run=dry)
