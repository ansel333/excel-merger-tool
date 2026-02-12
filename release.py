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
import argparse

import requests


def run(cmd: List[str]) -> str:
    return subprocess.check_output(cmd, stderr=subprocess.STDOUT).decode().strip()


def get_uncommitted_changes() -> List[str]:
    try:
        out = run(["git", "status", "--porcelain"]) or ""
        lines = [l.strip() for l in out.splitlines() if l.strip()]
        return lines
    except Exception:
        return []


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
    if r.status_code in (401, 403):
        # Authentication / permission error
        raise PermissionError(f"GitHub API authentication failed (status {r.status_code}).")
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
    changelog_file_exists = os.path.exists("CHANGELOG.md")
    if changelog_file_exists:
        with open("CHANGELOG.md", "r", encoding="utf-8") as f:
            existing = f.read()
    is_empty_changelog = changelog_file_exists and (not existing.strip())

    with open("CHANGELOG.md", "w", encoding="utf-8") as f:
        f.write(changelog + header + body + existing)


def create_release_on_github(owner: str, repo: str, tag: str, name: str, body: str, token: str, draft=False):
    url = f"https://api.github.com/repos/{owner}/{repo}/releases"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    payload = {"tag_name": tag, "name": name, "body": body, "draft": draft}
    r = requests.post(url, headers=headers, data=json.dumps(payload))
    if r.status_code in (401, 403):
        raise PermissionError(f"GitHub API authentication failed when creating release (status {r.status_code}).")
    r.raise_for_status()
    return r.json()


def gh_installed() -> bool:
    try:
        run(["gh", "--version"])
        return True
    except Exception:
        return False


def gh_get_latest_release(owner: str, repo: str) -> Optional[str]:
    try:
        out = run(["gh", "release", "view", "--repo", f"{owner}/{repo}", "--json", "tagName"]) 
        j = json.loads(out)
        return j.get("tagName")
    except subprocess.CalledProcessError:
        return None
    except Exception:
        return None


def get_latest_local_tag() -> Optional[str]:
    try:
        t = run(["git", "describe", "--tags", "--abbrev=0"]) 
        return t
    except Exception:
        return None


def get_local_tags() -> set:
    try:
        out = run(["git", "tag", "--list"]) or ""
        return set([t.strip() for t in out.splitlines() if t.strip()])
    except Exception:
        return set()


def get_local_tags_list() -> list:
    try:
        out = run(["git", "tag", "--list", "--sort=-v:refname"]) or ""
        return [t.strip() for t in out.splitlines() if t.strip()]
    except Exception:
        return []


def get_tag_date(tag: str) -> Optional[str]:
    try:
        d = run(["git", "log", "-1", "--format=%ad", "--date=short", tag])
        return d.strip()
    except Exception:
        return None


def get_tag_commit_summary(tag: str) -> str:
    try:
        s = run(["git", "show", "-s", "--format=%h %s", tag])
        return s.strip()
    except Exception:
        return ""


def ensure_tags_have_sections(existing: str, tag_list: list) -> str:
    # Parse existing versions
    present = set()
    parts = re.split(r'(?=^## )', existing, flags=re.M)
    for sec in parts[1:]:
        m = re.match(r"##\s*(v[0-9]+\.[0-9]+\.[0-9]+)", sec)
        if m:
            present.add(m.group(1))

    added = []
    for tag in tag_list:
        if tag not in present:
            date = get_tag_date(tag) or datetime.utcnow().strftime("%Y-%m-%d")
            summary = get_tag_commit_summary(tag)
            sec = f"## {tag} - {date}\n\n- (auto-added) tagged commit: {summary}\n\n"
            added.append(sec)

    # append added sections after existing text
    if added:
        return existing + "\n" + "".join(added)
    return existing


def get_github_release_tags(owner: str, repo: str, token: str) -> set:
    try:
        url = f"https://api.github.com/repos/{owner}/{repo}/releases?per_page=100"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        items = r.json()
        return set([i.get("tag_name") for i in items if i.get("tag_name")])
    except Exception:
        return set()


def get_gh_release_tags(owner: str, repo: str) -> set:
    try:
        out = run(["gh", "release", "list", "--repo", f"{owner}/{repo}", "--limit", "200", "--json", "tagName"]) 
        j = json.loads(out)
        return set([item.get("tagName") for item in j if item.get("tagName")])
    except Exception:
        return set()


def get_remote_releases_api(owner: str, repo: str, token: str) -> list:
    try:
        url = f"https://api.github.com/repos/{owner}/{repo}/releases?per_page=200"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        items = r.json()
        # normalize to dicts with tag_name, published_at, body
        out = []
        for i in items:
            out.append({
                "tag_name": i.get("tag_name"),
                "published_at": i.get("published_at") or "",
                "body": i.get("body") or "",
            })
        # sort by published_at desc (missing dates last)
        out.sort(key=lambda x: x.get("published_at") or "", reverse=True)
        return out
    except Exception:
        return []


def get_remote_releases_gh(owner: str, repo: str) -> list:
    try:
        out = run(["gh", "release", "list", "--repo", f"{owner}/{repo}", "--limit", "200", "--json", "tagName,publishedAt,body"]) 
        j = json.loads(out)
        out_list = []
        for item in j:
            out_list.append({
                "tag_name": item.get("tagName"),
                "published_at": item.get("publishedAt") or "",
                "body": item.get("body") or "",
            })
        out_list.sort(key=lambda x: x.get("published_at") or "", reverse=True)
        return out_list
    except Exception:
        return []


def parse_changelog_sections(text: str) -> Tuple[list, dict]:
    # returns (ordered_versions, mapping version->section body)
    parts = re.split(r'(?=^## )', text, flags=re.M)
    mapping = {}
    order = []
    if not parts:
        return order, mapping
    for sec in parts[1:]:
        m = re.match(r"##\s*(v[0-9]+\.[0-9]+\.[0-9]+)\s*-\s*([0-9]{4}-[0-9]{2}-[0-9]{2})\s*\n\n", sec)
        if m:
            ver = m.group(1)
            mapping[ver] = sec
            order.append(ver)
    return order, mapping


def clean_existing_changelog(existing: str, known_tags: set, keep_versions: set) -> str:
    # Keep header before first section
    parts = re.split(r'(?=^## )', existing, flags=re.M)
    if not parts:
        return existing
    header = parts[0] if not parts[0].strip().startswith('##') else "# Changelog\n\n"
    kept = []
    for sec in parts[1:]:
        m = re.match(r"##\s*(v[0-9]+\.[0-9]+\.[0-9]+)", sec)
        if m:
            ver = m.group(1)
            if ver in known_tags or ver in keep_versions:
                kept.append(sec)
            else:
                # skip this generated/invalid section
                continue
        else:
            # keep unknown-format sections
            kept.append(sec)
    return header + "".join(kept)


def dedupe_changelog(full_text: str) -> str:
    # Robustly parse sections, merge duplicate version sections and dedupe bullets.
    header = "# Changelog\n\n"
    sec_pat = re.compile(r"(?ms)^##\s*(v[0-9]+\.[0-9]+\.[0-9]+)\s*-\s*([0-9]{4}-[0-9]{2}-[0-9]{2})\s*\n(.*?)(?=^##\s*v|\Z)")
    sections = {}
    order = []
    misc_parts = []
    for m in sec_pat.finditer(full_text):
        ver = m.group(1)
        date = m.group(2)
        body = m.group(3).strip()
        bullets = [line.strip() for line in body.splitlines() if line.strip()]
        if ver not in sections:
            sections[ver] = {"date": date, "bullets": []}
            order.append(ver)
        for b in bullets:
            if b not in sections[ver]["bullets"]:
                sections[ver]["bullets"].append(b)
    # detect misc content before first section (ignore redundant '# Changelog')
    first_section = sec_pat.search(full_text)
    if first_section:
        pre = full_text[: first_section.start()].strip()
        if pre and not pre.strip().startswith('# Changelog'):
            misc_parts.append(pre)
    # build output: misc header, then versions in order (preserve order seen)
    out_lines = [header]
    for ver in order:
        date = sections[ver]["date"]
        out_lines.append(f"## {ver} - {date}\n\n")
        for b in sections[ver]["bullets"]:
            out_lines.append(b + "\n")
        out_lines.append("\n")
    # append misc parts if any
    if misc_parts:
        out_lines.append("\n".join(misc_parts))
        out_lines.append("\n")
    return "".join(out_lines)


def create_release_with_gh(owner: str, repo: str, tag: str, name: str, body: str, draft=False) -> dict:
    import tempfile
    with tempfile.NamedTemporaryFile(mode="w", delete=False, encoding="utf-8", suffix=".md") as tf:
        tf.write(body)
        tf.flush()
        tmpname = tf.name
    cmd = [
        "gh",
        "release",
        "create",
        tag,
        "-t",
        name,
        "-F",
        tmpname,
        "--repo",
        f"{owner}/{repo}",
    ]
    if draft:
        cmd.append("--draft")
    out = run(cmd)
    # gh prints URL on success; return a minimal dict
    return {"html_url": out.strip()}


def main(mode: Optional[str] = None):
    token = os.environ.get("GITHUB_TOKEN")
    # mode: one of 'preview' (dry-run), 'release' (real), 'dry-run' (alias)
    if mode is None:
        # interactive prompt
        print("No mode option provided. Choose action:")
        print("  1) preview  - show what will be released (safe dry-run)")
        print("  2) dry-run  - same as preview")
        print("  3) release  - perform the release (will create tag and push)")
        choice = input("Enter 1,2,3 (default 1): ").strip() or "1"
        if choice == "3":
            mode = "release"
        elif choice == "2":
            mode = "dry-run"
        else:
            mode = "preview"

    dry_run = (mode != "release")

    owner = repo = None
    latest = None
    latest_tag = None
    # Prefer local git tags (most reliable). If none found, try GitHub releases (API with token or gh CLI).
    latest_tag = get_latest_local_tag()
    if latest_tag:
        print(f"Latest release tag (local): {latest_tag}")
    else:
        if token:
            owner, repo = get_repo()
            latest = get_latest_release(owner, repo, token)
            latest_tag = latest["tag_name"] if latest else None
            print(f"Latest release tag (API): {latest_tag}")
        else:
            if gh_installed():
                try:
                    owner, repo = get_repo()
                    latest_tag = gh_get_latest_release(owner, repo)
                    print(f"Latest release tag (gh): {latest_tag}")
                except Exception:
                    latest_tag = None
            if not latest_tag:
                if dry_run:
                    print("No GitHub token and no tags found via gh CLI or locally; proceeding in dry-run mode using full commit history.")
                else:
                    print("Error: GITHUB_TOKEN not set and no tags found locally or via gh CLI.")
                    sys.exit(1)

    # If possible, fetch remote releases early and prefer remote latest as the base for computing commits.
    remote_releases = []
    try:
        if not owner or not repo:
            try:
                owner, repo = get_repo()
            except Exception:
                owner = repo = None
        if owner and repo:
            if token:
                remote_releases = get_remote_releases_api(owner, repo, token)
            elif gh_installed():
                remote_releases = get_remote_releases_gh(owner, repo)
            else:
                # fallback: use git ls-remote to list tags on origin
                ls = run(["git", "ls-remote", "--tags", "origin"]) or ""
                tags = []
                for line in ls.splitlines():
                    parts = line.split()
                    if len(parts) >= 2 and "refs/tags/" in parts[1]:
                        tagref = parts[1].split("refs/tags/")[-1]
                        # strip ^{} for annotated tags
                        tag = tagref.replace("^{}", "")
                        tags.append(tag)
                # pick semver-sorted latest
                def ver_key(t):
                    try:
                        return parse_version(t)
                    except Exception:
                        return (0, 0, 0)
                tags = sorted(set(tags), key=ver_key, reverse=True)
                if tags:
                    remote_releases = [{"tag_name": tags[0], "published_at": "", "body": ""}]
    except Exception:
        remote_releases = []

    # compute remote tag set for cleanup; if remote_releases empty, try ls-remote fallback
    if remote_releases:
        remote_tag_set = set([r.get("tag_name") for r in remote_releases if r.get("tag_name")])
    else:
        try:
            ls = run(["git", "ls-remote", "--tags", "origin"]) or ""
            tags = []
            for line in ls.splitlines():
                parts = line.split()
                if len(parts) >= 2 and "refs/tags/" in parts[1]:
                    tagref = parts[1].split("refs/tags/")[-1]
                    tag = tagref.replace("^{}", "")
                    tags.append(tag)
            remote_tag_set = set(tags)
        except Exception:
            remote_tag_set = set()

    local_tag_set = get_local_tags()
    # tags present locally but missing remotely should be removed (cleanup)
    local_only = local_tag_set - remote_tag_set
    if local_only:
        print(f"Local-only tags detected: {sorted(local_only)}")
        for t in sorted(local_only):
            if dry_run:
                print(f"Would delete local tag: {t}")
            else:
                try:
                    run(["git", "tag", "-d", t])
                    print(f"Deleted local tag {t}")
                except Exception as e:
                    print(f"Failed to delete local tag {t}: {e}")
        # refresh latest_tag after cleanup
        latest_tag = get_latest_local_tag()

    # determine remote_latest for re-evaluation
    if remote_releases:
        remote_latest = remote_releases[0].get("tag_name")
    elif remote_tag_set:
        def ver_key(t):
            try:
                return parse_version(t)
            except Exception:
                return (0, 0, 0)
        tags_sorted = sorted(remote_tag_set, key=ver_key, reverse=True)
        remote_latest = tags_sorted[0] if tags_sorted else None
    else:
        remote_latest = None

    if remote_latest and remote_latest != latest_tag:
        print(f"Remote latest release {remote_latest} differs from local latest {latest_tag}; re-evaluating from remote latest.")
        latest_tag = remote_latest

    file_version = read_version_md()
    if latest_tag:
        current_parsed = parse_version(latest_tag)
    elif file_version:
        current_parsed = parse_version(file_version)
    else:
        current_parsed = (0, 0, 0)

    # Read existing changelog early so we can detect if it's empty
    existing = ""
    changelog_file_exists = os.path.exists("CHANGELOG.md")
    if changelog_file_exists:
        with open("CHANGELOG.md", "r", encoding="utf-8") as f:
            existing = f.read()
    is_empty_changelog = changelog_file_exists and (not existing.strip())

    # Always prefer remote releases as source of truth when available.
    regenerate_all = False
    remote_releases = []
    try:
        if not owner:
            try:
                owner, repo = get_repo()
            except Exception:
                owner = repo = None

        if owner and repo:
            if token:
                remote_releases = get_remote_releases_api(owner, repo, token)
            elif gh_installed():
                remote_releases = get_remote_releases_gh(owner, repo)
            else:
                # fallback: attempt to list remote tags (no release bodies)
                ls = run(["git", "ls-remote", "--tags", "origin"]) or ""
                remote_releases = []
                for line in ls.splitlines():
                    parts = line.split()
                    if len(parts) >= 2 and "refs/tags/" in parts[1]:
                        tag = parts[1].split("refs/tags/")[-1]
                        remote_releases.append({"tag_name": tag, "published_at": "", "body": ""})
            # if we got remote releases, compare latest with local changelog
            if remote_releases:
                remote_latest = remote_releases[0].get("tag_name")
                order_map = parse_changelog_sections(existing)
                local_order, local_map = order_map
                local_has_latest = remote_latest in local_map
                remote_latest_body = (remote_releases[0].get("body") or "").strip()
                local_latest_body = ""
                if local_has_latest:
                    local_latest_body = local_map[remote_latest]
                # If latest remote release is missing or differs from local, regenerate entire changelog
                if (not local_has_latest) or (remote_latest_body and remote_latest_body not in local_latest_body):
                    print(f"Remote latest release {remote_latest} differs from local changelog — regenerating full changelog from remote state")
                    regenerate_all = True
    except Exception:
        # any failure fetching remote releases should not block local flow
        regenerate_all = False

    if regenerate_all:
        # build changelog from remote_releases; include body when present, else placeholder
        parts = []
        for r in remote_releases:
            tag = r.get("tag_name")
            date = r.get("published_at")[:10] if r.get("published_at") else (get_tag_date(tag) or datetime.utcnow().strftime("%Y-%m-%d"))
            body_text = r.get("body") or ""
            if not body_text:
                summary = get_tag_commit_summary(tag)
                body_text = f"- (auto-added) tagged commit: {summary}"
            sec = f"## {tag} - {date}\n\n{body_text}\n\n"
            parts.append(sec)
        full_text = "# Changelog\n\n" + "".join(parts)
        final_text = dedupe_changelog(full_text)
        if not dry_run:
            with open("CHANGELOG.md", "w", encoding="utf-8") as f:
                f.write(final_text)
        else:
            print("(preview) regenerated changelog from remote releases:\n")
            print(final_text[:2000])
        # update existing variable so later flow uses regenerated content
        existing = final_text
        is_empty_changelog = False

    commits = get_commits_since(latest_tag)
    if not commits and not is_empty_changelog:
        print("No commits since last release; nothing to do.")
        return

    bump = determine_bump(commits)
    new_version_tuple = bump_version(current_parsed, bump)
    new_version = f"v{new_version_tuple[0]}.{new_version_tuple[1]}.{new_version_tuple[2]}"

    print(f"Bump level: {bump}, new version: {new_version}")

    # determine known tags (local + remote via API or gh if available)
    known = get_local_tags()
    if token and owner and repo:
        known |= get_github_release_tags(owner, repo, token)
    elif gh_installed() and owner and repo:
        known |= get_gh_release_tags(owner, repo)

    # create changelog (clean old invalid entries)
    changelog_file_exists = os.path.exists("CHANGELOG.md")
    is_empty_changelog = changelog_file_exists and (not existing.strip())

    # build new top section
    date = datetime.utcnow().strftime("%Y-%m-%d")
    header = f"## {new_version} - {date}\n\n"
    body_lines = [f"- {msg} ({sha[:7]})" for sha, msg in commits]
    body = "\n".join(body_lines) + "\n\n"

    cleaned_existing = clean_existing_changelog(existing, known, keep_versions={new_version})

    # If changelog was empty, ensure we generate sections for existing tags
    if is_empty_changelog:
        local_tags_ordered = get_local_tags_list()
        cleaned_existing = ensure_tags_have_sections(cleaned_existing, local_tags_ordered)

    # ensure every known local tag has a changelog section (if not already done)
    local_tags_ordered = get_local_tags_list()
    augmented = ensure_tags_have_sections(cleaned_existing, local_tags_ordered)

    # If `augmented` already contains a leading '# Changelog', strip it to avoid duplication
    aug = augmented
    if aug.lstrip().startswith("# Changelog"):
        # remove first occurrence of the header and following blank lines
        aug = re.sub(r"^# Changelog\s*\n", "", aug.lstrip(), count=1)
    full_text = "# Changelog\n\n" + header + body + aug
    final_text = dedupe_changelog(full_text)
    # Ensure only a single '# Changelog' header appears
    final_text = re.sub(r'(?ms)(?:# Changelog\s*\n\s*)+', '# Changelog\n\n', final_text)
    # Remove any other stray '# Changelog' occurrences after the first header
    header_text = '# Changelog\n\n'
    first_idx = final_text.find(header_text)
    if first_idx != -1:
        prefix = final_text[: first_idx + len(header_text)]
        rest = final_text[first_idx + len(header_text):]
        rest = rest.replace('# Changelog', '')
        final_text = prefix + rest
    with open("CHANGELOG.md", "w", encoding="utf-8") as f:
        f.write(final_text)

    # update version.md
    write_version_md(new_version)

    # If there were no commits since latest tag but changelog was empty,
    # we regenerated changelog for existing tags and should re-create the release
    recreate_release_for_tag = False
    existing_tag_to_release = None
    if not commits and is_empty_changelog:
        # re-use the latest tag
        existing_tag_to_release = latest_tag or read_version_md()
        if existing_tag_to_release:
            print(f"No new commits, but changelog was empty — will regenerate changelog and recreate release for {existing_tag_to_release}")
            recreate_release_for_tag = True

    if not recreate_release_for_tag:
        # git add, commit, tag, push for a new release
        run(["git", "add", "version.md", "CHANGELOG.md"]) if not dry_run else print("git add version.md CHANGELOG.md")
        run(["git", "commit", "-m", f"chore(release): {new_version}"]) if not dry_run else print(f"git commit -m chore(release): {new_version}")
        run(["git", "tag", new_version]) if not dry_run else print(f"git tag {new_version}")
        run(["git", "push"]) if not dry_run else print("git push")
        run(["git", "push", "origin", new_version]) if not dry_run else print(f"git push origin {new_version}")

    # create release on GitHub (use token or gh CLI if available)
    with open("CHANGELOG.md", "r", encoding="utf-8") as f:
        changelog_body = f.read().splitlines()
    # only include top section
    body = "\n".join(changelog_body[:2000])

    if dry_run:
        print("Dry-run / preview: would create GitHub release with body:\n")
        print(body[:1000])
        return

    # Before performing a real release, check for uncommitted changes
    changes = get_uncommitted_changes()
    if changes:
        print("Found uncommitted/unstaged changes in the working tree:")
        for c in changes:
            print(f"  {c}")
        ans = input("You have uncommitted changes. Continue release anyway? [y/N]: ").strip().lower()
        if ans not in ("y", "yes"):
            print("Aborting release. Commit or stash your changes and retry.")
            sys.exit(1)

    if token:
        try:
            rel = create_release_on_github(owner, repo, new_version, new_version, body, token)
            print(f"Created release: {rel.get('html_url')}")
        except PermissionError as e:
            print(str(e))
            print("Permission error with GitHub API. Try: gh auth login or set GITHUB_TOKEN.")
            sys.exit(1)
    elif gh_installed():
        try:
            if not owner:
                owner, repo = get_repo()
            rel = create_release_with_gh(owner, repo, new_version, new_version, body)
            print(f"Created release (gh): {rel.get('html_url')}")
        except Exception as e:
            print(f"gh CLI release failed: {e}")
            print("If not authenticated, run: gh auth login")
            sys.exit(1)
    else:
        print("No GITHUB_TOKEN and gh CLI not available to create release.")
        sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Release helper: preview or perform releases")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--preview", action="store_true", help="Preview the release (dry-run)")
    group.add_argument("--dry-run", action="store_true", help="Alias for --preview")
    group.add_argument("--release", action="store_true", help="Perform the release")
    args = parser.parse_args()
    if args.release:
        mode = "release"
    elif args.preview or args.dry_run:
        mode = "preview"
    else:
        mode = None
    main(mode=mode)
