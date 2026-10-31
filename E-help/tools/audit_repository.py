"""Read-only privacy/secret triage for repository text and Git history; never prints values.

This is a heuristic, not a proof that no secret exists. Ignored runtime data and
credentials outside the repository are deliberately never read.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import subprocess


TEXT_SUFFIXES = {".py", ".md", ".txt", ".json", ".spec", ".ps1", ".toml", ".yml", ".yaml", ".ini", ".cfg",
                 ".js", ".ts", ".tsx", ".jsx", ".html", ".css", ".sh", ".bat", ".cmd", ".xml", ".sql"}


def noreply_identity(value):
    return value.lower().endswith(("@users.noreply.github.com", "@noreply.github.com"))


def protected_path(path):
    name = Path(path).name.lower()
    return (name in {"license.dat", ".env", "credentials.json", "id_rsa", "id_ed25519"}
            or (name.startswith(".env.") and name not in {".env.example", ".env.sample"})
            or Path(path).suffix.lower() in {".pem", ".key"})
PATTERNS = (
    ("token-shaped-literal", re.compile(r"\b(?:sk-[A-Za-z0-9_-]{24,}|gh[pousr]_[A-Za-z0-9_]{25,}|github_pat_[A-Za-z0-9_]{25,}|AKIA[A-Z0-9]{16})\b")),
    ("private-key-marker", re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----")),
    ("credential-literal-review", re.compile(r"(?i)\b(?:api[_-]?key|auth[_-]?token|access[_-]?token|password|secret[_-]?key)\s*[:=]\s*[\"'][^\"'\n]{12,}[\"']")),
    ("personal-user-path", re.compile(r"(?i)[A-Z]:[/\\]+Users[/\\]+(?!Public\b|Default\b|<|\$)[^/\\\s\"'`<>]+")),
    ("non-noreply-email-review", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")),
    ("phone-number-review", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
)


def git(root, *arguments):
    result = subprocess.run(["git", "-C", str(root), *arguments], stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, check=False)
    if result.returncode:
        raise RuntimeError("Git read-only operation failed; sensitive output withheld")
    return result.stdout


def scan_text(content, path, location, *, revision=None):
    findings = []
    for line_number, line in enumerate(content.splitlines(), 1):
        for category, pattern in PATTERNS:
            matches = list(pattern.finditer(line))
            if category == "non-noreply-email-review":
                matches = [match for match in matches if not noreply_identity(match.group())
                           and not match.group().lower().endswith(("@example.com", "@example.org", "@example.net"))]
            if matches:
                record = {"category": category, "path": path, "line": line_number, "location": location}
                if revision:
                    record["revision"] = revision
                findings.append(record)  # No excerpt, full value, or reversible digest.
    return findings


def is_text(path):
    name = Path(path).name
    return Path(path).suffix.lower() in TEXT_SUFFIXES or name in {"AGENTS.md", ".gitignore", ".gitattributes"}


def audit(root, *, history=False):
    root = Path(root).resolve()
    findings, counts = [], {"worktree_text_files": 0, "history_commits": 0, "unique_history_text_blobs": 0}
    paths = git(root, "ls-files", "--cached", "--others", "--exclude-standard", "-z").decode("utf-8").split("\0")
    for relative in filter(None, paths):
        candidate = root / relative
        if candidate.is_symlink() or not candidate.resolve().is_relative_to(root):
            continue
        if protected_path(relative):
            findings.append({"category": "protected-filename-review", "path": relative, "location": "index or untracked"})
            continue  # Report names only; never read authorization/private key files.
        if is_text(relative) and candidate.is_file():
            counts["worktree_text_files"] += 1
            findings.extend(scan_text(candidate.read_text(encoding="utf-8-sig", errors="replace"), relative, "worktree"))
    if history:
        commits = git(root, "rev-list", "--all").decode("ascii").splitlines()
        counts["history_commits"] = len(commits)
        seen = set()
        for commit in commits:
            identity_lines = git(root, "show", "-s", "--format=%ae%n%ce", commit).decode("utf-8").splitlines()
            if any(identity and not noreply_identity(identity) for identity in identity_lines):
                findings.append({"category": "non-noreply-commit-identity", "revision": commit, "location": "Git metadata"})
            for entry in git(root, "ls-tree", "-r", "-z", commit).split(b"\0"):
                if not entry:
                    continue
                metadata, path_bytes = entry.split(b"\t", 1)
                mode, kind, blob = metadata.decode("ascii").split()
                path = path_bytes.decode("utf-8", errors="replace")
                if kind != "blob" or mode == "120000" or blob in seen:
                    continue
                if protected_path(path):
                    seen.add(blob)
                    findings.append({"category": "protected-filename-review", "path": path,
                                     "location": "history", "revision": commit})
                    continue
                if not is_text(path):
                    continue
                seen.add(blob)
                counts["unique_history_text_blobs"] += 1
                content = git(root, "cat-file", "blob", blob).decode("utf-8-sig", errors="replace")
                findings.extend(scan_text(content, path, "history", revision=commit))
    return {"scope": "tracked and non-ignored text; optional local reachable Git history", "counts": counts,
            "findings": findings, "values_redacted": True,
            "limitations": "Heuristic triage only; ignored runtime data/external API settings not read; review matches before any remediation."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--history", action="store_true")
    options = parser.parse_args()
    root = Path(__file__).resolve().parents[2]
    print(json.dumps(audit(root, history=options.history), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
