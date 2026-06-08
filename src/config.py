from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional


@dataclass
class GitHubConfig:
    api_token: str
    api_base_url: str = "https://api.github.com"
    max_commits_to_fetch: int = 500


@dataclass
class RepoConfig:
    name: str
    github_org: str
    github_repo: str
    validation_type: Literal["branch-based", "develop-based"]
    release_version: str
    develop_branch: str = "develop"
    current_branch: str = ""
    previous_branch: str = ""
    max_commits_to_fetch: Optional[int] = None


@dataclass
class AppConfig:
    github: GitHubConfig
    repositories: list[RepoConfig]
    base_url: str
    non_dev_statuses: list[str]
    qe_statuses: list[str]


def _repo_from_dict(d: dict) -> RepoConfig:
    return RepoConfig(
        name=d["name"],
        github_org=d.get("githubOrg", ""),
        github_repo=d.get("githubRepo", ""),
        validation_type=d.get("validationType", "develop-based"),
        release_version=d.get("releaseVersion", ""),
        develop_branch=d.get("developBranch", "develop"),
        current_branch=d.get("currentBranch", ""),
        previous_branch=d.get("previousBranch", ""),
        max_commits_to_fetch=d.get("maxCommitsToFetch") or None,
    )


def load_config(
    script_dir: Path, config_arg: Optional[str] = None
) -> tuple[AppConfig, str]:
    """Load config, returning (AppConfig, path_used). Schema is backward-compatible with PS1 config."""
    if config_arg:
        config_path = Path(config_arg)
    else:
        local = script_dir / "config.local.json"
        default = script_dir / "config.json"
        if local.exists():
            config_path = local
        elif default.exists():
            config_path = default
        else:
            raise SystemExit(
                f"ERROR: No config file found at {local} or {default}\n"
                "Please create config.local.json with your GitHub token and repository settings."
            )

    try:
        with open(config_path, encoding="utf-8") as f:
            d = json.load(f)
    except Exception as e:
        raise SystemExit(f"ERROR: Failed to read config file {config_path}: {e}")

    gh = d.get("github", {})
    token = gh.get("apiToken", "")
    if not token or token == "YOUR_GITHUB_TOKEN_HERE":
        raise SystemExit(
            f"ERROR: GitHub API token not configured in {config_path}\n"
            "Please set the 'github.apiToken' value."
        )

    global_max = gh.get("maxCommitsToFetch", 500)
    if not global_max or global_max <= 0:
        raise SystemExit("ERROR: maxCommitsToFetch must be greater than 0")

    github_cfg = GitHubConfig(
        api_token=token,
        api_base_url=gh.get("apiBaseUrl", "https://api.github.com"),
        max_commits_to_fetch=int(global_max),
    )

    repos_raw = d.get("repositories", [])
    repos = []
    for r in repos_raw:
        org = r.get("githubOrg", "")
        if not org or org == "YOUR_GITHUB_ORG_HERE":
            raise SystemExit(
                f"ERROR: GitHub organization not configured for repository: {r.get('name')}"
            )
        repos.append(_repo_from_dict(r))

    return (
        AppConfig(
            github=github_cfg,
            repositories=repos,
            base_url=d.get("BaseURL", ""),
            non_dev_statuses=d.get("nonDevStatuses", []),
            qe_statuses=d.get("qeStatuses", []),
        ),
        str(config_path),
    )
