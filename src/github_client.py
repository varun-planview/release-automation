from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable, Optional

import httpx

PER_PAGE = 100


@dataclass
class Commit:
    message: str
    html_url: str
    sha: str


@dataclass
class RepoInfo:
    last_commit_date: str = ""
    open_pr_count: int = 0
    effective_branch: str = ""
    branch_fallback: bool = False


@dataclass
class RepoCommits:
    current: list[Commit] = field(default_factory=list)
    previous: list[Commit] = field(default_factory=list)
    develop: list[Commit] = field(default_factory=list)
    extra_in_develop: list[Commit] = field(default_factory=list)
    repo_info: RepoInfo = field(default_factory=RepoInfo)


def _safe(val: Any, default: Any) -> Any:
    return default if isinstance(val, BaseException) else val


async def _empty() -> list:
    return []


def _parse_commit(c: dict) -> Commit:
    return Commit(
        message=c.get("commit", {}).get("message", ""),
        html_url=c.get("html_url", ""),
        sha=c.get("sha", ""),
    )


class GitHubClient:
    def __init__(self, token: str, base_url: str = "https://api.github.com"):
        self._base = base_url.rstrip("/")
        self._client = httpx.AsyncClient(
            headers={
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json",
                "User-Agent": "Python-Release-Report",
            },
            limits=httpx.Limits(max_connections=20, max_keepalive_connections=10),
            timeout=30.0,
        )

    async def aclose(self):
        await self._client.aclose()

    async def branch_exists(self, owner: str, repo: str, branch: str) -> bool:
        if not branch:
            return False
        try:
            r = await self._client.get(f"{self._base}/repos/{owner}/{repo}/branches/{branch}")
            return r.status_code == 200
        except Exception:
            return False

    async def get_commits(
        self, owner: str, repo: str, branch: str, max_commits: int
    ) -> list[Commit]:
        commits: list[Commit] = []
        page = 1
        while True:
            url = (
                f"{self._base}/repos/{owner}/{repo}/commits"
                f"?sha={branch}&per_page={PER_PAGE}&page={page}"
            )
            try:
                r = await self._client.get(url)
                if r.status_code != 200:
                    break
                data: list[dict] = r.json()
                if not data:
                    break
                commits.extend(_parse_commit(c) for c in data)
                if len(data) < PER_PAGE or len(commits) >= max_commits:
                    break
                page += 1
            except Exception:
                break
        return commits[:max_commits]

    async def compare_branches(
        self, owner: str, repo: str, base: str, head: str
    ) -> list[Commit]:
        try:
            r = await self._client.get(
                f"{self._base}/repos/{owner}/{repo}/compare/{base}...{head}"
            )
            if r.status_code != 200:
                return []
            return [_parse_commit(c) for c in r.json().get("commits", [])]
        except Exception:
            return []

    async def get_repo_info(self, owner: str, repo: str, branch: str) -> RepoInfo:
        info = RepoInfo(effective_branch=branch)
        try:
            commits_url = f"{self._base}/repos/{owner}/{repo}/commits?sha={branch}&per_page=1"
            prs_url = f"{self._base}/repos/{owner}/{repo}/pulls?state=open&base={branch}&per_page=100"
            c_resp, p_resp = await asyncio.gather(
                self._client.get(commits_url),
                self._client.get(prs_url),
                return_exceptions=True,
            )
            if not isinstance(c_resp, BaseException) and c_resp.status_code == 200:
                data = c_resp.json()
                if data:
                    date_str = data[0].get("commit", {}).get("committer", {}).get("date", "")
                    if date_str:
                        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
                        info.last_commit_date = dt.strftime("%Y-%m-%d %H:%M")
            if not isinstance(p_resp, BaseException) and p_resp.status_code == 200:
                info.open_pr_count = len(p_resp.json())
        except Exception:
            pass
        return info

    async def fetch_repo(
        self,
        repo_cfg,
        global_max: int,
        log_fn: Optional[Callable[[str], None]] = None,
    ) -> RepoCommits:
        from src.config import RepoConfig

        def log(msg: str):
            if log_fn:
                log_fn(msg)

        result = RepoCommits()
        owner = repo_cfg.github_org
        repo = repo_cfg.github_repo
        max_commits = repo_cfg.max_commits_to_fetch or global_max

        try:
            if repo_cfg.validation_type == "branch-based":
                current_exists = await self.branch_exists(owner, repo, repo_cfg.current_branch)
                effective = repo_cfg.current_branch if current_exists else repo_cfg.develop_branch
                result.repo_info.branch_fallback = not current_exists
                if not current_exists:
                    log(
                        f"  {repo_cfg.name}: '{repo_cfg.current_branch}' not found,"
                        f" using '{repo_cfg.develop_branch}'"
                    )

                prev_exists = await self.branch_exists(owner, repo, repo_cfg.previous_branch)
                is_same_as_develop = effective == repo_cfg.develop_branch

                if is_same_as_develop:
                    curr_r, prev_r, info_r = await asyncio.gather(
                        self.get_commits(owner, repo, effective, max_commits),
                        self.get_commits(owner, repo, repo_cfg.previous_branch, max_commits)
                        if prev_exists
                        else _empty(),
                        self.get_repo_info(owner, repo, effective),
                        return_exceptions=True,
                    )
                    result.current = _safe(curr_r, [])
                    result.previous = _safe(prev_r, [])
                    result.develop = result.current
                    result.extra_in_develop = []
                    result.repo_info = _safe(info_r, RepoInfo())
                else:
                    curr_r, prev_r, dev_r, cmp_r, info_r = await asyncio.gather(
                        self.get_commits(owner, repo, effective, max_commits),
                        self.get_commits(owner, repo, repo_cfg.previous_branch, max_commits)
                        if prev_exists
                        else _empty(),
                        self.get_commits(owner, repo, repo_cfg.develop_branch, global_max),
                        self.compare_branches(owner, repo, effective, repo_cfg.develop_branch),
                        self.get_repo_info(owner, repo, effective),
                        return_exceptions=True,
                    )
                    result.current = _safe(curr_r, [])
                    result.previous = _safe(prev_r, [])
                    result.develop = _safe(dev_r, [])
                    result.extra_in_develop = _safe(cmp_r, [])
                    result.repo_info = _safe(info_r, RepoInfo())

                result.repo_info.effective_branch = effective
                result.repo_info.branch_fallback = not current_exists

            else:  # develop-based
                dev_r, info_r = await asyncio.gather(
                    self.get_commits(owner, repo, repo_cfg.develop_branch, max_commits),
                    self.get_repo_info(owner, repo, repo_cfg.develop_branch),
                    return_exceptions=True,
                )
                result.develop = _safe(dev_r, [])
                result.repo_info = _safe(info_r, RepoInfo())
                result.repo_info.effective_branch = repo_cfg.develop_branch

        except Exception as e:
            log(f"  Error fetching {repo_cfg.name}: {e}")

        return result

    async def fetch_all_repos(
        self,
        repos,
        global_max: int,
        log_fn: Optional[Callable[[str], None]] = None,
    ) -> dict[str, RepoCommits]:
        results = await asyncio.gather(
            *(self.fetch_repo(repo, global_max, log_fn) for repo in repos),
            return_exceptions=True,
        )
        return {
            repo.name: (_safe(r, RepoCommits()))
            for repo, r in zip(repos, results)
        }
