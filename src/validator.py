from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal

from src.config import RepoConfig
from src.excel_reader import Card
from src.github_client import Commit, RepoCommits

StatusLevel = Literal["success", "error", "notfound"]


@dataclass
class RepoResult:
    repo_name: str
    status_label: str
    status_level: StatusLevel


@dataclass
class ValidatedCard:
    card: Card
    per_repo_results: list[RepoResult]
    overall_level: StatusLevel


@dataclass
class OrphanCommit:
    card_id: str
    pap_id: str
    card_url: str
    commit_url: str


def extract_pap_ids(commits: list[Commit]) -> frozenset[str]:
    if not commits:
        return frozenset()
    ids: set[str] = set()
    for c in commits:
        for m in re.finditer(r"PAP-\d+", c.message):
            ids.add(m.group())
    return frozenset(ids)


def _validate_branch_based(pap_id: str, rc: RepoCommits) -> tuple[str, StatusLevel]:
    current_ids = extract_pap_ids(rc.current)
    previous_ids = extract_pap_ids(rc.previous)
    develop_ids = extract_pap_ids(rc.develop)
    extra_ids: frozenset[str] = frozenset(
        m.group()
        for c in rc.extra_in_develop
        for m in re.finditer(r"PAP-\d+", c.message)
    )

    in_current = pap_id in current_ids
    in_previous = pap_id in previous_ids
    in_develop = pap_id in develop_ids
    has_extra = pap_id in extra_ids

    if in_previous:
        return "❌ In previous", "error"
    if in_current and in_develop and has_extra:
        return "❌ Extra commits", "error"
    if in_current:
        return "✅ OK", "success"
    if in_develop:
        return "❌ Only in develop", "error"
    return "⚠️ Commit Missing", "notfound"


def _validate_develop_based(pap_id: str, rc: RepoCommits) -> tuple[str, StatusLevel]:
    if pap_id in extract_pap_ids(rc.develop):
        return "✅ OK", "success"
    return "⚠️ Commit Missing", "notfound"


def _apply_smart_filter(results: list[RepoResult]) -> list[RepoResult]:
    """If any success/error exists in the group, suppress notfound results."""
    has_match = any(r.status_level in ("success", "error") for r in results)
    if has_match:
        return [r for r in results if r.status_level != "notfound"]
    return results


def _overall_level(results: list[RepoResult]) -> StatusLevel:
    if not results:
        return "notfound"
    levels = {r.status_level for r in results}
    if "error" in levels:
        return "error"
    if "notfound" in levels:
        return "notfound"
    return "success"


def validate_card(
    card: Card,
    repositories: list[RepoConfig],
    all_repo_commits: dict[str, RepoCommits],
) -> ValidatedCard:
    pap_id = card.pap_id
    matching = [r for r in repositories if r.release_version in card.release_versions]
    non_matching = [r for r in repositories if r.release_version not in card.release_versions]

    all_results: list[RepoResult] = []

    if matching:
        # Group by release version and apply smart filter per group
        by_version: dict[str, list[RepoConfig]] = {}
        for repo in matching:
            by_version.setdefault(repo.release_version, []).append(repo)

        for group_repos in by_version.values():
            group_results: list[RepoResult] = []
            for repo in group_repos:
                rc = all_repo_commits.get(repo.name, RepoCommits())
                if repo.validation_type == "branch-based":
                    label, level = _validate_branch_based(pap_id, rc)
                else:
                    label, level = _validate_develop_based(pap_id, rc)
                group_results.append(RepoResult(repo.name, label, level))
            all_results.extend(_apply_smart_filter(group_results))

    # Check non-matching repos for unexpected commits
    for repo in non_matching:
        rc = all_repo_commits.get(repo.name, RepoCommits())
        if repo.validation_type == "branch-based":
            found = (
                pap_id in extract_pap_ids(rc.current)
                or pap_id in extract_pap_ids(rc.previous)
                or pap_id in extract_pap_ids(rc.develop)
            )
        else:
            found = pap_id in extract_pap_ids(rc.develop)
        if found:
            all_results.append(RepoResult(repo.name, "❌ Unexpected", "error"))

    return ValidatedCard(
        card=card,
        per_repo_results=all_results,
        overall_level=_overall_level(all_results),
    )


def detect_orphans(
    all_repo_commits: dict[str, RepoCommits],
    excel_pap_ids: frozenset[str],
    base_url: str,
    pve_web_repo_name: str = "PVE Web",
) -> list[OrphanCommit]:
    if pve_web_repo_name not in all_repo_commits:
        return []

    pve = all_repo_commits[pve_web_repo_name]
    current_ids = extract_pap_ids(pve.current)
    previous_ids = extract_pap_ids(pve.previous)

    orphan_ids = current_ids - previous_ids - excel_pap_ids
    orphans: list[OrphanCommit] = []

    for pap_id in sorted(orphan_ids):
        card_id = pap_id.replace("PAP-", "")
        commit_url = ""
        for c in pve.current:
            if re.search(re.escape(pap_id), c.message):
                commit_url = c.html_url
                break
        orphans.append(OrphanCommit(
            card_id=card_id,
            pap_id=pap_id,
            card_url=f"{base_url}/{card_id}",
            commit_url=commit_url,
        ))

    return orphans
