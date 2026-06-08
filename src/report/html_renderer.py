from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from src.config import AppConfig
from src.github_client import RepoCommits, RepoInfo
from src.validator import OrphanCommit, ValidatedCard

TEMPLATE_DIR = Path(__file__).parent / "templates"

_LEVEL_ROW_CLASS = {
    "success": "row-ok",
    "error": "row-error",
    "notfound": "row-warning",
}
_BADGE_CLASS = {
    "success": "badge-ok",
    "error": "badge-error",
    "notfound": "badge-warning",
}


def render_html(
    validated_cards: list[ValidatedCard],
    orphans: list[OrphanCommit],
    all_repo_commits: dict[str, RepoCommits],
    email_body: str,
    config: AppConfig,
    input_filename: str,
    config_path: str,
    generated_at: str,
) -> str:
    env = Environment(loader=FileSystemLoader(str(TEMPLATE_DIR)), autoescape=True)
    tmpl = env.get_template("report.html.jinja")

    ok_count = sum(1 for vc in validated_cards if vc.overall_level == "success")
    error_count = sum(1 for vc in validated_cards if vc.overall_level == "error")
    not_found_count = sum(1 for vc in validated_cards if vc.overall_level == "notfound")
    no_data_count = sum(1 for vc in validated_cards if not vc.per_repo_results)

    summary = {
        "total": len(validated_cards),
        "ok_count": ok_count,
        "error_count": error_count,
        "not_found_count": not_found_count,
        "no_data_count": no_data_count,
        "orphan_count": len(orphans),
    }

    repo_infos = []
    for repo in config.repositories:
        rc = all_repo_commits.get(repo.name, RepoCommits())
        info = rc.repo_info
        repo_infos.append({
            "name": repo.name,
            "branch": info.effective_branch or repo.develop_branch,
            "last_commit_date": info.last_commit_date or "—",
            "open_pr_count": info.open_pr_count,
            "branch_fallback": info.branch_fallback,
            "commit_count": len(rc.current) or len(rc.develop),
        })

    card_rows = []
    for vc in validated_cards:
        c = vc.card
        per_repo = [
            {
                "repo_name": r.repo_name,
                "status_label": r.status_label,
                "status_level": r.status_level,
                "badge_class": _BADGE_CLASS.get(r.status_level, "badge-warning"),
            }
            for r in vc.per_repo_results
        ]
        card_rows.append({
            "team": c.team,
            "card_title": c.card_title,
            "card_id": c.card_id,
            "card_type": c.card_type,
            "lane_title": c.lane_title,
            "release_version": c.planning_increment_label,
            "pap_id": c.pap_id,
            "card_url": c.card_url,
            "assignees": c.assignee_names,
            "per_repo_results": per_repo,
            "row_class": _LEVEL_ROW_CLASS.get(vc.overall_level, "row-warning"),
            "overall_level": vc.overall_level,
        })

    orphan_rows = [
        {
            "card_id": o.card_id,
            "pap_id": o.pap_id,
            "card_url": o.card_url,
            "commit_url": o.commit_url,
        }
        for o in orphans
    ]

    return tmpl.render(
        meta={
            "generated_at": generated_at,
            "input_filename": input_filename,
            "config_path": config_path,
        },
        summary=summary,
        repo_infos=repo_infos,
        card_rows=card_rows,
        orphan_rows=orphan_rows,
        email_body=email_body,
    )
