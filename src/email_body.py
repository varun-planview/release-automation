from __future__ import annotations

from src.config import RepoConfig
from src.validator import ValidatedCard


def _pve_web_release_version(repositories: list[RepoConfig]) -> str:
    for repo in repositories:
        if repo.name == "PVE Web":
            return repo.release_version
    return ""


def _format_assignee(vc: ValidatedCard, pve_web_rv: str) -> str:
    """Port of PS1 FormatAssignee: appends (releaseVersion) for non-PVE-Web cards."""
    assignee = f"{vc.card.card_url} - {vc.card.assignee_names}"
    if pve_web_rv and pve_web_rv in vc.card.release_versions:
        return assignee
    if vc.card.planning_increment_label:
        return f"{assignee} ({vc.card.planning_increment_label})"
    return assignee


def generate_email_body(
    validated_cards: list[ValidatedCard],
    repositories: list[RepoConfig],
    non_dev_statuses: list[str],
    qe_statuses: list[str],
) -> str:
    pve_web_rv = _pve_web_release_version(repositories)
    non_dev_upper = {s.upper() for s in non_dev_statuses}
    qe_upper = {s.upper() for s in qe_statuses}

    def lane(vc: ValidatedCard) -> str:
        return (vc.card.lane_title or "").upper()

    todo = [vc for vc in validated_cards if lane(vc) and lane(vc) not in non_dev_upper]
    qe = [vc for vc in validated_cards if lane(vc) and lane(vc) in qe_upper]
    final = [vc for vc in validated_cards if lane(vc) == "FINAL PM / TR / UX REVIEW"]

    def build_section(title: str, items: list[ValidatedCard]) -> str:
        teams: dict[str, list[ValidatedCard]] = {}
        for vc in items:
            teams.setdefault(vc.card.team, []).append(vc)
        lines = [f"{title}:\n\n"]
        for i, (team, group) in enumerate(teams.items()):
            if team:
                if i > 0:
                    lines.append("\n")
                lines.append(f"{team}:\n")
                for vc in group:
                    lines.append(f"• {_format_assignee(vc, pve_web_rv)}\n")
        return "".join(lines)

    return (
        build_section("TO DO ", todo)
        + "\n" + build_section("QE TASKS", qe)
        + "\n" + build_section("FINAL PM / TR / UX REVIEW", final)
    )
