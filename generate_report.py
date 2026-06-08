#!/usr/bin/env python3
"""
Release Validation Report Generator
Validates PAP IDs from a Planview Excel export against GitHub repositories
and produces a self-contained HTML report.

Usage:
  python generate_report.py
  python generate_report.py --input path/to/export.xlsx
  python generate_report.py --input export.xlsx --config config.local.json
"""
from __future__ import annotations

import argparse
import asyncio
import sys
from datetime import datetime, timezone
from pathlib import Path

# Ensure src/ is importable when run from the project root
sys.path.insert(0, str(Path(__file__).parent))

from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.table import Table
from rich import print as rprint

from src.config import load_config
from src.email_body import generate_email_body
from src.excel_reader import read_cards
from src.github_client import GitHubClient
from src.report.html_renderer import render_html
from src.validator import OrphanCommit, ValidatedCard, detect_orphans, validate_card

console = Console()


def pick_input_file() -> Path:
    """Open a file dialog if tkinter is available, otherwise prompt for path."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select the Planview export Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        root.destroy()
        if not path:
            console.print("[red]No file selected. Exiting.[/red]")
            sys.exit(0)
        return Path(path)
    except Exception:
        path = console.input("[cyan]Enter path to input Excel file: [/cyan]").strip().strip('"')
        return Path(path)


async def _run(input_path: Path, config_path_arg: str | None, script_dir: Path) -> None:
    # ── Load config ──
    config, config_used = load_config(script_dir, config_path_arg)
    console.print(f"[green]✓[/green] Config loaded: [dim]{config_used}[/dim]")
    console.print(f"[green]✓[/green] {len(config.repositories)} repositories configured")

    # ── Read Excel ──
    console.print(f"\n[cyan]Reading Excel:[/cyan] {input_path.name}")
    cards = read_cards(input_path, config.base_url)
    console.print(f"[green]✓[/green] {len(cards)} cards loaded (after filtering)")

    # ── Fetch GitHub data ──
    console.print(f"\n[cyan]Fetching GitHub data[/cyan] (all repos in parallel)...")
    log_lines: list[str] = []

    def log_fn(msg: str):
        log_lines.append(msg)
        console.print(f"  [dim]{msg.strip()}[/dim]")

    client = GitHubClient(config.github.api_token, config.github.api_base_url)
    try:
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            console=console,
            transient=True,
        ) as progress:
            task = progress.add_task("Fetching commits...", total=None)
            all_repo_commits = await client.fetch_all_repos(
                config.repositories, config.github.max_commits_to_fetch, log_fn
            )
            progress.update(task, completed=1, total=1)
    finally:
        await client.aclose()

    # Print repo summary
    tbl = Table(show_header=True, header_style="bold", box=None, padding=(0, 2))
    tbl.add_column("Repository")
    tbl.add_column("Branch")
    tbl.add_column("Commits", justify="right")
    tbl.add_column("Last Commit")
    tbl.add_column("Open PRs", justify="right")
    for repo in config.repositories:
        rc = all_repo_commits.get(repo.name)
        if rc:
            info = rc.repo_info
            n = len(rc.current) or len(rc.develop)
            fallback = " [yellow](fallback)[/yellow]" if info.branch_fallback else ""
            tbl.add_row(
                repo.name,
                f"{info.effective_branch}{fallback}",
                str(n),
                info.last_commit_date or "—",
                str(info.open_pr_count),
            )
    console.print(tbl)

    # ── Validate cards ──
    console.print(f"\n[cyan]Validating {len(cards)} cards...[/cyan]")
    validated: list[ValidatedCard] = [
        validate_card(card, config.repositories, all_repo_commits) for card in cards
    ]

    ok = sum(1 for v in validated if v.overall_level == "success")
    warn = sum(1 for v in validated if v.overall_level == "warning")
    miss = sum(1 for v in validated if v.overall_level == "notfound")
    nodata = sum(1 for v in validated if not v.per_repo_results)
    console.print(
        f"[green]✅ {ok} OK[/green]  "
        f"[yellow]⚠️ {warn} warnings[/yellow]  "
        f"[red]❌ {miss} missing[/red]  "
        f"[dim]— {nodata} no version[/dim]"
    )

    # ── Orphan detection ──
    excel_pap_ids: frozenset[str] = frozenset(c.pap_id for c in cards)
    orphans = detect_orphans(all_repo_commits, excel_pap_ids, config.base_url)
    if orphans:
        console.print(f"[yellow]👻 {len(orphans)} orphan commit(s) found[/yellow]")
    else:
        console.print("[green]✓[/green] No orphan commits detected")

    # ── Generate email body ──
    email_body = generate_email_body(
        validated, config.repositories, config.non_dev_statuses, config.qe_statuses
    )

    # ── Render HTML ──
    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    html = render_html(
        validated_cards=validated,
        orphans=orphans,
        all_repo_commits=all_repo_commits,
        email_body=email_body,
        config=config,
        input_filename=input_path.name,
        config_path=config_used,
        generated_at=generated_at,
    )

    # ── Write output ──
    output_path = input_path.parent / f"{input_path.stem}_report.html"
    output_path.write_text(html, encoding="utf-8")

    console.print(
        Panel(
            f"[bold green]Report generated successfully![/bold green]\n\n"
            f"📄 [link={output_path.as_uri()}]{output_path}[/link]\n\n"
            f"Open in your browser to view the interactive report.",
            title="Done",
            border_style="green",
        )
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a release validation HTML report from a Planview Excel export."
    )
    parser.add_argument("--input", "-i", help="Path to the input Excel file (.xlsx)")
    parser.add_argument("--config", "-c", help="Path to config JSON (default: config.local.json)")
    args = parser.parse_args()

    script_dir = Path(__file__).parent

    console.print(
        Panel(
            "[bold]Release Validation Report Generator[/bold]\n[dim]Validates PAP IDs against GitHub repositories[/dim]",
            border_style="blue",
        )
    )

    input_path = Path(args.input) if args.input else pick_input_file()
    if not input_path.exists():
        console.print(f"[red]ERROR: File not found: {input_path}[/red]")
        sys.exit(1)

    try:
        asyncio.run(_run(input_path, args.config, script_dir))
    except SystemExit:
        raise
    except KeyboardInterrupt:
        console.print("\n[yellow]Cancelled.[/yellow]")
        sys.exit(0)
    except Exception as e:
        console.print(f"\n[red]Unexpected error: {e}[/red]")
        raise


if __name__ == "__main__":
    main()
