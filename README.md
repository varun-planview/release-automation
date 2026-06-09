# Release Automation

## Overview

**generate_report.py** validates PAP IDs from a Planview Excel export against GitHub repositories and produces a self-contained interactive HTML report. It transforms manual release validation from hours of work into a single command.

- **Validates commit presence** across multiple GitHub repositories
- **Detects orphan commits** — commits in the current release branch whose card does not have the current release version assigned
- **Generates an interactive HTML report** with filtering, sorting, and status badges
- **Produces formatted email body text** for release status communications

## Prerequisites

- Python 3.10+
- Install dependencies:

```
pip install -r requirements.txt
```

## Configuration

Copy `config.json` to `config.local.json` and fill in your values. `config.local.json` is gitignored and is the preferred location for secrets.

### Configuration Parameters

| Parameter | Description |
|-----------|-------------|
| `github.apiToken` | GitHub personal access token with repo read access |
| `github.apiBaseUrl` | GitHub API base URL (default: `https://api.github.com`) |
| `github.maxCommitsToFetch` | Global default for max commits fetched per branch (default: 500) |
| `repositories[].name` | Display name for the repository |
| `repositories[].githubOrg` | GitHub organization name |
| `repositories[].githubRepo` | GitHub repository name |
| `repositories[].validationType` | `"branch-based"` or `"develop-based"` |
| `repositories[].releaseVersion` | Release version identifier matching the Excel `Planning Increment Label` |
| `repositories[].currentBranch` | Current release branch (branch-based only) |
| `repositories[].previousBranch` | Previous release branch (branch-based only) |
| `repositories[].developBranch` | Development branch name (default: `develop`) |
| `repositories[].maxCommitsToFetch` | Per-repo override for max commits fetched |
| `BaseURL` | Base URL for generating Planview card links |
| `nonDevStatuses` | Lane statuses excluded from the "TO DO" email section |
| `qeStatuses` | Lane statuses included in the "QE TASKS" email section |

## Usage

```
python generate_report.py
```

Without `--input`, a file dialog opens for you to select the Excel file. Alternatively:

```
python generate_report.py --input path/to/export.xlsx
python generate_report.py --input export.xlsx --config config.local.json
```

### Input

A Planview Excel export with these columns: `Card ID`, `Team`, `Card Title`, `Assignee(s)`, `Card Type`, `Current Lane Title`, `Planning Increment Label`, `Tags`

### Output

A single `[InputFileName]_report.html` file written alongside the input. Open it in any browser.

## Validation Logic

### Branch-Based Validation

For repositories with a dedicated release branch (e.g. PVE Web):

| Status | Meaning |
|--------|---------|
| ✅ OK | PAP ID found in the current release branch |
| ❌ In previous | Shipped in last release |
| ❌ Extra commits | Develop has additional commits beyond the release branch |
| ❌ Only in develop | Merged to develop but not yet in the release branch |
| ⚠️ Commit Missing | No commits found |

### Develop-Based Validation

For repositories validated against the develop branch only (e.g. ActionBoard):

| Status | Meaning |
|--------|---------|
| ✅ OK | PAP ID found in the develop branch |
| ⚠️ Commit Missing | Not found in the develop branch |

### Cross-Repository Validation

If a PAP ID is found in a repository that does not match the card's release version, it is flagged as **❌ Unexpected**.

### Smart Filtering

When a card has results across multiple repositories, `notfound` results are suppressed if any `success` or `error` result exists in the same release-version group — avoiding noise for cards that simply don't exist in certain repos.

### Orphan Detection

Commits found in the current release branch of PVE Web where the associated card does not have the current release version assigned are reported as orphan commits in the HTML report.

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Config file not found | Ensure `config.local.json` or `config.json` exists in the project directory |
| GitHub API token not configured | Set `github.apiToken` in `config.local.json` |
| GitHub API rate limiting | Lower `maxCommitsToFetch` in the config |
| Branch not found | Script automatically falls back to the develop branch and flags it in the report |
| `openpyxl` / `httpx` not found | Run `pip install -r requirements.txt` |
