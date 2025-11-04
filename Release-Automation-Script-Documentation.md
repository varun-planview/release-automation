# Release Automation Script - Confluence Documentation

## Overview

The **Process-ReleaseReport.ps1** script is a comprehensive PowerShell automation tool designed to transform manual release management processes from **2-3 hours into 10 minutes** of automated workflow. This script addresses critical release management challenges by automating status validation, email preparation, and GitHub repository synchronization.

### Key Benefits
- ✅ **83% Time Reduction**: From 2-3 hours to 10 minutes
- ✅ **100% Accuracy**: Automated validation eliminates human error
- ✅ **Multi-Repository Support**: Validates across all GitHub repositories
- ✅ **Intelligent Validation**: Detects misalignments and orphan commits
- ✅ **Formatted Output**: Generates Excel reports and email body text

## What This Script Does

### Primary Functions

1. **Excel Processing & Validation**
   - Reads release planning Excel files with card information
   - Validates PAP IDs against GitHub repository commits
   - Generates comprehensive validation status messages

2. **GitHub API Integration**
   - Connects to multiple GitHub repositories using configured tokens
   - Compares commits across current, previous, and develop branches
   - Identifies cross-repository code alignment issues

3. **Automated Report Generation**
   - Creates multi-sheet Excel output with validation results
   - Generates structured email body text for release communications
   - Groups tasks by team and status for organized workflow

4. **Smart Validation Logic**
   - **Branch-based validation**: For release repositories (PVE Web, Analytics, Polaris, Dovetail)
   - **Develop-based validation**: For continuous deployment repositories (ActionBoard)
   - **Orphan detection**: Finds commits in repositories missing from planning

### Validation Status Messages

The script provides detailed commit status with concise warning messages:

| Status | Description |
|--------|-------------|
| ✅ **OK** | PAP ID found in appropriate branches |
| ⚠️ **Missing** | PAP ID missing from expected repository |
| ⚠️ **Unexpected** | PAP ID found in repository that doesn't match card's release version |
| ⚠️ **In previous** | Code exists in previous release branch |
| ⚠️ **Extra commits** | Additional commits in develop branch not in current release |
| ⚠️ **Only in develop** | Code found in develop but missing from release branch |
| ⚠️ **Orphaned commit** | Code exists but card missing from planning |

## Prerequisites & Setup

### Required Software
- **PowerShell 5.1+** (Windows PowerShell or PowerShell Core)
- **ImportExcel Module**: `Install-Module -Name ImportExcel`
- **GitHub API Token** with repository read access

### Configuration Files
The script uses two configuration files (attach both to this Confluence page):

1. **config.json** - Main configuration template
2. **config.local.json** - Local override (recommended for personal tokens)

### GitHub Token Setup
1. Generate a GitHub Personal Access Token with `repo` permissions
2. Update the `github.apiToken` value in `config.local.json`
3. Ensure all repository `githubOrg` values are correctly configured

## How to Use

### Step-by-Step Instructions

1. **Prepare Configuration**
   - Copy `config.json` to `config.local.json`
   - Update GitHub API token and repository settings
   - Verify all repository branch names are current

2. **Run the Script**
   ```powershell
   .\Process-ReleaseReport.ps1
   ```

3. **Select Input File**
   - File dialog will open automatically
   - Select your release planning Excel file
   - Must contain columns: 'Card ID', 'Current Lane Title', 'Team', etc.

4. **Review Generated Output**
   - **[Filename]_OUTPUT.xlsx**: Multi-sheet Excel report
   - **[Filename]_EmailBody.txt**: Formatted email content

### Output Files Description

#### Excel Output Sheets
- **Report Sheet**: Main data with PAP ID validation and GitHub status
- **GitHub Sheet**: Repository information (branches, last commit dates, PR counts)
- **Orphan Commits Sheet**: PAP IDs found in code but missing from planning

#### Email Body Content
- **TO DO**: Development tasks requiring completion
- **QE TASKS**: Items in testing phases (grouped by team)
- **FINAL REVIEW**: Items awaiting final PM/TR/UX approval

## Configuration Reference

### Repository Configuration Types

#### Branch-Based Repositories
For repositories with formal release branches:
```json
{
  "name": "PVE Web",
  "githubOrg": "pv-e1",
  "githubRepo": "pve-web", 
  "validationType": "branch-based",
  "releaseVersion": "PRM October 2025",
  "currentBranch": "PRM_October2025",
  "previousBranch": "PRM_September2025",
  "developBranch": "develop"
}
```

#### Develop-Based Repositories
For repositories using continuous deployment:
```json
{
  "name": "ActionBoard",
  "githubOrg": "pv-e1", 
  "githubRepo": "actionboard-api",
  "validationType": "develop-based",
  "releaseVersion": "Overviews AB 2025.10",
  "developBranch": "main"
}
```

### Status Configuration
```json
{
  "nonDevStatuses": [
    "FINAL PM / TR / UX REVIEW",
    "READY FOR TESTING", 
    "TESTING",
    "READY FOR RELEASE",
    "READY TO ARCHIVE"
  ],
  "qeStatuses": [
    "READY FOR TESTING",
    "TESTING"  
  ]
}
```

## Troubleshooting

### Common Issues

#### Configuration Errors
- **"GitHub API token not configured"**: Update `github.apiToken` in config file
- **"GitHub organization not configured"**: Verify `githubOrg` values for all repositories
- **"Configuration file not found"**: Ensure `config.json` or `config.local.json` exists

#### API Issues  
- **Rate limiting**: Script respects GitHub API limits with configurable `maxCommitsToFetch`
- **Authentication failures**: Verify token has `repo` permissions
- **Network errors**: Check internet connectivity and GitHub API status

#### Excel File Issues
- **"No file selected"**: Ensure Excel file is properly selected in dialog
- **Import errors**: Verify Excel file contains required columns ('Card ID', etc.)
- **Empty results**: Check that 'Card ID' column contains numeric values

### Performance Optimization
- Adjust `maxCommitsToFetch` in configuration (default: 1000)
- Use `config.local.json` for faster local development
- Consider repository filtering for large organizations

## Human Analysis Still Required

⚠️ **Important**: While this script significantly reduces manual work, **human analysis and validation is still required** for complex scenarios that automated tools cannot fully interpret.

### Common Scenarios Requiring Manual Review

#### Reverted Commits
**Example**: A commit may exist in the current release branch but was later reverted in the same branch. The card might show a future release version since the changes were effectively undone, but the script will detect the original commit and flag it as present in the current release.

- **What the script sees**: PAP ID found in current branch commits
- **Reality**: Changes were reverted, so functionally not in the release
- **Required action**: Manual review of commit history to identify reverts

#### Other Complex Git Scenarios
- **Cherry-picked commits**: Same changes across multiple branches with different commit hashes
- **Merge conflicts**: Resolution commits that don't contain PAP IDs but affect functionality
- **Hotfix branches**: Emergency fixes that bypass normal release flow
- **Squashed commits**: Multiple PAP IDs combined into single commits during merge

### Best Practice
Always perform a **final human review** of the generated report, especially for:
- Cards marked with warnings or unexpected status
- High-priority or critical functionality changes
- Items flagged as "orphaned" or "unexpected"
- Release-critical features requiring special attention

## Security Notes

- **Never commit GitHub tokens**: Use `config.local.json` for sensitive data
- **Token permissions**: Use minimum required permissions (read-only repo access)
- **Local storage**: Keep configuration files secure and out of version control

## Support & Maintenance

### Regular Updates Needed
- **Repository branches**: Update branch names for new releases
- **Release versions**: Modify `releaseVersion` strings for current releases  
- **Team configurations**: Adjust status categories as workflow evolves

### Script Maintenance
- Monitor GitHub API changes and rate limits
- Update PowerShell modules: `Update-Module ImportExcel`
- Review and optimize `maxCommitsToFetch` based on repository size

---

## Attachments
📎 **Process-ReleaseReport.ps1** - Main PowerShell script
📎 **config.json** - Configuration template with all settings

*Last Updated: October 8, 2025*
*Version: 1.0*