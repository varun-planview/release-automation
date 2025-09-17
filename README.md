# Release Automation Script

## Overview

The **Process-ReleaseReport.ps1** script is a comprehensive PowerShell automation tool designed to streamline release management processes by:

- Processing Excel-based release reports from project management systems
- Validating PAP IDs across multiple GitHub repositories using GitHub API
- Generating multi-sheet Excel reports with validation status and GitHub repository information
- Creating formatted email body text for release communications
- Supporting multiple repository types with different validation strategies
- Detecting orphan commits that exist in repositories but are missing from reports

## Features

### 🔧 **Multi-Repository GitHub API Support**
- **Branch-based validation**: For primary repositories (PVE Web, PVE Analytics) - validates across current, previous, and develop branches
- **Develop-based validation**: For secondary repositories (Dovetail, ActionBoard) - validates against develop/main/master branches only
- **GitHub API integration**: Direct API calls for commit retrieval and repository information
- **Configurable commit limits**: Control how many commits to fetch per branch for performance optimization

### 📊 **Multi-Sheet Excel Processing**
- **Report Sheet**: Main report with PAP ID validation and assignee information
- **GitHub Sheet**: Repository information including branch, last commit date, and open PR counts
- **Orphan Commits Sheet**: PAP IDs found in repositories but missing from the report
- Automatically handles file path generation based on input file location

### 🔍 **Enhanced PAP ID Validation**
- Extracts PAP IDs from GitHub commit messages using pattern matching (`PAP-\d+`)
- Cross-references PAP IDs between Excel data and GitHub repositories
- Provides detailed validation status with smart filtering logic:
  - ✅ **OK**: PAP ID found in appropriate branches
  - ⚠️ **Not found**: PAP ID missing from repository (hidden when other statuses exist)
  - ⚠️ **Warnings**: Found in unexpected branches or extra commits detected
- Supports multiple release versions per task with intelligent repository matching

### 📧 **Email Body Generation**
- Automatically generates structured email content with:
  - **TO DO**: Development tasks (items not in completion statuses)
  - **QE TASKS**: Quality Engineering tasks (items in testing statuses)
  - **FINAL PM / TR / UX REVIEW**: Items requiring final review
- Groups tasks by team for organized communication
- Handles release version formatting for cross-repository tasks

### ⚙️ **Configuration-Driven with Local Support**
- External JSON configuration files (`config.json` for templates, `config.local.json` for local development)
- `config.local.json` takes priority and is git-ignored for security
- GitHub API token configuration with validation
- Configurable commit fetch limits for performance tuning
- No hardcoded values in the script

## Prerequisites

### Required PowerShell Module
```powershell
Install-Module -Name ImportExcel
```

### Required Files
1. **Process-ReleaseReport.ps1** - Main script file
2. **config.json** - Template configuration file
3. **config.local.json** - Local configuration file (recommended for development)

### GitHub Access
- GitHub API token with repository access
- Access to configured GitHub organizations and repositories

## Configuration

### Configuration Files

The script supports two configuration files with priority-based loading:

1. **config.local.json** (Priority 1) - Local development configuration
   - Git-ignored for security (contains real API tokens)
   - Used for personal/local development
   - Takes priority if present

2. **config.json** (Priority 2) - Template/shared configuration
   - Contains placeholder values
   - Safe to commit to repository
   - Used as fallback when local config doesn't exist

### config.json / config.local.json Structure

```json
{
  "github": {
    "apiToken": "YOUR_GITHUB_TOKEN_HERE",
    "apiBaseUrl": "https://api.github.com",
    "maxCommitsToFetch": 1000
  },
  "repositories": [
    {
      "name": "PVE Web",
      "githubOrg": "pv-e1",
      "githubRepo": "pve-web",
      "validationType": "branch-based",
      "releaseVersion": "PRM October 2025",
      "currentBranch": "PRM_October2025",
      "previousBranch": "PRM_September2025",
      "developBranch": "develop"
    },
    {
      "name": "PVE Analytics",
      "githubOrg": "pv-e1", 
      "githubRepo": "pve-analytics",
      "validationType": "branch-based",
      "releaseVersion": "PRM October 2025",
      "currentBranch": "PRM_October2025",
      "previousBranch": "PRM_September2025",
      "developBranch": "develop"
    },
    {
      "name": "Dovetail",
      "githubOrg": "pv-platforma",
      "githubRepo": "dovetail",
      "validationType": "develop-based",
      "releaseVersion": "Overviews DT 2025.10",
      "developBranch": "master"
    },
    {
      "name": "ActionBoard",
      "githubOrg": "pv-e1",
      "githubRepo": "actionboard-api",
      "validationType": "develop-based",
      "releaseVersion": "Overviews AB 2025.10",
      "developBranch": "main"
    }
  ],
  "BaseURL": "https://planview.leankit.com/card",
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

### Configuration Parameters

| Parameter | Description |
|-----------|-------------|
| **GitHub Configuration** | |
| `github.apiToken` | GitHub API token for repository access |
| `github.apiBaseUrl` | GitHub API base URL (usually https://api.github.com) |
| `github.maxCommitsToFetch` | Maximum commits to fetch per branch (default: 1000) |
| **Repository Configuration** | |
| `repositories` | Array of repository configurations |
| `name` | Display name for the repository |
| `githubOrg` | GitHub organization name |
| `githubRepo` | GitHub repository name |
| `validationType` | Either "branch-based" or "develop-based" |
| `releaseVersion` | Release version identifier that matches Excel data |
| `currentBranch` | Current release branch (branch-based only) |
| `previousBranch` | Previous release branch (branch-based only) |
| `developBranch` | Development branch name |
| **General Configuration** | |
| `BaseURL` | Base URL for generating card links |
| `nonDevStatuses` | Statuses considered non-development work |
| `qeStatuses` | Statuses considered QE/testing work |

## Usage

### Initial Setup

1. **Copy configuration template**:
   ```powershell
   # Create local configuration from template
   Copy-Item config.json config.local.json
   ```

2. **Configure GitHub API token** in `config.local.json`:
   ```json
   {
     "github": {
       "apiToken": "your_actual_github_token_here"
     }
   }
   ```

3. **Update repository configurations** as needed for your release versions and branches

### Running the Script

1. **Execute the script**:
   ```powershell
   .\Process-ReleaseReport.ps1
   ```

2. **Select input Excel file** using the file dialog that appears

3. **Review console output** for processing status and any warnings

4. **Check generated files**:
   - `[InputFileName]_OUTPUT.xlsx` - Multi-sheet Excel report with:
     - **Report**: Main validation results
     - **GitHub**: Repository information and PR counts
     - **Orphan Commits**: PAP IDs in repos but missing from report
   - `[InputFileName]_EmailBody.txt` - Formatted email content

### Expected Input Excel Format

The input Excel file should contain columns:
- `Card ID` - Numeric identifier for the task
- `Team` - Team responsible for the task
- `Card Title` - Task description
- `Assignee(s)` - Task assignees (format: "LastName, FirstName")
- `Card Type` - Type of task/card
- `Current Lane Title` - Current status/lane
- `Planning Increment Label` - Release version(s)
- `Tags` - Task tags

### Output Files

#### Multi-Sheet Excel Output
- **Report Sheet**: Original columns plus additional computed columns:
  - `PAP ID` - Formatted PAP identifier
  - `Card URL` - Direct link to the card
  - `Validation Comment` - Git validation status across all matching repositories
  - `Card Assignee(s)` - Formatted assignee information with URL
  - `Release Version` - Renamed from "Planning Increment Label"

- **GitHub Sheet**: Repository monitoring information:
  - `Repository Name` - Repository display name
  - `Branch` - Target branch being monitored
  - `Last Commit Date` - Date of most recent commit
  - `Open PR Count` - Number of open pull requests targeting the branch

- **Orphan Commits Sheet** (if orphans found): 
  - `Card ID` - Extracted card identifier
  - `PAP ID` - Full PAP identifier
  - `Card URL` - Link to the card
  - `Commit URL` - Direct link to the GitHub commit
  - `Validation Comment` - Description of orphan status

#### Email Body Output
Structured text file with three sections:
1. **TO DO** - Development tasks grouped by team
2. **QE TASKS** - Testing tasks grouped by team  
3. **FINAL PM / TR / UX REVIEW** - Review tasks grouped by team

## Validation Logic

### Smart Filtering Algorithm
The script uses intelligent filtering to show only relevant validation results:
- **Priority Display**: If any repositories show success (✅) or warnings (⚠️), "not found" results are hidden
- **Fallback Display**: If no repositories show success/warnings, all "not found" results are shown
- **Multiple Release Versions**: Each release version group is processed separately with its own filtering logic

### Branch-Based Validation (PVE Web, PVE Analytics)
- **✅ OK**: PAP ID found in current branch
- **⚠️ Warning**: PAP ID found in previous branch (indicates potential rollback issue)
- **⚠️ Warning**: Extra commits in develop branch not in current branch (indicates missing merge)
- **⚠️ Warning**: Found in develop but not in current branch (indicates incomplete merge)
- **⚠️ Not found**: PAP ID not found in any branch (hidden when other statuses exist)

### Develop-Based Validation (Dovetail, ActionBoard)
- **✅ OK**: PAP ID found in develop/main/master branch
- **⚠️ Not found**: PAP ID not found in develop/main/master branch (hidden when other statuses exist)

### GitHub Repository Information
- **Last Commit Date**: Retrieved from GitHub API for the monitored branch
- **Open PR Count**: Number of open pull requests targeting the monitored branch
- **Branch-specific**: Uses `currentBranch` for branch-based repos, `developBranch` for develop-based repos

### Orphan Detection
- Identifies PAP IDs that exist in current branch but are missing from the Excel report
- Only performed for branch-based repositories (PVE Web)
- Provides direct links to both the card and the GitHub commit
- Indicates potential missing tasks in release planning

## Troubleshooting

### Common Issues

1. **"Configuration file not found"**
   - Ensure `config.local.json` or `config.json` exists in the same directory as the script
   - Create `config.local.json` by copying `config.json` and updating with real values

2. **"GitHub API token not configured"**
   - Set a valid GitHub API token in `config.local.json`
   - Ensure the token has access to the configured repositories

3. **"GitHub organization not configured"**
   - Verify all repositories have valid `githubOrg` and `githubRepo` values
   - Check that the GitHub organization names are correct

4. **ImportExcel module errors**
   - Install the module: `Install-Module -Name ImportExcel`
   - Run PowerShell as Administrator if needed

5. **GitHub API rate limiting**
   - Reduce `maxCommitsToFetch` value for better performance
   - Ensure you're using a personal access token with appropriate permissions

6. **Incorrect PR counts**
   - Verify branch names in configuration match actual GitHub branches
   - Check that the API token has access to pull request information

### Performance Considerations

- **API Rate Limits**: GitHub API has rate limits; reduce `maxCommitsToFetch` if hitting limits
- **Processing Time**: Scales with number of commits, repositories, and API response time
- **Progress Indicators**: Console shows completion percentage and API call progress
- **Configurable Limits**: Adjust `maxCommitsToFetch` based on your repository size and performance needs

### Security Best Practices

- **Never commit `config.local.json`**: Contains sensitive API tokens
- **Use `.gitignore`**: Ensure `config.local.json` is properly ignored
- **Token Permissions**: Use minimal required GitHub token permissions
- **Template Configuration**: Keep `config.json` with placeholder values for sharing

## Version History & Development Evolution

### Current Version: v5.0
**Major GitHub API Integration Release**

#### Key Features:
- **GitHub API Integration**: Direct API calls replace local Git operations
- **Multi-Sheet Excel Export**: Report, GitHub, and Orphan Commits sheets
- **Configurable Performance**: Adjustable commit fetch limits
- **Local Configuration Support**: `config.local.json` for secure local development
- **Enhanced Repository Monitoring**: Real-time GitHub repository information
- **Smart Validation Filtering**: Intelligent result display logic

### Version History

- **v1.0**: Initial version with hardcoded configuration and local Git operations
- **v2.0**: Configuration externalization to JSON file
- **v3.0**: Multi-repository support with different validation types
- **v4.0**: Terminology updates and repository name improvements  
- **v4.1**: Enhanced validation logic and orphan detection
- **v5.0**: Complete GitHub API integration with multi-sheet reporting

### Development Evolution

#### Phase 1: Configuration Externalization (v1.0 → v2.0)
**Initial Request**: "Move CONFIGURABLE VARIABLES to configuration file"
- Externalized hardcoded variables (repository paths, URLs, status arrays) to JSON
- Improved maintainability and deployment flexibility

#### Phase 2: Multi-Repository Architecture (v2.0 → v3.0)  
- Added support for multiple repositories with different validation strategies
- Implemented branch-based validation (PVE Web) vs develop-based validation (Dovetail, ActionBoard)
- Enhanced PAP ID cross-referencing across repository boundaries

#### Phase 3: Professional Terminology (v3.0 → v4.0)
- Renamed "Planning Increment Label" to "Release Version" for clarity
- Updated repository display names from technical to user-friendly:
  - "PRM" → "PVE Web"
  - "Overviews DT" → "Dovetail"
  - "Overviews AB" → "ActionBoard"

#### Phase 4: Enhanced Validation (v4.0 → v4.1)
- Implemented comprehensive PAP ID validation with detailed status reporting
- Added orphan detection for commits missing from reports
- Improved email body generation with team grouping

#### Phase 5: GitHub API Integration (v4.1 → v5.0)
**Major Architecture Shift**: Local Git → GitHub API
- **API-First Approach**: Replaced local Git commands with direct GitHub API calls
- **Enhanced Performance**: Configurable commit limits and optimized API usage
- **Multi-Sheet Reporting**: Separated concerns into Report, GitHub, and Orphan sheets
- **Repository Monitoring**: Real-time branch information and PR counts
- **Security Enhancement**: Local configuration support with git-ignore protection
- **Smart Filtering**: Intelligent validation result display logic

#### Technical Evolution Summary:
1. **Configuration Management**: Hardcoded → JSON → Local + Template
2. **Data Source**: Local Git → GitHub API  
3. **Repository Support**: Single → Multiple with validation strategies
4. **Output Format**: Single sheet → Multi-sheet with specialized content
5. **Performance**: Fixed limits → Configurable optimization
6. **Security**: Basic → Token-based with local configuration protection

### Architecture Highlights

**Current Technical Stack**:
- **PowerShell Core**: Cross-platform automation scripting
- **GitHub API v3**: Direct repository and commit data access
- **ImportExcel Module**: Multi-sheet Excel processing
- **JSON Configuration**: Hierarchical configuration management
- **REST API Integration**: Paginated commit retrieval and repository information
- **Pattern Matching**: Regular expression PAP ID extraction
- **Smart Filtering**: Context-aware validation result display

**Integration Points**:
- GitHub Organizations and Repositories
- LeanKit Card Management System
- Excel-based Project Management Reports
- Email Communication Systems

This evolution demonstrates a progression from simple configuration externalization to a comprehensive, API-driven release automation platform supporting multiple repositories, validation strategies, and output formats while maintaining security and performance optimization.

## License

This script is provided as-is for internal use. Modify and distribute according to your organization's policies.

## Contributing

When contributing to this project:

1. **Configuration**: Always use `config.local.json` for development
2. **Security**: Never commit API tokens or sensitive configuration
3. **Testing**: Test with multiple repository types and release versions
4. **Documentation**: Update README for any new features or configuration changes

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Verify configuration file format and required fields
3. Test with minimal configuration to isolate issues
4. Review console output for specific error messages
