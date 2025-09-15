# Release Automation Script

## Overview

The **Process-ReleaseReport.ps1** script is a comprehensive PowerShell automation tool designed to streamline release management processes by:

- Processing Excel-based release reports from project management systems
- Validating PAP IDs (Project Action Points) across multiple Git repositories
- Generating augmented Excel reports with validation status
- Creating formatted email body text for release communications
- Supporting multiple repository types with different validation strategies

## Features

### 🔧 **Multi-Repository Support**
- **Branch-based validation**: For primary repositories (PVE Web) - validates across current, previous, and develop branches
- **Develop-based validation**: For secondary repositories (Dovetail, ActionBoard) - validates against develop/main/master branches only
- **Configurable repository paths** and branch names

### 📊 **Excel Processing**
- Imports Excel files containing release task information
- Filters and processes Card IDs, assignees, and release versions
- Exports augmented Excel files with validation comments
- Automatically handles file path generation based on input file location

### 🔍 **PAP ID Validation**
- Extracts PAP IDs from Git commit messages using pattern matching (`PAP-\d+`)
- Cross-references PAP IDs between Excel data and Git repositories
- Provides detailed validation status:
  - ✅ **OK**: PAP ID found in appropriate branches
  - ❌ **Not found**: PAP ID missing from repository
  - ⚠️ **Warnings**: Found in unexpected branches or extra commits detected

### 📧 **Email Body Generation**
- Automatically generates structured email content with:
  - **TO DO**: Development tasks (items not in completion statuses)
  - **QE TASKS**: Quality Engineering tasks (items in testing statuses)
  - **FINAL PM / TR / UX REVIEW**: Items requiring final review
- Groups tasks by team for organized communication
- Handles release version formatting for cross-repository tasks

### ⚙️ **Configuration-Driven**
- External JSON configuration file (`config.json`) for easy maintenance
- No hardcoded values in the script
- Supports adding new repositories without code changes

## Prerequisites

### Required PowerShell Module
```powershell
Install-Module -Name ImportExcel
```

### Required Files
1. **Process-ReleaseReport.ps1** - Main script file
2. **config.json** - Configuration file (must be in same directory as script)

### Git Repositories
- All configured repository paths must be accessible
- Repositories should be up-to-date with latest commits

## Configuration

### config.json Structure

```json
{
  "repositories": [
    {
      "name": "PVE Web",
      "path": "C:\\Dev\\pve-web",
      "validationType": "branch-based",
      "releaseVersion": "PRM October 2025",
      "currentBranch": "PRM_October2025",
      "previousBranch": "PRM_September2025", 
      "developBranch": "develop"
    },
    {
      "name": "Dovetail",
      "path": "C:\\Dev\\dovetail",
      "validationType": "develop-based",
      "releaseVersion": "Overviews DT 2025.10",
      "developBranch": "master"
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
| `repositories` | Array of repository configurations |
| `name` | Display name for the repository |
| `path` | Absolute path to the Git repository |
| `validationType` | Either "branch-based" or "develop-based" |
| `releaseVersion` | Release version identifier that matches Excel data |
| `currentBranch` | Current release branch (branch-based only) |
| `previousBranch` | Previous release branch (branch-based only) |
| `developBranch` | Development branch name |
| `BaseURL` | Base URL for generating card links |
| `nonDevStatuses` | Statuses considered non-development work |
| `qeStatuses` | Statuses considered QE/testing work |

## Usage

### Running the Script

1. **Execute the script**:
   ```powershell
   .\Process-ReleaseReport.ps1
   ```

2. **Select input Excel file** using the file dialog that appears

3. **Review console output** for processing status and any warnings

4. **Check generated files**:
   - `[InputFileName]_OUTPUT.xlsx` - Augmented Excel report
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

#### Excel Output
- **Original columns** plus additional computed columns:
  - `PAP ID` - Formatted PAP identifier
  - `Card URL` - Direct link to the card
  - `Validation Comment` - Git validation status
  - `Card Assignee(s)` - Formatted assignee information with URL
  - `Release Version` - Renamed from "Planning Increment Label"

#### Email Body Output
Structured text file with three sections:
1. **TO DO** - Development tasks grouped by team
2. **QE TASKS** - Testing tasks grouped by team  
3. **FINAL PM / TR / UX REVIEW** - Review tasks grouped by team

## Validation Logic

### Branch-Based Validation (PVE Web)
- **✅ OK**: PAP ID found in current branch
- **⚠️ Warning**: PAP ID found in previous branch (potential issue)
- **⚠️ Warning**: Extra commits in develop branch not in current branch
- **⚠️ Warning**: Found in develop but not in current branch
- **❌ Not found**: PAP ID not found in any branch

### Develop-Based Validation (Dovetail/ActionBoard)
- **✅ OK**: PAP ID found in develop/main/master branch
- **❌ Not found**: PAP ID not found in develop/main/master branch

### Orphan Detection
- Identifies PAP IDs that exist in current branch but are missing from the Excel report
- Only performed for branch-based repositories
- Adds orphan entries to the output Excel file

## Troubleshooting

### Common Issues

1. **"Configuration file not found"**
   - Ensure `config.json` exists in the same directory as the script

2. **"Repository path not found"**
   - Verify repository paths in config.json are correct and accessible
   - Update paths if repositories have been moved

3. **ImportExcel module errors**
   - Install the module: `Install-Module -Name ImportExcel`
   - Run PowerShell as Administrator if needed

4. **Git command failures**
   - Ensure Git is installed and accessible from PowerShell
   - Verify repository paths are valid Git repositories

### Performance Considerations

- Large repositories may take longer to process
- Processing time scales with number of commits and repositories
- Progress indicators show completion percentage

## Original Development Context

### Detailed Prompt Used to Generate This Script

**Initial Request**: 
"Move CONFIGURABLE VARIABLES to configuration file"

**Evolution Through Development**:

1. **Configuration Externalization**: Initially focused on moving hardcoded variables (repository paths, URLs, status arrays) to an external JSON configuration file for better maintainability.

2. **Multi-Repository Support**: Expanded to support multiple repositories with different validation approaches:
   - PVE Web repository using branch-based validation (current, previous, develop branches)
   - Dovetail and ActionBoard repositories using develop-based validation (single branch)

3. **Terminology Updates**: Renamed "Planning Increment Label" to "Release Version" throughout the system for better clarity and professional presentation.

4. **Repository Name Refinement**: Updated repository display names from technical identifiers to user-friendly names:
   - "PRM" → "PVE Web"
   - "Overviews DT" → "Dovetail" 
   - "Overviews AB" → "ActionBoard"

5. **Validation Enhancement**: Implemented comprehensive PAP ID validation across all repositories with detailed status reporting and orphan detection.

6. **Output Optimization**: Added email body generation with proper team grouping and release version handling for cross-repository tasks.

**Final Implementation Features**:
- JSON-driven configuration system
- Multi-repository Git commit analysis
- PAP ID pattern matching and validation
- Excel import/export with data augmentation
- Automated email body generation
- Progress tracking and error handling
- Flexible validation strategies per repository type

**Technical Approach**:
- PowerShell with ImportExcel module for Excel processing
- Git command-line integration for commit history analysis
- JSON configuration for maintainable deployment
- Regular expression pattern matching for PAP ID extraction
- Structured output generation for both Excel and email formats

This script evolved from a simple configuration externalization request into a comprehensive release automation solution supporting multiple repositories, validation strategies, and output formats while maintaining professional terminology and user-friendly repository naming conventions.

## Version History

- **v1.0**: Initial version with hardcoded configuration
- **v2.0**: Configuration externalization to JSON file
- **v3.0**: Multi-repository support with different validation types
- **v4.0**: Terminology updates and repository name improvements
- **v4.1**: Current version with comprehensive validation and reporting

## License

This script is provided as-is for internal use. Modify and distribute according to your organization's policies.
