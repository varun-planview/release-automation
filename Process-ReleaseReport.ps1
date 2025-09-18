# Requires: ImportExcel module (Install-Module -Name ImportExcel)

# ------------------ LOAD CONFIGURATION ------------------
$configPath = Join-Path $PSScriptRoot "config.local.json"
$fallbackConfigPath = Join-Path $PSScriptRoot "config.json"

# Check for local config first, then fall back to main config
if (Test-Path $configPath) {
    Write-Host "Using local configuration: $configPath" -ForegroundColor Green
} elseif (Test-Path $fallbackConfigPath) {
    $configPath = $fallbackConfigPath
    Write-Host "Using default configuration: $configPath" -ForegroundColor Yellow
} else {
    Write-Host "Configuration file not found at: $configPath or $fallbackConfigPath" -ForegroundColor Red
    Write-Host "Please ensure config.local.json or config.json exists in the same directory as this script." -ForegroundColor Red
    exit 1
}

try {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    
    # Load configuration variables
    $repositories = $config.repositories
    $BaseURL = $config.BaseURL
    $nonDevStatuses = $config.nonDevStatuses
    $qeStatuses = $config.qeStatuses
    $github = $config.github
    
    # Validate GitHub configuration
    if (-not $github.apiToken -or $github.apiToken -eq "YOUR_GITHUB_TOKEN_HERE") {
        Write-Host "ERROR: GitHub API token not configured in config.json" -ForegroundColor Red
        Write-Host "Please set the 'github.apiToken' value in config.json" -ForegroundColor Red
        exit 1
    }
    
    # Set default maxCommitsToFetch if not specified
    if (-not $github.maxCommitsToFetch) {
        $github.maxCommitsToFetch = 1000
        Write-Host "Using default maxCommitsToFetch: 1000" -ForegroundColor Yellow
    } elseif ($github.maxCommitsToFetch -le 0) {
        Write-Host "ERROR: maxCommitsToFetch must be greater than 0" -ForegroundColor Red
        exit 1
    }
    
    # Validate each repository has GitHub organization configured
    foreach ($repo in $repositories) {
        if (-not $repo.githubOrg -or $repo.githubOrg -eq "YOUR_GITHUB_ORG_HERE") {
            Write-Host "ERROR: GitHub organization not configured for repository: $($repo.name)" -ForegroundColor Red
            Write-Host "Please set the 'githubOrg' value for $($repo.name) in config.json" -ForegroundColor Red
            exit 1
        }
    }
    
    Write-Host "Configuration loaded successfully from: $configPath" -ForegroundColor Green
    Write-Host "Found $($repositories.Count) repositories configured" -ForegroundColor Green
    Write-Host "GitHub API configured with token" -ForegroundColor Green
    Write-Host "Max commits to fetch per branch: $($github.maxCommitsToFetch)" -ForegroundColor Green
} catch {
    Write-Host "Error reading configuration file: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
# --------------------------------------------------------

# Prompt user to select input file
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$OpenFileDialog.Title = "Select the input Excel file"
$null = $OpenFileDialog.ShowDialog()
$InputExcelPath = $OpenFileDialog.FileName

if (-not $InputExcelPath -or !(Test-Path $InputExcelPath)) {
    Write-Host "No file selected or file does not exist. Exiting script."
    exit
}

Write-Host "Selected input file: $InputExcelPath"

# Dynamically set output paths
$inputDir = Split-Path $InputExcelPath
$inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputExcelPath)
$OutputExcelPath = Join-Path $inputDir ("${inputBase}_OUTPUT.xlsx")
$EmailBodyPath = Join-Path $inputDir ("${inputBase}_EmailBody.txt")

Write-Host "Output Excel will be: $OutputExcelPath"
Write-Host "Email body will be: $EmailBodyPath"

Write-Host "Starting report generation..."

Write-Host "Reading Excel file..."
$rows = Import-Excel $InputExcelPath

Write-Host "Filtering unwanted rows..."
$rows = $rows | Where-Object { $_.'Card ID' -match "\d+" }

# ------------------ GITHUB API FUNCTIONS ------------------

function Get-GitHubCommits {
    param (
        [string]$owner,
        [string]$repo,
        [string]$branch,
        [string]$token,
        [int]$perPage = 100,
        [int]$maxCommits = 1000
    )
    
    $headers = @{
        "Authorization" = "token $token"
        "Accept" = "application/vnd.github.v3+json"
        "User-Agent" = "PowerShell-Release-Report"
    }
    
    $commits = @()
    $page = 1
    
    do {
        $uri = "$($github.apiBaseUrl)/repos/$owner/$repo/commits?sha=$branch&per_page=$perPage&page=$page"
        
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
            
            if ($response -and $response.Count -gt 0) {
                # Store both commit message and full commit object for later use
                $commits += $response | ForEach-Object { 
                    [PSCustomObject]@{
                        Message = $_.commit.message
                        HtmlUrl = $_.html_url
                        Sha = $_.sha
                    }
                }
                $page++
            } else {
                break
            }
        } catch {
            Write-Warning "GitHub API call failed for $uri : $_"
            break
        }
    } while ($response.Count -eq $perPage -and ($commits.Count -lt $maxCommits)) # Use configurable max commits limit
    
    return $commits
}

function Compare-GitHubBranches {
    param (
        [string]$owner,
        [string]$repo,
        [string]$baseBranch,
        [string]$headBranch,
        [string]$token
    )
    
    $headers = @{
        "Authorization" = "token $token"
        "Accept" = "application/vnd.github.v3+json"
        "User-Agent" = "PowerShell-Release-Report"
    }
    
    $uri = "$($github.apiBaseUrl)/repos/$owner/$repo/compare/$baseBranch...$headBranch"
    
    try {
        $comparison = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        
        if ($comparison -and $comparison.commits) {
            return $comparison.commits | ForEach-Object { 
                [PSCustomObject]@{
                    Message = $_.commit.message
                    HtmlUrl = $_.html_url
                    Sha = $_.sha
                }
            }
        }
    } catch {
        Write-Warning "GitHub branch comparison failed for $uri : $_"
    }
    
    return @()
}

function Get-GitHubRepoInfo {
    param (
        [string]$owner,
        [string]$repo,
        [string]$branch,
        [string]$token
    )
    
    $headers = @{
        "Authorization" = "token $token"
        "Accept" = "application/vnd.github.v3+json"
        "User-Agent" = "PowerShell-Release-Report"
    }
    
    $repoInfo = [PSCustomObject]@{
        LastCommitDate = ""
        OpenPRCount = 0
    }
    
    try {
        # Get last commit date for the branch
        $commitsUri = "$($github.apiBaseUrl)/repos/$owner/$repo/commits?sha=$branch&per_page=1"
        $commits = Invoke-RestMethod -Uri $commitsUri -Headers $headers -Method Get
        
        if ($commits -and $commits.Count -gt 0) {
            $repoInfo.LastCommitDate = [DateTime]::Parse($commits[0].commit.committer.date).ToString("yyyy-MM-dd HH:mm")
        }
        
        # Get open PR count for the branch
        $prsUri = "$($github.apiBaseUrl)/repos/$owner/$repo/pulls?state=open&base=$branch&per_page=100"
        $targetingPrs = Invoke-RestMethod -Uri $prsUri -Headers $headers -Method Get
        
        $repoInfo.OpenPRCount = $targetingPrs.Count
        
    } catch {
        Write-Warning "Could not retrieve repo info for $owner/$repo branch $branch : $_"
    }
    
    return $repoInfo
}

function Test-GitHubBranch {
    param (
        [string]$owner,
        [string]$repo,
        [string]$branch,
        [string]$token
    )
    
    $headers = @{
        "Authorization" = "token $token"
        "Accept" = "application/vnd.github.v3+json"
        "User-Agent" = "PowerShell-Release-Report"
    }
    
    try {
        $uri = "$($github.apiBaseUrl)/repos/$owner/$repo/branches/$branch"
        $branchInfo = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        return $true
    } catch {
        return $false
    }
}

Write-Host "Extracting commit messages from GitHub API for all repositories..."

# Initialize hash tables for PAP IDs from all repositories
$allRepositoryCommits = @{}
$gitHubRepoInfo = @()

foreach ($repo in $repositories) {
    Write-Host "Processing repository: $($repo.name) ($($repo.githubOrg)/$($repo.githubRepo))" -ForegroundColor Cyan
    
    # Validate GitHub repository configuration
    if (-not $repo.githubRepo) {
        Write-Host "Warning: GitHub repository name not configured for $($repo.name)" -ForegroundColor Yellow
        continue
    }
    
    if (-not $repo.githubOrg) {
        Write-Host "Warning: GitHub organization not configured for $($repo.name)" -ForegroundColor Yellow
        continue
    }
    
    $repoCommits = @{}
    
    try {
        if ($repo.validationType -eq "branch-based") {
            # Check if current branch exists, if not use develop branch as current
            $currentBranchExists = Test-GitHubBranch -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.currentBranch -token $github.apiToken
            $effectiveCurrentBranch = $repo.currentBranch
            
            if (-not $currentBranchExists) {
                Write-Host "  Current branch ($($repo.currentBranch)) not found - using develop branch ($($repo.developBranch)) as current" -ForegroundColor Yellow
                $effectiveCurrentBranch = $repo.developBranch
            }
            
            # For branch-based repos - extract from current, previous, and develop branches
            Write-Host "  Fetching commits from current branch: $effectiveCurrentBranch" -ForegroundColor DarkGray
            $repoCommits.current = Get-GitHubCommits -owner $repo.githubOrg -repo $repo.githubRepo -branch $effectiveCurrentBranch -token $github.apiToken -maxCommits $github.maxCommitsToFetch
            
            # Check if previous branch exists before fetching
            $previousBranchExists = Test-GitHubBranch -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.previousBranch -token $github.apiToken
            if ($previousBranchExists) {
                Write-Host "  Fetching commits from previous branch: $($repo.previousBranch)" -ForegroundColor DarkGray
                $repoCommits.previous = Get-GitHubCommits -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.previousBranch -token $github.apiToken -maxCommits $github.maxCommitsToFetch
            } else {
                Write-Host "  Previous branch ($($repo.previousBranch)) not found - skipping previous branch validation" -ForegroundColor Yellow
                $repoCommits.previous = @()
            }
            
            # Only fetch develop branch separately if it's different from the effective current branch
            if ($effectiveCurrentBranch -ne $repo.developBranch) {
                Write-Host "  Fetching commits from develop branch: $($repo.developBranch)" -ForegroundColor DarkGray
                $repoCommits.develop = Get-GitHubCommits -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.developBranch -token $github.apiToken -maxCommits $github.maxCommitsToFetch
                
                # Get extra commits in develop branch (not in current branch)
                Write-Host "  Comparing branches for extra commits..." -ForegroundColor DarkGray
                $repoCommits.extraInDevelop = Compare-GitHubBranches -owner $repo.githubOrg -repo $repo.githubRepo -baseBranch $effectiveCurrentBranch -headBranch $repo.developBranch -token $github.apiToken
            } else {
                # If current branch is same as develop, reuse the commits and no extra commits
                Write-Host "  Using current branch commits as develop branch (same branch)" -ForegroundColor DarkGray
                $repoCommits.develop = $repoCommits.current
                $repoCommits.extraInDevelop = @()
            }
            
            # Get repository info for the effective current branch
            Write-Host "  Getting repository info for current branch..." -ForegroundColor DarkGray
            $branchInfo = Get-GitHubRepoInfo -owner $repo.githubOrg -repo $repo.githubRepo -branch $effectiveCurrentBranch -token $github.apiToken
            
            # Add to GitHub info collection
            $gitHubRepoInfo += [PSCustomObject]@{
                'Repository Name' = $repo.name
                'Branch' = $effectiveCurrentBranch
                'Last Commit Date' = $branchInfo.LastCommitDate
                'Open PR Count' = $branchInfo.OpenPRCount
            }
            
        } elseif ($repo.validationType -eq "develop-based") {
            # For develop-based repos - extract from develop branch only
            Write-Host "  Fetching commits from develop branch: $($repo.developBranch)" -ForegroundColor DarkGray
            $repoCommits.develop = Get-GitHubCommits -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.developBranch -token $github.apiToken -maxCommits $github.maxCommitsToFetch
            
            # Get repository info for develop branch (develop-based uses develop branch)
            Write-Host "  Getting repository info for develop branch..." -ForegroundColor DarkGray
            $branchInfo = Get-GitHubRepoInfo -owner $repo.githubOrg -repo $repo.githubRepo -branch $repo.developBranch -token $github.apiToken
            
            # Add to GitHub info collection
            $gitHubRepoInfo += [PSCustomObject]@{
                'Repository Name' = $repo.name
                'Branch' = $repo.developBranch
                'Last Commit Date' = $branchInfo.LastCommitDate
                'Open PR Count' = $branchInfo.OpenPRCount
            }
        }
        
        Write-Host "  Retrieved $($repoCommits.develop.Count) commits from $($repo.name)" -ForegroundColor Green
        
    } catch {
        Write-Host "Error accessing GitHub repository $($repo.githubOrg)/$($repo.githubRepo): $_" -ForegroundColor Red
        # Initialize empty arrays to prevent errors
        $repoCommits.current = @()
        $repoCommits.previous = @()
        $repoCommits.develop = @()
        $repoCommits.extraInDevelop = @()
    }
    
    $allRepositoryCommits[$repo.name] = $repoCommits
}

Write-Host "Extracting PAP IDs from commits..."
function Get-PAPIDsFromCommits($commitObjects) {
    if (-not $commitObjects) { return @() }
    $commitObjects | Where-Object { $_.Message -match "PAP-\d+" } | ForEach-Object {
        $_.Message | Select-String -Pattern "PAP-\d+" -AllMatches | ForEach-Object {
            $_.Matches | ForEach-Object { $_.Value }
        }
    } | Select-Object -Unique
}

# Extract PAP IDs for PVE Web repository (maintaining backward compatibility)
$pveWebRepo = $repositories | Where-Object { $_.name -eq "PVE Web" }
if ($pveWebRepo -and $allRepositoryCommits.ContainsKey("PVE Web")) {
    $papCurrent = Get-PAPIDsFromCommits $allRepositoryCommits["PVE Web"].current
    $papPrevious = Get-PAPIDsFromCommits $allRepositoryCommits["PVE Web"].previous
    $papDevelop = Get-PAPIDsFromCommits $allRepositoryCommits["PVE Web"].develop
} else {
    $papCurrent = @()
    $papPrevious = @()
    $papDevelop = @()
}

Write-Host "Building lookup for Excel PAP IDs..."
$excelPapIDs = @{}
foreach ($row in $rows) {
    $papId = "PAP-$($row.'Card ID')"
    $excelPapIDs[$papId] = $true
}

Write-Host "Preparing output rows..."
$totalRows = $rows.Count
Write-Host -NoNewline "Progress(%): "
$outputRows = @()
$counter = 0
$nextPercent = 5
foreach ($row in $rows) {
    $counter++
    $percentComplete = [math]::Floor(($counter / $totalRows) * 100)
    if ($percentComplete -ge $nextPercent) {
        Write-Host -NoNewline "..$nextPercent"
        $nextPercent += 5
    }
    $papId = "PAP-$($row.'Card ID')"
    $cardUrl = "$BaseURL/$($row.'Card ID')"
    $assignees = @()
    if ($row.'Assignee(s)') {
        foreach ($a in $row.'Assignee(s)' -split ';') {
            $parts = $a -split ','
            if ($parts.Count -ge 2) {
                $assignees += $parts[1].Trim()
            }
        }
        $assigneeNames = $assignees -join ', '
    } else {
        $assigneeNames = ""
    }
    $cardAssignees = "$cardUrl - $assigneeNames"

    # Validation Comment - Check across all repositories
    $validation = ""
    $releaseVersions = @()
    if ($row.'Planning Increment Label') {
        $releaseVersions = $row.'Planning Increment Label' -split ',' | ForEach-Object { $_.Trim() }
    }
    
    # Find ALL repositories that match any release version
    $matchingRepos = @()
    foreach ($repo in $repositories) {
        if ($releaseVersions -contains $repo.releaseVersion) {
            $matchingRepos += $repo
        }
    }
    
    if ($matchingRepos.Count -gt 0) {
        $allValidationResults = @()
        
        # Group repositories by release version
        $reposByReleaseVersion = @{}
        foreach ($matchingRepo in $matchingRepos) {
            if (-not $reposByReleaseVersion.ContainsKey($matchingRepo.releaseVersion)) {
                $reposByReleaseVersion[$matchingRepo.releaseVersion] = @()
            }
            $reposByReleaseVersion[$matchingRepo.releaseVersion] += $matchingRepo
        }
        
        # Process each release version group separately
        foreach ($releaseVersion in $reposByReleaseVersion.Keys) {
            $reposInGroup = $reposByReleaseVersion[$releaseVersion]
            $successfulRepos = @()
            $warningRepos = @()
            $notFoundRepos = @()
            
            foreach ($matchingRepo in $reposInGroup) {
                $repoCommits = $allRepositoryCommits[$matchingRepo.name]
                $repoValidation = ""
                $status = ""
                
                if ($matchingRepo.validationType -eq "branch-based") {
                    # Enhanced branch-based repository validation using GitHub API
                    $papCurrentRepo = Get-PAPIDsFromCommits $repoCommits.current
                    $papPreviousRepo = Get-PAPIDsFromCommits $repoCommits.previous
                    $papDevelopRepo = Get-PAPIDsFromCommits $repoCommits.develop
                    
                    $inCurrent = $papCurrentRepo -contains $papId
                    $inPrevious = $papPreviousRepo -contains $papId
                    $inDevelop = $papDevelopRepo -contains $papId

                    # Check for extra commits using GitHub branch comparison
                    $extraCommitsInDevelop = @()
                    if ($repoCommits.extraInDevelop) {
                        $extraCommitsInDevelop = $repoCommits.extraInDevelop | Where-Object { $_.Message -match $papId }
                    }

                    if ($inPrevious) {
                        $repoValidation = "⚠️ Found in previous branch: $($matchingRepo.previousBranch)"
                        $status = "warning"
                    } elseif ($inCurrent -and $inDevelop) {
                        # Check for extra commits in develop using API comparison
                        if ($extraCommitsInDevelop.Count -gt 0) {
                            $repoValidation = "⚠️ Extra commit(s) in $($matchingRepo.developBranch) branch"
                            $status = "warning"
                        } else {
                            $repoValidation = "✅ OK"
                            $status = "success"
                        }
                    } elseif ($inCurrent) {
                        $repoValidation = "✅ OK"
                        $status = "success"
                    } elseif ($inDevelop -and -not $inCurrent) {
                        $repoValidation = "⚠️ Found in $($matchingRepo.developBranch) branch but not in $($matchingRepo.currentBranch) branch"
                        $status = "warning"
                    } else {
                        $repoValidation = "⚠️ Not found"
                        $status = "notfound"
                    }
                } elseif ($matchingRepo.validationType -eq "develop-based") {
                    # Dovetail and ActionBoard repositories validation (develop branch only)
                    $papDevelopRepo = Get-PAPIDsFromCommits $repoCommits.develop
                    $inDevelop = $papDevelopRepo -contains $papId
                    
                    if ($inDevelop) {
                        $repoValidation = "✅ OK"
                        $status = "success"
                    } else {
                        $repoValidation = "⚠️ Not found"
                        $status = "notfound"
                    }
                }
                
                # Categorize results by status within this release version group
                if ($repoValidation) {
                    $repoResult = "$($matchingRepo.name) - $repoValidation"
                    if ($status -eq "success") {
                        $successfulRepos += $repoResult
                    } elseif ($status -eq "warning") {
                        $warningRepos += $repoResult
                    } elseif ($status -eq "notfound") {
                        $notFoundRepos += $repoResult
                    }
                }
            }
            
            # Smart filtering logic per release version group:
            # 1. If any repos have success or warnings, show only those (hide "not found")
            # 2. If no repos have success/warnings, show all "not found" repos
            $groupValidationResults = @()
            if ($successfulRepos.Count -gt 0 -or $warningRepos.Count -gt 0) {
                # Show successful and warning repos, hide "not found" ones
                $groupValidationResults = $successfulRepos + $warningRepos
            } else {
                # No successful repos, show all "not found" repos
                $groupValidationResults = $notFoundRepos
            }
            
            # Add this group's results to the overall results
            $allValidationResults += $groupValidationResults
        }
        
        # Combine all validation results from all release version groups
        if ($allValidationResults.Count -gt 0) {
            $validation = $allValidationResults -join " | "
        }
    }

    $outputRows += [PSCustomObject]@{
        Team = $row.Team
        'Card Title' = $row.'Card Title'
        'Assignee(s)' = $row.'Assignee(s)'
        'Card ID' = $row.'Card ID'
        'Card Type' = $row.'Card Type'
        'Current Lane Title' = $row.'Current Lane Title'
        'Release Version' = $row.'Planning Increment Label'
        Tags = $row.Tags
        'PAP ID' = $papId
        'Card URL' = $cardUrl
        'Validation Comment' = $validation
        'Card Assignee(s)' = $cardAssignees
    }
}
# Ensure 100 is printed at the end
if ($nextPercent -le 100) {
    Write-Host -NoNewline "..100"
}
Write-Host "" # Move to next line after loop

Write-Host "Detecting orphan PAP IDs..."
# Only check for orphans in PVE Web repository (branch-based validation)
$orphanRows = @()
if ($pveWebRepo -and $allRepositoryCommits.ContainsKey("PVE Web")) {
    $orphanPapIDs = $papCurrent | Where-Object { 
        (-not $papPrevious.Contains($_)) -and (-not $excelPapIDs.ContainsKey($_))
    }
    $counter = 0
    foreach ($orphan in $orphanPapIDs) {
        $counter++
        $cardId = $orphan -replace "PAP-", ""
        $cardUrl = "$BaseURL/$cardId"
        
        # Find the commit URL from already fetched commit data
        $commitUrl = ""
        $currentCommits = $allRepositoryCommits["PVE Web"].current
        if ($currentCommits) {
            $matchingCommit = $currentCommits | Where-Object { $_.Message -match $orphan } | Select-Object -First 1
            if ($matchingCommit) {
                $commitUrl = $matchingCommit.HtmlUrl
            }
        }
        
        $orphanRows += [PSCustomObject]@{
            'Card ID' = $cardId
            'PAP ID' = $orphan
            'Card URL' = $cardUrl
            'Commit URL' = $commitUrl
            'Validation Comment' = "⚠️ Exists in $($pveWebRepo.currentBranch) branch but missing in report"
        }
    }
} else {
    Write-Host "No PVE Web repository configured for orphan detection" -ForegroundColor Yellow
}

Write-Host "Exporting to Excel..."
# Create main report sheet
$outputRows | Export-Excel -Path $OutputExcelPath -WorksheetName 'Report' -TableName 'ReleaseTasks' -TableStyle 'None'

# Create GitHub info sheet
Write-Host "Adding GitHub repository information to separate sheet"
$gitHubRepoInfo | Export-Excel -Path $OutputExcelPath -WorksheetName 'GitHub' -TableName 'GitHubInfo' -TableStyle 'None'

# Create orphan commits sheet if there are orphan rows
if ($orphanRows.Count -gt 0) {
    Write-Host "Found $($orphanRows.Count) orphan PAP IDs - adding to separate sheet"
    $orphanRows | Export-Excel -Path $OutputExcelPath -WorksheetName 'Orphan Commits' -TableName 'OrphanCommits' -TableStyle 'None'
} else {
    Write-Host "No orphan PAP IDs found"
}

Write-Host "Augmented Excel file created at $OutputExcelPath"

# Helper function to format Card Assignee(s) with release version if needed
function FormatAssignee($row, $allRepositories) {
    $assignee = $row.'Card Assignee(s)'
    $releaseVersions = @()
    if ($row.'Release Version') {
        $releaseVersions = $row.'Release Version' -split ',' | ForEach-Object { $_.Trim() }
    }
    
    # Find the PVE Web repository's release version
    $pveWebReleaseVersion = ""
    $pveWebRepo = $allRepositories | Where-Object { $_.name -eq "PVE Web" }
    if ($pveWebRepo) {
        $pveWebReleaseVersion = $pveWebRepo.releaseVersion
    }
    
    # Check if release versions contain PVE Web's release version
    $containsPveWebReleaseVersion = $releaseVersions -contains $pveWebReleaseVersion
    
    # If it contains PVE Web release version, show assignee without release version
    # If it contains only non-PVE Web release versions, show assignee with release version
    if ($containsPveWebReleaseVersion) {
        return $assignee
    } elseif ($row.'Release Version') {
        return "$assignee ($($row.'Release Version'))"
    } else {
        return $assignee
    }
}

Write-Host "Generating text email body (To Do, QE Tasks, Final PM / TR / UX Review)..."

# To Do: Items in development (Current Lane Title NOT in nonDevStatuses)
$todoRows = $outputRows | Where-Object {
    $_.'Current Lane Title' -and
    ($nonDevStatuses -notcontains $_.'Current Lane Title')
}
$todoGroups = $todoRows | Group-Object Team

# Build To Do section (plain text)
$emailBody = "TO DO :`r`n`r`n"
foreach ($group in $todoGroups) {
    if ($group.Name) {
        $emailBody += "$($group.Name):`r`n"
        foreach ($row in $group.Group) {
            $emailBody += "• $(FormatAssignee $row $repositories)`r`n"
        }
        $emailBody += "`r`n"
    }
}

# QE Tasks: Items in testing (Current Lane Title is in qeStatuses)
$qeRows = $outputRows | Where-Object {
    $_.'Current Lane Title' -and
    ($qeStatuses -contains $_.'Current Lane Title')
}
$qeGroups = $qeRows | Group-Object Team

# Build QE Tasks section (plain text)
$emailBody += "QE TASKS:`r`n`r`n"
foreach ($group in $qeGroups) {
    if ($group.Name) {
        $emailBody += "$($group.Name):`r`n"
        foreach ($row in $group.Group) {
            $emailBody += "• $(FormatAssignee $row $repositories)`r`n"
        }
        $emailBody += "`r`n"
    }
}

# Final PM / TR / UX Review: Items where Current Lane Title is FINAL PM / TR / UX REVIEW
$finalRows = $outputRows | Where-Object {
    $_.'Current Lane Title' -eq 'FINAL PM / TR / UX REVIEW'
}
$finalGroups = $finalRows | Group-Object Team

# Build Final PM / TR / UX Review section (plain text)
$emailBody += "FINAL PM / TR / UX REVIEW:`r`n`r`n"
foreach ($group in $finalGroups) {
    if ($group.Name) {
        $emailBody += "$($group.Name):`r`n"
        foreach ($row in $group.Group) {
            $emailBody += "• $(FormatAssignee $row $repositories)`r`n"
        }
        $emailBody += "`r`n"
    }
}

Write-Host "Saving text email body to $EmailBodyPath"
Set-Content -Path $EmailBodyPath -Value $emailBody

Write-Host "Report generation complete!" -ForegroundColor Green