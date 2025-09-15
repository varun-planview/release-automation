# Requires: ImportExcel module (Install-Module -Name ImportExcel)

# ------------------ LOAD CONFIGURATION ------------------
$configPath = Join-Path $PSScriptRoot "config.json"

if (-not (Test-Path $configPath)) {
    Write-Host "Configuration file not found at: $configPath" -ForegroundColor Red
    Write-Host "Please ensure config.json exists in the same directory as this script." -ForegroundColor Red
    exit 1
}

try {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    
    # Load configuration variables
    $repositories = $config.repositories
    $BaseURL = $config.BaseURL
    $nonDevStatuses = $config.nonDevStatuses
    $qeStatuses = $config.qeStatuses
    
    Write-Host "Configuration loaded successfully from: $configPath" -ForegroundColor Green
    Write-Host "Found $($repositories.Count) repositories configured" -ForegroundColor Green
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

Write-Host "Extracting git commit messages for all repositories..."

# Initialize hash tables for PAP IDs from all repositories
$allRepositoryCommits = @{}

foreach ($repo in $repositories) {
    Write-Host "Processing repository: $($repo.name)" -ForegroundColor Cyan
    
    if (-not (Test-Path $repo.path)) {
        Write-Host "Warning: Repository path not found: $($repo.path)" -ForegroundColor Yellow
        continue
    }
    
    $repoCommits = @{}
    
    if ($repo.validationType -eq "branch-based") {
        # For PRM repo - extract from current, previous, and develop branches
        $repoCommits.current = git -C $repo.path log $repo.currentBranch --pretty=format:"%s" 2>$null
        $repoCommits.previous = git -C $repo.path log $repo.previousBranch --pretty=format:"%s" 2>$null
        $repoCommits.develop = git -C $repo.path log $repo.developBranch --pretty=format:"%s" 2>$null
    } elseif ($repo.validationType -eq "develop-based") {
        # For Dovetail and ActionBoard repos - extract from develop branch only
        $repoCommits.develop = git -C $repo.path log $repo.developBranch --pretty=format:"%s" 2>$null
    }
    
    $allRepositoryCommits[$repo.name] = $repoCommits
}

Write-Host "Extracting PAP IDs from commits..."
function Get-PAPIDsFromCommits($commitMessages) {
    if (-not $commitMessages) { return @() }
    $commitMessages | Select-String -Pattern "PAP-\d+" -AllMatches | ForEach-Object {
        $_.Matches | ForEach-Object { $_.Value }
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
        $validationResults = @()
        
        foreach ($matchingRepo in $matchingRepos) {
            $repoCommits = $allRepositoryCommits[$matchingRepo.name]
            $repoValidation = ""
            
            if ($matchingRepo.validationType -eq "branch-based") {
                # PRM repository validation (existing logic)
                $papCurrentRepo = Get-PAPIDsFromCommits $repoCommits.current
                $papPreviousRepo = Get-PAPIDsFromCommits $repoCommits.previous
                $papDevelopRepo = Get-PAPIDsFromCommits $repoCommits.develop
                
                $inCurrent = $papCurrentRepo -contains $papId
                $inPrevious = $papPreviousRepo -contains $papId
                $inDevelop = $papDevelopRepo -contains $papId

                # Get all commit messages for this PAP ID in each branch
                $currentCommits = $repoCommits.current | Where-Object { $_ -match $papId }
                $developCommits = $repoCommits.develop | Where-Object { $_ -match $papId }

                if ($inPrevious) {
                    $repoValidation = "⚠️ Found in previous branch: $($matchingRepo.previousBranch)"
                } elseif ($inCurrent -and $inDevelop) {
                    # Check for extra commits in develop
                    $extraDevelopCommits = $developCommits | Where-Object { $currentCommits -notcontains $_ }
                    if ($extraDevelopCommits.Count -gt 0) {
                        $repoValidation = "⚠️ Extra commit(s) in $($matchingRepo.developBranch) branch"
                    } else {
                        $repoValidation = "✅ OK"
                    }
                } elseif ($inCurrent) {
                    $repoValidation = "✅ OK"
                } elseif ($inDevelop -and -not $inCurrent) {
                    $repoValidation = "⚠️ Found in $($matchingRepo.developBranch) branch but not in $($matchingRepo.currentBranch) branch"
                } else {
                    $repoValidation = "❌ Not found"
                }
            } elseif ($matchingRepo.validationType -eq "develop-based") {
                # Dovetail and ActionBoard repositories validation (develop branch only)
                $papDevelopRepo = Get-PAPIDsFromCommits $repoCommits.develop
                $inDevelop = $papDevelopRepo -contains $papId
                
                if ($inDevelop) {
                    $repoValidation = "✅ OK"
                } else {
                    $repoValidation = "❌ Not found"
                }
            }
            
            # Add repo-specific validation result
            if ($repoValidation) {
                $validationResults += "$($matchingRepo.name) - $repoValidation"
            }
        }
        
        # Combine all validation results
        if ($validationResults.Count -gt 0) {
            $validation = $validationResults -join " | "
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
if ($pveWebRepo -and $allRepositoryCommits.ContainsKey("PVE Web")) {
    $orphanPapIDs = $papCurrent | Where-Object { 
        (-not $papPrevious.Contains($_)) -and (-not $excelPapIDs.ContainsKey($_))
    }
    $counter = 0
    foreach ($orphan in $orphanPapIDs) {
        $counter++
        $cardId = $orphan -replace "PAP-", ""
        $cardUrl = "$BaseURL/$cardId"
        $outputRows += [PSCustomObject]@{
            Team = ""
            'Card Title' = ""
            'Assignee(s)' = ""
            'Card ID' = $cardId
            'Card Type' = ""
            'Current Lane Title' = ""
            'Release Version' = ""  # Leave blank for orphan rows
            Tags = ""
            'PAP ID' = $orphan
            'Card URL' = $cardUrl
            'Validation Comment' = "⚠️ Exists in $($pveWebRepo.currentBranch) branch but missing in report"
            'Card Assignee(s)' = ""
        }
    }
} else {
    Write-Host "No PVE Web repository configured for orphan detection" -ForegroundColor Yellow
}

Write-Host "Exporting to Excel..."
$outputRows | Export-Excel -Path $OutputExcelPath -TableName 'ReleaseTasks' -TableStyle 'None'
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

Write-Host "Report generation complete!"