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
    $GitRepoPath = $config.GitRepoPath
    $BaseURL = $config.BaseURL
    $CurrentPI = $config.CurrentPI
    $CurrentBranch = $config.CurrentBranch
    $PreviousBranch = $config.PreviousBranch
    $DevelopBranch = $config.DevelopBranch
    $nonDevStatuses = $config.nonDevStatuses
    $qeStatuses = $config.qeStatuses
    
    Write-Host "Configuration loaded successfully from: $configPath" -ForegroundColor Green
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

Write-Host "Extracting git commit messages for branches..."
$commitsCurrent = git -C $GitRepoPath log $CurrentBranch --pretty=format:"%s"
$commitsPrevious = git -C $GitRepoPath log $PreviousBranch --pretty=format:"%s"
$commitsDevelop = git -C $GitRepoPath log $DevelopBranch --pretty=format:"%s"

Write-Host "Extracting PAP IDs from commits..."
function Get-PAPIDsFromCommits($commitMessages) {
    $commitMessages | Select-String -Pattern "PAP-\d+" -AllMatches | ForEach-Object {
        $_.Matches | ForEach-Object { $_.Value }
    } | Select-Object -Unique
}
$papCurrent = Get-PAPIDsFromCommits $commitsCurrent
$papPrevious = Get-PAPIDsFromCommits $commitsPrevious
$papDevelop = Get-PAPIDsFromCommits $commitsDevelop

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

    # Validation Comment
    $validation = ""
    $piLabels = @()
    if ($row.'Planning Increment Label') {
        $piLabels = $row.'Planning Increment Label' -split ',' | ForEach-Object { $_.Trim() }
    }
    if ($piLabels -contains $CurrentPI) {
        $inCurrent = $papCurrent -contains $papId
        $inPrevious = $papPrevious -contains $papId
        $inDevelop = $papDevelop -contains $papId

        # Get all commit messages for this PAP ID in each branch
        $currentCommits = $commitsCurrent | Where-Object { $_ -match $papId }
        $developCommits = $commitsDevelop | Where-Object { $_ -match $papId }

        if ($inPrevious) {
            $validation = "⚠️ Found in previous branch: $PreviousBranch"
        } elseif ($inCurrent -and $inDevelop) {
            # Check for extra commits in develop
            $extraDevelopCommits = $developCommits | Where-Object { $currentCommits -notcontains $_ }
            if ($extraDevelopCommits.Count -gt 0) {
                $validation = "⚠️ Extra commit(s) in $DevelopBranch branch"
            } else {
                $validation = "✅ OK"
            }
        } elseif ($inCurrent) {
            $validation = "✅ OK"
        } elseif ($inDevelop -and -not $inCurrent) {
            $validation = "⚠️ Found in $DevelopBranch branch but not in $CurrentBranch branch"
        } else {
            $validation = ""
        }
    }

    $outputRows += [PSCustomObject]@{
        Team = $row.Team
        'Card Title' = $row.'Card Title'
        'Assignee(s)' = $row.'Assignee(s)'
        'Card ID' = $row.'Card ID'
        'Card Type' = $row.'Card Type'
        'Current Lane Title' = $row.'Current Lane Title'
        'Planning Increment Label' = $row.'Planning Increment Label'
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
        'Planning Increment Label' = ""  # Leave blank for orphan rows
        Tags = ""
        'PAP ID' = $orphan
        'Card URL' = $cardUrl
        'Validation Comment' = "⚠️ Exists in $CurrentBranch branch but missing in report"
        'Card Assignee(s)' = ""
    }
}

Write-Host "Exporting to Excel..."
$outputRows | Export-Excel -Path $OutputExcelPath -TableName 'ReleaseTasks' -TableStyle 'None'
Write-Host "Augmented Excel file created at $OutputExcelPath"

# Helper function to format Card Assignee(s) with PI label if needed
function FormatAssignee($row, $CurrentPI) {
    $assignee = $row.'Card Assignee(s)'
    $piLabels = @()
    if ($row.'Planning Increment Label') {
        $piLabels = $row.'Planning Increment Label' -split ',' | ForEach-Object { $_.Trim() }
    }
    if ($piLabels -contains $CurrentPI) {
        return $assignee
    } elseif ($row.'Planning Increment Label') {
        return "$assignee ($($row.'Planning Increment Label'))"
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
            $emailBody += "• $(FormatAssignee $row $CurrentPI)`r`n"
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
            $emailBody += "• $(FormatAssignee $row $CurrentPI)`r`n"
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
            $emailBody += "• $(FormatAssignee $row $CurrentPI)`r`n"
        }
        $emailBody += "`r`n"
    }
}

Write-Host "Saving text email body to $EmailBodyPath"
Set-Content -Path $EmailBodyPath -Value $emailBody

Write-Host "Report generation complete!"