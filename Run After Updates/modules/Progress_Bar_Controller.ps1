# =========================
# STEP-BASED PROGRESS ENGINE
# =========================

$script:ProgressSteps = @()
$script:CurrentStepIndex = 0
$script:Activity = "Processing"

# -------------------------------------
# Adds a Step to the progress bar
# -------------------------------------
function Add-Process {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $script:ProgressSteps += $Name
}

# -------------------------------------
# Initialize Progress bar
# -------------------------------------
function Initialize-ProgressBar {
    param(
        [string]$Activity = "Processing"
    )

    $script:Activity = $Activity
    $script:CurrentStepIndex = 0

    Write-Progress -Id 1 `
        -Activity $Activity `
        -Status "Starting..." `
        -PercentComplete 0
}

# -------------------------------------
# Advance to Next process
# -------------------------------------
function Move-Process {
    param(
        [string]$StatusOverride
    )

    if ($script:ProgressSteps.Count -eq 0) {
        throw "No steps defined. Use Add-Step first."
    }

    $script:CurrentStepIndex++

    $stepName = $script:ProgressSteps[$script:CurrentStepIndex - 1]

    $status = if ($StatusOverride) { $StatusOverride } else { $stepName }

    $percent = [math]::Round(
        ($script:CurrentStepIndex / $script:ProgressSteps.Count) * 100
    )

    Write-Progress -Id 1 `
        -Activity $script:Activity `
        -Status $status `
        -PercentComplete $percent
}

# -------------------------------------
# Run All Steps (optional helper)
# -------------------------------------
function Invoke-Steps {
    param(
        [scriptblock[]]$Steps
    )

    for ($i = 0; $i -lt $Steps.Count; $i++) {
        & $Steps[$i]
        Next-Step
    }
}

# -------------------------------------
# Complete
# -------------------------------------
function Complete-Progress {
    Write-Progress -Id 1 `
        -Activity $script:Activity `
        -Completed
}

# -------------------------------------
# Sub-progress context 
# -------------------------------------
function New-ProgressContext {
    param(
        [int]$ParentId = 1,
        [int]$Id = 2
    )

    return [PSCustomObject]@{
        Update = {
            param($activity, $status, $percent)

            Write-Progress -Id $Id `
                -ParentId $ParentId `
                -Activity $activity `
                -Status $status `
                -PercentComplete $percent
        }.GetNewClosure()

        Complete = {
            Write-Progress -Id $Id -Completed
        }.GetNewClosure()
    }
}