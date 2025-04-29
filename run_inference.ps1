###############################################################################
# Configuration
###############################################################################
# Parameter block for slide path
param (
    [Parameter(Mandatory = $true)]
    [string]$SlidePath
)

# Enable or disable AI inference (set to $false for test mode)
$ENABLE_AI_INFERENCE = $true  # Set to $true to run inference; $false to skip for testing

# Set to $true to copy SVS, logs, CSV, etc. into OneDrive
$ENABLE_ONEDRIVE_SYNC = $true

# Set to $true to copy SVS, logs, CSV, etc. into Google Drive
$ENABLE_GDRIVE_SYNC = $true

###############################################################################
# STEP 0: Enable Windows Forms and Visual Styles
###############################################################################
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

###############################################################################
# STEP 0.2: Validate Slide Path
###############################################################################
if (-not $SlidePath -or -not (Test-Path $SlidePath)) {
    [System.Windows.Forms.MessageBox]::Show("No or invalid slide path provided. Please select the slide file.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Slide File"
    $openFileDialog.Filter = "Image Files (*.png;*.jpg;*.tif;*.svs)|*.png;*.jpg;*.tif;*.svs|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $SlidePath = $openFileDialog.FileName
    } else {
        Write-Host "Slide path not provided. Exiting."
        return
    }
}
Add-Content $LogFile "Slide path: $SlidePath"

###############################################################################
# STEP 1: Prompt User for Confirmation
###############################################################################
$userChoice = [System.Windows.Forms.MessageBox]::Show(
    "Run AI-based filamentous fungus detection? This tool works best on KOH smear slides.",
    "Confirm Analysis",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)
if ($userChoice -ne [System.Windows.Forms.DialogResult]::Yes) {
    Write-Host "User declined to run the analysis."
    return
}

###############################################################################
# STEP 1.5: Get User Name
###############################################################################
$userNameForm = New-Object System.Windows.Forms.Form
$userNameForm.Text = "User Information"
$userNameForm.Size = New-Object System.Drawing.Size(350, 150)
$userNameForm.StartPosition = 'CenterScreen'
$userNameForm.TopMost = $true
$userNameForm.FormBorderStyle = 'FixedDialog'
$userNameForm.MaximizeBox = $false
$userNameForm.MinimizeBox = $false

$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Please enter your name:"
$nameLabel.AutoSize = $true
$nameLabel.Location = New-Object System.Drawing.Point(20,20)
$userNameForm.Controls.Add($nameLabel)

$nameTextBox = New-Object System.Windows.Forms.TextBox
$nameTextBox.Location = New-Object System.Drawing.Point(20, 50)
$nameTextBox.Size = New-Object System.Drawing.Size(300, 20)
$userNameForm.Controls.Add($nameTextBox)

$submitNameButton = New-Object System.Windows.Forms.Button
$submitNameButton.Text = "Submit"
$submitNameButton.Location = New-Object System.Drawing.Point(120, 80)
$submitNameButton.Size = New-Object System.Drawing.Size(100, 30)
$userNameForm.Controls.Add($submitNameButton)

$global:UserName = $null
$submitNameButton.Add_Click({
    if ($nameTextBox.Text.Trim()) {
        $global:UserName = $nameTextBox.Text.Trim()
        $userNameForm.Close()
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter your name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
    }
})

$userNameForm.ShowDialog() | Out-Null
if (-not $global:UserName) {
    Write-Host "User name not provided. Exiting."
    return
}
Add-Content $LogFile "User Name: $global:UserName"

###############################################################################
# STEP 2: Setup Paths, Logging, and Environment
###############################################################################
$TriggerTime = Get-Date  # Record trigger time for filtering overlay images

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ModelsDir  = Join-Path $ScriptDir "models"
$Executable = Join-Path $ScriptDir "inference.exe"
$OutputDir  = Join-Path $ScriptDir "heatmaps"
$LogDir     = Join-Path $ScriptDir "log"

# Setup OneDrive backup
if ($ENABLE_ONEDRIVE_SYNC) {
    $OneDriveFolder = $env:OneDrive
    if (-not $OneDriveFolder) {
        # Fallback: Query registry
        $OneDriveFolder = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive" -Name "UserFolder" -ErrorAction SilentlyContinue).UserFolder
    }

    if ($OneDriveFolder) {
        # Create KOH Analysis folder structure in OneDrive
        $OneDriveBackupRoot = Join-Path $OneDriveFolder "KOH_Analysis"
        $OneDriveHeatmaps = Join-Path $OneDriveBackupRoot "heatmaps"
        $OneDriveLogs = Join-Path $OneDriveBackupRoot "logs"
        $OneDriveScans = Join-Path $OneDriveBackupRoot "scans"
        
        # Create directories if they don't exist
        @($OneDriveBackupRoot, $OneDriveHeatmaps, $OneDriveLogs, $OneDriveScans) | ForEach-Object {
            if (-not (Test-Path $_)) {
                New-Item -ItemType Directory -Path $_ -Force | Out-Null
            }
        }
        Add-Content $LogFile "OneDrive backup enabled: $OneDriveBackupRoot"
    } else {
        Add-Content $LogFile "OneDrive backup disabled: OneDrive not found"
        $ENABLE_ONEDRIVE_SYNC = $false
    }
} else {
    Add-Content $LogFile "OneDrive backup disabled by configuration"
}

###############################################################################
# STEP 2.5: Detect & prepare Google Drive sync folder
###############################################################################
$GDriveRoot = "G:\My Drive"
if ($ENABLE_GDRIVE_SYNC) {
    if (Test-Path $GDriveRoot) {
        $GDriveBackupRoot = Join-Path $GDriveRoot 'KOH_Analysis'
        $GDriveLogs = Join-Path $GDriveBackupRoot 'logs'
        
        # Create directories if they don't exist
        try {
            if (-not (Test-Path $GDriveBackupRoot)) {
                New-Item -ItemType Directory -Path $GDriveBackupRoot -Force | Out-Null
                Add-Content $LogFile "Created Google Drive backup root: $GDriveBackupRoot"
            }
            if (-not (Test-Path $GDriveLogs)) {
                New-Item -ItemType Directory -Path $GDriveLogs -Force | Out-Null
                Add-Content $LogFile "Created Google Drive logs directory: $GDriveLogs"
            }
            Add-Content $LogFile "Google Drive backup enabled: $GDriveLogs"
        } catch {
            Add-Content $LogFile "Error creating Google Drive directories: $_"
            $ENABLE_GDRIVE_SYNC = $false
        }
    } else {
        Add-Content $LogFile "Google Drive not found at: $GDriveRoot"
        $ENABLE_GDRIVE_SYNC = $false
    }
}

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$LogFile    = Join-Path $LogDir "run_inference_log.txt"
$StdOutFile = Join-Path $LogDir "stdout.txt"
$StdErrFile = Join-Path $LogDir "stderr.txt"

Add-Content $LogFile "`n===== $(Get-Date) ====="
Add-Content $LogFile "Script directory: $ScriptDir"
Add-Content $LogFile "Slide path: $SlidePath"
Add-Content $LogFile "Executable path: $Executable"
if ($ENABLE_ONEDRIVE_SYNC -and $OneDriveFolder) {
    Add-Content $LogFile "OneDrive backup enabled: $OneDriveBackupRoot"
} else {
    Add-Content $LogFile "OneDrive backup disabled: OneDrive not found"
}

# Set MKL environment variable to avoid library collisions
$env:KMP_DUPLICATE_LIB_OK = "TRUE"
Add-Content $LogFile "Set KMP_DUPLICATE_LIB_OK=TRUE"

# Setup CSV logging
$CsvLogFile = Join-Path $LogDir "interpretation_history.csv"
# Prepare CSV file and header migration if needed
$expectedHeader = "Timestamp,Operator Name,Slide Path,Initial Human Interpretation,Other Interpretation Details,AI Prediction,Final Human Interpretation,Final Other Details"
if (-not (Test-Path $CsvLogFile)) {
    # No CSV exists: create fresh with header
    $expectedHeader | Out-File $CsvLogFile -Encoding utf8
} else {
    # Existing CSV: check if header matches expected
    $firstLine = Get-Content $CsvLogFile -TotalCount 1
    if ($firstLine -ne $expectedHeader) {
        # Backup old CSV
        Copy-Item -Path $CsvLogFile -Destination "${CsvLogFile}.bak" -Force
        # Read all old records (skip header)
        $oldRecords = Get-Content $CsvLogFile | Select-Object -Skip 1
        # Rewrite CSV with new header
        $expectedHeader | Out-File $CsvLogFile -Encoding utf8
        # Migrate old records: insert empty placeholders for Slide Path and Final Other Details
        foreach ($line in $oldRecords) {
            $cols = $line -split ','
            # Ensure we have at least 6 columns: Timestamp,Operator Name,Initial Human Interpretation,Other Interpretation Details,AI Prediction,Final Human Interpretation
            # Then build new row with empty Slide Path at pos 2 and empty Final Other Details at end
            $newCols = @(
                $cols[0],
                $cols[1],
                "",            # Slide Path placeholder
                $cols[2],
                $cols[3],
                $cols[4],
                $cols[5],
                ""             # Final Other Details placeholder
            )
            ($newCols -join ',') | Out-File $CsvLogFile -Append -Encoding utf8
        }
    }
}

$ArgsString = @(
    "--slide_path `"$SlidePath`"",
    "--embedder_low `"$ModelsDir\low_mag_embedder.pth`"",
    "--embedder_high `"$ModelsDir\high_mag_embedder.pth`"",
    "--aggregator `"$ModelsDir\aggregator.pth`"",
    "--device cpu",
    "--detection_threshold 0.4597581923007965",
    "--output `"$OutputDir`""
) -join " "

Add-Content $LogFile "Running command:"
Add-Content $LogFile "`"$Executable`" $ArgsString"

###############################################################################
# STEP 3: Get Preliminary Human Interpretation First
###############################################################################
$global:PrelimInterpretation = $null
$global:OtherInterpretation = $null
$prelimForm = New-Object System.Windows.Forms.Form
$prelimForm.Text = "Preliminary Interpretation"
$prelimForm.Size = New-Object System.Drawing.Size(400, 250)
$prelimForm.StartPosition = 'CenterScreen'
$prelimForm.TopMost = $true
$prelimForm.FormBorderStyle = 'FixedDialog'
$prelimForm.MaximizeBox = $false
$prelimForm.MinimizeBox = $false

$interpretLabel = New-Object System.Windows.Forms.Label
$interpretLabel.Text = "Select your initial interpretation:"
$interpretLabel.AutoSize = $true
$interpretLabel.Location = New-Object System.Drawing.Point(20,20)
$prelimForm.Controls.Add($interpretLabel)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(20, 50)
$comboBox.Size = New-Object System.Drawing.Size(350, 30)
$comboBox.Items.Add("Fungus Positive")
$comboBox.Items.Add("Fungus Negative")
$comboBox.Items.Add("Indeterminate")
$comboBox.Items.Add("Other (e.g., yeast, Pythium, Acanthamoeba)")
$comboBox.DropDownStyle = "DropDownList"
$prelimForm.Controls.Add($comboBox)

$otherLabel = New-Object System.Windows.Forms.Label
$otherLabel.Text = "Please specify:"
$otherLabel.AutoSize = $true
$otherLabel.Location = New-Object System.Drawing.Point(20, 90)
$otherLabel.Visible = $false
$prelimForm.Controls.Add($otherLabel)

$otherTextBox = New-Object System.Windows.Forms.TextBox
$otherTextBox.Location = New-Object System.Drawing.Point(20, 110)
$otherTextBox.Size = New-Object System.Drawing.Size(350, 20)
$otherTextBox.Visible = $false
$prelimForm.Controls.Add($otherTextBox)

$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(50, 150)
$submitButton.Size = New-Object System.Drawing.Size(100, 30)
$prelimForm.Controls.Add($submitButton)

$skipButton = New-Object System.Windows.Forms.Button
$skipButton.Text = "Skip"
$skipButton.Location = New-Object System.Drawing.Point(180, 150)
$skipButton.Size = New-Object System.Drawing.Size(100, 30)
$prelimForm.Controls.Add($skipButton)

$comboBox.Add_SelectedIndexChanged({
    if ($comboBox.SelectedItem -eq "Other (e.g., yeast, Pythium, Acanthamoeba)") {
        $otherLabel.Visible = $true
        $otherTextBox.Visible = $true
    } else {
        $otherLabel.Visible = $false
        $otherTextBox.Visible = $false
    }
})

$submitButton.Add_Click({
    if ($comboBox.SelectedItem) {
        if ($comboBox.SelectedItem -eq "Other (e.g., yeast, Pythium, Acanthamoeba)") {
            if ($otherTextBox.Text.Trim()) {
                $global:PrelimInterpretation = "Other"
                $global:OtherInterpretation = $otherTextBox.Text.Trim()
                $prelimForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please specify the other interpretation.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
            }
        } else {
            $global:PrelimInterpretation = $comboBox.SelectedItem
            $prelimForm.Close()
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select an interpretation or click Skip.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
    }
})
$skipButton.Add_Click({
    $global:PrelimInterpretation = "Skipped"
    $prelimForm.Close()
})

$prelimForm.ShowDialog() | Out-Null
Add-Content $LogFile "Preliminary Human Interpretation: $global:PrelimInterpretation"
if ($global:OtherInterpretation) {
    Add-Content $LogFile "Other Interpretation Details: $global:OtherInterpretation"
}

###############################################################################
# STEP 4: Start the Inference Process (Non-Blocking)
###############################################################################
if ($ENABLE_AI_INFERENCE) {
    try {
        $process = Start-Process -FilePath $Executable -ArgumentList $ArgsString -NoNewWindow -PassThru `
            -RedirectStandardOutput $StdOutFile -RedirectStandardError $StdErrFile
    } catch {
        Add-Content $LogFile "Error launching inference.exe: $_"
        return
    }
} else {
    Add-Content $LogFile "AI Inference disabled - running in test mode"
    # Create a dummy process object for the loading screen
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $process.StartInfo.FileName = "cmd.exe"
    $process.StartInfo.Arguments = "/c echo Skipping AI inference"
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.CreateNoWindow = $true
    $process.Start()
    Start-Sleep -Seconds 2  # Simulate some processing time
    $process.WaitForExit()
}

###############################################################################
# STEP 5: Show the Loading Screen
###############################################################################
$loadingForm = New-Object System.Windows.Forms.Form
$loadingForm.Text = "KOH Smear Analysis"
$loadingForm.Size = New-Object System.Drawing.Size(350, 180)
$loadingForm.StartPosition = 'CenterScreen'
$loadingForm.TopMost = $true
$loadingForm.FormBorderStyle = 'FixedDialog'
$loadingForm.MaximizeBox = $false
$loadingForm.MinimizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Text = if ($ENABLE_AI_INFERENCE) {
    "Analyzing KOH smear with AI, please wait..."
} else {
    "Testing mode - AI analysis disabled..."
}
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(40, 20)
$loadingForm.Controls.Add($label)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = 'Marquee'
$progressBar.Location = New-Object System.Drawing.Point(40, 60)
$progressBar.Size = New-Object System.Drawing.Size(250, 25)
$loadingForm.Controls.Add($progressBar)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(120, 110)
$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
$cancelButton.Add_Click({
    if (!$process.HasExited) {
       $process | Stop-Process -Force
    }
    $loadingForm.Close()
    Write-Host "User canceled the analysis."
    exit
})
$loadingForm.Controls.Add($cancelButton)

$loadingForm.Show()

###############################################################################
# STEP 6: Poll Until the Inference Process Finishes
###############################################################################
while (-not $process.HasExited) {
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.Application]::DoEvents()
}
Add-Content $LogFile "Process exited with code: $($process.ExitCode)"
if ($loadingForm -and !$loadingForm.IsDisposed) {
    $loadingForm.Close()
}

###############################################################################
# STEP 7: Process and Display the AI Inference Output
###############################################################################
if (-not $ENABLE_AI_INFERENCE) {
    $Output = "TEST MODE - AI Inference Disabled`n`nPrediction: Test Mode`n`nDISCLAIMER: This is a test run with AI inference disabled."
    $AIPrediction = "Test Mode"
} else {
    if (Test-Path $StdOutFile) {
        $Output = Get-Content $StdOutFile -Raw
        Add-Content $LogFile "`n--- Inference Output ---`n$Output`n"

        # Remove unwanted lines
        $Output = $Output -replace "(?m)^Using device:.*\r?\n?", ""
        $Output = $Output -replace "(?m)^Low magnification patches processed:.*\r?\n?", ""
        $Output = $Output -replace "(?m)^High magnification patches processed:.*\r?\n?", ""
        
        # Extract AI prediction
        $AIPrediction = if ($Output -match "Prediction: (.+?)(\r?\n|$)") {
            $matches[1]
        } else {
            "Error: No prediction found"
        }
        
        # Clarify that "Positive" means filamentous fungus
        $Output = $Output -replace "Prediction: Positive", "Prediction: Positive (indicative of filamentous fungus)"
        
        # Add disclaimer about AI
        $Output += "`r`n`r`nDISCLAIMER: This analysis is performed by an AI model that can make mistakes. Please interpret results accordingly."

        # Immediately open the newest overlay image (if available)
        $overlayImages = Get-ChildItem -Path $OutputDir -Filter "*_overlay.png" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $TriggerTime } |
            Sort-Object LastWriteTime -Descending
        if ($overlayImages -and $overlayImages.Count -gt 0) {
            $latestOverlay = $overlayImages[0].FullName
            Add-Content $LogFile "Opening overlay heatmap: $latestOverlay"
            Start-Process $latestOverlay

            # Backup to OneDrive if available
            if ($ENABLE_ONEDRIVE_SYNC -and $OneDriveFolder) {
                $overlayImages | ForEach-Object {
                    $destPath = Join-Path $OneDriveHeatmaps $_.Name
                    Copy-Item -Path $_.FullName -Destination $destPath -Force
                    Add-Content $LogFile "Backed up heatmap to OneDrive: $destPath"
                }
            }
        } else {
            Write-Host "No new overlay heatmap found in $OutputDir."
        }

        # Create a custom result form (Fixed Size)
        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "AI Inference Result"
        $resultForm.Size = New-Object System.Drawing.Size(450, 250)
        $resultForm.StartPosition = 'CenterScreen'
        $resultForm.TopMost = $true
        $resultForm.FormBorderStyle = 'FixedDialog'
        $resultForm.MaximizeBox = $false
        $resultForm.MinimizeBox = $false

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ReadOnly = $true
        $textBox.Size = New-Object System.Drawing.Size(400, 120)
        $textBox.Location = New-Object System.Drawing.Point(20, 20)
        $textBox.Text = $Output
        $textBox.SelectionStart = 0
        $textBox.SelectionLength = 0
        $textBox.HideSelection = $true
        $textBox.TabStop = $false
        $resultForm.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(80, 160)
        $okButton.Add_Click({ $resultForm.Close() })
        $resultForm.Controls.Add($okButton)

        $readMoreButton = New-Object System.Windows.Forms.Button
        $readMoreButton.Text = "Read More"
        $readMoreButton.Location = New-Object System.Drawing.Point(200, 160)
        $readMoreButton.Add_Click({
            Start-Process "https://www.sciencedirect.com/science/article/pii/S2666914524001891"
        })
        $resultForm.Controls.Add($readMoreButton)

        $resultForm.ShowDialog() | Out-Null
    } else {
        Add-Content $LogFile "No StdOut file found; possibly an error. Check $StdErrFile."
        [System.Windows.Forms.MessageBox]::Show("No StdOut file found; possibly an error. Check logs.", "Inference Result") | Out-Null
        return
    }
}

###############################################################################
# STEP 8: Prompt for Final Human Interpretation (After Viewing AI Output)
###############################################################################
$global:FinalInterpretation = $null
$global:FinalOtherInterpretation = $null
$finalForm = New-Object System.Windows.Forms.Form
$finalForm.Text = "Final Interpretation"
$finalForm.Size = New-Object System.Drawing.Size(450, 250)
$finalForm.StartPosition = 'CenterScreen'
$finalForm.TopMost = $true
$finalForm.FormBorderStyle = 'FixedDialog'
$finalForm.MaximizeBox = $false
$finalForm.MinimizeBox = $false
$finalForm.BackColor = [System.Drawing.SystemColors]::Control

# Create a panel to hold the content with proper padding
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)
$finalForm.Controls.Add($panel)

$finalLabel = New-Object System.Windows.Forms.Label
$finalLabel.Text = "Look at the prediction heatmap and the slide again.`nWhat is your final interpretation?"
$finalLabel.AutoSize = $false
$finalLabel.Size = New-Object System.Drawing.Size(390, 45)
$finalLabel.Location = New-Object System.Drawing.Point(20, 20)
$finalLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$finalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($finalLabel)

$finalCombo = New-Object System.Windows.Forms.ComboBox
$finalCombo.Location = New-Object System.Drawing.Point(20, 70)
$finalCombo.Size = New-Object System.Drawing.Size(390, 30)
$finalCombo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$finalCombo.Items.Add("Fungus Positive")
$finalCombo.Items.Add("Fungus Negative")
$finalCombo.Items.Add("Indeterminate")
$finalCombo.Items.Add("Other (e.g., yeast, Pythium, Acanthamoeba)")
$finalCombo.DropDownStyle = "DropDownList"
$panel.Controls.Add($finalCombo)

$finalOtherLabel = New-Object System.Windows.Forms.Label
$finalOtherLabel.Text = "Please specify:"
$finalOtherLabel.AutoSize = $true
$finalOtherLabel.Location = New-Object System.Drawing.Point(20, 110)
$finalOtherLabel.Visible = $false
$panel.Controls.Add($finalOtherLabel)

$finalOtherTextBox = New-Object System.Windows.Forms.TextBox
$finalOtherTextBox.Location = New-Object System.Drawing.Point(20, 130)
$finalOtherTextBox.Size = New-Object System.Drawing.Size(390, 20)
$finalOtherTextBox.Visible = $false
$panel.Controls.Add($finalOtherTextBox)

$finalSubmit = New-Object System.Windows.Forms.Button
$finalSubmit.Text = "Submit"
$finalSubmit.Location = New-Object System.Drawing.Point(80, 170)
$finalSubmit.Size = New-Object System.Drawing.Size(100, 30)
$finalSubmit.BackColor = [System.Drawing.SystemColors]::Control
$finalSubmit.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($finalSubmit)

$finalSkip = New-Object System.Windows.Forms.Button
$finalSkip.Text = "Skip"
$finalSkip.Location = New-Object System.Drawing.Point(210, 170)
$finalSkip.Size = New-Object System.Drawing.Size(100, 30)
$finalSkip.BackColor = [System.Drawing.SystemColors]::Control
$finalSkip.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($finalSkip)

$finalCombo.Add_SelectedIndexChanged({
    if ($finalCombo.SelectedItem -eq "Other (e.g., yeast, Pythium, Acanthamoeba)") {
        $finalOtherLabel.Visible = $true
        $finalOtherTextBox.Visible = $true
    } else {
        $finalOtherLabel.Visible = $false
        $finalOtherTextBox.Visible = $false
    }
})

$finalSubmit.Add_Click({
    if ($finalCombo.SelectedItem) {
        if ($finalCombo.SelectedItem -eq "Other (e.g., yeast, Pythium, Acanthamoeba)") {
            if ($finalOtherTextBox.Text.Trim()) {
                $global:FinalInterpretation = "Other"
                $global:FinalOtherInterpretation = $finalOtherTextBox.Text.Trim()
                $finalForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please specify the other interpretation.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
            }
        } else {
            $global:FinalInterpretation = $finalCombo.SelectedItem
            $finalForm.Close()
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a final interpretation.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
    }
})
$finalSkip.Add_Click({
    $global:FinalInterpretation = "Skipped"
    $finalForm.Close()
})
$finalForm.ShowDialog() | Out-Null
Add-Content $LogFile "Final Human Interpretation: $global:FinalInterpretation"
if ($global:FinalOtherInterpretation) {
    Add-Content $LogFile "Final Other Interpretation Details: $global:FinalOtherInterpretation"
}

# Log to CSV and backup to OneDrive
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$csvLine = [PSCustomObject]@{
    'Timestamp' = $timestamp
    'Operator Name' = $global:UserName
    'Slide Path' = $SlidePath
    'Initial Human Interpretation' = $global:PrelimInterpretation
    'Other Interpretation Details' = if ($global:OtherInterpretation) { $global:OtherInterpretation } else { "NA" }
    'AI Prediction' = $AIPrediction
    'Final Human Interpretation' = $global:FinalInterpretation
    'Final Other Details' = if ($global:FinalOtherInterpretation) { $global:FinalOtherInterpretation } else { "NA" }
}
$csvLine | Export-Csv -Path $CsvLogFile -Append -NoTypeInformation -Force

# Backup logs to OneDrive if available
if ($ENABLE_ONEDRIVE_SYNC -and $OneDriveFolder) {
    # Backup CSV
    Copy-Item -Path $CsvLogFile -Destination (Join-Path $OneDriveLogs "interpretation_history.csv") -Force
    
    # Backup run log
    Copy-Item -Path $LogFile -Destination (Join-Path $OneDriveLogs "run_inference_log.txt") -Force
    
    # Backup stdout/stderr
    Copy-Item -Path $StdOutFile -Destination (Join-Path $OneDriveLogs "stdout.txt") -Force
    Copy-Item -Path $StdErrFile -Destination (Join-Path $OneDriveLogs "stderr.txt") -Force
    
    Add-Content $LogFile "Backed up log files to OneDrive: $OneDriveLogs"
}

###############################################################################
# STEP 9: Backup logs & CSV to Google Drive (if enabled)
###############################################################################
if ($ENABLE_GDRIVE_SYNC -and $GDriveRoot) {
    try {
        # Test if Google Drive is accessible
        if (Test-Path $GDriveRoot) {
            Copy-Item $CsvLogFile -Destination (Join-Path $GDriveLogs 'interpretation_history.csv') -Force -ErrorAction Stop
            Copy-Item $LogFile -Destination (Join-Path $GDriveLogs 'run_inference_log.txt') -Force -ErrorAction Stop
            Copy-Item $StdOutFile -Destination (Join-Path $GDriveLogs 'stdout.txt') -Force -ErrorAction Stop
            Copy-Item $StdErrFile -Destination (Join-Path $GDriveLogs 'stderr.txt') -Force -ErrorAction Stop
            Add-Content $LogFile "Successfully backed up logs & CSV to Google Drive: $GDriveLogs"
        } else {
            Add-Content $LogFile "Google Drive path not accessible: $GDriveRoot"
        }
    } catch {
        Add-Content $LogFile "Error backing up to Google Drive: $_"
    }
}