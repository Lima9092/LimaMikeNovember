<#
.SYNOPSIS
Batch processor for large book catalogs with rate limiting, resume capability, and continuous processing

.DESCRIPTION
Processes large catalogs in batches of 100 items every 5+ minutes to respect API limits.
Includes resume capability, progress tracking, ISBN-10 validation, and continuous processing until all ISBNs are enriched.
#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ISBN-10 validation functions
function Test-ISBN10 {
    param([string]$isbn)
    
    if (-not $isbn) { return $false }
    
    # Remove any non-digit/X characters and convert to uppercase
    $cleanISBN = $isbn -replace '[^0-9Xx]', '' -replace 'x', 'X'
    
    # ISBN-10 must be exactly 10 characters
    if ($cleanISBN.Length -ne 10) { return $false }
    
    # Check if first 9 characters are digits
    if ($cleanISBN.Substring(0, 9) -notmatch '^[0-9]{9}$') { return $false }
    
    # Last character must be digit or X
    if ($cleanISBN.Substring(9, 1) -notmatch '^[0-9X]$') { return $false }
    
    # Calculate check digit
    $sum = 0
    for ($i = 0; $i -lt 9; $i++) {
        $digit = [int]$cleanISBN.Substring($i, 1)
        $sum += $digit * (10 - $i)
    }
    
    $checkDigit = $cleanISBN.Substring(9, 1)
    $expectedCheck = (11 - ($sum % 11)) % 11
    
    # Convert expected check digit to string (10 becomes 'X')
    $expectedCheckStr = if ($expectedCheck -eq 10) { 'X' } else { $expectedCheck.ToString() }
    
    return $checkDigit -eq $expectedCheckStr
}

function Get-ValidISBN10 {
    param([string]$rawISBN)
    
    if (-not $rawISBN) { return $null }
    
    # Clean the ISBN
    $cleanISBN = $rawISBN -replace '[^0-9Xx]', '' -replace 'x', 'X'
    
    # Test if it's a valid ISBN-10
    if (Test-ISBN10 -isbn $cleanISBN) {
        return $cleanISBN
    }
    
    return $null
}

# LMS Category mapping function
function Get-LMSCategory {
    param([array]$subjects, [string]$format = 'Books')
    
    if (-not $subjects) { return 'AF' }
    
    $subjectNames = $subjects | ForEach-Object { $_.name.ToLower() }
    $allSubjects = $subjectNames -join ' '
    
    $ageGroup = 'Adult'
    
    $explicitChildKeywords = @(
        'picture book', 'board book', 'early reader', 'kindergarten', 'preschool', 'toddler',
        'ages 0', 'ages 1', 'ages 2', 'ages 3', 'ages 4', 'ages 5', 'ages 6', 'ages 7', 'ages 8',
        'grade k', 'grade 1', 'grade 2', 'grade 3', 'grade 4', 'grade 5'
    )
    
    $hasExplicitChild = $false
    foreach ($keyword in $explicitChildKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword)) {
            $hasExplicitChild = $true
            break
        }
    }
    
    $classicLiteratureKeywords = @(
        'fiction classics', 'penguin classics', 'oxford classics', 'bantam classics', 'classical literature',
        'english literature, history and criticism', 'russian literature', 'french literature', 
        'continental european fiction', 'ancient philosophy', 'philosophy, ancient'
    )
    
    $isClassicLiterature = $false
    foreach ($keyword in $classicLiteratureKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword)) {
            $isClassicLiterature = $true
            break
        }
    }
    
    if ($isClassicLiterature -and (
        $allSubjects -match '\b(young adult fiction|young adult works|teen fiction|ya fiction|children''?s books|juvenile fiction|children''?s fiction)\b' -or
        $allSubjects -match '\b(harry potter|twilight|hunger games|dr\. seuss|narnia)\b'
    )) {
        $isClassicLiterature = $false
    }
    
    $adultThemeKeywords = @(
        'marriage', 'courtship', 'man-woman relationships', 'romance fiction', 'love stories',
        'social classes', 'upper class', 'manners and customs', 'psychological fiction',
        'adult themes', 'mature themes'
    )
    
    $hasAdultThemes = $false
    foreach ($keyword in $adultThemeKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword)) {
            $hasAdultThemes = $true
            break
        }
    }
    
    if ($hasExplicitChild -and -not $isClassicLiterature -and -not $hasAdultThemes) {
        $ageGroup = 'Child'
    }
    elseif ($allSubjects -match '\b(young adult|ya fiction|teen fiction|teenage fiction|adolescent|ages 1[2-9]|grade [6-9]|high school|middle school)\b' -and -not $isClassicLiterature -and -not $hasAdultThemes) {
        $ageGroup = 'Teen'
    }
    elseif ($isClassicLiterature -or $hasAdultThemes) {
        $ageGroup = 'Adult'
    }
    
    $isReference = $false
    $isFiction = $false
    $isGraphicNovel = $false
    
    $explicitReferenceKeywords = @(
        'dictionaries', 'encyclopedia', 'handbook', 'manual', 'textbook', 
        'almanac', 'atlas', 'directory', 'reference work', 'reference book',
        'dictionary', 'the oxford english dictionary'
    )
    
    foreach ($keyword in $explicitReferenceKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword) -and $allSubjects -notmatch '\b(fiction|novel|literature|story)\b') {
            $isReference = $true
            break
        }
    }
    
    if (($allSubjects -match '\b(graphic novels?|comic books?|manga)\b' -or 
         $allSubjects -match '\b(comics & graphic novels)\b') -and 
        $allSubjects -notmatch '\b(adaptations|comic books, strips, etc\.|illustrations)\b') {
        $isGraphicNovel = $true
    }
    
    if (-not $isReference -and -not $isGraphicNovel) {
        if ($allSubjects -match '\b(fiction|novels?|romance|mystery|thriller|horror|fantasy|science fiction|historical fiction|adventure fiction|drama|short stories|literature|literary fiction)\b' -and
            $allSubjects -notmatch '\b(non-fiction|nonfiction)\b') {
            $isFiction = $true
        }
    }
    
    if ($format -eq 'Audiobook' -or $format -eq 'Spoken Word') {
        if ($isFiction) {
            switch ($ageGroup) {
                'Child' { return 'AUD-CHILD' }
                'Teen'  { return 'AUD-TEEN' }
                default { return 'AUD-ADULT' }
            }
        } else {
            switch ($ageGroup) {
                'Child' { return 'AUDNF-CH' }
                'Teen'  { return 'AUDNF-T' }
                default { return 'AUDNF' }
            }
        }
    }
    
    if ($isReference) {
        switch ($ageGroup) {
            'Child' { return 'CREF' }
            'Teen'  { return 'TREF' }
            default { return 'AREF' }
        }
    }
    elseif ($isGraphicNovel) {
        if ($ageGroup -eq 'Teen') {
            return 'GRAFICT'
        } else {
            return 'GRAFIC'
        }
    }
    elseif ($isFiction) {
        switch ($ageGroup) {
            'Child' { return 'CF' }
            'Teen'  { return 'TF' }
            default { return 'AF' }
        }
    }
    else {
        switch ($ageGroup) {
            'Child' { return 'CNF' }
            'Teen'  { return 'TNF' }
            default { return 'ANF' }
        }
    }
}

# Debug function
function Get-CategoryDebug {
    param([array]$subjects, [string]$format = 'Books')
    
    if (-not $subjects) { return 'No subjects | Default: AF' }
    
    $subjectNames = $subjects | ForEach-Object { $_.name.ToLower() }
    $allSubjects = $subjectNames -join ' '
    
    $debug = "Subjects: [$($subjectNames -join '; ')] | "
    $debug += "Format: $format | "
    $debug += "Final: $(Get-LMSCategory -subjects $subjects -format $format)"
    
    return $debug
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Continuous Book Enrichment - Large Catalogs"
$form.Size = New-Object System.Drawing.Size(700, 750)
$form.StartPosition = 'CenterScreen'

# Input file controls
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = 'Source CSV:'
$lblInput.AutoSize = $true
$lblInput.Location = '10,15'

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = '100,12'
$txtInput.Width = 400
$txtInput.ReadOnly = $true

$btnBrowseInput = New-Object System.Windows.Forms.Button
$btnBrowseInput.Text = 'Browse...'
$btnBrowseInput.Location = '510,10'
$btnBrowseInput.Width = 75

# Lookup field dropdown
$lblField = New-Object System.Windows.Forms.Label
$lblField.Text = 'Lookup Field:'
$lblField.AutoSize = $true
$lblField.Location = '10,50'

$comboField = New-Object System.Windows.Forms.ComboBox
$comboField.Location = '100,47'
$comboField.Width = 200
$comboField.DropDownStyle = 'DropDownList'
$comboField.Enabled = $false

# Batch size
$lblBatchSize = New-Object System.Windows.Forms.Label
$lblBatchSize.Text = 'Batch Size:'
$lblBatchSize.AutoSize = $true
$lblBatchSize.Location = '320,50'

$numBatchSize = New-Object System.Windows.Forms.NumericUpDown
$numBatchSize.Location = '400,47'
$numBatchSize.Width = 80
$numBatchSize.Minimum = 50
$numBatchSize.Maximum = 100
$numBatchSize.Value = 95

# Batch interval
$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Text = 'Batch Interval (min):'
$lblInterval.AutoSize = $true
$lblInterval.Location = '10,85'

$numInterval = New-Object System.Windows.Forms.NumericUpDown
$numInterval.Location = '130,82'
$numInterval.Width = 80
$numInterval.Minimum = 5
$numInterval.Maximum = 60
$numInterval.Value = 6

# Resume from record
$lblResume = New-Object System.Windows.Forms.Label
$lblResume.Text = 'Resume from record:'
$lblResume.AutoSize = $true
$lblResume.Location = '230,85'

$numResume = New-Object System.Windows.Forms.NumericUpDown
$numResume.Location = '350,82'
$numResume.Width = 80
$numResume.Minimum = 0
$numResume.Maximum = 100000
$numResume.Value = 0

# Format detection
$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = 'Material Format:'
$lblFormat.AutoSize = $true
$lblFormat.Location = '10,120'

$comboFormat = New-Object System.Windows.Forms.ComboBox
$comboFormat.Location = '120,117'
$comboFormat.Width = 150
$comboFormat.DropDownStyle = 'DropDownList'
$comboFormat.Items.AddRange(@('Books', 'Audiobook/Spoken Word'))
$comboFormat.SelectedIndex = 0

# Cycle counter display
$lblCycleCount = New-Object System.Windows.Forms.Label
$lblCycleCount.Text = 'Processing Cycle: 1'
$lblCycleCount.AutoSize = $true
$lblCycleCount.Location = '300,120'

$lblTotalProcessed = New-Object System.Windows.Forms.Label
$lblTotalProcessed.Text = 'Total API Calls: 0'
$lblTotalProcessed.AutoSize = $true
$lblTotalProcessed.Location = '450,120'

# ISBN Validation Summary
$lblValidation = New-Object System.Windows.Forms.Label
$lblValidation.Text = 'ISBN Validation Summary:'
$lblValidation.AutoSize = $true
$lblValidation.Location = '10,155'

$txtValidation = New-Object System.Windows.Forms.TextBox
$txtValidation.Location = '10,175'
$txtValidation.Size = New-Object System.Drawing.Size(650,20)
$txtValidation.ReadOnly = $true
$txtValidation.Text = 'Load a file and select lookup field to see validation summary'

# Attributes checklist
$lblAttrs = New-Object System.Windows.Forms.Label
$lblAttrs.Text = 'Attributes:'
$lblAttrs.AutoSize = $true
$lblAttrs.Location = '10,210'

$clbAttrs = New-Object System.Windows.Forms.CheckedListBox
$clbAttrs.Location = '100,207'
$clbAttrs.Size = New-Object System.Drawing.Size(200,120)
$attributes = @('subjects','authors','publishers','publish_date','number_of_pages','enh_title','lms_category','category_debug')
foreach ($a in $attributes) { $clbAttrs.Items.Add($a) }
$clbAttrs.Enabled = $false

# Output file controls
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = 'Output CSV:'
$lblOutput.AutoSize = $true
$lblOutput.Location = '10,345'

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = '100,342'
$txtOutput.Width = 400
$txtOutput.ReadOnly = $true

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = 'Save As...'
$btnBrowseOutput.Location = '510,340'
$btnBrowseOutput.Width = 75
$btnBrowseOutput.Enabled = $false

# Processing controls
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Continuous Processing'
$btnStart.Location = '150,380'
$btnStart.Width = 180
$btnStart.Enabled = $false

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Location = '340,380'
$btnStop.Width = 80
$btnStop.Enabled = $false

$btnPause = New-Object System.Windows.Forms.Button
$btnPause.Text = 'Pause'
$btnPause.Location = '430,380'
$btnPause.Width = 80
$btnPause.Enabled = $false

# Status labels
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Status: Ready'
$lblStatus.AutoSize = $true
$lblStatus.Location = '10,420'

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = 'Progress: 0 / 0'
$lblProgress.AutoSize = $true
$lblProgress.Location = '200,420'

$lblETA = New-Object System.Windows.Forms.Label
$lblETA.Text = 'ETA: --'
$lblETA.AutoSize = $true
$lblETA.Location = '400,420'

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = '10,450'
$progressBar.Size = New-Object System.Drawing.Size(650,20)

# Log box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = '10,480'
$txtLog.Size = New-Object System.Drawing.Size(650,220)
$txtLog.ReadOnly = $true

# Add controls to form
$controls = @(
    $lblInput, $txtInput, $btnBrowseInput,
    $lblField, $comboField, $lblBatchSize, $numBatchSize,
    $lblInterval, $numInterval, $lblResume, $numResume,
    $lblFormat, $comboFormat, $lblCycleCount, $lblTotalProcessed,
    $lblValidation, $txtValidation,
    $lblAttrs, $clbAttrs,
    $lblOutput, $txtOutput, $btnBrowseOutput,
    $btnStart, $btnStop, $btnPause,
    $lblStatus, $lblProgress, $lblETA,
    $progressBar, $txtLog
)
foreach ($c in $controls) { $form.Controls.Add($c) }

# Global storage
$global:Books = $null
$global:ValidISBNs = $null
$global:ProcessingActive = $false
$global:ProcessingPaused = $false
$global:StartTime = $null
$global:ProcessingSettings = $null
$global:CycleCount = 1
$global:TotalAPICalls = 0
$global:ProcessedISBNs = @{}

# Function to validate ISBNs in the dataset
function Update-ISBNValidation {
    param([string]$fieldName)
    
    if (-not $global:Books -or -not $fieldName) { return }
    
    $txtValidation.Text = "Validating ISBNs..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $validCount = 0
    $invalidCount = 0
    $emptyCount = 0
    
    $global:ValidISBNs = @()
    
    foreach ($book in $global:Books) {
        $rawISBN = $book.$fieldName
        if ([string]::IsNullOrWhiteSpace($rawISBN)) {
            $emptyCount++
        } else {
            $validISBN = Get-ValidISBN10 -rawISBN $rawISBN
            if ($validISBN) {
                $validCount++
                $global:ValidISBNs += $validISBN
            } else {
                $invalidCount++
            }
        }
    }
    
    $totalRecords = $global:Books.Count
    $txtValidation.Text = "Total: $totalRecords | Valid ISBNs: $validCount | Invalid: $invalidCount | Empty: $emptyCount"
    
    if ($validCount -gt 0) {
        $totalBatches = [Math]::Ceiling($validCount / $numBatchSize.Value)
        $estimatedHours = [Math]::Round(($totalBatches * $numInterval.Value) / 60, 1)
        
        $txtLog.AppendText("ISBN Validation Complete:`r`n")
        $txtLog.AppendText("- Valid ISBN-10 codes: $validCount`r`n")
        $txtLog.AppendText("- Invalid/malformed: $invalidCount`r`n")
        $txtLog.AppendText("- Empty fields: $emptyCount`r`n")
        $txtLog.AppendText("Estimated processing time per cycle: $estimatedHours hours ($totalBatches batches)`r`n")
    }
}

# Function to start processing
function Start-ProcessingJob {
    if (-not $comboField.SelectedItem -or -not $txtOutput.Text) { return }
    
    $selected = $clbAttrs.CheckedItems | ForEach-Object { $_ }
    if ($selected.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Please select at least one attribute to enrich.")
        return 
    }

    if (-not $global:ValidISBNs) {
        Update-ISBNValidation -fieldName $comboField.SelectedItem
    }

    if ($global:ValidISBNs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid ISBN-10 codes found in the selected field.")
        return
    }

    $global:ProcessingSettings = @{
        KeyField = $comboField.SelectedItem
        SelectedFormat = if ($comboFormat.SelectedItem -eq 'Audiobook/Spoken Word') { 'Audiobook' } else { 'Books' }
        BatchSize = $numBatchSize.Value
        IntervalMinutes = $numInterval.Value
        StartFrom = $numResume.Value
        SelectedAttributes = $selected
        OutputPath = $txtOutput.Text
    }

    $global:ProcessingActive = $true
    $global:ProcessingPaused = $false
    if ($global:CycleCount -eq 1) {
        $global:StartTime = Get-Date
    }
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $btnPause.Enabled = $true
    
    $totalRecords = $global:Books.Count
    $progressBar.Maximum = $totalRecords
    $progressBar.Value = $global:ProcessingSettings.StartFrom
    
    $lblCycleCount.Text = "Processing Cycle: $global:CycleCount"
    $lblTotalProcessed.Text = "Total API Calls: $global:TotalAPICalls"
    
    if ($global:CycleCount -eq 1) {
        $txtLog.Clear()
        $txtLog.AppendText("Starting continuous processing until all ISBNs are enriched...`r`n")
    } else {
        $txtLog.AppendText("`r`n=== CYCLE $global:CycleCount STARTED ===`r`n")
    }
    
    $txtLog.AppendText("Batch size: $($global:ProcessingSettings.BatchSize), Interval: $($global:ProcessingSettings.IntervalMinutes) minutes`r`n")
    $txtLog.AppendText("Starting from record: $($global:ProcessingSettings.StartFrom + 1)`r`n")
    $txtLog.AppendText("Valid ISBNs to process: $($global:ValidISBNs.Count)`r`n")
    $txtLog.AppendText("Already processed ISBNs: $($global:ProcessedISBNs.Count)`r`n")
    
    $apiCallCount = 0
    $validISBNsProcessed = 0
    $newEnrichmentsThisCycle = 0
    
    # Process in batches
    for ($i = $global:ProcessingSettings.StartFrom; $i -lt $totalRecords -and $global:ProcessingActive; $i += $global:ProcessingSettings.BatchSize) {
        # Check for pause
        while ($global:ProcessingPaused -and $global:ProcessingActive) {
            $lblStatus.Text = "Processing paused - click Start to resume"
            Start-Sleep -Seconds 1
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        if (-not $global:ProcessingActive) { break }
        
        $batchEnd = [Math]::Min($i + $global:ProcessingSettings.BatchSize - 1, $totalRecords - 1)
        $batchNumber = [Math]::Floor($i / $global:ProcessingSettings.BatchSize) + 1
        $totalBatches = [Math]::Ceiling($totalRecords / $global:ProcessingSettings.BatchSize)
        
        $lblStatus.Text = "Cycle $global:CycleCount - Processing batch $batchNumber of $totalBatches"
        $txtLog.AppendText("`r`nProcessing batch $batchNumber ($($i + 1) to $($batchEnd + 1))...`r`n")
        
        $batchAPICallCount = 0
        
        # Process current batch
        for ($j = $i; $j -le $batchEnd -and $global:ProcessingActive; $j++) {
            # Check for pause during batch processing
            while ($global:ProcessingPaused -and $global:ProcessingActive) {
                $lblStatus.Text = "Processing paused - click Start to resume"
                Start-Sleep -Seconds 1
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            if (-not $global:ProcessingActive) { break }
            
            $book = $global:Books[$j]
            $idRaw = $book.($global:ProcessingSettings.KeyField)
            $validISBN = Get-ValidISBN10 -rawISBN $idRaw
            
            if ($validISBN -and -not $global:ProcessedISBNs.ContainsKey($validISBN)) {
                $url = "https://openlibrary.org/api/books?bibkeys=ISBN:$validISBN&jscmd=data&format=json"
                try {
                    $resp = Invoke-RestMethod -Uri $url -ErrorAction Stop
                    $entry = $resp."ISBN:$validISBN"
                    $apiCallCount++
                    $global:TotalAPICalls++
                    $batchAPICallCount++
                    $validISBNsProcessed++
                    
                    if ($entry) {
                        $newEnrichmentsThisCycle++
                        $global:ProcessedISBNs[$validISBN] = $true
                        
                        foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
                            $val = ''
                            switch ($attr) {
                                'subjects'   { if ($entry.subjects) { $val = ($entry.subjects | ForEach-Object { $_.name }) -join '; ' } }
                                'lms_category' { $val = Get-LMSCategory -subjects $entry.subjects -format $global:ProcessingSettings.SelectedFormat }
                                'category_debug' { $val = Get-CategoryDebug -subjects $entry.subjects -format $global:ProcessingSettings.SelectedFormat }
                                'authors'    { if ($entry.authors) { $val = ($entry.authors | ForEach-Object { $_.name }) -join '; ' } }
                                'publishers' { if ($entry.publishers) { $val = ($entry.publishers | ForEach-Object { $_.name }) -join '; ' } }
                                'enh_title'  { if ($entry.title) { $val = $entry.title } }
                                default      {
                                    if ($entry.PSObject.Properties.Name -contains $attr) {
                                        $x = $entry.$attr
                                        $val = if ($x -is [System.Array]) { $x -join '; ' } else { $x }
                                    }
                                }
                            }
                            $book | Add-Member -NotePropertyName $attr -NotePropertyValue $val -Force
                        }
                    }
                } catch {
                    # Silently continue on errors to avoid spam
                }
                
                Start-Sleep -Milliseconds 100
            }
            
            $progressBar.Value = $j + 1
            $lblProgress.Text = "Progress: $($j + 1) / $totalRecords (API calls: $global:TotalAPICalls)"
            $lblTotalProcessed.Text = "Total API Calls: $global:TotalAPICalls"
            
            # Calculate ETA based on remaining unprocessed ISBNs
            if ($validISBNsProcessed -gt 0) {
                $elapsed = (Get-Date) - $global:StartTime
                $remainingISBNs = $global:ValidISBNs.Count - $global:ProcessedISBNs.Count
                
                if ($remainingISBNs -gt 0) {
                    $avgTimePerValidISBN = $elapsed.TotalMinutes / $global:ProcessedISBNs.Count
                    $estimatedRemainingTime = $remainingISBNs * $avgTimePerValidISBN
                    $lblETA.Text = "ETA: $([Math]::Round($estimatedRemainingTime, 0)) minutes ($remainingISBNs ISBNs remaining)"
                } else {
                    $lblETA.Text = "ETA: Complete! All ISBNs processed"
                }
            }
            
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Save progress after each batch
        try {
            $global:Books | Export-Csv -Path $global:ProcessingSettings.OutputPath -NoTypeInformation -Encoding UTF8
            $txtLog.AppendText("Batch $batchNumber completed. API calls in batch: $batchAPICallCount. Progress saved.`r`n")
        } catch {
            $txtLog.AppendText("Failed to save progress: $_`r`n")
        }
        
        # Update resume point for potential restart
        $global:ProcessingSettings.StartFrom = $i + $global:ProcessingSettings.BatchSize
        $numResume.Value = $global:ProcessingSettings.StartFrom
        
        # Wait between batches (except for last batch)
        if ($batchEnd -lt $totalRecords - 1 -and $global:ProcessingActive) {
            $lblStatus.Text = "Waiting $($global:ProcessingSettings.IntervalMinutes) minutes before next batch..."
            $txtLog.AppendText("Waiting $($global:ProcessingSettings.IntervalMinutes) minutes before next batch...`r`n")
            
            $waitSeconds = $global:ProcessingSettings.IntervalMinutes * 60
            for ($wait = 0; $wait -lt $waitSeconds -and $global:ProcessingActive; $wait++) {
                # Check for pause during wait
                while ($global:ProcessingPaused -and $global:ProcessingActive) {
                    $lblStatus.Text = "Processing paused during wait - click Start to resume"
                    Start-Sleep -Seconds 1
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                if (-not $global:ProcessingActive) { break }
                
                # Calculate remaining time
                $remainingSeconds = $waitSeconds - $wait
                $remainingMinutes = [Math]::Floor($remainingSeconds / 60)
                $remainingSecondsDisplay = $remainingSeconds % 60
                
                # Update status with countdown every second
                $lblStatus.Text = "Waiting... ${remainingMinutes}:${remainingSecondsDisplay:D2} remaining before next batch"
                
                # Process Windows messages to keep UI responsive
                [System.Windows.Forms.Application]::DoEvents()
                
                # Sleep for 1 second
                Start-Sleep -Seconds 1
            }
            
            # Clear waiting status and continue
            if ($global:ProcessingActive) {
                $lblStatus.Text = "Wait complete. Starting next batch..."
                $txtLog.AppendText("Wait complete. Starting next batch...`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 500
            }
        }
    }
    
    # Processing completed or stopped
    if ($global:ProcessingActive) {
        $txtLog.AppendText("`r`nCycle $global:CycleCount completed! New enrichments this cycle: $newEnrichmentsThisCycle`r`n")
        $txtLog.AppendText("Total API calls in this cycle: $apiCallCount`r`n")
        $txtLog.AppendText("Total ISBNs processed so far: $($global:ProcessedISBNs.Count) / $($global:ValidISBNs.Count)`r`n")
        
        # Check if all ISBNs have been processed OR if we've gone through all records
        $allRecordsProcessed = ($progressBar.Value -ge $totalRecords)
        $allISBNsProcessed = ($global:ProcessedISBNs.Count -ge $global:ValidISBNs.Count)
        
        if ($allISBNsProcessed -or $allRecordsProcessed) {
            $lblStatus.Text = "PROCESSING COMPLETE! All records have been processed."
            $txtLog.AppendText("`r`n=== PROCESSING COMPLETE ===`r`n")
            $txtLog.AppendText("All records in the file have been processed.`r`n")
            $txtLog.AppendText("Total API calls made: $global:TotalAPICalls`r`n")
            $txtLog.AppendText("Total processing cycles: $global:CycleCount`r`n")
            $txtLog.AppendText("ISBNs successfully enriched: $($global:ProcessedISBNs.Count)`r`n")
            
            [System.Windows.Forms.MessageBox]::Show("Processing Complete!`r`n`r`nAll records have been processed.`r`nTotal API calls: $global:TotalAPICalls`r`nTotal cycles: $global:CycleCount`r`nISBNs enriched: $($global:ProcessedISBNs.Count)`r`nFile saved to: $($global:ProcessingSettings.OutputPath)")
            
            # Reset for potential new run
            $global:ProcessedISBNs = @{}
            $global:CycleCount = 1
            $global:TotalAPICalls = 0
        } elseif ($newEnrichmentsThisCycle -eq 0) {
            # If no new enrichments were found this cycle, stop to avoid endless cycling
            $lblStatus.Text = "NO NEW ENRICHMENTS FOUND - Processing stopped."
            $txtLog.AppendText("`r`n=== NO NEW ENRICHMENTS FOUND ===`r`n")
            $txtLog.AppendText("No new data was found in this cycle. Stopping to avoid endless cycling.`r`n")
            $txtLog.AppendText("You may want to try again later as the OpenLibrary database is updated frequently.`r`n")
            
            [System.Windows.Forms.MessageBox]::Show("No new enrichments found in this cycle.`r`n`r`nProcessing stopped to avoid endless cycling.`r`nTotal API calls: $global:TotalAPICalls`r`nISBNs enriched: $($global:ProcessedISBNs.Count)")
        } else {
            # Continue cycling only if we haven't processed all records AND we found new enrichments
            $lblStatus.Text = "Cycle complete. Starting next cycle immediately..."
            $txtLog.AppendText("Auto-starting next cycle immediately...`r`n")
            
            # Reset resume point for next cycle
            $global:ProcessingSettings.StartFrom = 0
            $numResume.Value = 0
            $global:CycleCount++
            
            # Brief pause to show completion message, then restart
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 1000
            
            if ($global:ProcessingActive) {
                Start-ProcessingJob  # Recursive call to restart immediately
                return
            }
        }
    } else {
        $lblStatus.Text = "Processing stopped by user"
        $txtLog.AppendText("`r`nProcessing stopped. You can resume from record $($progressBar.Value).`r`n")
    }
    
    # Reset UI state
    $global:ProcessingActive = $false
    $global:ProcessingPaused = $false
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnPause.Enabled = $false
}

# Browse Input event
$btnBrowseInput.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV Files (*.csv)|*.csv'
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtInput.Text = $ofd.FileName
        try {
            $global:Books = Import-Csv -Path $ofd.FileName
            if ($global:Books.Count -gt 0) {
                $comboField.Items.Clear()
                $global:Books[0].PSObject.Properties.Name | ForEach-Object { $comboField.Items.Add($_) }
                $comboField.Enabled = $true
                $clbAttrs.Enabled = $true
                $btnBrowseOutput.Enabled = $true
                
                # Properly set the maximum for the resume NumericUpDown
                $numResume.Maximum = [Math]::Max($global:Books.Count, 100000)
                $numResume.Value = 0  # Reset to 0 when loading new file
                
                $txtLog.AppendText("Loaded $($global:Books.Count) records.`r`n")
                $txtLog.AppendText("Select lookup field to validate ISBNs and see time estimate.`r`n")
                
                # Reset processing state when loading new file
                $global:ProcessedISBNs = @{}
                $global:CycleCount = 1
                $global:TotalAPICalls = 0
                $lblCycleCount.Text = "Processing Cycle: 1"
                $lblTotalProcessed.Text = "Total API Calls: 0"
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to load CSV: $_")
        }
    }
})

# Field selection event
$comboField.Add_SelectedIndexChanged({
    if ($comboField.SelectedItem -and $global:Books) {
        Update-ISBNValidation -fieldName $comboField.SelectedItem
    }
})

# Browse Output event
$btnBrowseOutput.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV Files (*.csv)|*.csv'
    if ($sfd.ShowDialog() -eq 'OK') { 
        $txtOutput.Text = $sfd.FileName
        $btnStart.Enabled = $true
    }
})

# Start Processing event
$btnStart.Add_Click({
    if ($global:ProcessingPaused) {
        # Resume from pause
        $global:ProcessingPaused = $false
        $btnPause.Enabled = $true
        $txtLog.AppendText("Processing resumed.`r`n")
    } else {
        # Start new processing job
        Start-ProcessingJob
    }
})

# Stop Processing event
$btnStop.Add_Click({
    $global:ProcessingActive = $false
    $global:ProcessingPaused = $false
    $lblStatus.Text = "Stopping..."
})

# Pause Processing event
$btnPause.Add_Click({
    if ($global:ProcessingActive) {
        $global:ProcessingPaused = $true
        $btnPause.Enabled = $false
        $txtLog.AppendText("Processing paused. Click Start to resume.`r`n")
        $lblStatus.Text = "Processing paused"
    }
})

# Show the form
[void]$form.ShowDialog()
