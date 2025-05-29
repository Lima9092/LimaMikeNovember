<#
.SYNOPSIS
Batch processor for large book catalogs with rate limiting and resume capability

.DESCRIPTION
Processes large catalogs in batches of 100 items every 5+ minutes to respect API limits.
Includes resume capability, progress tracking, and ISBN-10 validation.
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

# LMS Category mapping function - maps Open Library subjects to your specific LMS categories
function Get-LMSCategory {
    param([array]$subjects, [string]$format = 'Books')
    
    if (-not $subjects) { return 'ANF' }  # Default to Adult Non-Fiction instead of 'Books'
    
    # Convert subjects to lowercase for matching
    $subjectNames = $subjects | ForEach-Object { $_.name.ToLower() }
    $allSubjects = $subjectNames -join ' '
    
    # Age group determination - be more intelligent about classification
    $ageGroup = 'Adult'  # Default
    
    # Check for EXPLICIT child indicators (must be very specific to avoid false positives)
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
    
    # Check for classic literature indicators (these should NOT override teen/child for modern books)
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
    
    # Don't treat modern books as classic literature even if they have generic "classics" tag
    # Only treat as classic if published before 1950 OR has very specific classic literature indicators
    if ($isClassicLiterature -and (
        $allSubjects -match '\b(young adult fiction|young adult works|teen fiction|ya fiction|children''?s books|juvenile fiction|children''?s fiction)\b' -or
        $allSubjects -match '\b(harry potter|twilight|hunger games|dr\. seuss|narnia)\b'
    )) {
        $isClassicLiterature = $false
    }
    
    # Check for adult fiction themes that should override child classification
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
    
    # Check for classic literature indicators (these should NOT be children's books)
    $classicLiteratureKeywords = @(
        'classic literature', 'literary fiction', 'english literature', 'american literature',
        'fiction classics', 'penguin classics', 'oxford classics', 'bantam classics', 'classical literature',
        'english fiction', 'russian literature', 'french literature', 'continental european fiction',
        'literature and fiction (general)', 'english literature, history and criticism'
    )
    
    $isClassicLiterature = $false
    foreach ($keyword in $classicLiteratureKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword)) {
            $isClassicLiterature = $true
            break
        }
    }
    
    # Don't treat modern YA/teen books as classic literature even if they have "classics" tag
    if ($isClassicLiterature -and ($allSubjects -match '\b(young adult fiction|young adult works|teen fiction|ya fiction)\b')) {
        $isClassicLiterature = $false
    }
    
    # Determine age group with better logic
    if ($hasExplicitChild -and -not $isClassicLiterature -and -not $hasAdultThemes) {
        $ageGroup = 'Child'
    }
    # Check for teen indicators (after child check)
    elseif ($allSubjects -match '\b(young adult|ya fiction|teen fiction|teenage fiction|adolescent|ages 1[2-9]|grade [6-9]|high school|middle school)\b' -and -not $isClassicLiterature -and -not $hasAdultThemes) {
        $ageGroup = 'Teen'
    }
    # If it has classic literature or adult themes, treat as adult regardless of other tags
    elseif ($isClassicLiterature -or $hasAdultThemes) {
        $ageGroup = 'Adult'
    }
    
    # Content type determination
    $isReference = $false
    $isFiction = $false
    $isGraphicNovel = $false
    
    # Reference detection - be much more specific to avoid false positives
    $explicitReferenceKeywords = @(
        'dictionaries', 'encyclopedia', 'handbook', 'manual', 'textbook', 
        'almanac', 'atlas', 'directory', 'reference work', 'reference book',
        'dictionary', 'the oxford english dictionary'
    )
    
    $isReference = $false
    foreach ($keyword in $explicitReferenceKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword) -and $allSubjects -notmatch '\b(fiction|novel|literature|story)\b') {
            $isReference = $true
            break
        }
    }
    
    # Graphic novel detection - be more specific and exclude false positives
    $isGraphicNovel = $false
    if (($allSubjects -match '\b(graphic novels?|comic books?|manga)\b' -or 
         $allSubjects -match '\b(comics & graphic novels)\b') -and 
        $allSubjects -notmatch '\b(adaptations|comic books, strips, etc\.|illustrations)\b') {
        $isGraphicNovel = $true
    }
    
    # Fiction detection (only if not reference or graphic novel)
    # Be more specific about fiction keywords
    if (-not $isReference -and -not $isGraphicNovel) {
        if ($allSubjects -match '\b(fiction|novels?|romance|mystery|thriller|horror|fantasy|science fiction|historical fiction|adventure fiction|drama|short stories|literature|literary fiction)\b' -and
            $allSubjects -notmatch '\b(non-fiction|nonfiction)\b') {
            $isFiction = $true
        }
    }
    
    # Handle audiobook format
    if ($format -eq 'Audiobook' -or $format -eq 'Spoken Word') {
        if ($isFiction) {
            switch ($ageGroup) {
                'Child' { return 'AUD-CHILD' }
                'Teen'  { return 'AUD-TEEN' }
                default { return 'AUD-ADULT' }
            }
        } else {
            # Non-fiction audiobooks
            switch ($ageGroup) {
                'Child' { return 'AUDNF-CH' }
                'Teen'  { return 'AUDNF-T' }
                default { return 'AUDNF' }
            }
        }
    }
    
    # Handle print books
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
            return 'GRAFIC'  # Assumes adult/child graphic novels use same category
        }
    }
    elseif ($isFiction) {
        switch ($ageGroup) {
            'Child' { return 'CF' }
            'Teen'  { return 'TF' }
            default { return 'AF' }  # Adult Fiction
        }
    }
    else {
        # Non-fiction
        switch ($ageGroup) {
            'Child' { return 'CNF' }
            'Teen'  { return 'TNF' }
            default { return 'ANF' }
        }
    }
}

# Debug function to show categorization logic
function Get-CategoryDebug {
    param([array]$subjects, [string]$format = 'Books')
    
    if (-not $subjects) { return 'No subjects | Default: ANF' }
    
    $subjectNames = $subjects | ForEach-Object { $_.name.ToLower() }
    $allSubjects = $subjectNames -join ' '
    
    $debug = "Subjects: [$($subjectNames -join '; ')] | "
    
    # Check for explicit child indicators
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
    
    # Check for classic literature
    $classicLiteratureKeywords = @(
        'classic', 'literature', 'literary fiction', 'english literature', 'american literature',
        'fiction classics', 'penguin classics', 'oxford classics', 'bantam classics'
    )
    
    $isClassicLiterature = $false
    foreach ($keyword in $classicLiteratureKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword)) {
            $isClassicLiterature = $true
            break
        }
    }
    
    # Age group determination
    $ageGroup = 'Adult'
    if ($hasExplicitChild -and -not $isClassicLiterature) {
        $ageGroup = 'Child'
        $debug += "Age: Child (explicit child keywords found, not classic lit) | "
    } elseif ($allSubjects -match '\b(young adult|ya fiction|teen fiction|teenage fiction|adolescent|ages 1[2-9]|grade [6-9]|high school|middle school)\b' -and -not $isClassicLiterature) {
        $ageGroup = 'Teen'
        $debug += "Age: Teen (teen keywords found, not classic lit) | "
    } elseif ($isClassicLiterature) {
        $debug += "Age: Adult (classic literature override) | "
    } else {
        $debug += "Age: Adult (default) | "
    }
    
    $debug += "ExplicitChild: $hasExplicitChild | ClassicLit: $isClassicLiterature | "
    
    # Content type
    $explicitReferenceKeywords = @(
        'dictionaries', 'encyclopedia', 'handbook', 'manual', 'textbook', 
        'almanac', 'atlas', 'directory', 'reference work', 'reference book',
        'dictionary', 'the oxford english dictionary'
    )
    
    $isReference = $false
    foreach ($keyword in $explicitReferenceKeywords) {
        if ($allSubjects -match [regex]::Escape($keyword) -and $allSubjects -notmatch '\b(fiction|novel|literature|story)\b') {
            $isReference = $true
            break
        }
    }
    
    $isGraphicNovel = $false
    if ($allSubjects -match '\b(graphic novels?|comic books?|manga)\b' -and 
        $allSubjects -notmatch '\b(comics & graphic novels, adaptations|adaptations)\b') {
        $isGraphicNovel = $true
    }
    
    $isFiction = (-not $isReference -and -not $isGraphicNovel) -and ($allSubjects -match '\b(fiction|novels?|romance|mystery|thriller|horror|fantasy|science fiction|historical fiction|adventure fiction|drama|short stories|literature|literary fiction)\b' -and
        $allSubjects -notmatch '\b(non-fiction|nonfiction)\b')
    
    $debug += "Reference: $isReference | Fiction: $isFiction | Graphic: $isGraphicNovel | "
    $debug += "Format: $format | "
    $debug += "Final: $(Get-LMSCategory -subjects $subjects -format $format)"
    
    return $debug
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Book Enrichment - Large Catalogs"
$form.Size = New-Object System.Drawing.Size(700, 700)
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
$numBatchSize.Value = 95  # Slightly under limit for safety

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
$numInterval.Value = 6  # 6 minutes for safety margin

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

# Format detection (add after lookup field selection)
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

# ISBN Validation Summary (new section)
$lblValidation = New-Object System.Windows.Forms.Label
$lblValidation.Text = 'ISBN Validation Summary:'
$lblValidation.AutoSize = $true
$lblValidation.Location = '300,120'

$txtValidation = New-Object System.Windows.Forms.TextBox
$txtValidation.Location = '300,140'
$txtValidation.Size = New-Object System.Drawing.Size(360,20)
$txtValidation.ReadOnly = $true
$txtValidation.Text = 'Load a file and select lookup field to see validation summary'

# Attributes checklist (moved down)
$lblAttrs = New-Object System.Windows.Forms.Label
$lblAttrs.Text = 'Attributes:'
$lblAttrs.AutoSize = $true
$lblAttrs.Location = '10,175'

$clbAttrs = New-Object System.Windows.Forms.CheckedListBox
$clbAttrs.Location = '100,172'
$clbAttrs.Size = New-Object System.Drawing.Size(200,120)
$attributes = @('subjects','authors','publishers','publish_date','number_of_pages','title','lms_category','category_debug')
foreach ($a in $attributes) { $clbAttrs.Items.Add($a) }
$clbAttrs.Enabled = $false

# Output file controls
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = 'Output CSV:'
$lblOutput.AutoSize = $true
$lblOutput.Location = '10,310'

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = '100,307'
$txtOutput.Width = 400
$txtOutput.ReadOnly = $true

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = 'Save As...'
$btnBrowseOutput.Location = '510,305'
$btnBrowseOutput.Width = 75
$btnBrowseOutput.Enabled = $false

# Processing controls
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Batch Processing'
$btnStart.Location = '200,345'
$btnStart.Width = 150
$btnStart.Enabled = $false

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Location = '360,345'
$btnStop.Width = 80
$btnStop.Enabled = $false

# Status labels
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Status: Ready'
$lblStatus.AutoSize = $true
$lblStatus.Location = '10,385'

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = 'Progress: 0 / 0'
$lblProgress.AutoSize = $true
$lblProgress.Location = '200,385'

$lblETA = New-Object System.Windows.Forms.Label
$lblETA.Text = 'ETA: --'
$lblETA.AutoSize = $true
$lblETA.Location = '400,385'

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = '10,415'
$progressBar.Size = New-Object System.Drawing.Size(650,20)

# Log box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = '10,445'
$txtLog.Size = New-Object System.Drawing.Size(650,215)
$txtLog.ReadOnly = $true

# Add controls to form
$controls = @(
    $lblInput, $txtInput, $btnBrowseInput,
    $lblField, $comboField, $lblBatchSize, $numBatchSize,
    $lblInterval, $numInterval, $lblResume, $numResume,
    $lblFormat, $comboFormat,
    $lblValidation, $txtValidation,
    $lblAttrs, $clbAttrs,
    $lblOutput, $txtOutput, $btnBrowseOutput,
    $btnStart, $btnStop,
    $lblStatus, $lblProgress, $lblETA,
    $progressBar, $txtLog
)
foreach ($c in $controls) { $form.Controls.Add($c) }

# Global storage
$global:Books = $null
$global:ValidISBNs = $null
$global:ProcessingActive = $false
$global:StartTime = $null

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
    
    # Update time estimate based on valid ISBNs
    if ($validCount -gt 0) {
        $totalBatches = [Math]::Ceiling($validCount / $numBatchSize.Value)
        $estimatedHours = [Math]::Round(($totalBatches * $numInterval.Value) / 60, 1)
        
        $txtLog.AppendText("ISBN Validation Complete:`r`n")
        $txtLog.AppendText("- Valid ISBN-10 codes: $validCount`r`n")
        $txtLog.AppendText("- Invalid/malformed: $invalidCount`r`n")
        $txtLog.AppendText("- Empty fields: $emptyCount`r`n")
        $txtLog.AppendText("Estimated processing time: $estimatedHours hours ($totalBatches batches) for valid ISBNs only`r`n")
    }
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
                $numResume.Maximum = $global:Books.Count
                
                $txtLog.AppendText("Loaded $($global:Books.Count) records.`r`n")
                $txtLog.AppendText("Select lookup field to validate ISBNs and see time estimate.`r`n")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to load CSV: $_")
        }
    }
})

# Field selection event (new)
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
    if (-not $comboField.SelectedItem -or -not $txtOutput.Text) { return }
    
    $selected = $clbAttrs.CheckedItems | ForEach-Object { $_ }
    if ($selected.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Please select at least one attribute to enrich.")
        return 
    }

    # Validate ISBNs if not already done
    if (-not $global:ValidISBNs) {
        Update-ISBNValidation -fieldName $comboField.SelectedItem
    }

    if ($global:ValidISBNs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid ISBN-10 codes found in the selected field.")
        return
    }

    $global:ProcessingActive = $true
    $global:StartTime = Get-Date
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    
    $keyField = $comboField.SelectedItem
    $selectedFormat = if ($comboFormat.SelectedItem -eq 'Audiobook/Spoken Word') { 'Audiobook' } else { 'Books' }
    $batchSize = $numBatchSize.Value
    $intervalMinutes = $numInterval.Value
    $startFrom = $numResume.Value
    
    $totalRecords = $global:Books.Count
    $progressBar.Maximum = $totalRecords
    $progressBar.Value = $startFrom
    
    $txtLog.Clear()
    $txtLog.AppendText("Starting batch processing...`r`n")
    $txtLog.AppendText("Batch size: $batchSize, Interval: $intervalMinutes minutes`r`n")
    $txtLog.AppendText("Starting from record: $($startFrom + 1)`r`n")
    $txtLog.AppendText("Valid ISBNs to process: $($global:ValidISBNs.Count)`r`n")
    
    $apiCallCount = 0
    $validISBNsProcessed = 0
    
    # Process in batches
    for ($i = $startFrom; $i -lt $totalRecords -and $global:ProcessingActive; $i += $batchSize) {
        $batchEnd = [Math]::Min($i + $batchSize - 1, $totalRecords - 1)
        $batchNumber = [Math]::Floor($i / $batchSize) + 1
        $totalBatches = [Math]::Ceiling($totalRecords / $batchSize)
        
        $lblStatus.Text = "Processing batch $batchNumber of $totalBatches"
        $txtLog.AppendText("`r`nProcessing batch $batchNumber ($($i + 1) to $($batchEnd + 1))...`r`n")
        
        $batchAPICallCount = 0
        
        # Process current batch
        for ($j = $i; $j -le $batchEnd -and $global:ProcessingActive; $j++) {
            $book = $global:Books[$j]
            $idRaw = $book.$keyField
            $validISBN = Get-ValidISBN10 -rawISBN $idRaw
            
            if ($validISBN) {
                $url = "https://openlibrary.org/api/books?bibkeys=ISBN:$validISBN&jscmd=data&format=json"
                try {
                    $resp = Invoke-RestMethod -Uri $url -ErrorAction Stop
                    $entry = $resp."ISBN:$validISBN"
                    $apiCallCount++
                    $batchAPICallCount++
                    $validISBNsProcessed++
                    
                    if ($entry) {
                        foreach ($attr in $selected) {
                            $val = ''
                            switch ($attr) {
                                'subjects'   { if ($entry.subjects) { $val = ($entry.subjects | ForEach-Object { $_.name }) -join '; ' } }
                                'lms_category' { $val = Get-LMSCategory -subjects $entry.subjects -format $selectedFormat }
                                'category_debug' { $val = Get-CategoryDebug -subjects $entry.subjects -format $selectedFormat }
                                'authors'    { if ($entry.authors) { $val = ($entry.authors | ForEach-Object { $_.name }) -join '; ' } }
                                'publishers' { if ($entry.publishers) { $val = ($entry.publishers | ForEach-Object { $_.name }) -join '; ' } }
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
                
                Start-Sleep -Milliseconds 100  # Small delay between individual requests
            }
            
            $progressBar.Value = $j + 1
            $lblProgress.Text = "Progress: $($j + 1) / $totalRecords (API calls: $apiCallCount)"
            
            # Calculate ETA based on valid ISBNs processed AND batch intervals
            if ($validISBNsProcessed -gt 0) {
                $elapsed = (Get-Date) - $global:StartTime
                $remainingValidISBNs = $global:ValidISBNs.Count - $validISBNsProcessed
                
                # Calculate remaining batches
                $currentBatch = [Math]::Floor(($j - $startFrom) / $batchSize) + 1
                $totalBatches = [Math]::Ceiling($global:ValidISBNs.Count / $batchSize)
                $remainingBatches = $totalBatches - $currentBatch
                
                # Estimate processing time for remaining ISBNs
                $avgTimePerValidISBN = $elapsed.TotalMinutes / $validISBNsProcessed
                $remainingProcessingTime = $remainingValidISBNs * $avgTimePerValidISBN
                
                # Add time for remaining batch intervals (6 minutes each, but not after the last batch)
                $remainingIntervalTime = $remainingBatches * $intervalMinutes
                
                $totalRemainingTime = $remainingProcessingTime + $remainingIntervalTime
                $lblETA.Text = "ETA: $([Math]::Round($totalRemainingTime, 0)) minutes"
            }
            
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Save progress after each batch
        try {
            $global:Books | Export-Csv -Path $txtOutput.Text -NoTypeInformation -Encoding UTF8
            $txtLog.AppendText("Batch $batchNumber completed. API calls in batch: $batchAPICallCount. Progress saved.`r`n")
        } catch {
            $txtLog.AppendText("Failed to save progress: $_`r`n")
        }
        
        # Wait between batches (except for last batch)
        if ($batchEnd -lt $totalRecords - 1 -and $global:ProcessingActive) {
            $lblStatus.Text = "Waiting $intervalMinutes minutes before next batch..."
            $txtLog.AppendText("Waiting $intervalMinutes minutes before next batch...`r`n")
            
            $waitSeconds = $intervalMinutes * 60
            for ($wait = 0; $wait -lt $waitSeconds -and $global:ProcessingActive; $wait++) {
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
                Start-Sleep -Milliseconds 500  # Brief pause to show the message
            }
        }
    }
    
    if ($global:ProcessingActive) {
        $lblStatus.Text = "Processing completed!"
        $txtLog.AppendText("`r`nAll batches completed successfully!`r`n")
        $txtLog.AppendText("Total API calls made: $apiCallCount`r`n")
        [System.Windows.Forms.MessageBox]::Show("Enrichment completed! File saved to $($txtOutput.Text)`r`nTotal API calls: $apiCallCount")
    } else {
        $lblStatus.Text = "Processing stopped by user"
        $txtLog.AppendText("`r`nProcessing stopped. You can resume from record $($progressBar.Value).`r`n")
    }
    
    $global:ProcessingActive = $false
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
})

# Stop Processing event
$btnStop.Add_Click({
    $global:ProcessingActive = $false
    $lblStatus.Text = "Stopping..."
})

# Show the form
[void]$form.ShowDialog()
