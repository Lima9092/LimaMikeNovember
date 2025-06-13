<#
.SYNOPSIS
Batch processor for large book catalogs with local ISBN database, optimized processing, and rate limiting

.DESCRIPTION
Uses a two-phase approach: Phase 1 processes all local database hits instantly, Phase 2 handles API calls with rate limiting.
Includes resume capability, progress tracking, ISBN-10 validation, and continuous processing until all ISBNs are enriched.
Automatically updates the local ISBN database with new API results AND appends data to the original CSV file.
Includes "Local DB Only" mode for instant processing without API calls.
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

# Local ISBN Database functions
function Initialize-LocalISBNDatabase {
    param([string]$scriptPath)
    
    $dbPath = Join-Path $scriptPath "ISBN-Enhanced-DB.csv"
    
    if (-not (Test-Path $dbPath)) {
        # Create new database file with headers
        $headers = "RCN,subjects,authors,publishers,publish_date,number_of_pages,enh_title,lms_category,category_debug,last_updated"
        $headers | Out-File -FilePath $dbPath -Encoding UTF8
        $txtLog.AppendText("Created new ISBN database: $dbPath`r`n")
        return @{}
    }
    
    try {
        $dbData = Import-Csv -Path $dbPath -ErrorAction Stop
        $txtLog.AppendText("Loaded local ISBN database with $($dbData.Count) records from: $dbPath`r`n")
        
        # Convert to hashtable for fast lookups
        $dbHashTable = @{}
        $loadedCount = 0
        $skippedCount = 0
        
        foreach ($record in $dbData) {
            if ($record.RCN -and $record.RCN.Trim() -ne '') {
                $cleanRCN = $record.RCN.Trim()
                $dbHashTable[$cleanRCN] = $record
                $loadedCount++
            } else {
                $skippedCount++
            }
        }
        
        $txtLog.AppendText("Database loading summary: $loadedCount records loaded, $skippedCount skipped (empty RCN)`r`n")
        
        return $dbHashTable
    } catch {
        $txtLog.AppendText("Error loading ISBN database: $_`r`n")
        return @{}
    }
}

function Save-LocalISBNDatabase {
    param([hashtable]$dbHashTable, [string]$scriptPath)
    
    $dbPath = Join-Path $scriptPath "ISBN-Enhanced-DB.csv"
    
    try {
        # Convert hashtable back to array of objects
        $dbArray = @()
        foreach ($key in $dbHashTable.Keys) {
            $dbArray += $dbHashTable[$key]
        }
        
        # Only save if we have data
        if ($dbArray.Count -gt 0) {
            # Export to CSV
            $dbArray | Export-Csv -Path $dbPath -NoTypeInformation -Encoding UTF8
            $txtLog.AppendText("Successfully saved $($dbArray.Count) records to local database`r`n")
            return $true
        } else {
            $txtLog.AppendText("Warning: No data to save to local database`r`n")
            return $false
        }
    } catch {
        $txtLog.AppendText("Error saving ISBN database: $_`r`n")
        return $false
    }
}

function Get-ISBNFromLocalDB {
    param([string]$isbn, [hashtable]$dbHashTable)
    
    if ($dbHashTable.ContainsKey($isbn)) {
        return $dbHashTable[$isbn]
    }
    
    return $null
}

function Add-ISBNToLocalDB {
    param([string]$isbn, [object]$apiData, [hashtable]$dbHashTable, [string]$selectedFormat = 'Books')
    
    # Create new record for local database
    $dbRecord = [PSCustomObject]@{
        RCN = $isbn
        subjects = ''
        authors = ''
        publishers = ''
        publish_date = ''
        number_of_pages = ''
        enh_title = ''
        lms_category = ''
        category_debug = ''
        last_updated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    
    # Populate from API data
    if ($apiData.subjects) { 
        $dbRecord.subjects = ($apiData.subjects | ForEach-Object { $_.name }) -join '; ' 
        # Calculate LMS category and debug info
        $dbRecord.lms_category = Get-LMSCategory -subjects $apiData.subjects -format $selectedFormat
        $dbRecord.category_debug = Get-CategoryDebug -subjects $apiData.subjects -format $selectedFormat
    } else {
        $dbRecord.lms_category = 'AF'
        $dbRecord.category_debug = 'No subjects | Default: AF'
    }
    
    if ($apiData.authors) { 
        $dbRecord.authors = ($apiData.authors | ForEach-Object { $_.name }) -join '; ' 
    }
    if ($apiData.publishers) { 
        $dbRecord.publishers = ($apiData.publishers | ForEach-Object { $_.name }) -join '; ' 
    }
    if ($apiData.publish_date) { 
        $dbRecord.publish_date = $apiData.publish_date 
    }
    if ($apiData.number_of_pages) { 
        $dbRecord.number_of_pages = $apiData.number_of_pages 
    }
    if ($apiData.title) { 
        $dbRecord.enh_title = $apiData.title 
    }
    
    # Add to hashtable
    $dbHashTable[$isbn] = $dbRecord
    
    return $dbRecord
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

# Helper function to parse subjects from database string
function Parse-SubjectsFromDB {
    param([string]$subjectsString)
    
    if (-not $subjectsString) { return $null }
    
    $subjectNames = $subjectsString -split '; '
    $subjects = @()
    foreach ($name in $subjectNames) {
        if ($name.Trim()) {
            $subjects += [PSCustomObject]@{ name = $name.Trim() }
        }
    }
    
    return $subjects
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

# FIXED: Apply data to CSV row - Enhanced version
function Apply-DataToCSVRow {
    param([object]$book, [object]$dataEntry, [array]$selectedAttributes, [string]$selectedFormat)
    
    foreach ($attr in $selectedAttributes) {
        $val = ''
        switch ($attr) {
            'subjects' { 
                $val = if ($dataEntry.subjects) { $dataEntry.subjects } else { '' } 
            }
            'lms_category' { 
                $val = if ($dataEntry.lms_category -and $dataEntry.lms_category.Trim() -ne '') { 
                    $dataEntry.lms_category 
                } else {
                    $subjects = Parse-SubjectsFromDB -subjectsString $dataEntry.subjects
                    Get-LMSCategory -subjects $subjects -format $selectedFormat
                }
            }
            'category_debug' { 
                $val = if ($dataEntry.category_debug -and $dataEntry.category_debug.Trim() -ne '') { 
                    $dataEntry.category_debug 
                } else {
                    $subjects = Parse-SubjectsFromDB -subjectsString $dataEntry.subjects
                    Get-CategoryDebug -subjects $subjects -format $selectedFormat
                }
            }
            'authors' { 
                $val = if ($dataEntry.authors) { $dataEntry.authors } else { '' } 
            }
            'publishers' { 
                $val = if ($dataEntry.publishers) { $dataEntry.publishers } else { '' } 
            }
            'enh_title' { 
                $val = if ($dataEntry.enh_title) { $dataEntry.enh_title } else { '' } 
            }
            'publish_date' { 
                $val = if ($dataEntry.publish_date) { $dataEntry.publish_date } else { '' } 
            }
            'number_of_pages' { 
                $val = if ($dataEntry.number_of_pages) { $dataEntry.number_of_pages } else { '' } 
            }
        }
        
        # CRITICAL FIX: Ensure we're directly setting the property on the book object
        if ($book.PSObject.Properties.Name -contains $attr) {
            # Property exists, update it directly
            $book.$attr = $val
        } else {
            # Property doesn't exist, add it
            try {
                $book | Add-Member -NotePropertyName $attr -NotePropertyValue $val -Force
            } catch {
                # Fallback: try different method
                try {
                    $book | Add-Member -MemberType NoteProperty -Name $attr -Value $val -Force
                } catch {
                    # Last resort: create new property manually
                    $book.PSObject.Properties.Add((New-Object PSNoteProperty($attr, $val)))
                }
            }
        }
    }
}

# NEW FUNCTION: Process using Local DB Only
function Start-LocalDBOnlyProcessing {
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

    if ($global:LocalISBNDB.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Local database is empty. Please run regular processing first to build the database.")
        return
    }

    $global:ProcessingActive = $true
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $btnPause.Enabled = $false  # No pause for local-only processing
    
    $lblStatus.Text = "Processing with Local DB Only..."
    $txtLog.AppendText("`r`n=== LOCAL DB ONLY PROCESSING STARTED ===`r`n")
    
    $totalRecords = $global:Books.Count
    $processedCount = 0
    $localHitCount = 0
    $missedCount = 0
    $actualEnrichments = 0
    
    $progressBar.Maximum = $totalRecords
    $progressBar.Value = 0
    
    $selectedFormat = if ($comboFormat.SelectedItem -eq 'Audiobook/Spoken Word') { 'Audiobook' } else { 'Books' }
    $keyField = $comboField.SelectedItem
    
    # Create a copy of the Books array to ensure we're working with the right data
    $enrichedBooks = @()
    
    # Process all records instantly (no rate limiting needed)
    for ($i = 0; $i -lt $totalRecords -and $global:ProcessingActive; $i++) {
        $book = $global:Books[$i].PSObject.Copy()  # Create a copy to avoid reference issues
        $idRaw = $book.$keyField
        $validISBN = Get-ValidISBN10 -rawISBN $idRaw
        
        if ($validISBN) {
            # Check if Update Existing mode and already has category
            $shouldProcess = $true
            if ($chkUpdateExisting.Checked) {
                $currentCategory = $null
                if ($book.PSObject.Properties.Name -contains 'lms_category') {
                    $currentCategory = $book.lms_category
                }
                $shouldProcess = (-not $currentCategory -or $currentCategory.Trim() -eq '' -or $currentCategory.Trim() -eq 'NULL')
            }
            
            if ($shouldProcess) {
                # Check local database only
                $localEntry = Get-ISBNFromLocalDB -isbn $validISBN -dbHashTable $global:LocalISBNDB
                if ($localEntry) {
                    # Apply enriched data to the book record
                    Apply-DataToCSVRow -book $book -dataEntry $localEntry -selectedAttributes $selected -selectedFormat $selectedFormat
                    
                    $localHitCount++
                    $actualEnrichments++
                } else {
                    $missedCount++
                    
                    # Still add empty columns for consistency
                    foreach ($attr in $selected) {
                        if (-not ($book.PSObject.Properties.Name -contains $attr)) {
                            $book | Add-Member -NotePropertyName $attr -NotePropertyValue '' -Force
                        }
                    }
                }
            } else {
                # Add empty columns for records that shouldn't be processed
                foreach ($attr in $selected) {
                    if (-not ($book.PSObject.Properties.Name -contains $attr)) {
                        $book | Add-Member -NotePropertyName $attr -NotePropertyValue '' -Force
                    }
                }
            }
            $processedCount++
        } else {
            # Add empty columns for invalid ISBNs
            foreach ($attr in $selected) {
                if (-not ($book.PSObject.Properties.Name -contains $attr)) {
                    $book | Add-Member -NotePropertyName $attr -NotePropertyValue '' -Force
                }
            }
        }
        
        # Add the processed book to our enriched collection
        $enrichedBooks += $book
        
        $progressBar.Value = $i + 1
        $lblProgress.Text = "Progress: $($i + 1) / $totalRecords"
        
        if ($i % 100 -eq 0) {  # Update UI every 100 records
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    
    # Update the global Books array with our enriched data
    $global:Books = $enrichedBooks
    
    # Save the enriched CSV file
    if ($global:ProcessingActive) {
        try {
            $global:Books | Export-Csv -Path $txtOutput.Text -NoTypeInformation -Encoding UTF8
            $txtLog.AppendText("Successfully saved enriched file to: $($txtOutput.Text)`r`n")
            $txtLog.AppendText("Enriched $actualEnrichments records with new data columns.`r`n")
            
            # Verify the file was written correctly
            $testImport = Import-Csv -Path $txtOutput.Text
            $firstRecord = $testImport[0]
            $addedColumns = @()
            foreach ($attr in $selected) {
                if ($firstRecord.PSObject.Properties.Name -contains $attr) {
                    $addedColumns += $attr
                }
            }
            $txtLog.AppendText("Verified: Added columns [$($addedColumns -join ', ')] to output file.`r`n")
            
            $txtLog.AppendText("`r`n=== LOCAL DB ONLY PROCESSING COMPLETE ===`r`n")
            $txtLog.AppendText("Total records processed: $totalRecords`r`n")
            $txtLog.AppendText("Valid ISBNs found: $processedCount`r`n")
            $txtLog.AppendText("Local DB hits: $localHitCount`r`n")
            $txtLog.AppendText("Local DB misses: $missedCount`r`n")
            $txtLog.AppendText("Actual enrichments: $actualEnrichments`r`n")
            
            $efficiency = if ($processedCount -gt 0) { [Math]::Round(($localHitCount / $processedCount) * 100, 1) } else { 0 }
            $txtLog.AppendText("Local DB efficiency: $efficiency%`r`n")
            
            $lblStatus.Text = "Local DB processing complete!"
            
            [System.Windows.Forms.MessageBox]::Show("Local DB Only Processing Complete!`r`n`r`nTotal ISBNs: $processedCount`r`nLocal DB hits: $localHitCount (instant)`r`nLocal DB misses: $missedCount`r`nActual enrichments: $actualEnrichments`r`nEfficiency: $efficiency%`r`n`r`nColumns added: $($addedColumns -join ', ')`r`nFile saved to: $($txtOutput.Text)")
            
        } catch {
            $txtLog.AppendText("ERROR saving file: $_`r`n")
            [System.Windows.Forms.MessageBox]::Show("Error saving file: $_")
        }
    } else {
        $lblStatus.Text = "Local DB processing stopped by user"
    }
    
    # Reset UI state
    $global:ProcessingActive = $false
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnPause.Enabled = $false
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "ISBN Enrichment - Local DB + API"
$form.Size = New-Object System.Drawing.Size(700, 850)  # Made slightly taller
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

# Database statistics
$lblDBStats = New-Object System.Windows.Forms.Label
$lblDBStats.Text = 'Local DB: Not loaded'
$lblDBStats.AutoSize = $true
$lblDBStats.Location = '300,120'

# Update existing checkbox
$chkUpdateExisting = New-Object System.Windows.Forms.CheckBox
$chkUpdateExisting.Text = 'Update Existing (fill missing lms_category only)'
$chkUpdateExisting.Location = '10,155'
$chkUpdateExisting.AutoSize = $true

# NEW: Local DB Only checkbox
$chkLocalDBOnly = New-Object System.Windows.Forms.CheckBox
$chkLocalDBOnly.Text = 'Local DB Only (no API calls)'
$chkLocalDBOnly.Location = '350,155'
$chkLocalDBOnly.AutoSize = $true

# Cycle counter display
$lblCycleCount = New-Object System.Windows.Forms.Label
$lblCycleCount.Text = 'Processing Cycle: 1'
$lblCycleCount.AutoSize = $true
$lblCycleCount.Location = '10,185'

$lblTotalProcessed = New-Object System.Windows.Forms.Label
$lblTotalProcessed.Text = 'Total API Calls: 0'
$lblTotalProcessed.AutoSize = $true
$lblTotalProcessed.Location = '150,185'

$lblLocalHits = New-Object System.Windows.Forms.Label
$lblLocalHits.Text = 'Local DB Hits: 0'
$lblLocalHits.AutoSize = $true
$lblLocalHits.Location = '300,185'

# ISBN Validation Summary
$lblValidation = New-Object System.Windows.Forms.Label
$lblValidation.Text = 'ISBN Validation Summary:'
$lblValidation.AutoSize = $true
$lblValidation.Location = '10,215'

$txtValidation = New-Object System.Windows.Forms.TextBox
$txtValidation.Location = '10,235'
$txtValidation.Size = New-Object System.Drawing.Size(650,20)
$txtValidation.ReadOnly = $true
$txtValidation.Text = 'Load a file and select lookup field to see validation summary'

# Attributes checklist
$lblAttrs = New-Object System.Windows.Forms.Label
$lblAttrs.Text = 'Attributes:'
$lblAttrs.AutoSize = $true
$lblAttrs.Location = '10,270'

$clbAttrs = New-Object System.Windows.Forms.CheckedListBox
$clbAttrs.Location = '100,267'
$clbAttrs.Size = New-Object System.Drawing.Size(200,120)
$attributes = @('subjects','authors','publishers','publish_date','number_of_pages','enh_title','lms_category','category_debug')
foreach ($a in $attributes) { $clbAttrs.Items.Add($a) }
$clbAttrs.Enabled = $false

# Output file controls
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = 'Output CSV:'
$lblOutput.AutoSize = $true
$lblOutput.Location = '10,405'

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = '100,402'
$txtOutput.Width = 400
$txtOutput.ReadOnly = $true

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = 'Save As...'
$btnBrowseOutput.Location = '510,400'
$btnBrowseOutput.Width = 75
$btnBrowseOutput.Enabled = $false

# Processing controls
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Processing'
$btnStart.Location = '150,440'
$btnStart.Width = 180
$btnStart.Enabled = $false

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Location = '340,440'
$btnStop.Width = 80
$btnStop.Enabled = $false

$btnPause = New-Object System.Windows.Forms.Button
$btnPause.Text = 'Pause'
$btnPause.Location = '430,440'
$btnPause.Width = 80
$btnPause.Enabled = $false

# Status labels
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Status: Ready'
$lblStatus.AutoSize = $true
$lblStatus.Location = '10,480'

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = 'Progress: 0 / 0'
$lblProgress.AutoSize = $true
$lblProgress.Location = '200,480'

$lblETA = New-Object System.Windows.Forms.Label
$lblETA.Text = 'ETA: --'
$lblETA.AutoSize = $true
$lblETA.Location = '400,480'

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = '10,510'
$progressBar.Size = New-Object System.Drawing.Size(650,20)

# Log box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = '10,540'
$txtLog.Size = New-Object System.Drawing.Size(650,260)  # Made taller
$txtLog.ReadOnly = $true

# Add controls to form
$controls = @(
    $lblInput, $txtInput, $btnBrowseInput,
    $lblField, $comboField, $lblBatchSize, $numBatchSize,
    $lblInterval, $numInterval, $lblResume, $numResume,
    $lblFormat, $comboFormat, $lblDBStats,
    $chkUpdateExisting, $chkLocalDBOnly,  # Added new checkbox
    $lblCycleCount, $lblTotalProcessed, $lblLocalHits,
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
$global:LocalDBHits = 0
$global:ProcessedISBNs = @{}
$global:LocalISBNDB = @{}
$global:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Function to validate ISBNs in the dataset
function Update-ISBNValidation {
    param([string]$fieldName)
    
    if (-not $global:Books -or -not $fieldName) { return }
    
    $txtValidation.Text = "Validating ISBNs..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $validCount = 0
    $invalidCount = 0
    $emptyCount = 0
    $localDBCount = 0
    $alreadyProcessedCount = 0
    $hasLMSCategory = 0
    $missingLMSCategory = 0
    
    $global:ValidISBNs = @()
    
    # Check if lms_category column exists
    $hasLMSCategoryColumn = $global:Books[0].PSObject.Properties.Name -contains 'lms_category'
    
    foreach ($book in $global:Books) {
        $rawISBN = $book.$fieldName
        if ([string]::IsNullOrWhiteSpace($rawISBN)) {
            $emptyCount++
        } else {
            $validISBN = Get-ValidISBN10 -rawISBN $rawISBN
            if ($validISBN) {
                $validCount++
                $global:ValidISBNs += $validISBN
                
                # Check if it exists in local database
                if ($global:LocalISBNDB.ContainsKey($validISBN)) {
                    $localDBCount++
                }
                
                # Check if already processed in this session
                if ($global:ProcessedISBNs.ContainsKey($validISBN)) {
                    $alreadyProcessedCount++
                }
                
                # Check LMS category status if in update mode
                if ($chkUpdateExisting.Checked -and $hasLMSCategoryColumn) {
                    $currentCategory = $book.lms_category
                    if ($currentCategory -and $currentCategory.Trim() -ne '' -and $currentCategory.Trim() -ne 'NULL' -and $currentCategory.Trim() -ne '0') {
                        $hasLMSCategory++
                    } else {
                        $missingLMSCategory++
                    }
                }
            } else {
                $invalidCount++
            }
        }
    }
    
    $totalRecords = $global:Books.Count
    $needingAPICall = $validCount - $localDBCount
    $remainingToProcess = $validCount - $alreadyProcessedCount
    
    # Special handling for Local DB Only mode
    if ($chkLocalDBOnly.Checked) {
        $txtValidation.Text = "LOCAL DB ONLY MODE: Total: $totalRecords | Valid ISBNs: $validCount | In Local DB: $localDBCount | NOT in Local DB: $needingAPICall"
    } elseif ($chkUpdateExisting.Checked -and $hasLMSCategoryColumn) {
        # Adjust calculations for update existing mode
        $needingUpdate = $missingLMSCategory
        $missingNotInDB = 0
        foreach ($book in $global:Books) {
            $rawISBN = $book.$fieldName
            $validISBN = Get-ValidISBN10 -rawISBN $rawISBN
            if ($validISBN) {
                $currentCategory = $book.lms_category
                $isMissingCategory = (-not $currentCategory -or $currentCategory.Trim() -eq '')
                $isInLocalDB = $global:LocalISBNDB.ContainsKey($validISBN)
                if ($isMissingCategory -and -not $isInLocalDB) {
                    $missingNotInDB++
                }
            }
        }
        $needingAPICall = $missingNotInDB
        $txtValidation.Text = "UPDATE MODE: Total: $totalRecords | Valid ISBNs: $validCount | Has Category: $hasLMSCategory | Missing Category: $missingLMSCategory | Missing & Need API: $needingAPICall"
    } else {
        $txtValidation.Text = "Total: $totalRecords | Valid ISBNs: $validCount | Invalid: $invalidCount | Empty: $emptyCount | In Local DB: $localDBCount | Remaining: $remainingToProcess"
    }
    
    if ($validCount -gt 0) {
        $totalBatches = [Math]::Ceiling($needingAPICall / $numBatchSize.Value)
        $estimatedHours = [Math]::Round(($totalBatches * $numInterval.Value) / 60, 1)
        
        $txtLog.AppendText("ISBN Validation Complete:`r`n")
        $txtLog.AppendText("- Total records: $totalRecords`r`n")
        $txtLog.AppendText("- Valid ISBN-10 codes: $validCount`r`n")
        $txtLog.AppendText("- Already in local DB: $localDBCount`r`n")
        
        if ($chkLocalDBOnly.Checked) {
            $txtLog.AppendText("LOCAL DB ONLY MODE:`r`n")
            $txtLog.AppendText("- Can be enriched from local DB: $localDBCount`r`n")
            $txtLog.AppendText("- NOT in local DB (will be skipped): $needingAPICall`r`n")
            $txtLog.AppendText("- Processing time: INSTANT (no API calls)`r`n")
        } elseif ($chkUpdateExisting.Checked) {
            if ($hasLMSCategoryColumn) {
                $txtLog.AppendText("UPDATE EXISTING MODE:`r`n")
                $txtLog.AppendText("- Records with lms_category: $hasLMSCategory`r`n")
                $txtLog.AppendText("- Records missing lms_category: $missingLMSCategory`r`n")
                
                if ($missingLMSCategory -eq 0) {
                    $txtLog.AppendText("*** ALL RECORDS ALREADY HAVE CATEGORIES - NO PROCESSING NEEDED! ***`r`n")
                    [System.Windows.Forms.MessageBox]::Show("Update Existing Mode: All $hasLMSCategory records already have lms_category values.`r`n`r`nNo processing needed - stopping now!")
                    return
                }
                
                if ($needingAPICall -eq 0) {
                    $txtLog.AppendText("*** ALL MISSING CATEGORIES CAN BE FILLED FROM LOCAL DB - NO API CALLS NEEDED! ***`r`n")
                }
            } else {
                $txtLog.AppendText("UPDATE EXISTING MODE: lms_category column will be added`r`n")
            }
        }
        
        $txtLog.AppendText("- Requiring API calls: $needingAPICall`r`n")
        $txtLog.AppendText("- Remaining to process: $remainingToProcess`r`n")
        $txtLog.AppendText("- Invalid/malformed: $invalidCount`r`n")
        $txtLog.AppendText("- Empty fields: $emptyCount`r`n")
        
        if (-not $chkLocalDBOnly.Checked) {
            $txtLog.AppendText("Estimated API processing time per cycle: $estimatedHours hours ($totalBatches batches)`r`n")
        }
        
        if ($needingAPICall -eq 0 -and -not $chkLocalDBOnly.Checked) {
            $txtLog.AppendText("*** ALL ISBNs ARE ALREADY IN LOCAL DATABASE - NO API CALLS NEEDED! ***`r`n")
        }
        
        $txtLog.AppendText("Local DB contains $($global:LocalISBNDB.Count) total records`r`n")
        $efficiency = if ($validCount -gt 0) { [Math]::Round(($localDBCount / $validCount) * 100, 1) } else { 0 }
        $txtLog.AppendText("Local DB coverage: $efficiency% of valid ISBNs`r`n")
    }
}

# FIXED: Function to start processing with enhanced CSV updating
function Start-ProcessingJob {
    # Check if Local DB Only mode is selected
    if ($chkLocalDBOnly.Checked) {
        Start-LocalDBOnlyProcessing
        return
    }
    
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

    # Check if Update Existing mode and all categories are filled
    if ($chkUpdateExisting.Checked) {
        $hasLMSCategoryColumn = $global:Books[0].PSObject.Properties.Name -contains 'lms_category'
        if ($hasLMSCategoryColumn) {
            $actualMissingCount = 0
            
            foreach ($book in $global:Books) {
                $currentCategory = $book.lms_category
                if (-not $currentCategory -or $currentCategory.Trim() -eq '' -or $currentCategory.Trim() -eq 'NULL' -or $currentCategory.Trim() -eq '0') {
                    $actualMissingCount++
                }
            }
            
            if ($actualMissingCount -eq 0) {
                $txtLog.AppendText("*** ALL RECORDS HAVE CATEGORIES - NO PROCESSING NEEDED! ***`r`n")
                $lblStatus.Text = "UPDATE MODE: All categories filled - no work needed"
                [System.Windows.Forms.MessageBox]::Show("Update Existing Mode: All $($global:Books.Count) records already have lms_category values.`r`n`r`nNo processing needed!")
                return
            }
        }
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
    $lblLocalHits.Text = "Local DB Hits: $global:LocalDBHits"
    
    if ($global:CycleCount -eq 1) {
        $txtLog.Clear()
        $txtLog.AppendText("Starting processing - checking local DB first, then API calls...`r`n")
    } else {
        $txtLog.AppendText("`r`n=== CYCLE $global:CycleCount STARTED ===`r`n")
    }
    
    $txtLog.AppendText("Batch size: $($global:ProcessingSettings.BatchSize), Interval: $($global:ProcessingSettings.IntervalMinutes) minutes`r`n")
    $txtLog.AppendText("Starting from record: $($global:ProcessingSettings.StartFrom + 1)`r`n")
    $txtLog.AppendText("Valid ISBNs to process: $($global:ValidISBNs.Count)`r`n")
    $txtLog.AppendText("Already processed ISBNs: $($global:ProcessedISBNs.Count)`r`n")
    $txtLog.AppendText("Local database contains: $($global:LocalISBNDB.Count) ISBNs`r`n")
    
    # CRITICAL: Add empty columns to ALL records BEFORE processing begins
    $txtLog.AppendText("Adding selected attribute columns to all records...`r`n")
    foreach ($book in $global:Books) {
        foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
            if (-not ($book.PSObject.Properties.Name -contains $attr)) {
                $book | Add-Member -NotePropertyName $attr -NotePropertyValue '' -Force
            }
        }
    }
    $txtLog.AppendText("All records now have columns: $($global:ProcessingSettings.SelectedAttributes -join ', ')`r`n")
    
    $apiCallCount = 0
    $validISBNsProcessed = 0
    $newEnrichmentsThisCycle = 0
    $localHitsThisCycle = 0
    $newDBEntries = 0
    $dbUpdatesPending = $false
    
    # Process in batches (similar to working script)
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
                # In update existing mode, only process if lms_category is missing
                $shouldProcess = $true
                if ($chkUpdateExisting.Checked) {
                    $currentCategory = $null
                    if ($book.PSObject.Properties.Name -contains 'lms_category') {
                        $currentCategory = $book.lms_category
                    }
                    $shouldProcess = (-not $currentCategory -or $currentCategory.Trim() -eq '')
                }
                
                if ($shouldProcess) {
                    # FIRST: Check local database
                    $localEntry = Get-ISBNFromLocalDB -isbn $validISBN -dbHashTable $global:LocalISBNDB
                    if ($localEntry) {
                        # Use local database entry (instant) and apply to CSV
                        Apply-DataToCSVRow -book $book -dataEntry $localEntry -selectedAttributes $global:ProcessingSettings.SelectedAttributes -selectedFormat $global:ProcessingSettings.SelectedFormat
                        
                        $newEnrichmentsThisCycle++
                        $global:ProcessedISBNs[$validISBN] = $true
                        $localHitsThisCycle++
                        $global:LocalDBHits++
                        $validISBNsProcessed++
                        
                        # Verify local data was applied
                        $verifyValues = @()
                        foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
                            $verifyValues += "$attr='$($book.$attr)'"
                        }
                        $txtLog.AppendText("  Local DB hit for ISBN ${validISBN}: $($verifyValues -join ', ')`r`n")
                    } else {
                        # SECOND: Make API call if not in local database
                        $url = "https://openlibrary.org/api/books?bibkeys=ISBN:$validISBN&jscmd=data&format=json"
                        try {
                            $resp = Invoke-RestMethod -Uri $url -ErrorAction Stop
                            $apiEntry = $resp."ISBN:$validISBN"
                            $apiCallCount++
                            $global:TotalAPICalls++
                            $batchAPICallCount++
                            $validISBNsProcessed++
                            
                            if ($apiEntry) {
                                # Add to local database
                                $entry = Add-ISBNToLocalDB -isbn $validISBN -apiData $apiEntry -dbHashTable $global:LocalISBNDB -selectedFormat $global:ProcessingSettings.SelectedFormat
                                $newDBEntries++
                                $dbUpdatesPending = $true
                                
                                # CRITICAL: Apply enriched data to the CSV row
                                Apply-DataToCSVRow -book $book -dataEntry $entry -selectedAttributes $global:ProcessingSettings.SelectedAttributes -selectedFormat $global:ProcessingSettings.SelectedFormat
                                
                                # DEBUGGING: Verify the data was applied
                                $appliedValues = @()
                                foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
                                    $appliedValues += "$attr='$($book.$attr)'"
                                }
                                $txtLog.AppendText("  API enriched ISBN ${validISBN}: $($appliedValues -join ', ')`r`n")
                                
                                $newEnrichmentsThisCycle++
                                $global:ProcessedISBNs[$validISBN] = $true
                            } else {
                                # Mark as processed even if no data found
                                $global:ProcessedISBNs[$validISBN] = $true
                                $txtLog.AppendText("  API returned no data for ISBN ${validISBN}`r`n")
                            }
                        } catch {
                            $txtLog.AppendText("  API error for ISBN ${validISBN}: $_`r`n")
                        }
                        
                        # API rate limiting delay
                        Start-Sleep -Milliseconds 100
                    }
                }
            }
            
            $progressBar.Value = $j + 1
            $lblProgress.Text = "Progress: $($j + 1) / $totalRecords"
            $lblTotalProcessed.Text = "Total API Calls: $global:TotalAPICalls"
            $lblLocalHits.Text = "Local DB Hits: $global:LocalDBHits"
            
            # Calculate ETA
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
        
        # CRITICAL FIX: Save progress after each batch - this ensures CSV gets the enriched data
        try {
            $global:Books | Export-Csv -Path $global:ProcessingSettings.OutputPath -NoTypeInformation -Encoding UTF8
            $txtLog.AppendText("Batch $batchNumber completed. API calls: $batchAPICallCount, Local hits: $localHitsThisCycle. ENRICHED DATA SAVED to output file.`r`n")
            
            # Verify the save worked by checking the first enriched record
            if ($newEnrichmentsThisCycle -gt 0) {
                $testImport = Import-Csv -Path $global:ProcessingSettings.OutputPath -ErrorAction SilentlyContinue
                if ($testImport -and $testImport.Count -gt 0) {
                    $hasAllColumns = $true
                    foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
                        if (-not ($testImport[0].PSObject.Properties.Name -contains $attr)) {
                            $hasAllColumns = $false
                            break
                        }
                    }
                    if ($hasAllColumns) {
                        $txtLog.AppendText("VERIFIED: CSV file contains all selected columns.`r`n")
                    } else {
                        $txtLog.AppendText("WARNING: CSV file missing some selected columns!`r`n")
                    }
                }
            }
        } catch {
            $txtLog.AppendText("Failed to save progress: $_`r`n")
        }
        
        # Save local database updates if there are new entries
        if ($dbUpdatesPending -and $newDBEntries -gt 0) {
            if (Save-LocalISBNDatabase -dbHashTable $global:LocalISBNDB -scriptPath $global:ScriptPath) {
                $txtLog.AppendText("Updated local ISBN database with $newDBEntries new entries.`r`n")
                $lblDBStats.Text = "Local DB: $($global:LocalISBNDB.Count) records"
                $newDBEntries = 0
                $dbUpdatesPending = $false
            }
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
                while ($global:ProcessingPaused -and $global:ProcessingActive) {
                    $lblStatus.Text = "Processing paused during wait - click Start to resume"
                    Start-Sleep -Seconds 1
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                if (-not $global:ProcessingActive) { break }
                
                $remainingSeconds = $waitSeconds - $wait
                $remainingMinutes = [Math]::Floor($remainingSeconds / 60)
                $remainingSecondsDisplay = $remainingSeconds % 60
                
                $lblStatus.Text = "Waiting... ${remainingMinutes}:${remainingSecondsDisplay:D2} remaining before next batch"
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Seconds 1
            }
            
            if ($global:ProcessingActive) {
                $lblStatus.Text = "Wait complete. Starting next batch..."
                $txtLog.AppendText("Wait complete. Starting next batch...`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 500
            }
        }
    }
    
    # Final save of both CSV and local database
    try {
        $global:Books | Export-Csv -Path $global:ProcessingSettings.OutputPath -NoTypeInformation -Encoding UTF8
        $txtLog.AppendText("FINAL SAVE: All enriched data saved to CSV file.`r`n")
        
        # Final verification
        $finalImport = Import-Csv -Path $global:ProcessingSettings.OutputPath -ErrorAction SilentlyContinue
        if ($finalImport -and $finalImport.Count -gt 0) {
            $enrichedCount = 0
            foreach ($record in $finalImport) {
                foreach ($attr in $global:ProcessingSettings.SelectedAttributes) {
                    if ($record.$attr -and $record.$attr.Trim() -ne '') {
                        $enrichedCount++
                        break
                    }
                }
            }
            $txtLog.AppendText("FINAL VERIFICATION: $enrichedCount records have enriched data in the CSV file.`r`n")
        }
    } catch {
        $txtLog.AppendText("ERROR on final CSV save: $_`r`n")
    }
    
    # Final save of local database if there are pending updates
    if ($dbUpdatesPending -and $newDBEntries -gt 0) {
        if (Save-LocalISBNDatabase -dbHashTable $global:LocalISBNDB -scriptPath $global:ScriptPath) {
            $txtLog.AppendText("Final update to local ISBN database with $newDBEntries new entries.`r`n")
            $lblDBStats.Text = "Local DB: $($global:LocalISBNDB.Count) records"
        }
    }
    
    # Processing completed or stopped
    if ($global:ProcessingActive) {
        $txtLog.AppendText("`r`nCycle $global:CycleCount completed!`r`n")
        $txtLog.AppendText("EFFICIENCY SUMMARY:`r`n")
        $txtLog.AppendText("- Instant local DB hits: $localHitsThisCycle`r`n")
        $txtLog.AppendText("- API calls needed: $apiCallCount`r`n")
        $txtLog.AppendText("- New database entries: $newDBEntries`r`n")
        $txtLog.AppendText("- Total enrichments: $newEnrichmentsThisCycle`r`n")
        $txtLog.AppendText("- Total ISBNs processed so far: $($global:ProcessedISBNs.Count) / $($global:ValidISBNs.Count)`r`n")
        
        # Check if all ISBNs have been processed
        $allRecordsProcessed = ($progressBar.Value -ge $totalRecords)
        $allISBNsProcessed = ($global:ProcessedISBNs.Count -ge $global:ValidISBNs.Count)
        
        if ($allISBNsProcessed -or $allRecordsProcessed) {
            $lblStatus.Text = "PROCESSING COMPLETE! All records have been processed."
            $txtLog.AppendText("`r`n=== PROCESSING COMPLETE ===`r`n")
            $txtLog.AppendText("All records in the file have been processed.`r`n")
            $txtLog.AppendText("FINAL EFFICIENCY STATS:`r`n")
            $txtLog.AppendText("- Total API calls made: $global:TotalAPICalls`r`n")
            $txtLog.AppendText("- Total local DB hits: $global:LocalDBHits`r`n")
            $txtLog.AppendText("- Total processing cycles: $global:CycleCount`r`n")
            $txtLog.AppendText("- ISBNs successfully enriched: $($global:ProcessedISBNs.Count)`r`n")
            $txtLog.AppendText("- Local database now contains: $($global:LocalISBNDB.Count) ISBNs`r`n")
            
            $efficiency = if ($global:TotalAPICalls + $global:LocalDBHits -gt 0) {
                [Math]::Round(($global:LocalDBHits / ($global:TotalAPICalls + $global:LocalDBHits)) * 100, 1)
            } else { 0 }
            
            $txtLog.AppendText("- Local DB efficiency: $efficiency%`r`n")
            
            [System.Windows.Forms.MessageBox]::Show("Processing Complete!`r`n`r`nAll records have been processed and enriched data has been saved to the CSV file.`r`nLocal DB hits: $global:LocalDBHits (instant)`r`nAPI calls: $global:TotalAPICalls (rate limited)`r`nEfficiency: $efficiency%`r`nTotal cycles: $global:CycleCount`r`nFile saved to: $($global:ProcessingSettings.OutputPath)")
            
            # Reset for potential new run
            $global:ProcessedISBNs = @{}
            $global:CycleCount = 1
            $global:TotalAPICalls = 0
            $global:LocalDBHits = 0
        } elseif ($newEnrichmentsThisCycle -eq 0) {
            # If no new enrichments were found this cycle, stop to avoid cycling
            $lblStatus.Text = "NO NEW ENRICHMENTS FOUND - Processing stopped."
            $txtLog.AppendText("`r`n=== NO NEW ENRICHMENTS FOUND ===`r`n")
            $txtLog.AppendText("No new data was found in this cycle.`r`n")
            $txtLog.AppendText("You may want to try again later as the OpenLibrary database is updated frequently.`r`n")
            
            [System.Windows.Forms.MessageBox]::Show("No new enrichments found in this cycle.`r`n`r`nProcessing complete. Enriched data has been saved to the CSV file.`r`nTotal API calls: $global:TotalAPICalls`r`nLocal DB hits: $global:LocalDBHits`r`nISBNs enriched: $($global:ProcessedISBNs.Count)")
        } else {
            # Continue cycling only if we found new enrichments
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
        $txtLog.AppendText("`r`nProcessing stopped. You can resume from where you left off.`r`n")
    }
    
    # Reset UI state
    $global:ProcessingActive = $false
    $global:ProcessingPaused = $false
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnPause.Enabled = $false
}

# Event handlers for checkboxes
$chkUpdateExisting.Add_CheckedChanged({
    # Auto-select lms_category when in update mode
    if ($chkUpdateExisting.Checked) {
        for ($i = 0; $i -lt $clbAttrs.Items.Count; $i++) {
            if ($clbAttrs.Items[$i] -eq 'lms_category') {
                $clbAttrs.SetItemChecked($i, $true)
            } else {
                $clbAttrs.SetItemChecked($i, $false)
            }
        }
        $txtLog.AppendText("Update Existing mode enabled - will only update missing lms_category fields.`r`n")
        
        # Immediately check if there's any work to do
        if ($global:Books -and $comboField.SelectedItem) {
            $hasLMSCategoryColumn = $global:Books[0].PSObject.Properties.Name -contains 'lms_category'
            if ($hasLMSCategoryColumn) {
                $actualMissingCount = 0
                foreach ($book in $global:Books) {
                    $currentCategory = $book.lms_category
                    if (-not $currentCategory -or $currentCategory.Trim() -eq '' -or $currentCategory.Trim() -eq 'NULL' -or $currentCategory.Trim() -eq '0') {
                        $actualMissingCount++
                    }
                }
                
                if ($actualMissingCount -eq 0) {
                    $txtLog.AppendText("*** NOTICE: ALL $($global:Books.Count) RECORDS ALREADY HAVE CATEGORIES - NO WORK NEEDED! ***`r`n")
                    $txtValidation.Text = "UPDATE MODE: All $($global:Books.Count) records have categories - NO PROCESSING NEEDED"
                } else {
                    $txtLog.AppendText("Found $actualMissingCount records that need category updates.`r`n")
                }
            }
        }
    }
    
    # Re-validate when checkbox changes
    if ($comboField.SelectedItem -and $global:Books) {
        Update-ISBNValidation -fieldName $comboField.SelectedItem
    }
})

# Event handler for Local DB Only checkbox
$chkLocalDBOnly.Add_CheckedChanged({
    if ($chkLocalDBOnly.Checked) {
        $txtLog.AppendText("Local DB Only mode enabled - will use local database only, no API calls will be made.`r`n")
        
        # Disable batch controls since they're not needed for local-only processing
        $numBatchSize.Enabled = $false
        $numInterval.Enabled = $false
        $numResume.Enabled = $false
        $btnPause.Enabled = $false
        
        # Update button text
        $btnStart.Text = 'Start Local Processing'
        
        # Show expected efficiency
        if ($global:Books -and $comboField.SelectedItem -and $global:ValidISBNs) {
            $localAvailable = 0
            foreach ($isbn in $global:ValidISBNs) {
                if ($global:LocalISBNDB.ContainsKey($isbn)) {
                    $localAvailable++
                }
            }
            $txtLog.AppendText("Local DB Only: $localAvailable out of $($global:ValidISBNs.Count) valid ISBNs can be enriched instantly.`r`n")
        }
    } else {
        $txtLog.AppendText("Local DB Only mode disabled - will use local database first, then API calls for missing data.`r`n")
        
        # Re-enable batch controls
        $numBatchSize.Enabled = $true
        $numInterval.Enabled = $true
        $numResume.Enabled = $true
        $btnPause.Enabled = $true
        
        # Restore button text
        $btnStart.Text = 'Start Processing'
    }
    
    # Re-validate when checkbox changes
    if ($comboField.SelectedItem -and $global:Books) {
        Update-ISBNValidation -fieldName $comboField.SelectedItem
    }
})

# Browse Input event
$btnBrowseInput.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV Files (*.csv)|*.csv'
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtInput.Text = $ofd.FileName
        
        # Initialize local ISBN database from script directory
        $global:LocalISBNDB = Initialize-LocalISBNDatabase -scriptPath $global:ScriptPath
        $lblDBStats.Text = "Local DB: $($global:LocalISBNDB.Count) records"
        
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
                $txtLog.AppendText("Local ISBN database initialized with $($global:LocalISBNDB.Count) existing entries.`r`n")
                $txtLog.AppendText("Database location: $(Join-Path $global:ScriptPath 'ISBN-Enhanced-DB.csv')`r`n")
                $txtLog.AppendText("Select lookup field to validate ISBNs and see time estimate.`r`n")
                
                # Reset processing state when loading new file
                $global:ProcessedISBNs = @{}
                $global:CycleCount = 1
                $global:TotalAPICalls = 0
                $global:LocalDBHits = 0
                $lblCycleCount.Text = "Processing Cycle: 1"
                $lblTotalProcessed.Text = "Total API Calls: 0"
                $lblLocalHits.Text = "Local DB Hits: 0"
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

# Initialize the script path and local database on startup
$global:ScriptPath = if ($MyInvocation.MyCommand.Path) { 
    Split-Path -Parent $MyInvocation.MyCommand.Path 
} else { 
    Get-Location 
}

# Show the form
[void]$form.ShowDialog()
