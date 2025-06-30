Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Assurance Tool"
$form.Size = New-Object System.Drawing.Size(900, 730)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Add auto-detect controls at the top
$autoDetectPanel = New-Object System.Windows.Forms.Panel
$autoDetectPanel.Location = New-Object System.Drawing.Point(10, 10)
$autoDetectPanel.Size = New-Object System.Drawing.Size(860, 120)
$autoDetectPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($autoDetectPanel)

# Folder selection section
$lblFolders = New-Object System.Windows.Forms.Label
$lblFolders.Text = "Auto-Detect Files:"
$lblFolders.Location = New-Object System.Drawing.Point(10, 10)
$lblFolders.Size = New-Object System.Drawing.Size(120, 20)
$lblFolders.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$autoDetectPanel.Controls.Add($lblFolders)

# Source folder
$lblSourceFolder = New-Object System.Windows.Forms.Label
$lblSourceFolder.Text = "Source Folder:"
$lblSourceFolder.Location = New-Object System.Drawing.Point(10, 35)
$lblSourceFolder.Size = New-Object System.Drawing.Size(100, 20)
$autoDetectPanel.Controls.Add($lblSourceFolder)

$txtSourceFolder = New-Object System.Windows.Forms.TextBox
$txtSourceFolder.Location = New-Object System.Drawing.Point(115, 33)
$txtSourceFolder.Size = New-Object System.Drawing.Size(380, 20)
$autoDetectPanel.Controls.Add($txtSourceFolder)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(505, 32)
$btnBrowseSource.Size = New-Object System.Drawing.Size(80, 23)
$autoDetectPanel.Controls.Add($btnBrowseSource)

# Processed folder
$lblProcessedFolder = New-Object System.Windows.Forms.Label
$lblProcessedFolder.Text = "Processed Folder:"
$lblProcessedFolder.Location = New-Object System.Drawing.Point(10, 65)
$lblProcessedFolder.Size = New-Object System.Drawing.Size(100, 20)
$autoDetectPanel.Controls.Add($lblProcessedFolder)

$txtProcessedFolder = New-Object System.Windows.Forms.TextBox
$txtProcessedFolder.Location = New-Object System.Drawing.Point(115, 63)
$txtProcessedFolder.Size = New-Object System.Drawing.Size(380, 20)
$autoDetectPanel.Controls.Add($txtProcessedFolder)

$btnBrowseProcessed = New-Object System.Windows.Forms.Button
$btnBrowseProcessed.Text = "Browse..."
$btnBrowseProcessed.Location = New-Object System.Drawing.Point(505, 62)
$btnBrowseProcessed.Size = New-Object System.Drawing.Size(80, 23)
$autoDetectPanel.Controls.Add($btnBrowseProcessed)

# Auto-detect button
$btnAutoDetect = New-Object System.Windows.Forms.Button
$btnAutoDetect.Text = "Auto-Detect Files"
$btnAutoDetect.Location = New-Object System.Drawing.Point(600, 32)
$btnAutoDetect.Size = New-Object System.Drawing.Size(120, 25)
$btnAutoDetect.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$autoDetectPanel.Controls.Add($btnAutoDetect)

# Global controls section
$lblGlobal = New-Object System.Windows.Forms.Label
$lblGlobal.Text = "Global Actions:"
$lblGlobal.Location = New-Object System.Drawing.Point(10, 95)
$lblGlobal.Size = New-Object System.Drawing.Size(100, 20)
$lblGlobal.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$autoDetectPanel.Controls.Add($lblGlobal)

# Run all button
$btnRunAll = New-Object System.Windows.Forms.Button
$btnRunAll.Text = "Run All"
$btnRunAll.Location = New-Object System.Drawing.Point(600, 65)
$btnRunAll.Size = New-Object System.Drawing.Size(60, 25)
$btnRunAll.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$autoDetectPanel.Controls.Add($btnRunAll)

# Add a button to generate file analysis report
$btnAnalyzeFiles = New-Object System.Windows.Forms.Button
$btnAnalyzeFiles.Text = "Analyze Files"
$btnAnalyzeFiles.Location = New-Object System.Drawing.Point(760, 65)
$btnAnalyzeFiles.Size = New-Object System.Drawing.Size(80, 25)
$autoDetectPanel.Controls.Add($btnAnalyzeFiles)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(120, 95)
$progressBar.Size = New-Object System.Drawing.Size(200, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 3
$progressBar.Value = 0
$autoDetectPanel.Controls.Add($progressBar)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready - Select folders and click Auto-Detect"
$lblStatus.Location = New-Object System.Drawing.Point(330, 97)
$lblStatus.Size = New-Object System.Drawing.Size(300, 20)
$autoDetectPanel.Controls.Add($lblStatus)

# Create TabControl for different data types (moved below auto-detect panel)
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 140)
$tabControl.Size = New-Object System.Drawing.Size(860, 540)
$form.Controls.Add($tabControl)

# Function to browse for folders
function Browse-Folder {
    param([string]$Description = "Select Folder")
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowNewFolderButton = $false
    
    # Set the initial directory to where the script is being run from
    $scriptDirectory = Get-Location
    $folderBrowser.SelectedPath = $scriptDirectory.Path
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Function to find latest file by keyword in folder
function Find-LatestFileByKeyword {
    param([string]$FolderPath, [string]$Keyword)
    
    if ([string]::IsNullOrEmpty($FolderPath) -or -not (Test-Path $FolderPath)) {
        return $null
    }
    
    try {
        $files = Get-ChildItem -Path $FolderPath -Filter "*.csv" | 
                 Where-Object { $_.Name -match $Keyword } |
                 Sort-Object LastWriteTime -Descending
        
        if ($files.Count -gt 0) {
            return $files[0].FullName
        }
        
        return $null
    }
    catch {
        return $null
    }
}

# Function to auto-detect and load files for all data types
function Auto-DetectFiles {
    param([string]$SourceFolder, [string]$ProcessedFolder)
    
    $results = @{}
    
    # Simple direct approach - manually check for your specific files
    $results["Items"] = @{
        SourceFile = $null
        ProcessedFile = $null
        SourceFound = $false
        ProcessedFound = $false
    }
    
    $results["Users"] = @{
        SourceFile = $null
        ProcessedFile = $null
        SourceFound = $false
        ProcessedFound = $false
    }
    
    $results["Loans"] = @{
        SourceFile = $null
        ProcessedFile = $null
        SourceFound = $false
        ProcessedFound = $false
    }
    
    # Direct file search - bypass the keyword matching entirely
    try {
        # Find source files
        $allSourceFiles = Get-ChildItem -Path $SourceFolder -Filter "*.csv" -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $allSourceFiles) {
            if ($file.Name -match "item|catalogue|catalog") {
                $results["Items"].SourceFile = $file.FullName
                $results["Items"].SourceFound = $true
            }
            if ($file.Name -match "user|patron|member|borrower") {
                $results["Users"].SourceFile = $file.FullName
                $results["Users"].SourceFound = $true
            }
            if ($file.Name -match "loan|circulation|checkout|issue") {
                $results["Loans"].SourceFile = $file.FullName
                $results["Loans"].SourceFound = $true
            }
        }
        
        # Find processed files (CSV and TSV) - more specific matching
        $allProcessedFiles = Get-ChildItem -Path $ProcessedFolder -Include "*.csv", "*.tsv" -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $allProcessedFiles) {
            # More specific matching to avoid picking up wrong files
            if ($file.Name -match "^items_authx" -or $file.Name -match "^item\.") {
                $results["Items"].ProcessedFile = $file.FullName
                $results["Items"].ProcessedFound = $true
            }
            elseif ($file.Name -match "^borrowers_authx" -or $file.Name -match "^users_authx") {
                $results["Users"].ProcessedFile = $file.FullName
                $results["Users"].ProcessedFound = $true
            }
            elseif ($file.Name -match "^loans_authx" -or $file.Name -match "^loan\.") {
                $results["Loans"].ProcessedFile = $file.FullName
                $results["Loans"].ProcessedFound = $true
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error in auto-detection: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    
    return $results
}

# Function to load delimited file headers and populate key field dropdown
function Load-CsvHeaders {
    param([string]$FilePath, [System.Windows.Forms.ComboBox]$ComboBox)
    
    if ([string]::IsNullOrEmpty($FilePath) -or -not (Test-Path $FilePath)) {
        return
    }
    
    try {
        $delimiter = Get-FileDelimiter -FilePath $FilePath
        $firstLine = Get-Content $FilePath -First 1
        
        # Check if this is a TSV key=value format (headers on every row)
        $isKeyValueFormat = $delimiter -eq "`t" -and $firstLine -match "="
        
        if ($isKeyValueFormat) {
            # Get headers from key=value format
            $headers = Get-TsvKeyValueHeaders -FilePath $FilePath
        } else {
            # Standard delimited format
            $headers = Parse-DelimitedLine -Line $firstLine -Delimiter $delimiter
        }
        
        $processedHeaders = @()
        $headerCounts = @{}
        
        foreach ($header in $headers) {
            $cleanHeader = $header.Trim('"').Trim()
            
            if ($headerCounts.ContainsKey($cleanHeader)) {
                $headerCounts[$cleanHeader]++
                $processedHeaders += "$cleanHeader`_$($headerCounts[$cleanHeader])"
            } else {
                $headerCounts[$cleanHeader] = 1
                $processedHeaders += $cleanHeader
            }
        }
        
        $ComboBox.Items.Clear()
        foreach ($header in $processedHeaders) {
            $ComboBox.Items.Add($header)
        }
        
        # Auto-select barcode field if it exists (case insensitive)
        $barcodeIndex = -1
        for ($i = 0; $i -lt $processedHeaders.Count; $i++) {
            if ($processedHeaders[$i] -match "^barcode$|^Barcode$|^BARCODE$") {
                $barcodeIndex = $i
                break
            }
        }
        
        if ($barcodeIndex -ge 0) {
            $ComboBox.SelectedIndex = $barcodeIndex
        } elseif ($ComboBox.Items.Count -gt 0) {
            $ComboBox.SelectedIndex = 0
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading file headers: $($_.Exception.Message)`n`nFile: $FilePath", "Error", "OK", "Error")
    }
}

# Function to update UI with detected files
function Update-FilesFromDetection {
    param($DetectionResults)
    
    $detectionReport = @()
    $detectionReport += "AUTO-DETECTION RESULTS:"
    $detectionReport += "=" * 40
    
    foreach ($dataType in @("Items", "Users", "Loans")) {
        $result = $DetectionResults[$dataType]
        
        $txtOriginal = $form.Controls.Find("txtOriginal$dataType", $true)[0]
        $txtMigrated = $form.Controls.Find("txtMigrated$dataType", $true)[0]
        $cmbOriginalKeyField = $form.Controls.Find("cmbOriginalKeyField$dataType", $true)[0]
        $cmbMigratedKeyField = $form.Controls.Find("cmbMigratedKeyField$dataType", $true)[0]
        
        if ($result.SourceFile) {
            $txtOriginal.Text = $result.SourceFile
            $fileName = [System.IO.Path]::GetFileName($result.SourceFile)
            $detectionReport += "$dataType Source: $fileName (Found)"
            
            try {
                Load-CsvHeaders -FilePath $result.SourceFile -ComboBox $cmbOriginalKeyField
                $detectionReport += "  - Headers loaded successfully"
            }
            catch {
                $detectionReport += "  - ERROR loading headers: $($_.Exception.Message)"
            }
        } else {
            $txtOriginal.Text = ""
            $detectionReport += "$dataType Source: Not found"
        }
        
        if ($result.ProcessedFile) {
            $txtMigrated.Text = $result.ProcessedFile
            $fileName = [System.IO.Path]::GetFileName($result.ProcessedFile)
            $extension = [System.IO.Path]::GetExtension($result.ProcessedFile)
            $detectionReport += "$dataType Processed: $fileName ($extension format) (Found)"
            
            try {
                Load-CsvHeaders -FilePath $result.ProcessedFile -ComboBox $cmbMigratedKeyField
                $headerCount = $cmbMigratedKeyField.Items.Count
                $detectionReport += "  - Headers loaded successfully ($headerCount headers found)"
                if ($headerCount -gt 0) {
                    $firstFewHeaders = @()
                    for ($i = 0; $i -lt [Math]::Min(5, $headerCount); $i++) {
                        $firstFewHeaders += $cmbMigratedKeyField.Items[$i]
                    }
                    $detectionReport += "  - Sample headers: $($firstFewHeaders -join ', ')"
                }
            }
            catch {
                $detectionReport += "  - ERROR loading headers: $($_.Exception.Message)"
            }
        } else {
            $txtMigrated.Text = ""
            $detectionReport += "$dataType Processed: Not found"
        }
        
        $detectionReport += ""
    }
    
    # Add file finding test for processed folder - inline logic
    $detectionReport += "FILE FINDING TEST - PROCESSED FOLDER:"
    $processedFolder = $txtProcessedFolder.Text
    if (Test-Path $processedFolder) {
        $testKeywords = @("item", "borrower", "borrowers", "loan")
        
        foreach ($keyword in $testKeywords) {
            $found = Find-LatestFileByKeyword -FolderPath $processedFolder -Keyword $keyword
            if ($found) {
                $fileName = [System.IO.Path]::GetFileName($found)
                $detectionReport += "  Keyword '$keyword': Found $fileName"
            } else {
                $detectionReport += "  Keyword '$keyword': Not found"
            }
        }
    } else {
        $detectionReport += "  Processed folder path is invalid"
    }
    $detectionReport += ""
    
    # Add all files listing
    $detectionReport += "ALL FILES IN PROCESSED FOLDER:"
    if (Test-Path $processedFolder) {
        $allProcessedFiles = Get-ChildItem -Path $processedFolder -Include "*.csv", "*.tsv" -Recurse -ErrorAction SilentlyContinue
        if ($allProcessedFiles.Count -gt 0) {
            foreach ($file in $allProcessedFiles | Sort-Object Name) {
                $detectionReport += "  - $($file.Name) ($($file.Extension))"
            }
        } else {
            $detectionReport += "  - No CSV/TSV files found"
        }
    }
    
    $detectionText = $detectionReport -join "`r`n"
    [System.Windows.Forms.MessageBox]::Show($detectionText, "File Detection Results", "OK", "Information")
}

# Function to safely import delimited files (CSV/TSV) with duplicate headers
function Import-CsvSafe {
    param([string]$FilePath)
    
    try {
        $lines = Get-Content $FilePath
        if ($lines.Count -eq 0) {
            return @()
        }
        
        $delimiter = Get-FileDelimiter -FilePath $FilePath
        $firstLine = $lines[0]
        
        # Check if this is a TSV key=value format (headers on every row)
        $isKeyValueFormat = $delimiter -eq "`t" -and $firstLine -match "="
        
        if ($isKeyValueFormat) {
            # Handle TSV key=value format where headers are on every row
            $allHeaders = Get-TsvKeyValueHeaders -FilePath $FilePath
            
            # Process headers for duplicates
            $processedHeaders = @()
            $headerCounts = @{}
            
            foreach ($header in $allHeaders) {
                $cleanHeader = $header.Trim()
                if ($headerCounts.ContainsKey($cleanHeader)) {
                    $headerCounts[$cleanHeader]++
                    $processedHeaders += "$cleanHeader`_$($headerCounts[$cleanHeader])"
                } else {
                    $headerCounts[$cleanHeader] = 1
                    $processedHeaders += $cleanHeader
                }
            }
            
            # Convert TSV key=value to CSV format
            $newContent = @()
            $newContent += ($processedHeaders -join ",")
            
            foreach ($line in $lines) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    $keyValuePairs = Parse-TsvKeyValueLine -Line $line
                    $values = @()
                    
                    # Get values in the same order as headers
                    foreach ($header in $allHeaders) {
                        $value = ""
                        if ($keyValuePairs.ContainsKey($header)) {
                            $value = $keyValuePairs[$header]
                        }
                        
                        if ($value -match '[",\r\n]') {
                            $values += "`"$($value.Replace('"', '""'))`""
                        } else {
                            $values += $value
                        }
                    }
                    
                    $newContent += ($values -join ",")
                }
            }
        } else {
            # Standard delimited format processing
            $originalHeaders = Parse-DelimitedLine -Line $firstLine -Delimiter $delimiter
            
            $processedHeaders = @()
            $headerCounts = @{}
            
            for ($i = 0; $i -lt $originalHeaders.Length; $i++) {
                $header = $originalHeaders[$i].Trim('"').Trim()
                
                if ($headerCounts.ContainsKey($header)) {
                    $headerCounts[$header]++
                    $processedHeaders += "$header`_$($headerCounts[$header])"
                } else {
                    $headerCounts[$header] = 1
                    $processedHeaders += $header
                }
            }
            
            $newContent = @()
            $newContent += ($processedHeaders -join ",")
            
            # Convert data lines to CSV format for Import-Csv
            for ($i = 1; $i -lt $lines.Count; $i++) {
                if ($delimiter -eq "`t") {
                    # Convert TSV to CSV for processing
                    $fields = Parse-DelimitedLine -Line $lines[$i] -Delimiter $delimiter
                    $csvLine = ($fields | ForEach-Object { 
                        if ($_ -match '[",\r\n]' -or $_.Contains("`t")) { 
                            "`"$($_.Replace('"', '""'))`"" 
                        } else { 
                            $_ 
                        } 
                    }) -join ","
                    $newContent += $csvLine
                } else {
                    $newContent += $lines[$i]
                }
            }
        }
        
        $tempFile = [System.IO.Path]::GetTempFileName() + ".csv"
        $newContent | Out-File $tempFile -Encoding UTF8
        
        $data = Import-Csv $tempFile
        Remove-Item $tempFile -Force
        
        return $data
    }
    catch {
        throw "Error importing delimited file: $($_.Exception.Message)"
    }
}

# Function to detect file format and get appropriate delimiter
function Get-FileDelimiter {
    param([string]$FilePath)
    
    try {
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        # Try extension first
        if ($extension -eq ".tsv") {
            return "`t"
        }
        if ($extension -eq ".csv") {
            return ","
        }
        
        # Auto-detect by examining first few lines
        $sampleLines = Get-Content $FilePath -First 3
        $tabCount = 0
        $commaCount = 0
        
        foreach ($line in $sampleLines) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $tabCount += ($line -split "`t").Count - 1
                $commaCount += ($line -split ",").Count - 1
            }
        }
        
        if ($tabCount -gt $commaCount) {
            return "`t"
        } else {
            return ","
        }
    }
    catch {
        return ","  # Default to CSV on error
    }
}

# Function to parse TSV key=value format and extract headers from any line
function Parse-TsvKeyValueLine {
    param([string]$Line)
    
    if ([string]::IsNullOrEmpty($Line)) {
        return @{}
    }
    
    $keyValuePairs = @{}
    $fields = $Line -split "`t"
    
    foreach ($field in $fields) {
        if ($field -match "^(.+?)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $keyValuePairs[$key] = $value
        }
    }
    
    return $keyValuePairs
}

# Function to get headers from TSV key=value format
function Get-TsvKeyValueHeaders {
    param([string]$FilePath)
    
    try {
        # Read first few lines to find all possible keys
        $sampleLines = Get-Content $FilePath -First 10
        $allKeys = @()
        
        foreach ($line in $sampleLines) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $keyValuePairs = Parse-TsvKeyValueLine -Line $line
                foreach ($key in $keyValuePairs.Keys) {
                    if ($key -notin $allKeys) {
                        $allKeys += $key
                    }
                }
            }
        }
        
        return $allKeys | Sort-Object
    }
    catch {
        return @()
    }
}

# Function to properly parse delimited line (handles quoted fields for both CSV and TSV)
function Parse-DelimitedLine {
    param([string]$Line, [string]$Delimiter = ",", [bool]$IsKeyValue = $false)
    
    if ([string]::IsNullOrEmpty($Line)) {
        return @()
    }
    
    # Handle TSV key=value format
    if ($IsKeyValue -and $Delimiter -eq "`t") {
        $keyValuePairs = Parse-TsvKeyValueLine -Line $Line
        return $keyValuePairs.Values
    }
    
    # Standard delimited parsing
    $fields = @()
    $currentField = ""
    $inQuotes = $false
    $chars = $Line.ToCharArray()
    
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $char = $chars[$i]
        
        if ($char -eq '"') {
            if ($inQuotes -and $i + 1 -lt $chars.Length -and $chars[$i + 1] -eq '"') {
                $currentField += '"'
                $i++
            } else {
                $inQuotes = -not $inQuotes
            }
        } elseif ($char -eq $Delimiter -and -not $inQuotes) {
            $fields += $currentField
            $currentField = ""
        } else {
            $currentField += $char
        }
    }
    
    $fields += $currentField
    return $fields
}

# Function to analyze file structure issues
function Analyze-FileStructureIssues {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        return @{
            EmptyLines = 0
            MalformedLines = 0
            ValidDataLines = 0
            Issues = @()
        }
    }
    
    try {
        $delimiter = Get-FileDelimiter -FilePath $FilePath
        $allLines = Get-Content $FilePath
        $emptyLineCount = 0
        $malformedLineCount = 0
        $validDataLineCount = 0
        $issues = @()
        $expectedColumns = 0
        
        for ($i = 0; $i -lt $allLines.Count; $i++) {
            $lineNumber = $i + 1
            $line = $allLines[$i]
            
            if ([string]::IsNullOrWhiteSpace($line)) {
                $emptyLineCount++
                continue
            }
            
            if ($expectedColumns -eq 0) {
                # First non-empty line is header
                $headerFields = Parse-DelimitedLine -Line $line -Delimiter $delimiter
                $expectedColumns = $headerFields.Count
                continue
            }
            
            # Analyze data line
            $parsedColumns = Parse-DelimitedLine -Line $line -Delimiter $delimiter
            
            if ($parsedColumns.Count -ne $expectedColumns) {
                $malformedLineCount++
                if ($issues.Count -lt 5) {
                    $issues += "Line $lineNumber`: Expected $expectedColumns columns, found $($parsedColumns.Count)"
                }
            } else {
                $quoteCount = ($line.ToCharArray() | Where-Object { $_ -eq [char]34 }).Count
                if ($quoteCount % 2 -ne 0) {
                    $malformedLineCount++
                    if ($issues.Count -lt 5) {
                        $issues += "Line $lineNumber`: Unmatched quotes detected"
                    }
                } else {
                    $validDataLineCount++
                }
            }
        }
        
        return @{
            EmptyLines = $emptyLineCount
            MalformedLines = $malformedLineCount
            ValidDataLines = $validDataLineCount
            Issues = $issues
        }
    }
    catch {
        return @{
            EmptyLines = 0
            MalformedLines = 0
            ValidDataLines = 0
            Issues = @("Error analyzing file structure: $($_.Exception.Message)")
        }
    }
}
function Normalize-Key {
    param([string]$Key)
    
    if ([string]::IsNullOrEmpty($Key)) {
        return ""
    }
    
    $normalized = $Key -replace '[^a-zA-Z0-9]', ''
    return $normalized.Trim()
}

# Function to perform data comparison
function Compare-Data {
    param([string]$DataType, [bool]$IsGlobalRun = $false)
    
    $txtOriginal = $form.Controls.Find("txtOriginal$DataType", $true)[0]
    $txtMigrated = $form.Controls.Find("txtMigrated$DataType", $true)[0]
    $cmbOriginalKeyField = $form.Controls.Find("cmbOriginalKeyField$DataType", $true)[0]
    $cmbMigratedKeyField = $form.Controls.Find("cmbMigratedKeyField$DataType", $true)[0]
    $chkStripCode = $form.Controls.Find("chkStripCode$DataType", $true)[0]
    $chkNormalizeKeys = $form.Controls.Find("chkNormalizeKeys$DataType", $true)[0]
    $chkRowCountOnly = $form.Controls.Find("chkRowCountOnly$DataType", $true)[0]
    $txtResults = $form.Controls.Find("txtResults$DataType", $true)[0]
    
    $originalPath = $txtOriginal.Text
    $migratedPath = $txtMigrated.Text
    $originalKeyField = $cmbOriginalKeyField.SelectedItem
    $migratedKeyField = $cmbMigratedKeyField.SelectedItem
    $stripCode = $chkStripCode.Checked
    $normalizeKeys = $chkNormalizeKeys.Checked
    $rowCountOnly = $chkRowCountOnly.Checked
    
    if ([string]::IsNullOrEmpty($originalPath) -or [string]::IsNullOrEmpty($migratedPath)) {
        if (-not $IsGlobalRun) {
            [System.Windows.Forms.MessageBox]::Show("Please select both original and migrated files for $DataType.", "Missing Files", "OK", "Warning")
        }
        return $null
    }
    
    if (-not $rowCountOnly -and ([string]::IsNullOrEmpty($originalKeyField) -or [string]::IsNullOrEmpty($migratedKeyField))) {
        if (-not $IsGlobalRun) {
            [System.Windows.Forms.MessageBox]::Show("Please select key fields for both original and migrated files for $DataType.", "Missing Key Fields", "OK", "Warning")
        }
        return $null
    }
    
    try {
        if (-not $IsGlobalRun) {
            $txtResults.Text = "Loading and processing CSV files..."
            $form.Refresh()
        }
        
        $originalData = Import-CsvSafe -FilePath $originalPath
        $migratedData = Import-CsvSafe -FilePath $migratedPath
        
        if (-not $IsGlobalRun) {
            $txtResults.Text = "Processing data comparison..."
            $form.Refresh()
        }
        
        $report = @()
        $report += "=" * 60
        $report += "$DataType Data Assurance Report"
        $report += "=" * 60
        $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $report += ""
        $report += "Files Compared:"
        $report += "  Original: $originalPath"
        $report += "  Migrated: $migratedPath"
        
        if ($rowCountOnly) {
            $report += "  Mode: Row count comparison only"
            $report += ""
            $report += "Summary:"
            $report += "  Original total rows: $($originalData.Count)"
            $report += "  Migrated total rows: $($migratedData.Count)"
            
            $totalDiff = $migratedData.Count - $originalData.Count
            $report += "  Row difference: $totalDiff"
            $report += ""
            
            $report += "ASSESSMENT:"
            if ($totalDiff -eq 0) {
                $report += "[PASS] ROW COUNT VERIFIED - Same number of rows ($($originalData.Count))"
            } elseif ($totalDiff -lt 0) {
                $report += "[WARN] POTENTIAL DATA LOSS - Fewer rows in migrated data ($totalDiff)"
            } else {
                $report += "[INFO] MORE ROWS IN MIGRATED DATA - Additional rows: $totalDiff"
            }
        } else {
            $report += "  Original Key Field: $originalKeyField"
            $report += "  Migrated Key Field: $migratedKeyField"
            $report += "  Strip 3-char code: $stripCode"
            $report += "  Normalize keys: $normalizeKeys"
            $report += ""
            
            if ($originalData.Count -gt 0) {
                $originalHeaders = $originalData[0].PSObject.Properties.Name
                if ($originalKeyField -notin $originalHeaders) {
                    if (-not $IsGlobalRun) {
                        [System.Windows.Forms.MessageBox]::Show("Selected original key field '$originalKeyField' not found in file headers for $DataType.", "Key Field Error", "OK", "Error")
                    }
                    return $null
                }
            }
            
            if ($migratedData.Count -gt 0) {
                $migratedHeaders = $migratedData[0].PSObject.Properties.Name
                if ($migratedKeyField -notin $migratedHeaders) {
                    if (-not $IsGlobalRun) {
                        [System.Windows.Forms.MessageBox]::Show("Selected migrated key field '$migratedKeyField' not found in file headers for $DataType.", "Key Field Error", "OK", "Error")
                    }
                    return $null
                }
            }
            
            if ($stripCode) {
                foreach ($row in $migratedData) {
                    $fieldValue = $row.$migratedKeyField
                    if ($fieldValue -and $fieldValue.ToString().Length -gt 3) {
                        $row.$migratedKeyField = $fieldValue.ToString().Substring(0, $fieldValue.ToString().Length - 3)
                    }
                }
            }
            
            $originalKeys = @{}
            $migratedKeys = @{}
            $originalRawKeys = @{}
            $migratedRawKeys = @{}
            
            foreach ($row in $originalData) {
                $rawKey = $row.$originalKeyField
                if ($rawKey -and $rawKey.ToString().Trim() -ne "") {
                    $rawKeyStr = $rawKey.ToString().Trim()
                    $keyStr = if ($normalizeKeys) { Normalize-Key -Key $rawKeyStr } else { $rawKeyStr }
                    
                    if ($originalKeys.ContainsKey($keyStr)) {
                        $originalKeys[$keyStr]++
                    } else {
                        $originalKeys[$keyStr] = 1
                        $originalRawKeys[$keyStr] = $rawKeyStr
                    }
                }
            }
            
            foreach ($row in $migratedData) {
                $rawKey = $row.$migratedKeyField
                if ($rawKey -and $rawKey.ToString().Trim() -ne "") {
                    $rawKeyStr = $rawKey.ToString().Trim()
                    $keyStr = if ($normalizeKeys) { Normalize-Key -Key $rawKeyStr } else { $rawKeyStr }
                    
                    if ($migratedKeys.ContainsKey($keyStr)) {
                        $migratedKeys[$keyStr]++
                    } else {
                        $migratedKeys[$keyStr] = 1
                        $migratedRawKeys[$keyStr] = $rawKeyStr
                    }
                }
            }
            
            $originalFirstLine = (Get-Content $originalPath -First 1) -split ","
            $migratedFirstLine = (Get-Content $migratedPath -First 1) -split ","
            $originalDupeHeaders = ($originalFirstLine | Group-Object | Where-Object { $_.Count -gt 1 }).Name
            $migratedDupeHeaders = ($migratedFirstLine | Group-Object | Where-Object { $_.Count -gt 1 }).Name
            
            if ($originalDupeHeaders -or $migratedDupeHeaders) {
                $report += "HEADER WARNINGS:"
                if ($originalDupeHeaders) {
                    $report += "  Original file has duplicate headers: $($originalDupeHeaders -join ', ')"
                    $report += "  These have been automatically renamed with suffixes (_2, _3, etc.)"
                }
                if ($migratedDupeHeaders) {
                    $report += "  Migrated file has duplicate headers: $($migratedDupeHeaders -join ', ')"
                    $report += "  These have been automatically renamed with suffixes (_2, _3, etc.)"
                }
                $report += ""
            }
            
            $report += "Summary:"
            $report += "  Original total rows: $($originalData.Count)"
            $report += "  Migrated total rows: $($migratedData.Count)"
            
            # Add debug information for raw line counts
            try {
                $originalRawLines = (Get-Content $originalPath).Count
                $migratedRawLines = (Get-Content $migratedPath).Count
                
                $report += ""
                $report += "Raw File Information:"
                $report += "  Original raw lines (inc. header): $originalRawLines"
                $report += "  Migrated raw lines (inc. header): $migratedRawLines"
                $report += "  Original data rows (exc. header): $($originalData.Count)"
                $report += "  Migrated data rows (exc. header): $($migratedData.Count)"
                
                # Calculate and explain discrepancies with detailed analysis
                $originalDiscrepancy = $originalRawLines - $originalData.Count - 1
                $migratedDiscrepancy = $migratedRawLines - $migratedData.Count - 1
                
                if ($originalDiscrepancy -gt 0 -or $migratedDiscrepancy -gt 0) {
                    $report += ""
                    $report += "LINE DISCREPANCY ANALYSIS:"
                    
                    # Analyze original file structure
                    if ($originalDiscrepancy -gt 0) {
                        $originalAnalysis = Analyze-FileStructureIssues -FilePath $originalPath
                        $report += "  Original file has $originalDiscrepancy non-data lines breakdown:"
                        $report += "    - Empty lines: $($originalAnalysis.EmptyLines)"
                        $report += "    - Malformed/multi-line records: $($originalAnalysis.MalformedLines)"
                        $report += "    - Header line: 1"
                        
                        if ($originalAnalysis.Issues.Count -gt 0) {
                            $report += "  Original file issues (examples):"
                            foreach ($issue in $originalAnalysis.Issues) {
                                $report += "    - $issue"
                            }
                        }
                        
                        # Explain the math
                        $expectedProcessed = $originalAnalysis.ValidDataLines
                        if ($expectedProcessed -ne $originalData.Count) {
                            $adjustment = $originalData.Count - $expectedProcessed
                            $report += "  CSV processing adjustment: $adjustment records"
                            $report += "  (Multi-line records merged or malformed data filtered)"
                        }
                    }
                    
                    # Analyze migrated file structure
                    if ($migratedDiscrepancy -gt 0) {
                        $migratedAnalysis = Analyze-FileStructureIssues -FilePath $migratedPath
                        $report += "  Migrated file has $migratedDiscrepancy non-data lines breakdown:"
                        $report += "    - Empty lines: $($migratedAnalysis.EmptyLines)"
                        $report += "    - Malformed/multi-line records: $($migratedAnalysis.MalformedLines)"
                        $report += "    - Header line: 1"
                    }
                    
                    $report += ""
                    $report += "NOTES:"
                    $report += "  - Raw line counts include headers, empty lines, and malformed data"
                    $report += "  - Data row counts represent actual processable records"
                    $report += "  - CSV processing automatically filters invalid/empty lines"
                    $report += "  - CSV processing merges multi-line records into single rows"
                    $report += "  - The data row comparison is the accurate measure of migration success"
                    
                    if ($originalData.Count -eq $migratedData.Count) {
                        $report += "  - CONCLUSION: No actual data loss - line differences are formatting artifacts"
                    } else {
                        $report += "  - CONCLUSION: Data count differences detected - review required"
                    }
                }
                $report += ""
            }
            catch {
                # Skip if there's an error reading raw lines
            }
            
            $report += "  Original unique keys: $($originalKeys.Count)"
            $report += "  Migrated unique keys: $($migratedKeys.Count)"
            
            $totalDiff = $migratedData.Count - $originalData.Count
            $keyDiff = $migratedKeys.Count - $originalKeys.Count
            
            $report += "  Row difference: $totalDiff"
            $report += "  Unique key difference: $keyDiff"
            $report += ""
            
            $missingKeys = @()
            $extraKeys = @()
            
            foreach ($key in $originalKeys.Keys) {
                if (-not $migratedKeys.ContainsKey($key)) {
                    $missingKeys += $key
                }
            }
            
            foreach ($key in $migratedKeys.Keys) {
                if (-not $originalKeys.ContainsKey($key)) {
                    $extraKeys += $key
                }
            }
            
            if ($missingKeys.Count -gt 0) {
                $report += "MISSING KEYS IN MIGRATED DATA ($($missingKeys.Count)):"
                foreach ($key in $missingKeys | Sort-Object) {
                    $originalFormat = if ($originalRawKeys.ContainsKey($key)) { $originalRawKeys[$key] } else { $key }
                    $displayKey = if ($normalizeKeys -and $originalFormat -ne $key) { "$originalFormat -> $key" } else { $originalFormat }
                    $report += "  - $displayKey (appeared $($originalKeys[$key]) times in original)"
                }
                $report += ""
            }
            
            if ($extraKeys.Count -gt 0) {
                $report += "EXTRA KEYS IN MIGRATED DATA ($($extraKeys.Count)):"
                foreach ($key in $extraKeys | Sort-Object) {
                    $migratedFormat = if ($migratedRawKeys.ContainsKey($key)) { $migratedRawKeys[$key] } else { $key }
                    $displayKey = if ($normalizeKeys -and $migratedFormat -ne $key) { "$migratedFormat -> $key" } else { $migratedFormat }
                    $report += "  - $displayKey (appears $($migratedKeys[$key]) times in migrated)"
                }
                $report += ""
            }
            
            $originalDupes = @($originalKeys.Keys | Where-Object { $originalKeys[$_] -gt 1 })
            $migratedDupes = @($migratedKeys.Keys | Where-Object { $migratedKeys[$_] -gt 1 })
            
            if ($originalDupes.Count -gt 0) {
                $report += "DUPLICATE KEYS IN ORIGINAL DATA ($($originalDupes.Count)):"
                foreach ($key in $originalDupes | Sort-Object) {
                    $originalFormat = if ($originalRawKeys.ContainsKey($key)) { $originalRawKeys[$key] } else { $key }
                    $displayKey = if ($normalizeKeys -and $originalFormat -ne $key) { "$originalFormat -> $key" } else { $originalFormat }
                    $report += "  - $displayKey (appears $($originalKeys[$key]) times)"
                }
                $report += ""
            }
            
            if ($migratedDupes.Count -gt 0) {
                $report += "DUPLICATE KEYS IN MIGRATED DATA ($($migratedDupes.Count)):"
                foreach ($key in $migratedDupes | Sort-Object) {
                    $migratedFormat = if ($migratedRawKeys.ContainsKey($key)) { $migratedRawKeys[$key] } else { $key }
                    $displayKey = if ($normalizeKeys -and $migratedFormat -ne $key) { "$migratedFormat -> $key" } else { $migratedFormat }
                    $report += "  - $displayKey (appears $($migratedKeys[$key]) times)"
                }
                $report += ""
            }
            
            $report += "ASSESSMENT:"
            if ($missingKeys.Count -eq 0 -and $extraKeys.Count -eq 0 -and $totalDiff -eq 0) {
                $report += "[PASS] DATA INTEGRITY VERIFIED - No data loss detected"
            } elseif ($missingKeys.Count -gt 0) {
                $report += "[WARN] POTENTIAL DATA LOSS - Missing keys detected"
            } elseif ($totalDiff -lt 0) {
                $report += "[WARN] POTENTIAL DATA LOSS - Fewer rows in migrated data"
            } else {
                $report += "[INFO] DATA DIFFERENCES DETECTED - Review details above"
            }
        }
        
        $reportText = $report -join "`r`n"
        if (-not $IsGlobalRun) {
            $txtResults.Text = $reportText
        }
        
        return $reportText
        
    }
    catch {
        $errorMsg = "Error during comparison: $($_.Exception.Message)"
        if (-not $IsGlobalRun) {
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
        }
        return $errorMsg
    }
}

# Function to create tab for each data type
function Create-DataTab {
    param([string]$DataType)
    
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $DataType
    $tab.UseVisualStyleBackColor = $true
    
    $lblOriginal = New-Object System.Windows.Forms.Label
    $lblOriginal.Text = "Original $DataType File:"
    $lblOriginal.Location = New-Object System.Drawing.Point(20, 20)
    $lblOriginal.Size = New-Object System.Drawing.Size(150, 20)
    $tab.Controls.Add($lblOriginal)
    
    $txtOriginal = New-Object System.Windows.Forms.TextBox
    $txtOriginal.Location = New-Object System.Drawing.Point(20, 45)
    $txtOriginal.Size = New-Object System.Drawing.Size(500, 20)
    $txtOriginal.Name = "txtOriginal$DataType"
    $tab.Controls.Add($txtOriginal)
    
    $btnOriginal = New-Object System.Windows.Forms.Button
    $btnOriginal.Text = "Browse..."
    $btnOriginal.Location = New-Object System.Drawing.Point(530, 44)
    $btnOriginal.Size = New-Object System.Drawing.Size(80, 23)
    $btnOriginal.Name = "btnOriginal$DataType"
    $tab.Controls.Add($btnOriginal)
    
    $lblMigrated = New-Object System.Windows.Forms.Label
    $lblMigrated.Text = "Migrated $DataType File:"
    $lblMigrated.Location = New-Object System.Drawing.Point(20, 80)
    $lblMigrated.Size = New-Object System.Drawing.Size(150, 20)
    $tab.Controls.Add($lblMigrated)
    
    $txtMigrated = New-Object System.Windows.Forms.TextBox
    $txtMigrated.Location = New-Object System.Drawing.Point(20, 105)
    $txtMigrated.Size = New-Object System.Drawing.Size(500, 20)
    $txtMigrated.Name = "txtMigrated$DataType"
    $tab.Controls.Add($txtMigrated)
    
    $btnMigrated = New-Object System.Windows.Forms.Button
    $btnMigrated.Text = "Browse..."
    $btnMigrated.Location = New-Object System.Drawing.Point(530, 104)
    $btnMigrated.Size = New-Object System.Drawing.Size(80, 23)
    $btnMigrated.Name = "btnMigrated$DataType"
    $tab.Controls.Add($btnMigrated)
    
    $lblOriginalKeyField = New-Object System.Windows.Forms.Label
    $lblOriginalKeyField.Text = "Original Key Field:"
    $lblOriginalKeyField.Location = New-Object System.Drawing.Point(20, 140)
    $lblOriginalKeyField.Size = New-Object System.Drawing.Size(120, 20)
    $tab.Controls.Add($lblOriginalKeyField)
    
    $cmbOriginalKeyField = New-Object System.Windows.Forms.ComboBox
    $cmbOriginalKeyField.Location = New-Object System.Drawing.Point(20, 165)
    $cmbOriginalKeyField.Size = New-Object System.Drawing.Size(180, 20)
    $cmbOriginalKeyField.DropDownStyle = "DropDownList"
    $cmbOriginalKeyField.Name = "cmbOriginalKeyField$DataType"
    $tab.Controls.Add($cmbOriginalKeyField)
    
    $lblMigratedKeyField = New-Object System.Windows.Forms.Label
    $lblMigratedKeyField.Text = "Migrated Key Field:"
    $lblMigratedKeyField.Location = New-Object System.Drawing.Point(220, 140)
    $lblMigratedKeyField.Size = New-Object System.Drawing.Size(120, 20)
    $tab.Controls.Add($lblMigratedKeyField)
    
    $cmbMigratedKeyField = New-Object System.Windows.Forms.ComboBox
    $cmbMigratedKeyField.Location = New-Object System.Drawing.Point(220, 165)
    $cmbMigratedKeyField.Size = New-Object System.Drawing.Size(180, 20)
    $cmbMigratedKeyField.DropDownStyle = "DropDownList"
    $cmbMigratedKeyField.Name = "cmbMigratedKeyField$DataType"
    $tab.Controls.Add($cmbMigratedKeyField)
    
    $chkStripCode = New-Object System.Windows.Forms.CheckBox
    $chkStripCode.Text = "Strip 3-character code from migrated data"
    $chkStripCode.Location = New-Object System.Drawing.Point(420, 150)
    $chkStripCode.Size = New-Object System.Drawing.Size(250, 20)
    $chkStripCode.Name = "chkStripCode$DataType"
    $chkStripCode.Checked = $true
    $tab.Controls.Add($chkStripCode)
    
    $chkNormalizeKeys = New-Object System.Windows.Forms.CheckBox
    $chkNormalizeKeys.Text = "Normalize keys (remove spaces & special characters)"
    $chkNormalizeKeys.Location = New-Object System.Drawing.Point(420, 175)
    $chkNormalizeKeys.Size = New-Object System.Drawing.Size(300, 20)
    $chkNormalizeKeys.Name = "chkNormalizeKeys$DataType"
    $chkNormalizeKeys.Checked = $true
    $tab.Controls.Add($chkNormalizeKeys)
    
    $chkRowCountOnly = New-Object System.Windows.Forms.CheckBox
    $chkRowCountOnly.Text = "Row count only (ignore key fields)"
    $chkRowCountOnly.Location = New-Object System.Drawing.Point(420, 200)
    $chkRowCountOnly.Size = New-Object System.Drawing.Size(200, 20)
    $chkRowCountOnly.Name = "chkRowCountOnly$DataType"
    
    # Set default checked state based on data type
    if ($DataType -eq "Users" -or $DataType -eq "Loans") {
        $chkRowCountOnly.Checked = $true
    } else {
        $chkRowCountOnly.Checked = $false
    }
    
    $tab.Controls.Add($chkRowCountOnly)
    
    $chkRowCountOnly.Add_CheckedChanged({
        $dataType = $this.Name.Substring(15)
        $cmbOriginal = $form.Controls.Find("cmbOriginalKeyField$dataType", $true)[0]
        $cmbMigrated = $form.Controls.Find("cmbMigratedKeyField$dataType", $true)[0]
        $chkStrip = $form.Controls.Find("chkStripCode$dataType", $true)[0]
        $chkNormalize = $form.Controls.Find("chkNormalizeKeys$dataType", $true)[0]
        
        if ($this.Checked) {
            $cmbOriginal.Enabled = $false
            $cmbMigrated.Enabled = $false
            $chkStrip.Enabled = $false
            $chkNormalize.Enabled = $false
        } else {
            $cmbOriginal.Enabled = $true
            $cmbMigrated.Enabled = $true
            $chkStrip.Enabled = $true
            $chkNormalize.Enabled = $true
        }
    })
    
    # Set initial state based on checkbox
    if ($chkRowCountOnly.Checked) {
        $cmbOriginalKeyField.Enabled = $false
        $cmbMigratedKeyField.Enabled = $false
        $chkStripCode.Enabled = $false
        $chkNormalizeKeys.Enabled = $false
    }
    
    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "Compare Data"
    $btnCompare.Location = New-Object System.Drawing.Point(20, 230)
    $btnCompare.Size = New-Object System.Drawing.Size(100, 30)
    $btnCompare.Name = "btnCompare$DataType"
    $tab.Controls.Add($btnCompare)
    
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Location = New-Object System.Drawing.Point(20, 270)
    $txtResults.Size = New-Object System.Drawing.Size(800, 220)
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = "Vertical"
    $txtResults.ReadOnly = $true
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Location = New-Object System.Drawing.Point(20, 270)
    $txtResults.Size = New-Object System.Drawing.Size(800, 220)
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = "Vertical"
    $txtResults.ReadOnly = $true
    $txtResults.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtResults.Name = "txtResults$DataType"
    $tab.Controls.Add($txtResults)
    
    return $tab
}

# Create tabs for each data type
$itemsTab = Create-DataTab -DataType "Items"
$usersTab = Create-DataTab -DataType "Users"
$loansTab = Create-DataTab -DataType "Loans"

$tabControl.TabPages.Add($itemsTab)
$tabControl.TabPages.Add($usersTab)
$tabControl.TabPages.Add($loansTab)

# Global variable to store combined report
$script:combinedReport = ""

# Function to run all comparisons
function Run-AllComparisons {
    $script:combinedReport = ""
    $progressBar.Value = 0
    $btnSaveReport.Enabled = $false
    
    $dataTypes = @("Items", "Users", "Loans")
    $reports = @()
    $summaryResults = @()
    
    for ($i = 0; $i -lt $dataTypes.Count; $i++) {
        $dataType = $dataTypes[$i]
        $lblStatus.Text = "Processing $dataType..."
        $form.Refresh()
        
        $report = Compare-Data -DataType $dataType -IsGlobalRun $true
        if ($report) {
            $reports += $report
            $txtResults = $form.Controls.Find("txtResults$dataType", $true)[0]
            $txtResults.Text = $report
            
            # Extract summary information for popup
            $summaryInfo = @{
                DataType = $dataType
                Status = "PROCESSED"
                OriginalRows = 0
                MigratedRows = 0
                Assessment = "UNKNOWN"
            }
            
            # Parse the report for key information
            if ($report -match "Original total rows: (\d+)") {
                $summaryInfo.OriginalRows = [int]$matches[1]
            }
            if ($report -match "Migrated total rows: (\d+)") {
                $summaryInfo.MigratedRows = [int]$matches[1]
            }
            if ($report -match "\[PASS\]") {
                $summaryInfo.Assessment = "PASS"
            } elseif ($report -match "\[WARN\]") {
                $summaryInfo.Assessment = "WARNING"
            } elseif ($report -match "\[INFO\]") {
                $summaryInfo.Assessment = "INFO"
            }
            
            $summaryResults += $summaryInfo
        } else {
            $reports += "SKIPPED: $dataType - Missing files or configuration"
            $summaryResults += @{
                DataType = $dataType
                Status = "SKIPPED"
                OriginalRows = 0
                MigratedRows = 0
                Assessment = "SKIPPED"
            }
        }
        
        $progressBar.Value = $i + 1
        $form.Refresh()
    }
    
    # Show popup summary
    Show-ComparisonSummary -SummaryResults $summaryResults
    
    # Create combined report (existing code)
    $script:combinedReport = @()
    $script:combinedReport += "LIBRARY DATA ASSURANCE - COMBINED REPORT"
    $script:combinedReport += "=" * 80
    $script:combinedReport += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $script:combinedReport += "Tool: Library Data Assurance Tool"
    $script:combinedReport += ""
    $script:combinedReport += "EXECUTIVE SUMMARY:"
    
    $totalOriginal = 0
    $totalMigrated = 0
    $issues = @()
    
    foreach ($report in $reports) {
        if ($report -like "*SKIPPED*") {
            $issues += $report
            continue
        }
        
        if ($report -match "Original total rows: (\d+)") {
            $totalOriginal += [int]$matches[1]
        }
        if ($report -match "Migrated total rows: (\d+)") {
            $totalMigrated += [int]$matches[1]
        }
        
        if ($report -match "\[WARN\]") {
            $dataTypeName = ($report -split "`n")[1] -replace " Data Assurance Report", ""
            $issues += "WARNING in $dataTypeName"
        }
    }
    
    $script:combinedReport += "  Total Original Records: $totalOriginal"
    $script:combinedReport += "  Total Migrated Records: $totalMigrated"
    $script:combinedReport += "  Overall Difference: $($totalMigrated - $totalOriginal)"
    
    if ($issues.Count -eq 0) {
        $script:combinedReport += "  Status: [PASS] All data types verified successfully"
    } else {
        $script:combinedReport += "  Status: [REVIEW] Issues detected - see detailed reports below"
        $script:combinedReport += "  Issues Found:"
        foreach ($issue in $issues) {
            $script:combinedReport += "    - $issue"
        }
    }
    
    $script:combinedReport += ""
    $script:combinedReport += "DETAILED REPORTS:"
    $script:combinedReport += "=" * 80
    
    foreach ($report in $reports) {
        $script:combinedReport += ""
        $script:combinedReport += $report
        $script:combinedReport += ""
    }
    
    $script:combinedReport = $script:combinedReport -join "`r`n"
    
    $lblStatus.Text = "All comparisons completed"
    $btnSaveReport.Enabled = $true
}

# Save report button
$btnSaveReport = New-Object System.Windows.Forms.Button
$btnSaveReport.Text = "Save Report"
$btnSaveReport.Location = New-Object System.Drawing.Point(670, 65)
$btnSaveReport.Size = New-Object System.Drawing.Size(80, 25)
$btnSaveReport.Enabled = $false
$autoDetectPanel.Controls.Add($btnSaveReport)

# Function to analyze file structure issues
function Analyze-FileStructure {
    $sourceFolder = $txtSourceFolder.Text
    $processedFolder = $txtProcessedFolder.Text
    
    if ([string]::IsNullOrEmpty($sourceFolder) -or [string]::IsNullOrEmpty($processedFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select both source and processed folders first.", "Missing Folders", "OK", "Warning")
        return
    }
    
    $analysisForm = New-Object System.Windows.Forms.Form
    $analysisForm.Text = "File Structure Analysis"
    $analysisForm.Size = New-Object System.Drawing.Size(700, 500)
    $analysisForm.StartPosition = "CenterParent"
    
    $txtAnalysis = New-Object System.Windows.Forms.TextBox
    $txtAnalysis.Location = New-Object System.Drawing.Point(10, 10)
    $txtAnalysis.Size = New-Object System.Drawing.Size(665, 420)
    $txtAnalysis.Multiline = $true
    $txtAnalysis.ScrollBars = "Vertical"
    $txtAnalysis.ReadOnly = $true
    $txtAnalysis.Font = New-Object System.Drawing.Font("Consolas", 9)
    $analysisForm.Controls.Add($txtAnalysis)
    
    $btnCloseAnalysis = New-Object System.Windows.Forms.Button
    $btnCloseAnalysis.Text = "Close"
    $btnCloseAnalysis.Location = New-Object System.Drawing.Point(300, 440)
    $btnCloseAnalysis.Size = New-Object System.Drawing.Size(100, 30)
    $btnCloseAnalysis.Add_Click({ $analysisForm.Close() })
    $analysisForm.Controls.Add($btnCloseAnalysis)
    
    # Detect files and analyze
    $detectionResults = Auto-DetectFiles -SourceFolder $sourceFolder -ProcessedFolder $processedFolder
    
    $analysis = @()
    $analysis += "FILE STRUCTURE ANALYSIS FOR TESTERS"
    $analysis += "=" * 60
    $analysis += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $analysis += ""
    $analysis += "PURPOSE: Explain line count discrepancies between Notepad++ and data processing tools"
    $analysis += ""
    
    foreach ($dataType in @("Items", "Users", "Loans")) {
        $result = $detectionResults[$dataType]
        
        if ($result.SourceFile -and $result.ProcessedFile) {
            $analysis += "$dataType FILE ANALYSIS:"
            $analysis += "-" * 30
            
            try {
                # Read raw files
                $sourceLines = Get-Content $result.SourceFile
                $processedLines = Get-Content $result.ProcessedFile
                
                $analysis += "Source File: $([System.IO.Path]::GetFileName($result.SourceFile))"
                $analysis += "  Raw lines: $($sourceLines.Count)"
                $analysis += "  Empty lines: $(($sourceLines | Where-Object { $_.Trim() -eq '' }).Count)"
                
                $analysis += "Processed File: $([System.IO.Path]::GetFileName($result.ProcessedFile))"
                $analysis += "  Raw lines: $($processedLines.Count)"
                $analysis += "  Empty lines: $(($processedLines | Where-Object { $_.Trim() -eq '' }).Count)"
                
                # Check for data integrity
                $sourceData = Import-CsvSafe -FilePath $result.SourceFile
                $processedData = Import-CsvSafe -FilePath $result.ProcessedFile
                
                $analysis += "Data Analysis:"
                $analysis += "  Source processable rows: $($sourceData.Count)"
                $analysis += "  Processed processable rows: $($processedData.Count)"
                
                $sourceDiff = $sourceLines.Count - $sourceData.Count - 1
                $processedDiff = $processedLines.Count - $processedData.Count - 1
                
                if ($sourceDiff -gt 0) {
                    $analysis += "  Source file has $sourceDiff non-data lines (header + empty/malformed)"
                }
                if ($processedDiff -gt 0) {
                    $analysis += "  Processed file has $processedDiff non-data lines (header + empty/malformed)"
                }
                
                if ($sourceData.Count -eq $processedData.Count) {
                    $analysis += "  RESULT: No data loss - same number of valid records"
                } else {
                    $analysis += "  RESULT: Difference in valid records detected"
                }
                
                $analysis += ""
            }
            catch {
                $analysis += "  Error analyzing files: $($_.Exception.Message)"
                $analysis += ""
            }
        }
    }
    
    $analysis += "NOTES:"
    $analysis += "=" * 40
    $analysis += "1. Text editors (Notepad++) count ALL lines including:"
    $analysis += "   - Header rows"
    $analysis += "   - Empty lines at end of file"
    $analysis += "   - Malformed CSV rows"
    $analysis += "   - Trailing whitespace lines"
    $analysis += ""
    $analysis += "2. CSV processing tools count only VALID DATA rows:"
    $analysis += "   - Exclude header row"
    $analysis += "   - Skip empty lines"
    $analysis += "   - Filter malformed data"
    $analysis += ""
    $analysis += "3. For migration validation, DATA ROW counts are the accurate measure"
    $analysis += "4. Line count discrepancies are typically formatting artifacts, not data loss"
    $analysis += ""
    $analysis += "RECOMMENDATION: Use 'Data rows' counts for testing validation, not raw line counts"
    
    $txtAnalysis.Text = $analysis -join "`r`n"
    $analysisForm.ShowDialog()
}
function Show-ComparisonSummary {
    param($SummaryResults)
    
    # Create summary popup form
    $summaryForm = New-Object System.Windows.Forms.Form
    $summaryForm.Text = "Data Assurance Summary"
    $summaryForm.Size = New-Object System.Drawing.Size(500, 400)
    $summaryForm.StartPosition = "CenterParent"
    $summaryForm.FormBorderStyle = "FixedDialog"
    $summaryForm.MaximizeBox = $false
    $summaryForm.MinimizeBox = $false
    
    # Title label
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "MIGRATION DATA ASSURANCE SUMMARY"
    $lblTitle.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(450, 25)
    $lblTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.TextAlign = "MiddleCenter"
    $summaryForm.Controls.Add($lblTitle)
    
    # Summary text box
    $txtSummary = New-Object System.Windows.Forms.TextBox
    $txtSummary.Location = New-Object System.Drawing.Point(20, 60)
    $txtSummary.Size = New-Object System.Drawing.Size(450, 250)
    $txtSummary.Multiline = $true
    $txtSummary.ScrollBars = "Vertical"
    $txtSummary.ReadOnly = $true
    $txtSummary.Font = New-Object System.Drawing.Font("Consolas", 10)
    $summaryForm.Controls.Add($txtSummary)
    
    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(200, 325)
    $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.Add_Click({ $summaryForm.Close() })
    $summaryForm.Controls.Add($btnClose)
    
    # Generate summary content
    $summaryContent = @()
    $summaryContent += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $summaryContent += "=" * 55
    $summaryContent += ""
    
    $totalOriginal = 0
    $totalMigrated = 0
    $hasWarnings = $false
    $hasSkipped = $false
    
    foreach ($result in $SummaryResults) {
        $totalOriginal += $result.OriginalRows
        $totalMigrated += $result.MigratedRows
        
        $statusIcon = switch ($result.Assessment) {
            "PASS" { "[PASS]" }
            "WARNING" { "[WARN]"; $hasWarnings = $true }
            "INFO" { "[INFO]" }
            "SKIPPED" { "[SKIP]"; $hasSkipped = $true }
            default { "[?]" }
        }
        
        if ($result.Status -eq "SKIPPED") {
            $summaryContent += "$($result.DataType): $statusIcon Missing files/config"
        } else {
            $difference = $result.MigratedRows - $result.OriginalRows
            $diffText = if ($difference -eq 0) { "(No difference)" } 
                       elseif ($difference -gt 0) { "(+$difference rows)" }
                       else { "($difference rows)" }
            
            $summaryContent += "$($result.DataType): $statusIcon $($result.OriginalRows) -> $($result.MigratedRows) $diffText"
        }
    }
    
    $summaryContent += ""
    $summaryContent += "TOTALS:"
    $summaryContent += "  Original Records: $totalOriginal"
    $summaryContent += "  Migrated Records: $totalMigrated"
    $summaryContent += "  Net Difference: $($totalMigrated - $totalOriginal)"
    $summaryContent += ""
    
    # Overall assessment
    if ($hasSkipped -and $hasWarnings) {
        $summaryContent += "OVERALL STATUS: [REVIEW] Warnings + Missing Data"
    } elseif ($hasWarnings) {
        $summaryContent += "OVERALL STATUS: [REVIEW] Warnings Detected"
    } elseif ($hasSkipped) {
        $summaryContent += "OVERALL STATUS: [PARTIAL] Some Data Missing"
    } elseif ($totalOriginal -eq $totalMigrated -and $totalOriginal -gt 0) {
        $summaryContent += "OVERALL STATUS: [SUCCESS] All Data Verified"
    } else {
        $summaryContent += "OVERALL STATUS: [REVIEW] Check Details"
    }
    
    $summaryContent += ""
    $summaryContent += "Legend:"
    $summaryContent += "  [PASS] = Data integrity verified"
    $summaryContent += "  [WARN] = Potential issues detected"
    $summaryContent += "  [INFO] = Differences noted"
    $summaryContent += "  [SKIP] = Files not found/configured"
    $summaryContent += ""
    $summaryContent += "Click individual tabs for detailed reports."
    
    $txtSummary.Text = $summaryContent -join "`r`n"
    
    # Show the summary dialog
    $summaryForm.ShowDialog()
}
function Save-CombinedReport {
    if ([string]::IsNullOrEmpty($script:combinedReport)) {
        [System.Windows.Forms.MessageBox]::Show("No report to save. Please run comparisons first.", "No Report", "OK", "Warning")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Combined Report"
    $saveFileDialog.FileName = "Library_Data_Assurance_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $script:combinedReport | Out-File $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Report saved successfully to: $($saveFileDialog.FileName)", "Report Saved", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving report: $($_.Exception.Message)", "Save Error", "OK", "Error")
        }
    }
}

# Add event handlers for all tabs
foreach ($dataType in @("Items", "Users", "Loans")) {
    $btnOriginal = $form.Controls.Find("btnOriginal$dataType", $true)[0]
    $btnMigrated = $form.Controls.Find("btnMigrated$dataType", $true)[0]
    $btnCompare = $form.Controls.Find("btnCompare$dataType", $true)[0]
    
    $btnOriginal.Add_Click({
        $dataType = $this.Name.Substring(11)
        $txtOriginal = $form.Controls.Find("txtOriginal$dataType", $true)[0]
        $cmbOriginalKeyField = $form.Controls.Find("cmbOriginalKeyField$dataType", $true)[0]
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Delimited files (*.csv;*.tsv)|*.csv;*.tsv|CSV files (*.csv)|*.csv|TSV files (*.tsv)|*.tsv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Original Delimited File"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            $txtOriginal.Text = $openFileDialog.FileName
            Load-CsvHeaders -FilePath $openFileDialog.FileName -ComboBox $cmbOriginalKeyField
        }
    })
    
    $btnMigrated.Add_Click({
        $dataType = $this.Name.Substring(11)
        $txtMigrated = $form.Controls.Find("txtMigrated$dataType", $true)[0]
        $cmbMigratedKeyField = $form.Controls.Find("cmbMigratedKeyField$dataType", $true)[0]
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Delimited files (*.csv;*.tsv)|*.csv;*.tsv|CSV files (*.csv)|*.csv|TSV files (*.tsv)|*.tsv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Migrated Delimited File"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            $txtMigrated.Text = $openFileDialog.FileName
            Load-CsvHeaders -FilePath $openFileDialog.FileName -ComboBox $cmbMigratedKeyField
        }
    })
    
    $btnCompare.Add_Click({
        $dataType = $this.Name.Substring(10)
        Compare-Data -DataType $dataType
    })
}

# Add event handlers for global buttons
$btnBrowseSource.Add_Click({
    $folderPath = Browse-Folder -Description "Select Source Folder"
    if ($folderPath) {
        $txtSourceFolder.Text = $folderPath
    }
})

$btnBrowseProcessed.Add_Click({
    $folderPath = Browse-Folder -Description "Select Processed Folder"
    if ($folderPath) {
        $txtProcessedFolder.Text = $folderPath
    }
})

# Initialize folder paths to script directory on form load
$scriptDirectory = Get-Location
$txtSourceFolder.Text = $scriptDirectory.Path
$txtProcessedFolder.Text = $scriptDirectory.Path

$btnAutoDetect.Add_Click({
    $sourceFolder = $txtSourceFolder.Text
    $processedFolder = $txtProcessedFolder.Text
    
    if ([string]::IsNullOrEmpty($sourceFolder) -or [string]::IsNullOrEmpty($processedFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select both source and processed folders first.", "Missing Folders", "OK", "Warning")
        return
    }
    
    if (-not (Test-Path $sourceFolder) -or -not (Test-Path $processedFolder)) {
        [System.Windows.Forms.MessageBox]::Show("One or both selected folders do not exist.", "Invalid Folders", "OK", "Error")
        return
    }
    
    $lblStatus.Text = "Auto-detecting files..."
    $form.Refresh()
    
    $detectionResults = Auto-DetectFiles -SourceFolder $sourceFolder -ProcessedFolder $processedFolder
    Update-FilesFromDetection -DetectionResults $detectionResults
    
    $lblStatus.Text = "Auto-detection completed"
})

$btnAnalyzeFiles.Add_Click({
    Analyze-FileStructure
})

$btnRunAll.Add_Click({
    Run-AllComparisons
})

$btnSaveReport.Add_Click({
    Save-CombinedReport
})

# Show the form
$form.ShowDialog()
