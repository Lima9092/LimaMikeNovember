Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Assurance Tool - Sampling Method"
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

# Updated Perform-UsersSampling
function Perform-UsersSampling {
    param($OriginalData, $MigratedData, $ColumnMappings, $SampleSize)
    
    $report = @()
    $report += "USERS SAMPLING ANALYSIS:"
    $report += "-" * 30
    
    # Get column names from mappings
    $sourceUsernameCol = $ColumnMappings["cmbSourceUsernameUsers"]
    $sourcePrisonNumberCol = $ColumnMappings["cmbSourcePrisonNumberUsers"]
    $sourceFullName1Col = $ColumnMappings["cmbSourceFullName1Users"]
    $sourceFullName2Col = $ColumnMappings["cmbSourceFullName2Users"]
    $migratedBarcodeCol = "barcode"     # Fixed: maps to source Username
    $migratedPrisonIDCol = "prison_id"  # Fixed: maps to source Prison Number
    $migratedForenameCol = "forename"   # Fixed: maps to source Full Name
    $migratedSurnameCol = "surname"     # Fixed: maps to source Full Name
    
    if (-not $sourceUsernameCol) {
        $report += "ERROR: Source username column not selected"
        return $report
    }
    
    # Filter for users with complete required data (username + prison number)
    $completeSourceUsers = @()
    foreach ($user in $OriginalData) {
        $hasUsername = $user.$sourceUsernameCol -and $user.$sourceUsernameCol.ToString().Trim() -ne ""
        $hasPrisonNumber = $user.$sourcePrisonNumberCol -and $user.$sourcePrisonNumberCol.ToString().Trim() -ne "" -and $user.$sourcePrisonNumberCol.ToString().Trim() -ne "N/A"
        
        if ($hasUsername -and $hasPrisonNumber) {
            $completeSourceUsers += $user
        }
    }
    
    if ($completeSourceUsers.Count -eq 0) {
        $report += "ERROR: No users found with complete required data (username AND prison number)"
        $report += "Total source users: $($OriginalData.Count)"
        
        # Show breakdown of missing data
        $missingPrisonNumber = 0
        $missingUsername = 0
        $missingBoth = 0
        
        foreach ($user in $OriginalData) {
            $hasUsername = $user.$sourceUsernameCol -and $user.$sourceUsernameCol.ToString().Trim() -ne ""
            $hasPrisonNumber = $user.$sourcePrisonNumberCol -and $user.$sourcePrisonNumberCol.ToString().Trim() -ne "" -and $user.$sourcePrisonNumberCol.ToString().Trim() -ne "N/A"
            
            if (-not $hasUsername -and -not $hasPrisonNumber) { $missingBoth++ }
            elseif (-not $hasUsername) { $missingUsername++ }
            elseif (-not $hasPrisonNumber) { $missingPrisonNumber++ }
        }
        
        $report += "Users missing Prison Number: $missingPrisonNumber"
        $report += "Users missing Username: $missingUsername"
        $report += "Users missing both: $missingBoth"
        return $report
    }
    
    # Get sample from users with complete data
    $sampleUsers = $completeSourceUsers | Get-Random -Count ([Math]::Min($SampleSize, $completeSourceUsers.Count))
    
    $report += "Sample of $($sampleUsers.Count) users selected from $($completeSourceUsers.Count) users with complete data:"
    $report += "Total source users: $($OriginalData.Count)"
    $report += "Users with complete data (Username + Prison Number): $($completeSourceUsers.Count)"
    $report += "NOTE: Migrated usernames expected to have 3-character suffix (e.g., USERNAME -> USERNAMELNI)"
    $report += ""
    
    foreach ($user in $sampleUsers) {
        $username = $user.$sourceUsernameCol.ToString().Trim()
        $prisonNumber = $user.$sourcePrisonNumberCol.ToString().Trim()
        $fullName1 = if ($sourceFullName1Col -and $user.$sourceFullName1Col) { $user.$sourceFullName1Col } else { "N/A" }
        $fullName2 = if ($sourceFullName2Col -and $user.$sourceFullName2Col) { $user.$sourceFullName2Col } else { "N/A" }
        
        $report += "USER: $username"
        $report += "  Prison Number: $prisonNumber"
        $report += "  Full Name (forename): $fullName1"
        $report += "  Full Name (surname): $fullName2"
        
        # Check for migrated data by barcode prefix (username with 3-character suffix)
        $migratedUsers = @()
        foreach ($migratedUser in $MigratedData) {
            if ($migratedUser.$migratedBarcodeCol -and 
                $migratedUser.$migratedBarcodeCol.ToString().Trim().StartsWith($username, [System.StringComparison]::OrdinalIgnoreCase)) {
                $migratedUsers += $migratedUser
            }
        }
        
        $migratedCount = $migratedUsers.Count
        $report += "  Migrated occurrences (by username prefix): $migratedCount"
        
        # Show migrated user details if found
        if ($migratedCount -gt 0) {
            $migratedUser = $migratedUsers[0]  # Take first match
            
            $migratedBarcode = $migratedUser.$migratedBarcodeCol.ToString().Trim()
            $migratedPrisonID = if ($migratedUser.$migratedPrisonIDCol) { $migratedUser.$migratedPrisonIDCol.ToString().Trim() } else { "N/A" }
            $migratedForename = if ($migratedUser.$migratedForenameCol) { $migratedUser.$migratedForenameCol.ToString().Trim() } else { "N/A" }
            $migratedSurname = if ($migratedUser.$migratedSurnameCol) { $migratedUser.$migratedSurnameCol.ToString().Trim() } else { "N/A" }
            
            $report += "  Migrated Username: $migratedBarcode"
            $report += "  Migrated Prison ID: $migratedPrisonID"
            $report += "  Migrated Name: $migratedForename $migratedSurname"
            
            # Compare prison numbers
            if ($migratedPrisonID.ToLower() -eq $prisonNumber.ToLower()) {
                $report += "  Prison Number Match: [PASS] ($prisonNumber)"
            } elseif ($migratedPrisonID -eq "N/A") {
                $report += "  Prison Number Match: [WARN] Missing in migrated data"
            } else {
                $report += "  Prison Number Match: [WARN] Different (Source: $prisonNumber, Migrated: $migratedPrisonID)"
            }
            
            # Show case transformation details
            $originalCase = $username
            $migratedCase = $migratedBarcode
            if ($originalCase -cne $migratedCase.Substring(0, [Math]::Min($originalCase.Length, $migratedCase.Length))) {
                $report += "  Case Transform: $originalCase -> $migratedCase"
            }
        } else {
            $report += "  Migrated user: NOT FOUND"
        }
        
        # Assessment
        if ($migratedCount -gt 0) {
            $report += "  STATUS: [PASS] User found in migrated data"
        } else {
            $report += "  STATUS: [FAIL] User not found in migrated data"
        }
        
        $report += ""
    }
    
    return $report
}

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

# Function to load delimited file headers and populate dropdown
function Load-CsvHeaders {
    param([string]$FilePath, [System.Windows.Forms.ComboBox]$ComboBox, [string[]]$PreferredHeaders = @())
    
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
        $ComboBox.Items.Add("") # Add empty option
        foreach ($header in $processedHeaders) {
            $ComboBox.Items.Add($header)
        }
        
        # Auto-select preferred headers if they exist (case insensitive)
        $selectedIndex = 0 # Default to empty
        foreach ($preferred in $PreferredHeaders) {
            for ($i = 1; $i -lt $ComboBox.Items.Count; $i++) {
                if ($ComboBox.Items[$i].ToString() -eq $preferred) {
                    $selectedIndex = $i
                    break
                }
            }
            if ($selectedIndex -gt 0) { break }
        }
        
        # If no exact match, try case-insensitive match
        if ($selectedIndex -eq 0) {
            foreach ($preferred in $PreferredHeaders) {
                for ($i = 1; $i -lt $ComboBox.Items.Count; $i++) {
                    if ($ComboBox.Items[$i].ToString().ToLower() -eq $preferred.ToLower()) {
                        $selectedIndex = $i
                        break
                    }
                }
                if ($selectedIndex -gt 0) { break }
            }
        }
        
        $ComboBox.SelectedIndex = $selectedIndex
        
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
        
        if ($result.SourceFile) {
            $txtOriginal.Text = $result.SourceFile
            $fileName = [System.IO.Path]::GetFileName($result.SourceFile)
            $detectionReport += "$dataType Source: $fileName (Found)"
            
            try {
                # Load headers for all source dropdowns for this data type
                Load-HeadersForDataType -DataType $dataType -FilePath $result.SourceFile -FileType "Source"
                $detectionReport += "  - Source headers loaded successfully"
            }
            catch {
                $detectionReport += "  - ERROR loading source headers: $($_.Exception.Message)"
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
            $detectionReport += "  - Migrated columns are hardcoded"
        } else {
            $txtMigrated.Text = ""
            $detectionReport += "$dataType Processed: Not found"
        }
        
        $detectionReport += ""
    }
    
    $detectionText = $detectionReport -join "`r`n"
    [System.Windows.Forms.MessageBox]::Show($detectionText, "File Detection Results", "OK", "Information")
}

# Function to load headers for all dropdowns of a specific data type
function Load-HeadersForDataType {
    param([string]$DataType, [string]$FilePath, [string]$FileType)
    
    # Only load source file headers - migrated columns are now hardcoded labels
    if ($FileType -eq "Source") {
        # Define which dropdowns to populate based on data type
        $dropdownsToLoad = @()
        
        switch ($DataType) {
            "Items" {
                $dropdownsToLoad = @("cmbSourceBarcodeItems", "cmbSourceISBNItems", "cmbSourceTitleItems", "cmbSourcePublisherItems")
            }
            "Users" {
                $dropdownsToLoad = @("cmbSourceUsernameUsers", "cmbSourcePrisonNumberUsers", "cmbSourceFullName1Users", "cmbSourceFullName2Users")
            }
            "Loans" {
                $dropdownsToLoad = @("cmbSourceItemBarcodeLoans", "cmbSourceUserBarcodeLoans", "cmbSourceIssueDateLoans", "cmbSourceReturnDateLoans")
            }
        }
        
        # Load headers for each dropdown
        foreach ($dropdownName in $dropdownsToLoad) {
            $dropdown = $form.Controls.Find($dropdownName, $true)[0]
            if ($dropdown) {
                $preferredHeaders = Get-PreferredHeaders -DataType $DataType -FieldType $FileType -FieldName $dropdownName
                Load-CsvHeaders -FilePath $FilePath -ComboBox $dropdown -PreferredHeaders $preferredHeaders
            }
        }
    }
}

# Updated Get-PreferredHeaders function with correct duplicate naming patterns
function Get-PreferredHeaders {
    param([string]$DataType, [string]$FieldType, [string]$FieldName)
    
    $preferences = @{
        "Items" = @{
            "Source" = @{
                "cmbSourceBarcodeItems" = @("Barcode", "barcode", "BARCODE")
                "cmbSourceISBNItems" = @("ISBN-10", "ISBN10", "ISBN", "isbn")
                "cmbSourceTitleItems" = @("Title", "title", "TITLE")
                "cmbSourcePublisherItems" = @("Publisher", "publisher", "PUBLISHER")
            }
        }
        "Users" = @{
            "Source" = @{
                "cmbSourceUsernameUsers" = @("Username", "username", "USER_NAME")
                "cmbSourcePrisonNumberUsers" = @("Prison Number", "PRISON_NUMBER", "PRN", "Number", "PrisonNumber")
                "cmbSourceFullName1Users" = @("Full Name", "FULL_NAME", "Fullname", "Name", "FullName")
                "cmbSourceFullName2Users" = @("Full Name", "FULL_NAME", "Fullname", "Name", "FullName")
            }
        }
        "Loans" = @{
            "Source" = @{
                "cmbSourceItemBarcodeLoans" = @("Borrowed", "Item Barcode", "ITEM_BARCODE", "Barcode")
                "cmbSourceUserBarcodeLoans" = @("Borrowed By", "User Barcode", "USER_BARCODE", "Patron")
                "cmbSourceIssueDateLoans" = @("Borrowed_2", "Borrowed-2", "Issue Date", "ISSUE_DATE", "Date Out")  # FIXED: Added Borrowed_2
                "cmbSourceReturnDateLoans" = @("Date Return Due", "Return Date", "RETURN_DATE", "Due Date", "Date Due")  # FIXED: Added "Date Return By"
            }
        }
    }
    
    if ($preferences.ContainsKey($DataType) -and 
        $preferences[$DataType].ContainsKey($FieldType) -and 
        $preferences[$DataType][$FieldType].ContainsKey($FieldName)) {
        return $preferences[$DataType][$FieldType][$FieldName]
    }
    
    return @()
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

# Function to get column mappings by exact control names - THIS IS THE KEY FIX
function Get-ColumnMappings {
    param([string]$DataType)
    
    $columnMappings = @{}
    
    # Define the exact control names for each data type
    $controlNames = switch ($DataType) {
        "Items" {
            @("cmbSourceBarcodeItems", "cmbSourceISBNItems", "cmbSourceTitleItems", "cmbSourcePublisherItems")
        }
        "Users" {
            @("cmbSourceUsernameUsers", "cmbSourcePrisonNumberUsers", "cmbSourceFullName1Users", "cmbSourceFullName2Users")
        }
        "Loans" {
            @("cmbSourceItemBarcodeLoans", "cmbSourceUserBarcodeLoans", "cmbSourceIssueDateLoans", "cmbSourceReturnDateLoans")
        }
        default { @() }
    }
    
    # Get the selected values from each dropdown
    foreach ($controlName in $controlNames) {
        $control = $form.Controls.Find($controlName, $true)[0]
        if ($control -and $control.SelectedItem -and $control.SelectedItem.ToString() -ne "") {
            $columnMappings[$controlName] = $control.SelectedItem.ToString()
        }
    }
    
    return $columnMappings
}

# Function to perform sampling comparison - FIXED VERSION
function Compare-DataSampling {
    param([string]$DataType, [bool]$IsGlobalRun = $false, [int]$SampleSize = 5)
    
    $txtOriginal = $form.Controls.Find("txtOriginal$DataType", $true)[0]
    $txtMigrated = $form.Controls.Find("txtMigrated$DataType", $true)[0]
    $txtResults = $form.Controls.Find("txtResults$DataType", $true)[0]
    $numSampleSize = $form.Controls.Find("numSampleSize$DataType", $true)[0]
    
    $originalPath = $txtOriginal.Text
    $migratedPath = $txtMigrated.Text
    $sampleSize = if ($numSampleSize) { $numSampleSize.Value } else { 5 }
    
    if ([string]::IsNullOrEmpty($originalPath) -or [string]::IsNullOrEmpty($migratedPath)) {
        if (-not $IsGlobalRun) {
            [System.Windows.Forms.MessageBox]::Show("Please select both original and migrated files for $DataType.", "Missing Files", "OK", "Warning")
        }
        return $null
    }
    
    try {
        if (-not $IsGlobalRun) {
            $txtResults.Text = "Loading and processing files for sampling..."
            $form.Refresh()
        }
        
        $originalData = Import-CsvSafe -FilePath $originalPath
        $migratedData = Import-CsvSafe -FilePath $migratedPath
        
        if (-not $IsGlobalRun) {
            $txtResults.Text = "Performing sampling analysis..."
            $form.Refresh()
        }
        
        $report = @()
        $report += "=" * 60
        $report += "$DataType Sampling Report"
        $report += "=" * 60
        $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $report += "Sample Size: $sampleSize"
        $report += ""
        $report += "Files Analyzed:"
        $report += "  Original: $originalPath"
        $report += "  Migrated: $migratedPath"
        $report += ""
        
        # Get selected column mappings using the fixed function
        $columnMappings = Get-ColumnMappings -DataType $DataType
        
        $report += "Column Mappings:"
        foreach ($mapping in $columnMappings.GetEnumerator()) {
            $report += "  $($mapping.Key): $($mapping.Value)"
        }
        $report += ""
        
        # Perform sampling based on data type
        switch ($DataType) {
            "Items" {
                $sampleReport = Perform-ItemsSampling -OriginalData $originalData -MigratedData $migratedData -ColumnMappings $columnMappings -SampleSize $sampleSize
            }
            "Users" {
                $sampleReport = Perform-UsersSampling -OriginalData $originalData -MigratedData $migratedData -ColumnMappings $columnMappings -SampleSize $sampleSize
            }
            "Loans" {
                $sampleReport = Perform-LoansSampling -OriginalData $originalData -MigratedData $migratedData -ColumnMappings $columnMappings -SampleSize $sampleSize
            }
        }
        
        $report += $sampleReport
        
        $reportText = $report -join "`r`n"
        if (-not $IsGlobalRun) {
            $txtResults.Text = $reportText
        }
        
        return $reportText
        
    }
    catch {
        $errorMsg = "Error during sampling: $($_.Exception.Message)"
        if (-not $IsGlobalRun) {
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
        }
        return $errorMsg
    }
}

# Updated Perform-ItemsSampling - Shows sample migrated RCN
function Perform-ItemsSampling {
    param($OriginalData, $MigratedData, $ColumnMappings, $SampleSize)
    
    $report = @()
    $report += "ITEMS SAMPLING ANALYSIS:"
    $report += "-" * 30
    
    # Get column names from mappings
    $sourceBarcodeCol = $ColumnMappings["cmbSourceBarcodeItems"]
    $sourceISBNCol = $ColumnMappings["cmbSourceISBNItems"]  # This is the RCN equivalent
    $sourceTitleCol = $ColumnMappings["cmbSourceTitleItems"]
    $sourcePublisherCol = $ColumnMappings["cmbSourcePublisherItems"]
    $migratedBarcodeCol = "barcode"  # Fixed migrated column
    $migratedRCNCol = "rcn"          # Fixed migrated column
    
    if (-not $sourceBarcodeCol) {
        $report += "ERROR: Source barcode column not selected"
        return $report
    }
    
    if (-not $sourceISBNCol) {
        $report += "ERROR: Source ISBN/RCN column not selected"
        return $report
    }
    
    # Filter for items with BOTH valid barcode AND valid ISBN (complete required data)
    $completeSourceItems = @()
    foreach ($item in $OriginalData) {
        $hasBarcode = $item.$sourceBarcodeCol -and $item.$sourceBarcodeCol.ToString().Trim() -ne ""
        $hasISBN = $item.$sourceISBNCol -and $item.$sourceISBNCol.ToString().Trim() -ne "" -and $item.$sourceISBNCol.ToString().Trim() -ne "N/A"
        
        if ($hasBarcode -and $hasISBN) {
            $completeSourceItems += $item
        }
    }
    
    if ($completeSourceItems.Count -eq 0) {
        $report += "ERROR: No items found with complete required data (barcode AND ISBN)"
        $report += "Total source items: $($OriginalData.Count)"
        
        # Show breakdown of missing data
        $missingISBN = 0
        $missingBarcode = 0
        $missingBoth = 0
        
        foreach ($item in $OriginalData) {
            $hasBarcode = $item.$sourceBarcodeCol -and $item.$sourceBarcodeCol.ToString().Trim() -ne ""
            $hasISBN = $item.$sourceISBNCol -and $item.$sourceISBNCol.ToString().Trim() -ne "" -and $item.$sourceISBNCol.ToString().Trim() -ne "N/A"
            
            if (-not $hasBarcode -and -not $hasISBN) { $missingBoth++ }
            elseif (-not $hasBarcode) { $missingBarcode++ }
            elseif (-not $hasISBN) { $missingISBN++ }
        }
        
        $report += "Items missing ISBN: $missingISBN"
        $report += "Items missing Barcode: $missingBarcode"
        $report += "Items missing both: $missingBoth"
        return $report
    }
    
    # Get sample from items with complete data
    $sampleItems = $completeSourceItems | Get-Random -Count ([Math]::Min($SampleSize, $completeSourceItems.Count))
    
    $report += "Sample of $($sampleItems.Count) items selected from $($completeSourceItems.Count) items with complete data:"
    $report += "Total source items: $($OriginalData.Count)"
    $report += "Items with complete data (Barcode + ISBN): $($completeSourceItems.Count)"
    $report += "NOTE: Only sampling items with both Barcode AND ISBN for RCN-based comparison"
    $report += ""
    
    foreach ($item in $sampleItems) {
        $barcode = $item.$sourceBarcodeCol.ToString().Trim()
        $isbn = $item.$sourceISBNCol.ToString().Trim()
        $title = if ($sourceTitleCol -and $item.$sourceTitleCol) { $item.$sourceTitleCol } else { "N/A" }
        $publisher = if ($sourcePublisherCol -and $item.$sourcePublisherCol) { $item.$sourcePublisherCol } else { "N/A" }
        
        $report += "ITEM: $barcode"
        $report += "  ISBN-10/RCN: $isbn"
        $report += "  Title: $title"
        $report += "  Publisher: $publisher"
        
        # Count occurrences in SOURCE by ISBN/RCN
        $sourceCountByRCN = 0
        foreach ($sourceItem in $OriginalData) {
            if ($sourceItem.$sourceISBNCol -and $sourceItem.$sourceISBNCol.ToString().Trim() -eq $isbn) {
                $sourceCountByRCN++
            }
        }
        
        # Count occurrences in MIGRATED by RCN field
        $migratedCountByRCN = 0
        $sampleMigratedRCN = $null
        foreach ($migratedItem in $MigratedData) {
            if ($migratedItem.$migratedRCNCol -and $migratedItem.$migratedRCNCol.ToString().Trim() -eq $isbn) {
                $migratedCountByRCN++
                if (-not $sampleMigratedRCN) {
                    $sampleMigratedRCN = $migratedItem.$migratedRCNCol.ToString().Trim()
                }
            }
        }
        
        # Count occurrences in MIGRATED by barcode prefix (supplementary)
        $migratedCountByBarcode = 0
        $sampleMigratedBarcode = $null
        foreach ($migratedItem in $MigratedData) {
            if ($migratedItem.$migratedBarcodeCol -and $migratedItem.$migratedBarcodeCol.ToString().Trim().StartsWith($barcode)) {
                $migratedCountByBarcode++
                if (-not $sampleMigratedBarcode) {
                    $sampleMigratedBarcode = $migratedItem.$migratedBarcodeCol.ToString().Trim()
                }
            }
        }
        
        # Report the counts
        $report += "  Source occurrences (by ISBN/RCN): $sourceCountByRCN"
        $report += "  Migrated occurrences (by RCN): $migratedCountByRCN"
        $report += "  Migrated occurrences (by barcode prefix): $migratedCountByBarcode"
        
        # Show sample migrated values
        if ($sampleMigratedBarcode) {
            $report += "  Sample migrated barcode: $sampleMigratedBarcode"
        }
        if ($sampleMigratedRCN) {
            $report += "  Sample migrated RCN: $sampleMigratedRCN"
        }
        
        # Assessment - RCN-based comparison for complete data
        if ($sourceCountByRCN -eq $migratedCountByRCN -and $migratedCountByRCN -gt 0) {
            $report += "  STATUS: [PASS] RCN counts match ($sourceCountByRCN each)"
        } elseif ($migratedCountByRCN -eq 0 -and $migratedCountByBarcode -gt 0) {
            $report += "  STATUS: [WARN] Found by barcode but not by RCN (RCN may be missing/different in migrated)"
        } elseif ($migratedCountByRCN -eq 0 -and $migratedCountByBarcode -eq 0) {
            $report += "  STATUS: [FAIL] Item not found in migrated data (neither RCN nor barcode)"
        } else {
            $report += "  STATUS: [WARN] RCN count mismatch (Source: $sourceCountByRCN, Migrated: $migratedCountByRCN)"
        }
        
        $report += ""
    }
    
    return $report
}

# Enhanced Perform-LoansSampling with additional user-based analysis
function Perform-LoansSampling {
    param($OriginalData, $MigratedData, $ColumnMappings, $SampleSize)
    
    $report = @()
    $report += "LOANS SAMPLING ANALYSIS:"
    $report += "-" * 30
    
    # Get column names from mappings
    $sourceItemBarcodeCol = $ColumnMappings["cmbSourceItemBarcodeLoans"]
    $sourceUserBarcodeCol = $ColumnMappings["cmbSourceUserBarcodeLoans"]
    $sourceIssueDateCol = $ColumnMappings["cmbSourceIssueDateLoans"]
    $sourceReturnDateCol = $ColumnMappings["cmbSourceReturnDateLoans"]
    $migratedItemBarcodeCol = "itembarcode"     # From debug output
    $migratedUserBarcodeCol = "borrbarcode"     # From debug output
    $migratedIssueDateCol = "issuedate"         # From debug output
    $migratedReturnDateCol = "duedate"          # From debug output
    
    if (-not $sourceUserBarcodeCol -or -not $sourceItemBarcodeCol) {
        $report += "ERROR: Required barcode columns not selected"
        return $report
    }
    
    # Enhanced barcode extraction function using your proven logic
    function Extract-ItemBarcode {
        param([string]$borrowed)
        
        $barcode = ''
        $borrowedTitle = $borrowed
        
        if ($borrowed -match '^(.*?)\s*:\s*Copy\s+\d+\s*\((\d{6})\s*-') {
            $borrowedTitle = $matches[1].Trim()
            $barcode = $matches[2]
        } elseif ($borrowed -match '(\d{6})') {
            $barcode = $matches[1]
            if ($borrowed -match '^(.*?)\s*\([^)]*\d{6}') {
                $borrowedTitle = $matches[1].Trim()
            } else {
                if ($borrowed -match "^(.*?)\s*$barcode") {
                    $borrowedTitle = $matches[1].Trim()
                }
            }
        } else {
            if ($borrowed -match '^(.*?)\s*\(([^)]+)\)') {
                $borrowedTitle = $matches[1].Trim()
                $barcodeRaw = $matches[2]
                if ($barcodeRaw -match '^(\S+)') {
                    $barcode = $matches[1]
                } else {
                    $barcode = $barcodeRaw
                }
            }
        }
        
        return @{
            Barcode = $barcode
            Title = $borrowedTitle
        }
    }
    
    # PART 1: INDIVIDUAL LOAN SAMPLING (existing logic)
    $report += "PART 1: INDIVIDUAL LOAN SAMPLING"
    $report += "=" * 35
    
    # Filter for loans with extractable item barcodes
    $validSourceLoans = @()
    foreach ($loan in $OriginalData) {
        $itemText = if ($loan.$sourceItemBarcodeCol) { $loan.$sourceItemBarcodeCol.ToString() } else { "" }
        $userText = if ($loan.$sourceUserBarcodeCol) { $loan.$sourceUserBarcodeCol.ToString() } else { "" }
        
        if ($itemText -ne "" -and $userText -ne "") {
            # Try to extract item barcode using proven logic
            $extractionResult = Extract-ItemBarcode -borrowed $itemText
            if ($extractionResult.Barcode -ne "" -and $extractionResult.Barcode -match '^\d+$') {
                $validSourceLoans += $loan
            }
        }
    }
    
    if ($validSourceLoans.Count -eq 0) {
        $report += "ERROR: No loans found with extractable item barcodes"
        $report += "Total source loans: $($OriginalData.Count)"
        return $report
    }
    
    # Get sample from valid loans
    $sampleLoans = $validSourceLoans | Get-Random -Count ([Math]::Min($SampleSize, $validSourceLoans.Count))
    
    $report += "Sample of $($sampleLoans.Count) individual loans:"
    $report += ""
    
    foreach ($loan in $sampleLoans) {
        $sourceItemText = $loan.$sourceItemBarcodeCol.ToString()
        $sourceUserText = $loan.$sourceUserBarcodeCol.ToString()
        
        # Extract the actual item barcode using proven logic
        $extractionResult = Extract-ItemBarcode -borrowed $sourceItemText
        $extractedItemBarcode = $extractionResult.Barcode
        $extractedTitle = $extractionResult.Title
        
        $report += "LOAN:"
        $report += "  Source Item Text: $sourceItemText"
        $report += "  Extracted Item Barcode: $extractedItemBarcode"
        $report += "  Extracted Title: $extractedTitle"
        $report += "  Source User Text: $sourceUserText"
        
        # Find matching loans in migrated data by item barcode (with 3-char suffix)
        $matchingLoans = @()
        foreach ($migratedLoan in $MigratedData) {
            if ($migratedLoan.$migratedItemBarcodeCol -and 
                $migratedLoan.$migratedItemBarcodeCol.ToString().Trim().StartsWith($extractedItemBarcode)) {
                $matchingLoans += $migratedLoan
            }
        }
        
        $migratedCount = $matchingLoans.Count
        $report += "  Migrated loans (by item barcode): $migratedCount"
        
        # Show details of first matching migrated loan
        if ($migratedCount -gt 0) {
            $migratedLoan = $matchingLoans[0]
            $migratedItemBarcode = $migratedLoan.$migratedItemBarcodeCol.ToString().Trim()
            $migratedUserBarcode = $migratedLoan.$migratedUserBarcodeCol.ToString().Trim()
            
            $report += "  Migrated Item: $migratedItemBarcode, User: $migratedUserBarcode"
            
            # Validate the barcode transformation
            if ($migratedItemBarcode.StartsWith($extractedItemBarcode)) {
                $suffix = $migratedItemBarcode.Substring($extractedItemBarcode.Length)
                $report += "  Barcode Transform: [PASS] $extractedItemBarcode - $migratedItemBarcode (suffix: $suffix)"
            }
            
            if ($migratedCount -eq 1) {
                $report += "  STATUS: [PASS] Item found with single loan"
            } else {
                $report += "  STATUS: [INFO] Item found with multiple loans ($migratedCount total)"
            }
        } else {
            $report += "  STATUS: [FAIL] Item barcode '$extractedItemBarcode' not found in migrated data"
        }
        
        $report += ""
    }
    
    # PART 2: USER-BASED LOAN TOTALS (new analysis)
    $report += ""
    $report += "PART 2: USER LOAN TOTALS ANALYSIS"
    $report += "=" * 35
    
    # Get unique users from source data (by user barcode/name)
    $uniqueSourceUsers = @{}
    foreach ($loan in $OriginalData) {
        $userText = if ($loan.$sourceUserBarcodeCol) { $loan.$sourceUserBarcodeCol.ToString().Trim() } else { "" }
        if ($userText -ne "") {
            if (-not $uniqueSourceUsers.ContainsKey($userText)) {
                $uniqueSourceUsers[$userText] = @()
            }
            $uniqueSourceUsers[$userText] += $loan
        }
    }
    
    if ($uniqueSourceUsers.Count -eq 0) {
        $report += "ERROR: No users found in source loan data"
        return $report
    }
    
    # Get random sample of 5 users
    $userNames = $uniqueSourceUsers.Keys | Get-Random -Count ([Math]::Min(5, $uniqueSourceUsers.Count))
    
    $report += "Sample of $($userNames.Count) users selected from $($uniqueSourceUsers.Count) unique users:"
    $report += ""
    
    foreach ($userName in $userNames) {
        $userSourceLoans = $uniqueSourceUsers[$userName]
        $sourceLoansCount = $userSourceLoans.Count
        
        $report += "USER: $userName"
        $report += "  Source loans total: $sourceLoansCount"
        
        # Try to find this user in migrated data
        # Since source has names and migrated has barcodes, we'll try different approaches
        $migratedUserLoans = @()
        
        # Approach 1: Look for exact match (unlikely but possible)
        foreach ($migratedLoan in $MigratedData) {
            if ($migratedLoan.$migratedUserBarcodeCol -and 
                $migratedLoan.$migratedUserBarcodeCol.ToString().Trim() -eq $userName) {
                $migratedUserLoans += $migratedLoan
            }
        }
        
        # Approach 2: Look for partial name match in migrated user barcode (if it contains name info)
        if ($migratedUserLoans.Count -eq 0) {
            $nameWords = $userName -split '\s+' | Where-Object { $_.Length -gt 2 }  # Words longer than 2 chars
            foreach ($migratedLoan in $MigratedData) {
                if ($migratedLoan.$migratedUserBarcodeCol) {
                    $migratedUserBarcode = $migratedLoan.$migratedUserBarcodeCol.ToString().Trim()
                    foreach ($word in $nameWords) {
                        if ($migratedUserBarcode -match [regex]::Escape($word)) {
                            $migratedUserLoans += $migratedLoan
                            break
                        }
                    }
                }
            }
        }
        
        $migratedLoansCount = $migratedUserLoans.Count
        $report += "  Migrated loans total: $migratedLoansCount"
        
        # Show sample migrated user barcode if found
        if ($migratedLoansCount -gt 0) {
            $sampleMigratedUser = ($migratedUserLoans | Group-Object -Property $migratedUserBarcodeCol | Sort-Object Count -Descending)[0]
            $report += "  Sample migrated user barcode: $($sampleMigratedUser.Name) ($($sampleMigratedUser.Count) loans)"
        }
        
        # Assessment
        if ($sourceLoansCount -eq $migratedLoansCount -and $migratedLoansCount -gt 0) {
            $report += "  STATUS: [PASS] Loan counts match ($sourceLoansCount each)"
        } elseif ($migratedLoansCount -eq 0) {
            $report += "  STATUS: [FAIL] User not found in migrated data"
            $report += "  NOTE: User matching requires name barcode lookup for accurate validation"
        } else {
            $report += "  STATUS: [WARN] Loan count mismatch (Source: $sourceLoansCount, Migrated: $migratedLoansCount)"
        }
        
        $report += ""
    }
    
    return $report
}

# Function to create tab for each data type with sampling controls
function Create-DataTab {
    param([string]$DataType)
    
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $DataType
    $tab.UseVisualStyleBackColor = $true
    
    # File selection section
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
    
    # Sample size control
    $lblSampleSize = New-Object System.Windows.Forms.Label
    $lblSampleSize.Text = "Sample Size:"
    $lblSampleSize.Location = New-Object System.Drawing.Point(620, 20)
    $lblSampleSize.Size = New-Object System.Drawing.Size(80, 20)
    $tab.Controls.Add($lblSampleSize)
    
    $numSampleSize = New-Object System.Windows.Forms.NumericUpDown
    $numSampleSize.Location = New-Object System.Drawing.Point(620, 45)
    $numSampleSize.Size = New-Object System.Drawing.Size(60, 20)
    $numSampleSize.Minimum = 1
    $numSampleSize.Maximum = 50
    $numSampleSize.Value = 5
    $numSampleSize.Name = "numSampleSize$DataType"
    $tab.Controls.Add($numSampleSize)
    
    # Column mapping section
    $lblColumnMapping = New-Object System.Windows.Forms.Label
    $lblColumnMapping.Text = "Column Mappings:"
    $lblColumnMapping.Location = New-Object System.Drawing.Point(20, 140)
    $lblColumnMapping.Size = New-Object System.Drawing.Size(120, 20)
    $lblColumnMapping.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $tab.Controls.Add($lblColumnMapping)
    
    $yPos = 165
    
    # Create column dropdowns based on data type
    switch ($DataType) {
        "Items" {
            # Barcode mapping: barcode ← Barcode
            $lblBarcode = New-Object System.Windows.Forms.Label
            $lblBarcode.Text = "Barcode:"
            $lblBarcode.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblBarcode.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblBarcode)
            
            $lblBarcodeFixed = New-Object System.Windows.Forms.Label
            $lblBarcodeFixed.Text = "barcode"
            $lblBarcodeFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblBarcodeFixed.Size = New-Object System.Drawing.Size(60, 20)
            $lblBarcodeFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblBarcodeFixed)
            
            $cmbSourceBarcode = New-Object System.Windows.Forms.ComboBox
            $cmbSourceBarcode.Location = New-Object System.Drawing.Point(150, $yPos)
            $cmbSourceBarcode.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceBarcode.DropDownStyle = "DropDownList"
            $cmbSourceBarcode.Name = "cmbSourceBarcodeItems"
            $tab.Controls.Add($cmbSourceBarcode)
            $yPos += 30
            
            # RCN mapping: rcn ← ISBN-10
            $lblRCN = New-Object System.Windows.Forms.Label
            $lblRCN.Text = "RCN:"
            $lblRCN.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblRCN.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblRCN)
            
            $lblRCNFixed = New-Object System.Windows.Forms.Label
            $lblRCNFixed.Text = "rcn"
            $lblRCNFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblRCNFixed.Size = New-Object System.Drawing.Size(60, 20)
            $lblRCNFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblRCNFixed)
            
            $cmbSourceISBN = New-Object System.Windows.Forms.ComboBox
            $cmbSourceISBN.Location = New-Object System.Drawing.Point(150, $yPos)
            $cmbSourceISBN.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceISBN.DropDownStyle = "DropDownList"
            $cmbSourceISBN.Name = "cmbSourceISBNItems"
            $tab.Controls.Add($cmbSourceISBN)
            $yPos += 30
            
            # Title (display only)
            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Text = "Title:"
            $lblTitle.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblTitle.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblTitle)
            
            $lblTitleFixed = New-Object System.Windows.Forms.Label
            $lblTitleFixed.Text = "(display)"
            $lblTitleFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblTitleFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblTitleFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $lblTitleFixed.ForeColor = [System.Drawing.Color]::Gray
            $tab.Controls.Add($lblTitleFixed)
            
            $cmbSourceTitle = New-Object System.Windows.Forms.ComboBox
            $cmbSourceTitle.Location = New-Object System.Drawing.Point(160, $yPos)
            $cmbSourceTitle.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceTitle.DropDownStyle = "DropDownList"
            $cmbSourceTitle.Name = "cmbSourceTitleItems"
            $tab.Controls.Add($cmbSourceTitle)
            $yPos += 30
            
            # Publisher (display only)
            $lblPublisher = New-Object System.Windows.Forms.Label
            $lblPublisher.Text = "Publisher:"
            $lblPublisher.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblPublisher.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblPublisher)
            
            $lblPublisherFixed = New-Object System.Windows.Forms.Label
            $lblPublisherFixed.Text = "(display)"
            $lblPublisherFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblPublisherFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblPublisherFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $lblPublisherFixed.ForeColor = [System.Drawing.Color]::Gray
            $tab.Controls.Add($lblPublisherFixed)
            
            $cmbSourcePublisher = New-Object System.Windows.Forms.ComboBox
            $cmbSourcePublisher.Location = New-Object System.Drawing.Point(160, $yPos)
            $cmbSourcePublisher.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourcePublisher.DropDownStyle = "DropDownList"
            $cmbSourcePublisher.Name = "cmbSourcePublisherItems"
            $tab.Controls.Add($cmbSourcePublisher)
        }
        
        "Users" {
            # Barcode mapping: barcode ← Username
            $lblBarcode = New-Object System.Windows.Forms.Label
            $lblBarcode.Text = "Barcode:"
            $lblBarcode.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblBarcode.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblBarcode)
            
            $lblBarcodeFixed = New-Object System.Windows.Forms.Label
            $lblBarcodeFixed.Text = "barcode"
            $lblBarcodeFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblBarcodeFixed.Size = New-Object System.Drawing.Size(60, 20)
            $lblBarcodeFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblBarcodeFixed)
            
            $cmbSourceUsername = New-Object System.Windows.Forms.ComboBox
            $cmbSourceUsername.Location = New-Object System.Drawing.Point(150, $yPos)
            $cmbSourceUsername.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceUsername.DropDownStyle = "DropDownList"
            $cmbSourceUsername.Name = "cmbSourceUsernameUsers"
            $tab.Controls.Add($cmbSourceUsername)
            $yPos += 30
            
            # Prison ID mapping: prison_id ← Prison Number
            $lblPrisonID = New-Object System.Windows.Forms.Label
            $lblPrisonID.Text = "Prison ID:"
            $lblPrisonID.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblPrisonID.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblPrisonID)
            
            $lblPrisonIDFixed = New-Object System.Windows.Forms.Label
            $lblPrisonIDFixed.Text = "prison_id"
            $lblPrisonIDFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblPrisonIDFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblPrisonIDFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblPrisonIDFixed)
            
            $cmbSourcePrisonNumber = New-Object System.Windows.Forms.ComboBox
            $cmbSourcePrisonNumber.Location = New-Object System.Drawing.Point(160, $yPos)
            $cmbSourcePrisonNumber.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourcePrisonNumber.DropDownStyle = "DropDownList"
            $cmbSourcePrisonNumber.Name = "cmbSourcePrisonNumberUsers"
            $tab.Controls.Add($cmbSourcePrisonNumber)
            $yPos += 30
            
            # Forename mapping: forename ← Full Name
            $lblForename = New-Object System.Windows.Forms.Label
            $lblForename.Text = "Forename:"
            $lblForename.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblForename.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblForename)
            
            $lblForenameFixed = New-Object System.Windows.Forms.Label
            $lblForenameFixed.Text = "forename"
            $lblForenameFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblForenameFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblForenameFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblForenameFixed)
            
            $cmbSourceFullName1 = New-Object System.Windows.Forms.ComboBox
            $cmbSourceFullName1.Location = New-Object System.Drawing.Point(160, $yPos)
            $cmbSourceFullName1.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceFullName1.DropDownStyle = "DropDownList"
            $cmbSourceFullName1.Name = "cmbSourceFullName1Users"
            $tab.Controls.Add($cmbSourceFullName1)
            $yPos += 30
            
            # Surname mapping: surname ← Full Name
            $lblSurname = New-Object System.Windows.Forms.Label
            $lblSurname.Text = "Surname:"
            $lblSurname.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblSurname.Size = New-Object System.Drawing.Size(60, 20)
            $tab.Controls.Add($lblSurname)
            
            $lblSurnameFixed = New-Object System.Windows.Forms.Label
            $lblSurnameFixed.Text = "surname"
            $lblSurnameFixed.Location = New-Object System.Drawing.Point(85, $yPos)
            $lblSurnameFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblSurnameFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblSurnameFixed)
            
            $cmbSourceFullName2 = New-Object System.Windows.Forms.ComboBox
            $cmbSourceFullName2.Location = New-Object System.Drawing.Point(160, $yPos)
            $cmbSourceFullName2.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceFullName2.DropDownStyle = "DropDownList"
            $cmbSourceFullName2.Name = "cmbSourceFullName2Users"
            $tab.Controls.Add($cmbSourceFullName2)
        }
        
        "Loans" {
            # Item Barcode mapping: itemBarcode ← Borrowed
            $lblItemBarcode = New-Object System.Windows.Forms.Label
            $lblItemBarcode.Text = "Item Barcode:"
            $lblItemBarcode.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblItemBarcode.Size = New-Object System.Drawing.Size(80, 20)
            $tab.Controls.Add($lblItemBarcode)
            
            $lblItemBarcodeFixed = New-Object System.Windows.Forms.Label
            $lblItemBarcodeFixed.Text = "itemBarcode"
            $lblItemBarcodeFixed.Location = New-Object System.Drawing.Point(105, $yPos)
            $lblItemBarcodeFixed.Size = New-Object System.Drawing.Size(90, 20)
            $lblItemBarcodeFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblItemBarcodeFixed)
            
            $cmbSourceItemBarcode = New-Object System.Windows.Forms.ComboBox
            $cmbSourceItemBarcode.Location = New-Object System.Drawing.Point(200, $yPos)
            $cmbSourceItemBarcode.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceItemBarcode.DropDownStyle = "DropDownList"
            $cmbSourceItemBarcode.Name = "cmbSourceItemBarcodeLoans"
            $tab.Controls.Add($cmbSourceItemBarcode)
            $yPos += 30
            
            # User Barcode mapping: borrbarcode ← Borrowed By
            $lblUserBarcode = New-Object System.Windows.Forms.Label
            $lblUserBarcode.Text = "User Barcode:"
            $lblUserBarcode.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblUserBarcode.Size = New-Object System.Drawing.Size(80, 20)
            $tab.Controls.Add($lblUserBarcode)
            
            $lblUserBarcodeFixed = New-Object System.Windows.Forms.Label
            $lblUserBarcodeFixed.Text = "borrbarcode"
            $lblUserBarcodeFixed.Location = New-Object System.Drawing.Point(105, $yPos)
            $lblUserBarcodeFixed.Size = New-Object System.Drawing.Size(90, 20)
            $lblUserBarcodeFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblUserBarcodeFixed)
            
            $cmbSourceUserBarcode = New-Object System.Windows.Forms.ComboBox
            $cmbSourceUserBarcode.Location = New-Object System.Drawing.Point(200, $yPos)
            $cmbSourceUserBarcode.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceUserBarcode.DropDownStyle = "DropDownList"
            $cmbSourceUserBarcode.Name = "cmbSourceUserBarcodeLoans"
            $tab.Controls.Add($cmbSourceUserBarcode)
            $yPos += 30
            
            # Issue Date mapping: issuedate ← Borrowed-2
            $lblIssueDate = New-Object System.Windows.Forms.Label
            $lblIssueDate.Text = "Issue Date:"
            $lblIssueDate.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblIssueDate.Size = New-Object System.Drawing.Size(80, 20)
            $tab.Controls.Add($lblIssueDate)
            
            $lblIssueDateFixed = New-Object System.Windows.Forms.Label
            $lblIssueDateFixed.Text = "issuedate"
            $lblIssueDateFixed.Location = New-Object System.Drawing.Point(105, $yPos)
            $lblIssueDateFixed.Size = New-Object System.Drawing.Size(80, 20)
            $lblIssueDateFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblIssueDateFixed)
            
            $cmbSourceIssueDate = New-Object System.Windows.Forms.ComboBox
            $cmbSourceIssueDate.Location = New-Object System.Drawing.Point(190, $yPos)
            $cmbSourceIssueDate.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceIssueDate.DropDownStyle = "DropDownList"
            $cmbSourceIssueDate.Name = "cmbSourceIssueDateLoans"
            $tab.Controls.Add($cmbSourceIssueDate)
            $yPos += 30
            
            # Due Date mapping: duedate ← Return Date
            $lblDueDate = New-Object System.Windows.Forms.Label
            $lblDueDate.Text = "Due Date:"
            $lblDueDate.Location = New-Object System.Drawing.Point(20, $yPos)
            $lblDueDate.Size = New-Object System.Drawing.Size(80, 20)
            $tab.Controls.Add($lblDueDate)
            
            $lblDueDateFixed = New-Object System.Windows.Forms.Label
            $lblDueDateFixed.Text = "duedate"
            $lblDueDateFixed.Location = New-Object System.Drawing.Point(105, $yPos)
            $lblDueDateFixed.Size = New-Object System.Drawing.Size(70, 20)
            $lblDueDateFixed.Font = New-Object System.Drawing.Font("Consolas", 8)
            $tab.Controls.Add($lblDueDateFixed)
            
            $cmbSourceReturnDate = New-Object System.Windows.Forms.ComboBox
            $cmbSourceReturnDate.Location = New-Object System.Drawing.Point(180, $yPos)
            $cmbSourceReturnDate.Size = New-Object System.Drawing.Size(150, 20)
            $cmbSourceReturnDate.DropDownStyle = "DropDownList"
            $cmbSourceReturnDate.Name = "cmbSourceReturnDateLoans"
            $tab.Controls.Add($cmbSourceReturnDate)
        }
    }
    
    $yPos += 40
    
    # Sample button
    $btnSample = New-Object System.Windows.Forms.Button
    $btnSample.Text = "Run Sampling"
    $btnSample.Location = New-Object System.Drawing.Point(20, $yPos)
    $btnSample.Size = New-Object System.Drawing.Size(100, 30)
    $btnSample.Name = "btnSample$DataType"
    $tab.Controls.Add($btnSample)
    
    $yPos += 40
    
    # Results text box - FIXED: Better scrollbars and larger size
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Location = New-Object System.Drawing.Point(20, $yPos)
    $txtResults.Size = New-Object System.Drawing.Size(800, 180)
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = "Both"
    $txtResults.ReadOnly = $true
    $txtResults.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtResults.Name = "txtResults$DataType"
    $txtResults.WordWrap = $false
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

# Function to run all sampling comparisons
function Run-AllSampling {
    $script:combinedReport = ""
    $progressBar.Value = 0
    $btnSaveReport.Enabled = $false
    
    $dataTypes = @("Items", "Users", "Loans")
    $reports = @()
    $summaryResults = @()
    
    for ($i = 0; $i -lt $dataTypes.Count; $i++) {
        $dataType = $dataTypes[$i]
        $lblStatus.Text = "Processing $dataType sampling..."
        $form.Refresh()
        
        $report = Compare-DataSampling -DataType $dataType -IsGlobalRun $true
        if ($report) {
            $reports += $report
            $txtResults = $form.Controls.Find("txtResults$datatype", $true)[0]
            $txtResults.Text = $report
            
            # Extract summary information for popup
            $summaryInfo = @{
                DataType = $dataType
                Status = "PROCESSED"
                SampleSize = 5
                Assessment = "COMPLETED"
            }
            
            # Parse the report for assessment
            if ($report -match "\[PASS\]") {
                $summaryInfo.Assessment = "PASS"
            } elseif ($report -match "\[FAIL\]") {
                $summaryInfo.Assessment = "FAIL"
            } elseif ($report -match "\[WARN\]") {
                $summaryInfo.Assessment = "WARNING"
            }
            
            $summaryResults += $summaryInfo
        } else {
            $reports += "SKIPPED: $dataType - Missing files or configuration"
            $summaryResults += @{
                DataType = $dataType
                Status = "SKIPPED"
                SampleSize = 0
                Assessment = "SKIPPED"
            }
        }
        
        $progressBar.Value = $i + 1
        $form.Refresh()
    }
    
    # Show popup summary
    Show-SamplingSummary -SummaryResults $summaryResults
    
    # Create combined report
    $script:combinedReport = @()
    $script:combinedReport += "LIBRARY DATA ASSURANCE - SAMPLING REPORT"
    $script:combinedReport += "=" * 80
    $script:combinedReport += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $script:combinedReport += "Method: Random Sampling Analysis"
    $script:combinedReport += ""
    $script:combinedReport += "EXECUTIVE SUMMARY:"
    
    $issues = @()
    foreach ($report in $reports) {
        if ($report -like "*SKIPPED*") {
            $issues += $report
            continue
        }
        
        if ($report -match "\[FAIL\]") {
            $dataTypeName = ($report -split "`n")[1] -replace " Sampling Report", ""
            $issues += "FAILURES detected in $dataTypeName"
        } elseif ($report -match "\[WARN\]") {
            $dataTypeName = ($report -split "`n")[1] -replace " Sampling Report", ""
            $issues += "WARNINGS detected in $dataTypeName"
        }
    }
    
    if ($issues.Count -eq 0) {
        $script:combinedReport += "  Status: [SUCCESS] All sampling tests passed"
    } else {
        $script:combinedReport += "  Status: [REVIEW] Issues detected in sampling"
        $script:combinedReport += "  Issues Found:"
        foreach ($issue in $issues) {
            $script:combinedReport += "    - $issue"
        }
    }
    
    $script:combinedReport += ""
    $script:combinedReport += "DETAILED SAMPLING REPORTS:"
    $script:combinedReport += "=" * 80
    
    foreach ($report in $reports) {
        $script:combinedReport += ""
        $script:combinedReport += $report
        $script:combinedReport += ""
    }
    
    $script:combinedReport = $script:combinedReport -join "`r`n"
    
    $lblStatus.Text = "All sampling completed"
    $btnSaveReport.Enabled = $true
}

# Function to show sampling summary
function Show-SamplingSummary {
    param($SummaryResults)
    
    # Create summary popup form
    $summaryForm = New-Object System.Windows.Forms.Form
    $summaryForm.Text = "Sampling Analysis Summary"
    $summaryForm.Size = New-Object System.Drawing.Size(500, 400)
    $summaryForm.StartPosition = "CenterParent"
    $summaryForm.FormBorderStyle = "FixedDialog"
    $summaryForm.MaximizeBox = $false
    $summaryForm.MinimizeBox = $false
    
    # Title label
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "SAMPLING ANALYSIS SUMMARY"
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
    
    $hasWarnings = $false
    $hasFailures = $false
    $hasSkipped = $false
    
    foreach ($result in $SummaryResults) {
        $statusIcon = switch ($result.Assessment) {
            "PASS" { "[PASS]" }
            "FAIL" { "[FAIL]"; $hasFailures = $true }
            "WARNING" { "[WARN]"; $hasWarnings = $true }
            "SKIPPED" { "[SKIP]"; $hasSkipped = $true }
            default { "[DONE]" }
        }
        
        if ($result.Status -eq "SKIPPED") {
            $summaryContent += "$($result.DataType): $statusIcon Missing files/config"
        } else {
            $summaryContent += "$($result.DataType): $statusIcon Sample size $($result.SampleSize) completed"
        }
    }
    
    $summaryContent += ""
    
    # Overall assessment
    if ($hasFailures) {
        $summaryContent += "OVERALL STATUS: [CRITICAL] Data integrity failures detected"
    } elseif ($hasWarnings) {
        $summaryContent += "OVERALL STATUS: [REVIEW] Data discrepancies detected"
    } elseif ($hasSkipped) {
        $summaryContent += "OVERALL STATUS: [PARTIAL] Some data types not analyzed"
    } else {
        $summaryContent += "OVERALL STATUS: [SUCCESS] All samples verified"
    }
    
    $summaryContent += ""
    $summaryContent += "Legend:"
    $summaryContent += "  [PASS] = All sampled items verified"
    $summaryContent += "  [WARN] = Some discrepancies found"
    $summaryContent += "  [FAIL] = Missing or incorrect data"
    $summaryContent += "  [SKIP] = Files not configured"
    $summaryContent += ""
    $summaryContent += "Click individual tabs for detailed sampling reports."
    
    $txtSummary.Text = $summaryContent -join "`r`n"
    
    # Show the summary dialog
    $summaryForm.ShowDialog()
}

# Save report button
$btnSaveReport = New-Object System.Windows.Forms.Button
$btnSaveReport.Text = "Save Report"
$btnSaveReport.Location = New-Object System.Drawing.Point(670, 65)
$btnSaveReport.Size = New-Object System.Drawing.Size(80, 25)
$btnSaveReport.Enabled = $false
$autoDetectPanel.Controls.Add($btnSaveReport)

# Function to save combined report
function Save-CombinedReport {
    if ([string]::IsNullOrEmpty($script:combinedReport)) {
        [System.Windows.Forms.MessageBox]::Show("No report to save. Please run sampling first.", "No Report", "OK", "Warning")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Sampling Report"
    $saveFileDialog.FileName = "Library_Sampling_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
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
    $btnSample = $form.Controls.Find("btnSample$dataType", $true)[0]
    
    $btnOriginal.Add_Click({
        $dataType = $this.Name.Substring(11)
        $txtOriginal = $form.Controls.Find("txtOriginal$dataType", $true)[0]
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Delimited files (*.csv;*.tsv)|*.csv;*.tsv|CSV files (*.csv)|*.csv|TSV files (*.tsv)|*.tsv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Original Delimited File"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            $txtOriginal.Text = $openFileDialog.FileName
            
            # Load headers for all source dropdowns for this data type
            Load-HeadersForDataType -DataType $dataType -FilePath $openFileDialog.FileName -FileType "Source"
        }
    })
    
    $btnMigrated.Add_Click({
        $dataType = $this.Name.Substring(11)
        $txtMigrated = $form.Controls.Find("txtMigrated$dataType", $true)[0]
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Delimited files (*.csv;*.tsv)|*.csv;*.tsv|CSV files (*.csv)|*.csv|TSV files (*.tsv)|*.tsv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Migrated Delimited File"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            $txtMigrated.Text = $openFileDialog.FileName
        }
    })
    
    $btnSample.Add_Click({
        $dataType = $this.Name.Substring(9)
        Compare-DataSampling -DataType $dataType
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

$btnRunAll.Add_Click({
    Run-AllSampling
})

$btnSaveReport.Add_Click({
    Save-CombinedReport
})

# Show the form
$form.ShowDialog()
