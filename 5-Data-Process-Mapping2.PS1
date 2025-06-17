# Simple Library Data Transformation Tool with Data Validation

# Load required assemblies FIRST
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Minimize the PowerShell command prompt
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();
    }
"@

try {
    $consolePtr = [Win32]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Win32]::ShowWindow($consolePtr, 6)  # 6 = Minimize
    }
} catch {
    Write-Warning "Could not minimize console window: $_"
}

# Global variables
$global:ErrorCells = @{}
$global:Mapping = $null
$global:ProcessedData = $null
$global:RawData = $null
$global:ShowOnlyMappedFields = $false
$global:UseMappingOrder = $false
$global:OriginalColumnOrder = @()
$global:MappingColumnOrder = @()
$global:LogData = @()
$global:SourceToNewFieldMap = @{}
$global:DataValidationResults = @()
$global:OriginalRowCount = 0
$global:ProcessedRowCount = 0
$global:ProgressForm = $null
$global:ProgressBar = $null
$global:ProgressLabel = $null

# Load Transformation Functions
$transformFunctionsFile = ".\Transform-Functions.ps1"
if (Test-Path $transformFunctionsFile) {
    try {
        . $transformFunctionsFile
        Write-Host "Loaded transformation functions from $transformFunctionsFile"
    } catch {
        Write-Warning "Error loading transformation functions: $_"
        $global:TransformFunctions = @{}
    }
} else {
    Write-Warning "Transformation functions file not found. Continuing without external transforms."
    $global:TransformFunctions = @{}
}

# Progress bar functions
function Show-ProgressBar($title, $maxValue) {
    $global:ProgressForm = New-Object System.Windows.Forms.Form
    $global:ProgressForm.Text = $title
    $global:ProgressForm.Size = New-Object System.Drawing.Size(400, 120)
    $global:ProgressForm.StartPosition = "CenterParent"
    $global:ProgressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $global:ProgressForm.MaximizeBox = $false
    $global:ProgressForm.MinimizeBox = $false
    $global:ProgressForm.ControlBox = $false
    
    $global:ProgressLabel = New-Object System.Windows.Forms.Label
    $global:ProgressLabel.Text = "Initializing..."
    $global:ProgressLabel.Size = New-Object System.Drawing.Size(360, 20)
    $global:ProgressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $global:ProgressForm.Controls.Add($global:ProgressLabel)
    
    $global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $global:ProgressBar.Size = New-Object System.Drawing.Size(360, 25)
    $global:ProgressBar.Location = New-Object System.Drawing.Point(20, 50)
    $global:ProgressBar.Minimum = 0
    $global:ProgressBar.Maximum = $maxValue
    $global:ProgressBar.Value = 0
    $global:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $global:ProgressForm.Controls.Add($global:ProgressBar)
    
    $global:ProgressForm.Show()
    $global:ProgressForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-ProgressBar($value, $text) {
    if ($global:ProgressBar -and $global:ProgressLabel) {
        $global:ProgressBar.Value = [Math]::Min($value, $global:ProgressBar.Maximum)
        $global:ProgressLabel.Text = $text
        $global:ProgressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Hide-ProgressBar {
    if ($global:ProgressForm) {
        $global:ProgressForm.Close()
        $global:ProgressForm.Dispose()
        $global:ProgressForm = $null
        $global:ProgressBar = $null
        $global:ProgressLabel = $null
    }
}

# Optimized Data Validation Functions
function Initialize-DataValidation($inputData) {
    $global:DataValidationResults = @()
    $global:OriginalRowCount = if ($inputData) { $inputData.Count } else { 0 }
    
    Update-ProgressBar 0 "Creating data fingerprints..."
    
    $fingerprints = @()
    for ($i = 0; $i -lt $inputData.Count; $i++) {
        if ($i % 100 -eq 0) {
            Update-ProgressBar ($i * 10 / $inputData.Count) "Creating fingerprint $($i + 1) of $($inputData.Count)..."
        }
        
        $row = $inputData[$i]
        # Simplified fingerprint for better performance
        $fingerprint = ($row.PSObject.Properties | ForEach-Object { "$($_.Name):$($_.Value)" }) -join "|"
        
        $fingerprints += @{
            Index = $i
            Fingerprint = $fingerprint
            OriginalRow = $row
        }
    }
    return $fingerprints
}

function Validate-DataIntegrity($originalFingerprints, $processedData) {
    Update-ProgressBar 10 "Starting data integrity validation..."
    
    $validationResults = @()
    $global:ProcessedRowCount = if ($processedData) { $processedData.Count } else { 0 }
    
    Update-ProgressBar 15 "Checking row count changes..."
    
    # Check row count changes
    if ($global:OriginalRowCount -ne $global:ProcessedRowCount) {
        $countChange = @{
            Type = "RowCountChange"
            OriginalCount = $global:OriginalRowCount
            ProcessedCount = $global:ProcessedRowCount
            Difference = $global:ProcessedRowCount - $global:OriginalRowCount
            Reason = if ($global:ProcessedRowCount -gt $global:OriginalRowCount) { 
                "Rows were added during processing" 
            } else { 
                "Rows were removed during processing" 
            }
        }
        $validationResults += $countChange
    }
    
    Update-ProgressBar 20 "Matching processed rows to originals..."
    
    # Track which original rows are represented in processed data
    $processedRowTracker = @{}
    $unmatchedProcessedRows = @()
    
    # For each processed row, try to find its original (optimized)
    for ($i = 0; $i -lt $processedData.Count; $i++) {
        if ($i % 50 -eq 0) {
            $progress = 20 + ($i * 40 / $processedData.Count)
            Update-ProgressBar $progress "Matching row $($i + 1) of $($processedData.Count)..."
        }
        
        $processedRow = $processedData[$i]
        $matchFound = $false
        
        # Quick fingerprint check first
        $processedFingerprint = ($processedRow.PSObject.Properties | ForEach-Object { "$($_.Name):$($_.Value)" }) -join "|"
        
        for ($j = 0; $j -lt $originalFingerprints.Count; $j++) {
            $originalFingerprint = $originalFingerprints[$j]
            
            # Fast fingerprint comparison first
            if ($originalFingerprint.Fingerprint -eq $processedFingerprint) {
                $processedRowTracker[$j] = $i
                $matchFound = $true
                break
            }
        }
        
        # If no exact match, try intelligent matching (slower but more thorough)
        if (-not $matchFound) {
            for ($j = 0; $j -lt $originalFingerprints.Count; $j++) {
                if ($processedRowTracker.ContainsKey($j)) { continue }  # Skip already matched
                
                $originalFingerprint = $originalFingerprints[$j]
                
                if (Test-RowMatch $originalFingerprint.OriginalRow $processedRow) {
                    $processedRowTracker[$j] = $i
                    $matchFound = $true
                    break
                }
            }
        }
        
        if (-not $matchFound) {
            $unmatchedProcessedRows += @{
                ProcessedIndex = $i
                Row = $processedRow
                Reason = "No corresponding original row found - may be a generated/duplicate row"
            }
        }
    }
    
    Update-ProgressBar 60 "Identifying missing rows..."
    
    # Find original rows that don't appear in processed data
    $missingRows = @()
    for ($i = 0; $i -lt $originalFingerprints.Count; $i++) {
        if ($i % 100 -eq 0) {
            $progress = 60 + ($i * 20 / $originalFingerprints.Count)
            Update-ProgressBar $progress "Checking original row $($i + 1) of $($originalFingerprints.Count)..."
        }
        
        if (-not $processedRowTracker.ContainsKey($i)) {
            $missingRows += @{
                OriginalIndex = $i
                Row = $originalFingerprints[$i].OriginalRow
                Reason = "Original row not found in processed data - may have been filtered out"
            }
        }
    }
    
    Update-ProgressBar 80 "Compiling validation results..."
    
    # Add missing rows to validation results
    foreach ($missing in $missingRows) {
        $validationResults += @{
            Type = "MissingRow"
            OriginalIndex = $missing.OriginalIndex + 1
            Reason = $missing.Reason
            RowData = $missing.Row
        }
    }
    
    # Add unmatched processed rows to validation results
    foreach ($unmatched in $unmatchedProcessedRows) {
        $validationResults += @{
            Type = "UnmatchedProcessedRow"
            ProcessedIndex = $unmatched.ProcessedIndex + 1
            Reason = $unmatched.Reason
            RowData = $unmatched.Row
        }
    }
    
    Update-ProgressBar 90 "Finalizing validation summary..."
    
    # Summary validation result
    $summary = @{
        Type = "ValidationSummary"
        OriginalRowCount = $global:OriginalRowCount
        ProcessedRowCount = $global:ProcessedRowCount
        MissingRowCount = $missingRows.Count
        UnmatchedRowCount = $unmatchedProcessedRows.Count
        DataIntegrityStatus = if ($validationResults.Count -eq 1) { "PASSED" } else { "ISSUES_DETECTED" }
    }
    $validationResults += $summary
    
    Update-ProgressBar 95 "Validation complete!"
    
    return $validationResults
}

function Test-RowMatch($originalRow, $processedRow) {
    $originalProps = $originalRow.PSObject.Properties
    $matchCount = 0
    $totalChecked = 0
    
    foreach ($origProp in $originalProps) {
        $origFieldName = $origProp.Name
        $origValue = $origProp.Value
        
        if ([string]::IsNullOrEmpty($origValue)) {
            continue
        }
        
        $totalChecked++
        
        # Check if this field exists directly in processed data
        if ($processedRow.PSObject.Properties.Name -contains $origFieldName) {
            $processedValue = $processedRow.$origFieldName
            if ($origValue -eq $processedValue) {
                $matchCount++
                continue
            }
        }
        
        # Check if this field was renamed via mapping
        if ($global:SourceToNewFieldMap.ContainsKey($origFieldName)) {
            foreach ($newFieldName in $global:SourceToNewFieldMap[$origFieldName]) {
                if ($processedRow.PSObject.Properties.Name -contains $newFieldName) {
                    $processedValue = $processedRow.$newFieldName
                    if (Test-ValueSimilarity $origValue $processedValue) {
                        $matchCount++
                        break
                    }
                }
            }
        }
    }
    
    if ($totalChecked -eq 0) {
        return $false
    }
    
    $matchPercentage = ($matchCount / $totalChecked) * 100
    return $matchPercentage -ge 60
}

function Test-ValueSimilarity($originalValue, $processedValue) {
    if ($originalValue -eq $processedValue) {
        return $true
    }
    
    $origStr = "$originalValue".Trim()
    $procStr = "$processedValue".Trim()
    
    if ($origStr -eq $procStr) {
        return $true
    }
    
    # Date format changes - using simpler regex pattern
    if ($origStr -match '\d{2}/\d{2}/\d{4}' -and $procStr -match '\d{1,2}/\d{1,2}/\d{4}') {
        try {
            $origDate = [datetime]::Parse($origStr)
            $procDate = [datetime]::Parse($procStr)
            return $origDate -eq $procDate
        } catch { }
    }
    
    # Numeric conversions
    $origNum = $null
    $procNum = $null
    if ([decimal]::TryParse($origStr, [ref]$origNum) -and [decimal]::TryParse($procStr, [ref]$procNum)) {
        return $origNum -eq $procNum
    }
    
    # Case insensitive
    if ($origStr.ToLower() -eq $procStr.ToLower()) {
        return $true
    }
    
    return $false
}

function Export-DataValidationReport($validationResults, $filePath) {
    try {
        $reportLines = @()
        $reportLines += "Data Validation Report - Generated $(Get-Date)"
        $reportLines += "=" * 60
        $reportLines += ""
        
        $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary) {
            $reportLines += "VALIDATION SUMMARY"
            $reportLines += "-" * 20
            $reportLines += "Original Row Count: $($summary.OriginalRowCount)"
            $reportLines += "Processed Row Count: $($summary.ProcessedRowCount)"
            $reportLines += "Missing Rows: $($summary.MissingRowCount)"
            $reportLines += "Unmatched Processed Rows: $($summary.UnmatchedRowCount)"
            $reportLines += "Data Integrity Status: $($summary.DataIntegrityStatus)"
            $reportLines += ""
        }
        
        $countChanges = $validationResults | Where-Object { $_.Type -eq "RowCountChange" }
        if ($countChanges) {
            $reportLines += "ROW COUNT CHANGES"
            $reportLines += "-" * 20
            foreach ($change in $countChanges) {
                $reportLines += "Original: $($change.OriginalCount), Processed: $($change.ProcessedCount)"
                $reportLines += "Difference: $($change.Difference) rows"
                $reportLines += "Reason: $($change.Reason)"
                $reportLines += ""
            }
        }
        
        $missingRows = $validationResults | Where-Object { $_.Type -eq "MissingRow" }
        if ($missingRows) {
            $reportLines += "MISSING ROWS"
            $reportLines += "-" * 15
            foreach ($missing in $missingRows) {
                $reportLines += "Original Row $($missing.OriginalIndex): $($missing.Reason)"
                $properties = $missing.RowData.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }
                $reportLines += "Data: $($properties -join '; ')"
                $reportLines += ""
            }
        }
        
        $unmatchedRows = $validationResults | Where-Object { $_.Type -eq "UnmatchedProcessedRow" }
        if ($unmatchedRows) {
            $reportLines += "UNMATCHED PROCESSED ROWS"
            $reportLines += "-" * 25
            foreach ($unmatched in $unmatchedRows) {
                $reportLines += "Processed Row $($unmatched.ProcessedIndex): $($unmatched.Reason)"
                $properties = $unmatched.RowData.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }
                $reportLines += "Data: $($properties -join '; ')"
                $reportLines += ""
            }
        }
        
        $reportContent = $reportLines -join "`r`n"
        $reportContent | Out-File -FilePath $filePath -Encoding UTF8
        return $true
    } catch {
        Write-Error "Failed to export validation report: $_"
        return $false
    }
}

# Data loading functions
function Load-Mapping($path) {
    try {
        $mapping = Import-Csv -Path $path
        
        $fieldCount = @{}
        foreach ($m in $mapping) {
            if (-not [string]::IsNullOrEmpty($m.SourceField)) {
                if (-not $fieldCount.ContainsKey($m.SourceField)) {
                    $fieldCount[$m.SourceField] = 0
                }
                $fieldCount[$m.SourceField]++
            }
        }
        
        foreach ($field in $fieldCount.Keys) {
            if ($fieldCount[$field] -gt 1) {
                Write-Host "Field '$field' has $($fieldCount[$field]) mappings"
            }
        }
        
        return $mapping
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading mapping file: $_")
        return $null
    }
}

function Load-Data($path) {
    try {
        $rawContent = Get-Content -Path $path -Raw
        $headerLine = $rawContent.Split("`n")[0].Trim()
        
        $quotedColumnPattern = '(?:^|,)(?:"([^"]*(?:""[^"]*)*)"|([^,"]*))'
        $matches = [regex]::Matches($headerLine, $quotedColumnPattern)
        $global:OriginalColumnOrder = @()
        
        foreach ($match in $matches) {
            $columnName = if ($match.Groups[1].Success) { 
                $match.Groups[1].Value.Replace('""', '"') 
            } else { 
                $match.Groups[2].Value 
            }
            $global:OriginalColumnOrder += $columnName
        }
        
        Write-Host "Original CSV columns: $($global:OriginalColumnOrder -join ', ')"
        
        $data = Import-Csv -Path $path
        return $data
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading data file: $_")
        return $null
    }
}

# Process data according to mapping rules
function Process-Data($data, $mapping) {
    # Calculate total steps for progress tracking
    $totalSteps = $data.Count + ($data.Count * 0.1) + 100  # Processing + validation overhead
    Show-ProgressBar "Processing Data" $totalSteps
    
    $result = @()
    $log = @()
    $global:ErrorCells.Clear()
    $global:RawData = $data
	
    Update-ProgressBar 0 "Initializing data validation..."
    
    # Initialize data validation
    $originalFingerprints = Initialize-DataValidation $data
	
    Update-ProgressBar 10 "Building field mappings..."
    
	# Build field mappings
	$global:SourceToNewFieldMap.Clear()
	foreach ($m in $mapping) {
		if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
            if (-not $global:SourceToNewFieldMap.ContainsKey($m.SourceField)) {
                $global:SourceToNewFieldMap[$m.SourceField] = @()
            }
            $global:SourceToNewFieldMap[$m.SourceField] += $m.NewField
		}
	}
    
    # Capture mapping column order
    if ($mapping.Count -gt 0) {
        $global:MappingColumnOrder = @()
        foreach ($m in $mapping) {
            if (-not [string]::IsNullOrEmpty($m.NewField)) {
                if ($global:MappingColumnOrder -notcontains $m.NewField) {
                    $global:MappingColumnOrder += $m.NewField
                }
            } elseif (-not [string]::IsNullOrEmpty($m.SourceField)) {
                if ($global:MappingColumnOrder -notcontains $m.SourceField) {
                    $global:MappingColumnOrder += $m.SourceField
                }
            }
        }
    }

    Update-ProgressBar 15 "Creating mapping lookup tables..."

    # Create lookup dictionary from mapping
    $mapLookup = @{}
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField)) {
            if (-not $mapLookup.ContainsKey($m.SourceField)) {
                $mapLookup[$m.SourceField] = @()
            }
            $mapLookup[$m.SourceField] += $m
        }
    }
    
    $mandatoryFields = @{}
    foreach ($m in $mapping) {
        if ($m.Mandatory -eq "Y") {
            $mandatoryFields[$m.SourceField] = $true
        }
    }

    Update-ProgressBar 20 "Processing data rows..."

    # Process each row
    for ($i = 0; $i -lt $data.Count; $i++) {
        # Update progress every 10 rows for better performance
        if ($i % 10 -eq 0) {
            $progress = 20 + ($i * 60 / $data.Count)
            Update-ProgressBar $progress "Processing row $($i + 1) of $($data.Count)..."
        }
        
        $row = $data[$i]
        $newRow = @{}
        $processedFields = @{}
        
        foreach ($prop in $row.PSObject.Properties) {
            $sourceField = $prop.Name
            $value = $prop.Value
            
            if ($mapLookup.ContainsKey($sourceField)) {
                foreach ($map in $mapLookup[$sourceField]) {
                    $dataType = $map.DataType
                    $mandatory = $map.Mandatory -eq "Y"
                    $doTransform = $map.Transformation -eq "Y"
                    $transformFunction = $map.TransformFunction
                    $error = $null
                    
                    $processedFields[$sourceField] = $true
                    $originalValue = $value
                    $mappedValue = $value
                    
                    # Apply transformation
                    if ($doTransform -and $transformFunction -and $global:TransformFunctions.ContainsKey($transformFunction)) {
                        try {
                            $mappedValue = & $global:TransformFunctions[$transformFunction] $mappedValue
                        } catch {
                            $error = "Error in transformation: $_"
                        }
                    }
                    
                    # Mandatory check
                    if ($mandatory) {
                        if ([string]::IsNullOrEmpty($originalValue)) {
                            $error = "Mandatory field missing (original value)"
                        }
                        elseif ([string]::IsNullOrEmpty($mappedValue)) {
                            $error = "Mandatory field empty after transformation"
                        }
                    }
                    
                    # Type conversion
                    if ($dataType -and $mappedValue -ne $null -and $mappedValue -ne "") {
                        try {
                            switch ($dataType.ToLower()) {
                                "int" { $mappedValue = [int]$mappedValue }
                                "decimal" { $mappedValue = [decimal]$mappedValue }
                                "datetime" { 
                                    if ($mappedValue -match '^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}                    # Type conversion
                    if ($dataType -and $mappedValue -ne $null -and $mappedValue -ne "") {
                        try {
                            switch ($dataType.ToLower()) {
                                "int" { $mappedValue = [int]$mappedValue }
                                "decimal" { $mappedValue = [decimal]$mappedValue }
                                "datetime" { 
                                    if ($mappedValue -match '^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    return @{ Data = $result; Log = $log }
}

# UI Creation
Write-Host "Creating UI components..."

$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Transformation Tool with Validation"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)

$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($tableLayout)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($buttonPanel, 0, 0)

# Create buttons
$btnLoadFiles = New-Object System.Windows.Forms.Button
$btnLoadFiles.Text = "Load Files"
$btnLoadFiles.Width = 120
$btnLoadFiles.Location = New-Object System.Drawing.Point(10, 10)
$buttonPanel.Controls.Add($btnLoadFiles)

$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Text = "Show Errors Only"
$btnFilter.Width = 120
$btnFilter.Location = New-Object System.Drawing.Point(140, 10)
$buttonPanel.Controls.Add($btnFilter)

$btnMappedOnly = New-Object System.Windows.Forms.Button
$btnMappedOnly.Text = "Show Mapped Only"
$btnMappedOnly.Width = 120
$btnMappedOnly.Location = New-Object System.Drawing.Point(270, 10)
$buttonPanel.Controls.Add($btnMappedOnly)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Show All Data"
$btnReset.Width = 120
$btnReset.Location = New-Object System.Drawing.Point(400, 10)
$buttonPanel.Controls.Add($btnReset)

$btnSortByData = New-Object System.Windows.Forms.Button
$btnSortByData.Text = "Sort by Data"
$btnSortByData.Width = 120
$btnSortByData.Location = New-Object System.Drawing.Point(530, 10)
$buttonPanel.Controls.Add($btnSortByData)

$btnSortByMapping = New-Object System.Windows.Forms.Button
$btnSortByMapping.Text = "Sort by Mapping"
$btnSortByMapping.Width = 120
$btnSortByMapping.Location = New-Object System.Drawing.Point(660, 10)
$buttonPanel.Controls.Add($btnSortByMapping)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Width = 120
$btnExport.Location = New-Object System.Drawing.Point(790, 10)
$buttonPanel.Controls.Add($btnExport)

$btnExportErrors = New-Object System.Windows.Forms.Button
$btnExportErrors.Text = "Export Errors"
$btnExportErrors.Width = 120
$btnExportErrors.Location = New-Object System.Drawing.Point(920, 10)
$buttonPanel.Controls.Add($btnExportErrors)

$btnValidationReport = New-Object System.Windows.Forms.Button
$btnValidationReport.Text = "Validation Report"
$btnValidationReport.Width = 120
$btnValidationReport.Location = New-Object System.Drawing.Point(1050, 10)
$buttonPanel.Controls.Add($btnValidationReport)

# Create content panels
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($contentPanel, 0, 1)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 500
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$splitContainer.SplitterWidth = 5
$contentPanel.Controls.Add($splitContainer)

$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel1.Controls.Add($gridPanel)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Text = "Ready. Click 'Load Files' to begin."

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackColor = [System.Drawing.Color]::White
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.RowHeadersWidth = 60
$grid.RowHeadersVisible = $true
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 30
$grid.DefaultCellStyle.NullValue = ""
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true

# Event handlers
$grid.Add_DataError({
    param($sender, $e)
    if ($txtLog) {
        $txtLog.AppendText("`r`nGrid display error in cell [Row $($e.RowIndex), Column $($e.ColumnIndex)]: $($e.Exception.Message)")
    }
    $e.ThrowException = $false
})

$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.RowHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.RowHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridPanel.Controls.Add($grid)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel2.Controls.Add($logPanel)
$logPanel.Controls.Add($txtLog)

# Additional grid event handlers
$grid.Add_RowPostPaint({
    param($sender, $e)
    $rowNumber = ($e.RowIndex + 1).ToString()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($rowNumber, $grid.Font)
    
    $x = $e.RowBounds.Location.X + ($grid.RowHeadersWidth - $size.Width) / 2
    $y = $e.RowBounds.Location.Y + ($e.RowBounds.Height - $size.Height) / 2
    
    $location = New-Object System.Drawing.PointF($x, $y)
    $e.Graphics.DrawString($rowNumber, $grid.Font, $brush, $location)
})

$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    $hitTestInfo = $grid.HitTest($e.X, $e.Y)
    if ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
        $grid.ClearSelection()
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $grid.Rows[$i].Cells[$hitTestInfo.ColumnIndex].Selected = $true
        }
    }
})

$grid.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $grid.ClearSelection()
        $grid.Rows[$e.RowIndex].Selected = $true
        $txtLog.AppendText("`r`nSelected row $($e.RowIndex + 1) via row header")
    }
})

$grid.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    
    $rowIndex = $e.RowIndex
    $colName = if ($e.ColumnIndex -lt $grid.Columns.Count) { 
        $grid.Columns[$e.ColumnIndex].Name 
    } else { 
        return 
    }
    
    $key = "$rowIndex-$colName"
    if ($global:ErrorCells.ContainsKey($key)) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightPink
        $e.FormattingApplied = $true
    }
})

# Function definitions
function Load-Files {
    $global:ErrorCells.Clear()
    $global:ShowOnlyMappedFields = $false   
    $txtLog.Text = "Starting file loading process..."

    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    
    # Select Mapping CSV
    $openFileDialog.Title = "Select Mapping CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled mapping file selection."
        return
    }
    
    $mappingPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading mapping file: $mappingPath"
    
    $mapping = Load-Mapping $mappingPath
    if (-not $mapping -or $mapping.Count -eq 0) {
        $txtLog.Text += "`r`nMapping file is null or empty. Loading failed."
        return
    }
    
    $global:Mapping = $mapping
    $txtLog.Text += "`r`nLoaded mapping with $($mapping.Count) row(s)"
    
    # Build the source-to-new field map
    $global:SourceToNewFieldMap.Clear()
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
            if (-not $global:SourceToNewFieldMap.ContainsKey($m.SourceField)) {
                $global:SourceToNewFieldMap[$m.SourceField] = @()
            }
            $global:SourceToNewFieldMap[$m.SourceField] += $m.NewField
            $txtLog.Text += "`r`nField mapping: '$($m.SourceField)' => '$($m.NewField)'"
        }
    }
    
    $multiMappedFields = 0
    foreach ($sourceField in $global:SourceToNewFieldMap.Keys) {
        if ($global:SourceToNewFieldMap[$sourceField].Count -gt 1) {
            $multiMappedFields++
            $txtLog.Text += "`r`nMulti-mapped field: '$sourceField' => " + ($global:SourceToNewFieldMap[$sourceField] -join ", ")
        }
    }
    
    $txtLog.Text += "`r`nTotal source fields with mappings: $($global:SourceToNewFieldMap.Count) (including $multiMappedFields with multiple mappings)"
    $txtLog.Text += "`r`nOriginal column order from CSV: $($global:OriginalColumnOrder -join ', ')"
    
    # Extract column order from mapping
    $global:MappingColumnOrder = @()
    foreach ($m in $mapping) {
        $fieldName = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $m.SourceField }
        if (-not [string]::IsNullOrEmpty($fieldName) -and $global:MappingColumnOrder -notcontains $fieldName) {
            $global:MappingColumnOrder += $fieldName
        }
    }
    
    $txtLog.Text += "`r`nMapping column order: $($global:MappingColumnOrder -join ', ')"
    
    # Select Data CSV
    $openFileDialog.Title = "Select Data CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled data file selection."
        return
    }
    
    $dataPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading data file: $dataPath"
    
    $data = Load-Data $dataPath
    if (-not $data -or $data.Count -eq 0) {
        $txtLog.Text += "`r`nData file is null or empty."
        return
    }
    
    $txtLog.Text += "`r`nLoaded data with $($data.Count) row(s)"
    
    # Process data with progress bar
    $txtLog.Text += "`r`nProcessing data with validation..."
    try {
        # Disable the form during processing to prevent interference
        $form.Enabled = $false
        
        $result = Process-Data $data $global:Mapping
        $global:ProcessedData = $result.Data
        $global:LogData = $result.Log
        
        # Re-enable the form
        $form.Enabled = $true
        
        if ($global:LogData.Count -gt 0) {
            $txtLog.Text += "`r`n--- Processing Results ---"
            foreach ($error in $global:LogData) {
                $txtLog.Text += "`r`n$error"
            }
        } else {
            $txtLog.Text += "`r`nNo validation errors found."
        }
    } catch {
        # Make sure to re-enable form and hide progress bar on error
        $form.Enabled = $true
        Hide-ProgressBar
        $txtLog.Text += "`r`nError during processing: $_"
        $global:ProcessedData = $data
    }
    
    # Display in grid
    if ($global:ProcessedData.Count -gt 0) {
        $firstRow = $global:ProcessedData[0]
        $columns = $firstRow.PSObject.Properties.Name
        
        $txtLog.Text += "`r`nDetected columns in processed data: $($columns -join ', ')"
        
        Show-MappedOnly
        
        $txtLog.Text += "`r`nData loaded successfully. Found $($global:ProcessedData.Count) records."
        
        # Show validation summary in a more prominent way
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary) {
            $statusColor = if ($summary.DataIntegrityStatus -eq "PASSED") { "GREEN" } else { "ORANGE" }
            $txtLog.Text += "`r`n"
            $txtLog.Text += "`r`n=== DATA VALIDATION SUMMARY ==="
            $txtLog.Text += "`r`nStatus: $($summary.DataIntegrityStatus)"
            $txtLog.Text += "`r`nOriginal rows: $($summary.OriginalRowCount)"
            $txtLog.Text += "`r`nProcessed rows: $($summary.ProcessedRowCount)"
            if ($summary.MissingRowCount -gt 0) {
                $txtLog.Text += "`r`nMissing rows: $($summary.MissingRowCount)"
            }
            if ($summary.UnmatchedRowCount -gt 0) {
                $txtLog.Text += "`r`nUnmatched rows: $($summary.UnmatchedRowCount)"
            }
            $txtLog.Text += "`r`n=== END VALIDATION SUMMARY ==="
        }
    } else {
        $txtLog.Text += "`r`nNo data to display"
    }
}

function Show-ErrorsOnly {
    $txtLog.Text += "`r`nFiltering to show only rows with errors..."
    
    if ($global:ErrorCells.Count -eq 0) {
        $txtLog.Text += "`r`nNo error cells detected. Nothing to filter."
        return
    }
    
    $errorRows = @{}
    foreach ($key in $global:ErrorCells.Keys) {
        $parts = $key -split "-"
        if ($parts.Count -eq 2) {
            $rowIndex = [int]$parts[0]
            $errorRows[$rowIndex] = $true
        }
    }
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $errorRows.ContainsKey($i)
    }
    
    $txtLog.Text += "`r`nShowing only rows with errors."
}

function Show-MappedOnly {
    $txtLog.Text += "`r`nFiltering to show only fields specified in the mapping..."
    
    if ($null -eq $global:Mapping -or $global:Mapping.Count -eq 0) {
        $txtLog.Text += "`r`nNo mapping loaded. Nothing to filter."
        return
    }
    
    $global:ShowOnlyMappedFields = $true
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing only mapped fields."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nNo data loaded to reprocess."
    }
}

function Show-AllData {
    $txtLog.Text += "`r`nShowing all rows..."
    
    $global:ShowOnlyMappedFields = $false
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $true
    }
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing all fields and rows."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nAll rows are now visible."
    }
}

function Export-Data {
    if ($null -eq $global:ProcessedData -or $global:ProcessedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.")
        return
    }
    
    if ($global:ErrorCells.Count -gt 0) {
        $errorCount = $global:ErrorCells.Count
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            "Warning: Your data contains $errorCount error(s). Do you still want to proceed with the export?",
            "Export Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
            $txtLog.Text += "`r`nExport cancelled due to validation errors."
            return
        }
    }
    
    if ($global:DataValidationResults.Count -gt 0) {
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary -and $summary.DataIntegrityStatus -ne "PASSED") {
            $validationWarning = [System.Windows.Forms.MessageBox]::Show(
                "Warning: Data validation detected potential issues:`n" +
                "- Original rows: $($summary.OriginalRowCount)`n" +
                "- Processed rows: $($summary.ProcessedRowCount)`n" +
                "- Missing rows: $($summary.MissingRowCount)`n" +
                "- Unmatched rows: $($summary.UnmatchedRowCount)`n`n" +
                "Do you still want to proceed with the export?",
                "Data Validation Warning",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($validationWarning -eq [System.Windows.Forms.DialogResult]::No) {
                $txtLog.Text += "`r`nExport cancelled due to data validation issues."
                return
            }
        }
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Processed Data"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:ProcessedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nData exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting data: $_"
        }
    }
}

function Export-Errors {
    if ($null -eq $global:LogData -or $global:LogData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No errors to export.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Error Log"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $errorObjects = @()
            foreach ($error in $global:LogData) {
                $errorObjects += [PSCustomObject]@{ ErrorMessage = $error }
            }
            
            $errorObjects | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Errors exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nErrors exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting errors: $_"
        }
    }
}

function Export-ValidationReport {
    if ($null -eq $global:DataValidationResults -or $global:DataValidationResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No validation results available. Please load and process data first.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Data Validation Report"
    $saveFileDialog.FileName = "DataValidationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $success = Export-DataValidationReport $global:DataValidationResults $saveFileDialog.FileName
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Validation report exported successfully to $($saveFileDialog.FileName)")
                $txtLog.Text += "`r`nValidation report exported successfully to $($saveFileDialog.FileName)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error exporting validation report.")
                $txtLog.Text += "`r`nError exporting validation report."
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting validation report: $_")
            $txtLog.Text += "`r`nError exporting validation report: $_"
        }
    }
}

function Update-GridView {
    param([bool]$useMappingOrder = $false)
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $grid.Columns.Clear()
        
        $firstRow = $global:ProcessedData[0]
        $availableColumns = $firstRow.PSObject.Properties.Name
        
        $orderedColumns = @()
        
        if ($useMappingOrder -and $global:MappingColumnOrder.Count -gt 0) {
            $txtLog.AppendText("`r`nUsing mapping file column order")
            
            foreach ($col in $global:MappingColumnOrder) {
                if ($availableColumns -contains $col) {
                    $orderedColumns += $col
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        } 
        else {
            $txtLog.AppendText("`r`nUsing data file column order with renamed fields")
            
            foreach ($origCol in $global:OriginalColumnOrder) {
                $targetCol = $origCol
                
                if ($global:SourceToNewFieldMap.ContainsKey($origCol)) {
                    $targetCol = $global:SourceToNewFieldMap[$origCol]
                    $txtLog.AppendText("`r`nTranslating column '$origCol' to '$targetCol'")
                }
                
                if ($availableColumns -contains $targetCol -and $orderedColumns -notcontains $targetCol) {
                    $orderedColumns += $targetCol
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        }
        
        if ($orderedColumns.Count -eq 0) {
            $orderedColumns = $availableColumns
        }
        
        $txtLog.AppendText("`r`nFinal column order: $($orderedColumns -join ', ')")
        
        # Create columns
        foreach ($colName in $orderedColumns) {
            $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $column.HeaderText = $colName
            $column.Name = $colName
            $column.Width = 120
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $grid.Columns.Add($column)
        }
        
        # Add rows
        for ($rowIdx = 0; $rowIdx -lt $global:ProcessedData.Count; $rowIdx++) {
            $rowObj = $global:ProcessedData[$rowIdx]
            
            $grid.Rows.Add() | Out-Null
            $currentRow = $grid.Rows[$grid.Rows.Count - 1]
            
            if ($rowIdx % 2 -eq 1) {
                $currentRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
            }
            
            for ($colIdx = 0; $colIdx -lt $orderedColumns.Count; $colIdx++) {
                $colName = $orderedColumns[$colIdx]
                
                if ($null -eq $rowObj.$colName) {
                    $currentRow.Cells[$colIdx].Value = [DBNull]::Value
                } else {
                    $currentRow.Cells[$colIdx].Value = "$($rowObj.$colName)"
                }
            }
        }
        
        # Mark error cells
        foreach ($key in $global:ErrorCells.Keys) {
            $parts = $key -split "-"
            if ($parts.Count -eq 2) {
                $rowIndex = [int]$parts[0]
                $colName = $parts[1]
                
                $colIndex = -1
                for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                    if ($grid.Columns[$i].Name -eq $colName) {
                        $colIndex = $i
                        break
                    }
                }
                
                if ($rowIndex -lt $grid.Rows.Count -and $colIndex -ge 0) {
                    $grid.Rows[$rowIndex].Cells[$colIndex].Style.BackColor = [System.Drawing.Color]::LightPink
                }
            }
        }
    }
}

function Sort-ByDataOrder {
    $txtLog.AppendText("`r`nSorting by original data file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $false
    }
}

function Sort-ByMappingOrder {
    $txtLog.AppendText("`r`nSorting by mapping file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $true
    }
}

# Button event handlers
$btnLoadFiles.Add_Click({ Load-Files })
$btnFilter.Add_Click({ Show-ErrorsOnly })
$btnMappedOnly.Add_Click({ Show-MappedOnly })
$btnReset.Add_Click({ Show-AllData })
$btnSortByData.Add_Click({ Sort-ByDataOrder })
$btnSortByMapping.Add_Click({ Sort-ByMappingOrder })
$btnExport.Add_Click({ Export-Data })
$btnExportErrors.Add_Click({ Export-Errors })
$btnValidationReport.Add_Click({ Export-ValidationReport })

$form.Add_Shown({
    $txtLog.AppendText("`r`nApplication ready. Please click 'Load Files' to begin.")
})

Write-Host "Starting application..."
[void]$form.ShowDialog()) {
                                        try {
                                            $mappedValue = [datetime]::ParseExact(
                                                $mappedValue, 
                                                "dd/MM/yyyy HH:mm:ss", 
                                                [System.Globalization.CultureInfo]::InvariantCulture
                                            )
                                        }
                                        catch {
                                            $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                        }
                                    }
                                    else {
                                        $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                    }
                                }
                                default { }
                            }
                        } catch {
                            $error = "Data type conversion failed for $dataType"
                        }
                    }

                    # Use the mapped field name
                    $newField = $map.NewField
                    if ([string]::IsNullOrEmpty($newField)) {
                        $newField = $sourceField
                    }
                    
                    $newRow[$newField] = $mappedValue

                    # Record errors
                    if ($error) {
                        $log += "Row $($i+1) Field '$sourceField' (mapped to '$newField'): $error"
                        $global:ErrorCells["$i-$newField"] = $error
                    }
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    Update-ProgressBar 80 "Starting data integrity validation..."
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    Update-ProgressBar 95 "Compiling results..."
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    Update-ProgressBar 100 "Processing complete!"
    
    # Small delay to show completion
    Start-Sleep -Milliseconds 500
    Hide-ProgressBar
    
    return @{ Data = $result; Log = $log }
}# Type conversion
                    if ($dataType -and $mappedValue -ne $null -and $mappedValue -ne "") {
                        try {
                            switch ($dataType.ToLower()) {
                                "int" { $mappedValue = [int]$mappedValue }
                                "decimal" { $mappedValue = [decimal]$mappedValue }
                                "datetime" { 
                                    # Check for datetime format dd/mm/yyyy HH:MM:SS
                                    if ($mappedValue -match '\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}') {
                                        try {
                                            $mappedValue = [datetime]::ParseExact(
                                                $mappedValue, 
                                                "dd/MM/yyyy HH:mm:ss", 
                                                [System.Globalization.CultureInfo]::InvariantCulture
                                            )
                                        }
                                        catch {
                                            $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                        }
                                    }
                                    else {
                                        $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                    }
                                }
                                default { }
                            }
                        } catch {
                            $error = "Data type conversion failed for $dataType"
                        }
                    }

                    # Use the mapped field name
                    $newField = $map.NewField
                    if ([string]::IsNullOrEmpty($newField)) {
                        $newField = $sourceField
                    }
                    
                    $newRow[$newField] = $mappedValue

                    # Record errors
                    if ($error) {
                        $log += "Row $($i+1) Field '$sourceField' (mapped to '$newField'): $error"
                        $global:ErrorCells["$i-$newField"] = $error
                    }
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    return @{ Data = $result; Log = $log }
}

# UI Creation
Write-Host "Creating UI components..."

$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Transformation Tool with Validation"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)

$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($tableLayout)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($buttonPanel, 0, 0)

# Create buttons
$btnLoadFiles = New-Object System.Windows.Forms.Button
$btnLoadFiles.Text = "Load Files"
$btnLoadFiles.Width = 120
$btnLoadFiles.Location = New-Object System.Drawing.Point(10, 10)
$buttonPanel.Controls.Add($btnLoadFiles)

$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Text = "Show Errors Only"
$btnFilter.Width = 120
$btnFilter.Location = New-Object System.Drawing.Point(140, 10)
$buttonPanel.Controls.Add($btnFilter)

$btnMappedOnly = New-Object System.Windows.Forms.Button
$btnMappedOnly.Text = "Show Mapped Only"
$btnMappedOnly.Width = 120
$btnMappedOnly.Location = New-Object System.Drawing.Point(270, 10)
$buttonPanel.Controls.Add($btnMappedOnly)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Show All Data"
$btnReset.Width = 120
$btnReset.Location = New-Object System.Drawing.Point(400, 10)
$buttonPanel.Controls.Add($btnReset)

$btnSortByData = New-Object System.Windows.Forms.Button
$btnSortByData.Text = "Sort by Data"
$btnSortByData.Width = 120
$btnSortByData.Location = New-Object System.Drawing.Point(530, 10)
$buttonPanel.Controls.Add($btnSortByData)

$btnSortByMapping = New-Object System.Windows.Forms.Button
$btnSortByMapping.Text = "Sort by Mapping"
$btnSortByMapping.Width = 120
$btnSortByMapping.Location = New-Object System.Drawing.Point(660, 10)
$buttonPanel.Controls.Add($btnSortByMapping)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Width = 120
$btnExport.Location = New-Object System.Drawing.Point(790, 10)
$buttonPanel.Controls.Add($btnExport)

$btnExportErrors = New-Object System.Windows.Forms.Button
$btnExportErrors.Text = "Export Errors"
$btnExportErrors.Width = 120
$btnExportErrors.Location = New-Object System.Drawing.Point(920, 10)
$buttonPanel.Controls.Add($btnExportErrors)

$btnValidationReport = New-Object System.Windows.Forms.Button
$btnValidationReport.Text = "Validation Report"
$btnValidationReport.Width = 120
$btnValidationReport.Location = New-Object System.Drawing.Point(1050, 10)
$buttonPanel.Controls.Add($btnValidationReport)

# Create content panels
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($contentPanel, 0, 1)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 500
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$splitContainer.SplitterWidth = 5
$contentPanel.Controls.Add($splitContainer)

$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel1.Controls.Add($gridPanel)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Text = "Ready. Click 'Load Files' to begin."

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackColor = [System.Drawing.Color]::White
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.RowHeadersWidth = 60
$grid.RowHeadersVisible = $true
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 30
$grid.DefaultCellStyle.NullValue = ""
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true

# Event handlers
$grid.Add_DataError({
    param($sender, $e)
    if ($txtLog) {
        $txtLog.AppendText("`r`nGrid display error in cell [Row $($e.RowIndex), Column $($e.ColumnIndex)]: $($e.Exception.Message)")
    }
    $e.ThrowException = $false
})

$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.RowHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.RowHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridPanel.Controls.Add($grid)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel2.Controls.Add($logPanel)
$logPanel.Controls.Add($txtLog)

# Additional grid event handlers
$grid.Add_RowPostPaint({
    param($sender, $e)
    $rowNumber = ($e.RowIndex + 1).ToString()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($rowNumber, $grid.Font)
    
    $x = $e.RowBounds.Location.X + ($grid.RowHeadersWidth - $size.Width) / 2
    $y = $e.RowBounds.Location.Y + ($e.RowBounds.Height - $size.Height) / 2
    
    $location = New-Object System.Drawing.PointF($x, $y)
    $e.Graphics.DrawString($rowNumber, $grid.Font, $brush, $location)
})

$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    $hitTestInfo = $grid.HitTest($e.X, $e.Y)
    if ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
        $grid.ClearSelection()
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $grid.Rows[$i].Cells[$hitTestInfo.ColumnIndex].Selected = $true
        }
    }
})

$grid.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $grid.ClearSelection()
        $grid.Rows[$e.RowIndex].Selected = $true
        $txtLog.AppendText("`r`nSelected row $($e.RowIndex + 1) via row header")
    }
})

$grid.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    
    $rowIndex = $e.RowIndex
    $colName = if ($e.ColumnIndex -lt $grid.Columns.Count) { 
        $grid.Columns[$e.ColumnIndex].Name 
    } else { 
        return 
    }
    
    $key = "$rowIndex-$colName"
    if ($global:ErrorCells.ContainsKey($key)) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightPink
        $e.FormattingApplied = $true
    }
})

# Function definitions
function Load-Files {
    $global:ErrorCells.Clear()
    $global:ShowOnlyMappedFields = $false   
    $txtLog.Text = "Starting file loading process..."

    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    
    # Select Mapping CSV
    $openFileDialog.Title = "Select Mapping CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled mapping file selection."
        return
    }
    
    $mappingPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading mapping file: $mappingPath"
    
    $mapping = Load-Mapping $mappingPath
    if (-not $mapping -or $mapping.Count -eq 0) {
        $txtLog.Text += "`r`nMapping file is null or empty. Loading failed."
        return
    }
    
    $global:Mapping = $mapping
    $txtLog.Text += "`r`nLoaded mapping with $($mapping.Count) row(s)"
    
    # Build the source-to-new field map
    $global:SourceToNewFieldMap.Clear()
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
            if (-not $global:SourceToNewFieldMap.ContainsKey($m.SourceField)) {
                $global:SourceToNewFieldMap[$m.SourceField] = @()
            }
            $global:SourceToNewFieldMap[$m.SourceField] += $m.NewField
            $txtLog.Text += "`r`nField mapping: '$($m.SourceField)' => '$($m.NewField)'"
        }
    }
    
    $multiMappedFields = 0
    foreach ($sourceField in $global:SourceToNewFieldMap.Keys) {
        if ($global:SourceToNewFieldMap[$sourceField].Count -gt 1) {
            $multiMappedFields++
            $txtLog.Text += "`r`nMulti-mapped field: '$sourceField' => " + ($global:SourceToNewFieldMap[$sourceField] -join ", ")
        }
    }
    
    $txtLog.Text += "`r`nTotal source fields with mappings: $($global:SourceToNewFieldMap.Count) (including $multiMappedFields with multiple mappings)"
    $txtLog.Text += "`r`nOriginal column order from CSV: $($global:OriginalColumnOrder -join ', ')"
    
    # Extract column order from mapping
    $global:MappingColumnOrder = @()
    foreach ($m in $mapping) {
        $fieldName = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $m.SourceField }
        if (-not [string]::IsNullOrEmpty($fieldName) -and $global:MappingColumnOrder -notcontains $fieldName) {
            $global:MappingColumnOrder += $fieldName
        }
    }
    
    $txtLog.Text += "`r`nMapping column order: $($global:MappingColumnOrder -join ', ')"
    
    # Select Data CSV
    $openFileDialog.Title = "Select Data CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled data file selection."
        return
    }
    
    $dataPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading data file: $dataPath"
    
    $data = Load-Data $dataPath
    if (-not $data -or $data.Count -eq 0) {
        $txtLog.Text += "`r`nData file is null or empty."
        return
    }
    
    $txtLog.Text += "`r`nLoaded data with $($data.Count) row(s)"
    
    # Process data
    $txtLog.Text += "`r`nProcessing data..."
    try {
        $result = Process-Data $data $global:Mapping
        $global:ProcessedData = $result.Data
        $global:LogData = $result.Log
        
        if ($global:LogData.Count -gt 0) {
            $txtLog.Text += "`r`n--- Processing Results ---"
            foreach ($error in $global:LogData) {
                $txtLog.Text += "`r`n$error"
            }
        } else {
            $txtLog.Text += "`r`nNo validation errors found."
        }
    } catch {
        $txtLog.Text += "`r`nError during processing: $_"
        $global:ProcessedData = $data
    }
    
    # Display in grid
    if ($global:ProcessedData.Count -gt 0) {
        $firstRow = $global:ProcessedData[0]
        $columns = $firstRow.PSObject.Properties.Name
        
        $txtLog.Text += "`r`nDetected columns in processed data: $($columns -join ', ')"
        
        Show-MappedOnly
        
        $txtLog.Text += "`r`nData loaded successfully. Found $($global:ProcessedData.Count) records."
    } else {
        $txtLog.Text += "`r`nNo data to display"
    }
}

function Show-ErrorsOnly {
    $txtLog.Text += "`r`nFiltering to show only rows with errors..."
    
    if ($global:ErrorCells.Count -eq 0) {
        $txtLog.Text += "`r`nNo error cells detected. Nothing to filter."
        return
    }
    
    $errorRows = @{}
    foreach ($key in $global:ErrorCells.Keys) {
        $parts = $key -split "-"
        if ($parts.Count -eq 2) {
            $rowIndex = [int]$parts[0]
            $errorRows[$rowIndex] = $true
        }
    }
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $errorRows.ContainsKey($i)
    }
    
    $txtLog.Text += "`r`nShowing only rows with errors."
}

function Show-MappedOnly {
    $txtLog.Text += "`r`nFiltering to show only fields specified in the mapping..."
    
    if ($null -eq $global:Mapping -or $global:Mapping.Count -eq 0) {
        $txtLog.Text += "`r`nNo mapping loaded. Nothing to filter."
        return
    }
    
    $global:ShowOnlyMappedFields = $true
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing only mapped fields."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nNo data loaded to reprocess."
    }
}

function Show-AllData {
    $txtLog.Text += "`r`nShowing all rows..."
    
    $global:ShowOnlyMappedFields = $false
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $true
    }
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing all fields and rows."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nAll rows are now visible."
    }
}

function Export-Data {
    if ($null -eq $global:ProcessedData -or $global:ProcessedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.")
        return
    }
    
    if ($global:ErrorCells.Count -gt 0) {
        $errorCount = $global:ErrorCells.Count
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            "Warning: Your data contains $errorCount error(s). Do you still want to proceed with the export?",
            "Export Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
            $txtLog.Text += "`r`nExport cancelled due to validation errors."
            return
        }
    }
    
    if ($global:DataValidationResults.Count -gt 0) {
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary -and $summary.DataIntegrityStatus -ne "PASSED") {
            $validationWarning = [System.Windows.Forms.MessageBox]::Show(
                "Warning: Data validation detected potential issues:`n" +
                "- Original rows: $($summary.OriginalRowCount)`n" +
                "- Processed rows: $($summary.ProcessedRowCount)`n" +
                "- Missing rows: $($summary.MissingRowCount)`n" +
                "- Unmatched rows: $($summary.UnmatchedRowCount)`n`n" +
                "Do you still want to proceed with the export?",
                "Data Validation Warning",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($validationWarning -eq [System.Windows.Forms.DialogResult]::No) {
                $txtLog.Text += "`r`nExport cancelled due to data validation issues."
                return
            }
        }
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Processed Data"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:ProcessedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nData exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting data: $_"
        }
    }
}

function Export-Errors {
    if ($null -eq $global:LogData -or $global:LogData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No errors to export.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Error Log"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $errorObjects = @()
            foreach ($error in $global:LogData) {
                $errorObjects += [PSCustomObject]@{ ErrorMessage = $error }
            }
            
            $errorObjects | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Errors exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nErrors exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting errors: $_"
        }
    }
}

function Export-ValidationReport {
    if ($null -eq $global:DataValidationResults -or $global:DataValidationResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No validation results available. Please load and process data first.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Data Validation Report"
    $saveFileDialog.FileName = "DataValidationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $success = Export-DataValidationReport $global:DataValidationResults $saveFileDialog.FileName
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Validation report exported successfully to $($saveFileDialog.FileName)")
                $txtLog.Text += "`r`nValidation report exported successfully to $($saveFileDialog.FileName)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error exporting validation report.")
                $txtLog.Text += "`r`nError exporting validation report."
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting validation report: $_")
            $txtLog.Text += "`r`nError exporting validation report: $_"
        }
    }
}

function Update-GridView {
    param([bool]$useMappingOrder = $false)
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $grid.Columns.Clear()
        
        $firstRow = $global:ProcessedData[0]
        $availableColumns = $firstRow.PSObject.Properties.Name
        
        $orderedColumns = @()
        
        if ($useMappingOrder -and $global:MappingColumnOrder.Count -gt 0) {
            $txtLog.AppendText("`r`nUsing mapping file column order")
            
            foreach ($col in $global:MappingColumnOrder) {
                if ($availableColumns -contains $col) {
                    $orderedColumns += $col
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        } 
        else {
            $txtLog.AppendText("`r`nUsing data file column order with renamed fields")
            
            foreach ($origCol in $global:OriginalColumnOrder) {
                $targetCol = $origCol
                
                if ($global:SourceToNewFieldMap.ContainsKey($origCol)) {
                    $targetCol = $global:SourceToNewFieldMap[$origCol]
                    $txtLog.AppendText("`r`nTranslating column '$origCol' to '$targetCol'")
                }
                
                if ($availableColumns -contains $targetCol -and $orderedColumns -notcontains $targetCol) {
                    $orderedColumns += $targetCol
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        }
        
        if ($orderedColumns.Count -eq 0) {
            $orderedColumns = $availableColumns
        }
        
        $txtLog.AppendText("`r`nFinal column order: $($orderedColumns -join ', ')")
        
        # Create columns
        foreach ($colName in $orderedColumns) {
            $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $column.HeaderText = $colName
            $column.Name = $colName
            $column.Width = 120
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $grid.Columns.Add($column)
        }
        
        # Add rows
        for ($rowIdx = 0; $rowIdx -lt $global:ProcessedData.Count; $rowIdx++) {
            $rowObj = $global:ProcessedData[$rowIdx]
            
            $grid.Rows.Add() | Out-Null
            $currentRow = $grid.Rows[$grid.Rows.Count - 1]
            
            if ($rowIdx % 2 -eq 1) {
                $currentRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
            }
            
            for ($colIdx = 0; $colIdx -lt $orderedColumns.Count; $colIdx++) {
                $colName = $orderedColumns[$colIdx]
                
                if ($null -eq $rowObj.$colName) {
                    $currentRow.Cells[$colIdx].Value = [DBNull]::Value
                } else {
                    $currentRow.Cells[$colIdx].Value = "$($rowObj.$colName)"
                }
            }
        }
        
        # Mark error cells
        foreach ($key in $global:ErrorCells.Keys) {
            $parts = $key -split "-"
            if ($parts.Count -eq 2) {
                $rowIndex = [int]$parts[0]
                $colName = $parts[1]
                
                $colIndex = -1
                for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                    if ($grid.Columns[$i].Name -eq $colName) {
                        $colIndex = $i
                        break
                    }
                }
                
                if ($rowIndex -lt $grid.Rows.Count -and $colIndex -ge 0) {
                    $grid.Rows[$rowIndex].Cells[$colIndex].Style.BackColor = [System.Drawing.Color]::LightPink
                }
            }
        }
    }
}

function Sort-ByDataOrder {
    $txtLog.AppendText("`r`nSorting by original data file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $false
    }
}

function Sort-ByMappingOrder {
    $txtLog.AppendText("`r`nSorting by mapping file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $true
    }
}

# Button event handlers
$btnLoadFiles.Add_Click({ Load-Files })
$btnFilter.Add_Click({ Show-ErrorsOnly })
$btnMappedOnly.Add_Click({ Show-MappedOnly })
$btnReset.Add_Click({ Show-AllData })
$btnSortByData.Add_Click({ Sort-ByDataOrder })
$btnSortByMapping.Add_Click({ Sort-ByMappingOrder })
$btnExport.Add_Click({ Export-Data })
$btnExportErrors.Add_Click({ Export-Errors })
$btnValidationReport.Add_Click({ Export-ValidationReport })

$form.Add_Shown({
    $txtLog.AppendText("`r`nApplication ready. Please click 'Load Files' to begin.")
})

Write-Host "Starting application..."
[void]$form.ShowDialog()) {
                                        try {
                                            $mappedValue = [datetime]::ParseExact(
                                                $mappedValue, 
                                                "dd/MM/yyyy HH:mm:ss", 
                                                [System.Globalization.CultureInfo]::InvariantCulture
                                            )
                                        }
                                        catch {
                                            $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                        }
                                    }
                                    else {
                                        $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                    }
                                }
                                default { }
                            }
                        } catch {
                            $error = "Data type conversion failed for $dataType"
                        }
                    }

                    # Use the mapped field name
                    $newField = $map.NewField
                    if ([string]::IsNullOrEmpty($newField)) {
                        $newField = $sourceField
                    }
                    
                    $newRow[$newField] = $mappedValue

                    # Record errors
                    if ($error) {
                        $log += "Row $($i+1) Field '$sourceField' (mapped to '$newField'): $error"
                        $global:ErrorCells["$i-$newField"] = $error
                    }
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    return @{ Data = $result; Log = $log }
}

# UI Creation
Write-Host "Creating UI components..."

$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Transformation Tool with Validation"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)

$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($tableLayout)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($buttonPanel, 0, 0)

# Create buttons
$btnLoadFiles = New-Object System.Windows.Forms.Button
$btnLoadFiles.Text = "Load Files"
$btnLoadFiles.Width = 120
$btnLoadFiles.Location = New-Object System.Drawing.Point(10, 10)
$buttonPanel.Controls.Add($btnLoadFiles)

$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Text = "Show Errors Only"
$btnFilter.Width = 120
$btnFilter.Location = New-Object System.Drawing.Point(140, 10)
$buttonPanel.Controls.Add($btnFilter)

$btnMappedOnly = New-Object System.Windows.Forms.Button
$btnMappedOnly.Text = "Show Mapped Only"
$btnMappedOnly.Width = 120
$btnMappedOnly.Location = New-Object System.Drawing.Point(270, 10)
$buttonPanel.Controls.Add($btnMappedOnly)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Show All Data"
$btnReset.Width = 120
$btnReset.Location = New-Object System.Drawing.Point(400, 10)
$buttonPanel.Controls.Add($btnReset)

$btnSortByData = New-Object System.Windows.Forms.Button
$btnSortByData.Text = "Sort by Data"
$btnSortByData.Width = 120
$btnSortByData.Location = New-Object System.Drawing.Point(530, 10)
$buttonPanel.Controls.Add($btnSortByData)

$btnSortByMapping = New-Object System.Windows.Forms.Button
$btnSortByMapping.Text = "Sort by Mapping"
$btnSortByMapping.Width = 120
$btnSortByMapping.Location = New-Object System.Drawing.Point(660, 10)
$buttonPanel.Controls.Add($btnSortByMapping)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Width = 120
$btnExport.Location = New-Object System.Drawing.Point(790, 10)
$buttonPanel.Controls.Add($btnExport)

$btnExportErrors = New-Object System.Windows.Forms.Button
$btnExportErrors.Text = "Export Errors"
$btnExportErrors.Width = 120
$btnExportErrors.Location = New-Object System.Drawing.Point(920, 10)
$buttonPanel.Controls.Add($btnExportErrors)

$btnValidationReport = New-Object System.Windows.Forms.Button
$btnValidationReport.Text = "Validation Report"
$btnValidationReport.Width = 120
$btnValidationReport.Location = New-Object System.Drawing.Point(1050, 10)
$buttonPanel.Controls.Add($btnValidationReport)

# Create content panels
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($contentPanel, 0, 1)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 500
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$splitContainer.SplitterWidth = 5
$contentPanel.Controls.Add($splitContainer)

$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel1.Controls.Add($gridPanel)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Text = "Ready. Click 'Load Files' to begin."

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackColor = [System.Drawing.Color]::White
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.RowHeadersWidth = 60
$grid.RowHeadersVisible = $true
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 30
$grid.DefaultCellStyle.NullValue = ""
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true

# Event handlers
$grid.Add_DataError({
    param($sender, $e)
    if ($txtLog) {
        $txtLog.AppendText("`r`nGrid display error in cell [Row $($e.RowIndex), Column $($e.ColumnIndex)]: $($e.Exception.Message)")
    }
    $e.ThrowException = $false
})

$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.RowHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.RowHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridPanel.Controls.Add($grid)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel2.Controls.Add($logPanel)
$logPanel.Controls.Add($txtLog)

# Additional grid event handlers
$grid.Add_RowPostPaint({
    param($sender, $e)
    $rowNumber = ($e.RowIndex + 1).ToString()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($rowNumber, $grid.Font)
    
    $x = $e.RowBounds.Location.X + ($grid.RowHeadersWidth - $size.Width) / 2
    $y = $e.RowBounds.Location.Y + ($e.RowBounds.Height - $size.Height) / 2
    
    $location = New-Object System.Drawing.PointF($x, $y)
    $e.Graphics.DrawString($rowNumber, $grid.Font, $brush, $location)
})

$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    $hitTestInfo = $grid.HitTest($e.X, $e.Y)
    if ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
        $grid.ClearSelection()
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $grid.Rows[$i].Cells[$hitTestInfo.ColumnIndex].Selected = $true
        }
    }
})

$grid.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $grid.ClearSelection()
        $grid.Rows[$e.RowIndex].Selected = $true
        $txtLog.AppendText("`r`nSelected row $($e.RowIndex + 1) via row header")
    }
})

$grid.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    
    $rowIndex = $e.RowIndex
    $colName = if ($e.ColumnIndex -lt $grid.Columns.Count) { 
        $grid.Columns[$e.ColumnIndex].Name 
    } else { 
        return 
    }
    
    $key = "$rowIndex-$colName"
    if ($global:ErrorCells.ContainsKey($key)) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightPink
        $e.FormattingApplied = $true
    }
})

# Function definitions
function Load-Files {
    $global:ErrorCells.Clear()
    $global:ShowOnlyMappedFields = $false   
    $txtLog.Text = "Starting file loading process..."

    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    
    # Select Mapping CSV
    $openFileDialog.Title = "Select Mapping CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled mapping file selection."
        return
    }
    
    $mappingPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading mapping file: $mappingPath"
    
    $mapping = Load-Mapping $mappingPath
    if (-not $mapping -or $mapping.Count -eq 0) {
        $txtLog.Text += "`r`nMapping file is null or empty. Loading failed."
        return
    }
    
    $global:Mapping = $mapping
    $txtLog.Text += "`r`nLoaded mapping with $($mapping.Count) row(s)"
    
    # Build the source-to-new field map
    $global:SourceToNewFieldMap.Clear()
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
            if (-not $global:SourceToNewFieldMap.ContainsKey($m.SourceField)) {
                $global:SourceToNewFieldMap[$m.SourceField] = @()
            }
            $global:SourceToNewFieldMap[$m.SourceField] += $m.NewField
            $txtLog.Text += "`r`nField mapping: '$($m.SourceField)' => '$($m.NewField)'"
        }
    }
    
    $multiMappedFields = 0
    foreach ($sourceField in $global:SourceToNewFieldMap.Keys) {
        if ($global:SourceToNewFieldMap[$sourceField].Count -gt 1) {
            $multiMappedFields++
            $txtLog.Text += "`r`nMulti-mapped field: '$sourceField' => " + ($global:SourceToNewFieldMap[$sourceField] -join ", ")
        }
    }
    
    $txtLog.Text += "`r`nTotal source fields with mappings: $($global:SourceToNewFieldMap.Count) (including $multiMappedFields with multiple mappings)"
    $txtLog.Text += "`r`nOriginal column order from CSV: $($global:OriginalColumnOrder -join ', ')"
    
    # Extract column order from mapping
    $global:MappingColumnOrder = @()
    foreach ($m in $mapping) {
        $fieldName = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $m.SourceField }
        if (-not [string]::IsNullOrEmpty($fieldName) -and $global:MappingColumnOrder -notcontains $fieldName) {
            $global:MappingColumnOrder += $fieldName
        }
    }
    
    $txtLog.Text += "`r`nMapping column order: $($global:MappingColumnOrder -join ', ')"
    
    # Select Data CSV
    $openFileDialog.Title = "Select Data CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled data file selection."
        return
    }
    
    $dataPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading data file: $dataPath"
    
    $data = Load-Data $dataPath
    if (-not $data -or $data.Count -eq 0) {
        $txtLog.Text += "`r`nData file is null or empty."
        return
    }
    
    $txtLog.Text += "`r`nLoaded data with $($data.Count) row(s)"
    
    # Process data with progress bar
    $txtLog.Text += "`r`nProcessing data with validation..."
    try {
        # Disable the form during processing to prevent interference
        $form.Enabled = $false
        
        $result = Process-Data $data $global:Mapping
        $global:ProcessedData = $result.Data
        $global:LogData = $result.Log
        
        # Re-enable the form
        $form.Enabled = $true
        
        if ($global:LogData.Count -gt 0) {
            $txtLog.Text += "`r`n--- Processing Results ---"
            foreach ($error in $global:LogData) {
                $txtLog.Text += "`r`n$error"
            }
        } else {
            $txtLog.Text += "`r`nNo validation errors found."
        }
    } catch {
        # Make sure to re-enable form and hide progress bar on error
        $form.Enabled = $true
        Hide-ProgressBar
        $txtLog.Text += "`r`nError during processing: $_"
        $global:ProcessedData = $data
    }
    
    # Display in grid
    if ($global:ProcessedData.Count -gt 0) {
        $firstRow = $global:ProcessedData[0]
        $columns = $firstRow.PSObject.Properties.Name
        
        $txtLog.Text += "`r`nDetected columns in processed data: $($columns -join ', ')"
        
        Show-MappedOnly
        
        $txtLog.Text += "`r`nData loaded successfully. Found $($global:ProcessedData.Count) records."
        
        # Show validation summary in a more prominent way
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary) {
            $statusColor = if ($summary.DataIntegrityStatus -eq "PASSED") { "GREEN" } else { "ORANGE" }
            $txtLog.Text += "`r`n"
            $txtLog.Text += "`r`n=== DATA VALIDATION SUMMARY ==="
            $txtLog.Text += "`r`nStatus: $($summary.DataIntegrityStatus)"
            $txtLog.Text += "`r`nOriginal rows: $($summary.OriginalRowCount)"
            $txtLog.Text += "`r`nProcessed rows: $($summary.ProcessedRowCount)"
            if ($summary.MissingRowCount -gt 0) {
                $txtLog.Text += "`r`nMissing rows: $($summary.MissingRowCount)"
            }
            if ($summary.UnmatchedRowCount -gt 0) {
                $txtLog.Text += "`r`nUnmatched rows: $($summary.UnmatchedRowCount)"
            }
            $txtLog.Text += "`r`n=== END VALIDATION SUMMARY ==="
        }
    } else {
        $txtLog.Text += "`r`nNo data to display"
    }
}

function Show-ErrorsOnly {
    $txtLog.Text += "`r`nFiltering to show only rows with errors..."
    
    if ($global:ErrorCells.Count -eq 0) {
        $txtLog.Text += "`r`nNo error cells detected. Nothing to filter."
        return
    }
    
    $errorRows = @{}
    foreach ($key in $global:ErrorCells.Keys) {
        $parts = $key -split "-"
        if ($parts.Count -eq 2) {
            $rowIndex = [int]$parts[0]
            $errorRows[$rowIndex] = $true
        }
    }
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $errorRows.ContainsKey($i)
    }
    
    $txtLog.Text += "`r`nShowing only rows with errors."
}

function Show-MappedOnly {
    $txtLog.Text += "`r`nFiltering to show only fields specified in the mapping..."
    
    if ($null -eq $global:Mapping -or $global:Mapping.Count -eq 0) {
        $txtLog.Text += "`r`nNo mapping loaded. Nothing to filter."
        return
    }
    
    $global:ShowOnlyMappedFields = $true
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing only mapped fields."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nNo data loaded to reprocess."
    }
}

function Show-AllData {
    $txtLog.Text += "`r`nShowing all rows..."
    
    $global:ShowOnlyMappedFields = $false
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $true
    }
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing all fields and rows."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nAll rows are now visible."
    }
}

function Export-Data {
    if ($null -eq $global:ProcessedData -or $global:ProcessedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.")
        return
    }
    
    if ($global:ErrorCells.Count -gt 0) {
        $errorCount = $global:ErrorCells.Count
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            "Warning: Your data contains $errorCount error(s). Do you still want to proceed with the export?",
            "Export Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
            $txtLog.Text += "`r`nExport cancelled due to validation errors."
            return
        }
    }
    
    if ($global:DataValidationResults.Count -gt 0) {
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary -and $summary.DataIntegrityStatus -ne "PASSED") {
            $validationWarning = [System.Windows.Forms.MessageBox]::Show(
                "Warning: Data validation detected potential issues:`n" +
                "- Original rows: $($summary.OriginalRowCount)`n" +
                "- Processed rows: $($summary.ProcessedRowCount)`n" +
                "- Missing rows: $($summary.MissingRowCount)`n" +
                "- Unmatched rows: $($summary.UnmatchedRowCount)`n`n" +
                "Do you still want to proceed with the export?",
                "Data Validation Warning",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($validationWarning -eq [System.Windows.Forms.DialogResult]::No) {
                $txtLog.Text += "`r`nExport cancelled due to data validation issues."
                return
            }
        }
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Processed Data"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:ProcessedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nData exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting data: $_"
        }
    }
}

function Export-Errors {
    if ($null -eq $global:LogData -or $global:LogData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No errors to export.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Error Log"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $errorObjects = @()
            foreach ($error in $global:LogData) {
                $errorObjects += [PSCustomObject]@{ ErrorMessage = $error }
            }
            
            $errorObjects | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Errors exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nErrors exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting errors: $_"
        }
    }
}

function Export-ValidationReport {
    if ($null -eq $global:DataValidationResults -or $global:DataValidationResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No validation results available. Please load and process data first.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Data Validation Report"
    $saveFileDialog.FileName = "DataValidationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $success = Export-DataValidationReport $global:DataValidationResults $saveFileDialog.FileName
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Validation report exported successfully to $($saveFileDialog.FileName)")
                $txtLog.Text += "`r`nValidation report exported successfully to $($saveFileDialog.FileName)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error exporting validation report.")
                $txtLog.Text += "`r`nError exporting validation report."
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting validation report: $_")
            $txtLog.Text += "`r`nError exporting validation report: $_"
        }
    }
}

function Update-GridView {
    param([bool]$useMappingOrder = $false)
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $grid.Columns.Clear()
        
        $firstRow = $global:ProcessedData[0]
        $availableColumns = $firstRow.PSObject.Properties.Name
        
        $orderedColumns = @()
        
        if ($useMappingOrder -and $global:MappingColumnOrder.Count -gt 0) {
            $txtLog.AppendText("`r`nUsing mapping file column order")
            
            foreach ($col in $global:MappingColumnOrder) {
                if ($availableColumns -contains $col) {
                    $orderedColumns += $col
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        } 
        else {
            $txtLog.AppendText("`r`nUsing data file column order with renamed fields")
            
            foreach ($origCol in $global:OriginalColumnOrder) {
                $targetCol = $origCol
                
                if ($global:SourceToNewFieldMap.ContainsKey($origCol)) {
                    $targetCol = $global:SourceToNewFieldMap[$origCol]
                    $txtLog.AppendText("`r`nTranslating column '$origCol' to '$targetCol'")
                }
                
                if ($availableColumns -contains $targetCol -and $orderedColumns -notcontains $targetCol) {
                    $orderedColumns += $targetCol
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        }
        
        if ($orderedColumns.Count -eq 0) {
            $orderedColumns = $availableColumns
        }
        
        $txtLog.AppendText("`r`nFinal column order: $($orderedColumns -join ', ')")
        
        # Create columns
        foreach ($colName in $orderedColumns) {
            $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $column.HeaderText = $colName
            $column.Name = $colName
            $column.Width = 120
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $grid.Columns.Add($column)
        }
        
        # Add rows
        for ($rowIdx = 0; $rowIdx -lt $global:ProcessedData.Count; $rowIdx++) {
            $rowObj = $global:ProcessedData[$rowIdx]
            
            $grid.Rows.Add() | Out-Null
            $currentRow = $grid.Rows[$grid.Rows.Count - 1]
            
            if ($rowIdx % 2 -eq 1) {
                $currentRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
            }
            
            for ($colIdx = 0; $colIdx -lt $orderedColumns.Count; $colIdx++) {
                $colName = $orderedColumns[$colIdx]
                
                if ($null -eq $rowObj.$colName) {
                    $currentRow.Cells[$colIdx].Value = [DBNull]::Value
                } else {
                    $currentRow.Cells[$colIdx].Value = "$($rowObj.$colName)"
                }
            }
        }
        
        # Mark error cells
        foreach ($key in $global:ErrorCells.Keys) {
            $parts = $key -split "-"
            if ($parts.Count -eq 2) {
                $rowIndex = [int]$parts[0]
                $colName = $parts[1]
                
                $colIndex = -1
                for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                    if ($grid.Columns[$i].Name -eq $colName) {
                        $colIndex = $i
                        break
                    }
                }
                
                if ($rowIndex -lt $grid.Rows.Count -and $colIndex -ge 0) {
                    $grid.Rows[$rowIndex].Cells[$colIndex].Style.BackColor = [System.Drawing.Color]::LightPink
                }
            }
        }
    }
}

function Sort-ByDataOrder {
    $txtLog.AppendText("`r`nSorting by original data file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $false
    }
}

function Sort-ByMappingOrder {
    $txtLog.AppendText("`r`nSorting by mapping file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $true
    }
}

# Button event handlers
$btnLoadFiles.Add_Click({ Load-Files })
$btnFilter.Add_Click({ Show-ErrorsOnly })
$btnMappedOnly.Add_Click({ Show-MappedOnly })
$btnReset.Add_Click({ Show-AllData })
$btnSortByData.Add_Click({ Sort-ByDataOrder })
$btnSortByMapping.Add_Click({ Sort-ByMappingOrder })
$btnExport.Add_Click({ Export-Data })
$btnExportErrors.Add_Click({ Export-Errors })
$btnValidationReport.Add_Click({ Export-ValidationReport })

$form.Add_Shown({
    $txtLog.AppendText("`r`nApplication ready. Please click 'Load Files' to begin.")
})

Write-Host "Starting application..."
[void]$form.ShowDialog()) {
                                        try {
                                            $mappedValue = [datetime]::ParseExact(
                                                $mappedValue, 
                                                "dd/MM/yyyy HH:mm:ss", 
                                                [System.Globalization.CultureInfo]::InvariantCulture
                                            )
                                        }
                                        catch {
                                            $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                        }
                                    }
                                    else {
                                        $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                    }
                                }
                                default { }
                            }
                        } catch {
                            $error = "Data type conversion failed for $dataType"
                        }
                    }

                    # Use the mapped field name
                    $newField = $map.NewField
                    if ([string]::IsNullOrEmpty($newField)) {
                        $newField = $sourceField
                    }
                    
                    $newRow[$newField] = $mappedValue

                    # Record errors
                    if ($error) {
                        $log += "Row $($i+1) Field '$sourceField' (mapped to '$newField'): $error"
                        $global:ErrorCells["$i-$newField"] = $error
                    }
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    Update-ProgressBar 80 "Starting data integrity validation..."
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    Update-ProgressBar 95 "Compiling results..."
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    Update-ProgressBar 100 "Processing complete!"
    
    # Small delay to show completion
    Start-Sleep -Milliseconds 500
    Hide-ProgressBar
    
    return @{ Data = $result; Log = $log }
}# Type conversion
                    if ($dataType -and $mappedValue -ne $null -and $mappedValue -ne "") {
                        try {
                            switch ($dataType.ToLower()) {
                                "int" { $mappedValue = [int]$mappedValue }
                                "decimal" { $mappedValue = [decimal]$mappedValue }
                                "datetime" { 
                                    if ($mappedValue -match '^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$') {
                                        try {
                                            $mappedValue = [datetime]::ParseExact(
                                                $mappedValue, 
                                                "dd/MM/yyyy HH:mm:ss", 
                                                [System.Globalization.CultureInfo]::InvariantCulture
                                            )
                                        }
                                        catch {
                                            $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                        }
                                    }
                                    else {
                                        $error = "Invalid date format. Required format: dd/mm/yyyy HH:MM:SS"
                                    }
                                }
                                default { }
                            }
                        } catch {
                            $error = "Data type conversion failed for $dataType"
                        }
                    }

                    # Use the mapped field name
                    $newField = $map.NewField
                    if ([string]::IsNullOrEmpty($newField)) {
                        $newField = $sourceField
                    }
                    
                    $newRow[$newField] = $mappedValue

                    # Record errors
                    if ($error) {
                        $log += "Row $($i+1) Field '$sourceField' (mapped to '$newField'): $error"
                        $global:ErrorCells["$i-$newField"] = $error
                    }
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }

        # Handle Required fields (no SourceField)
        foreach ($m in $mapping | Where-Object { $_.Required -eq 'Y' -and [string]::IsNullOrEmpty($_.SourceField) }) {
            $value = $null
            $mappedValue = $null
            $error = $null

            if ($m.Transformation -eq 'Y' -and $m.TransformFunction) {
                if ($global:TransformFunctions.ContainsKey($m.TransformFunction)) {
                    try {
                        $mappedValue = & $global:TransformFunctions[$m.TransformFunction] $value
                    }
                    catch {
                        $error = "Error in transformation for required field (no SourceField): $_"
                    }
                }
                else {
                    Write-Warning "Transform function '$($m.TransformFunction)' not found in `$global:TransformFunctions; using null value."
                }
            }

            if (-not $mappedValue) {
                $mappedValue = ""
            }

            if ([string]::IsNullOrEmpty($mappedValue)) {
                $error = "Required field (no SourceField) is empty after transformation"
            }

            $newField = $m.NewField
            if ([string]::IsNullOrEmpty($newField)) {
                $newField = "UNDEFINED_FIELD"
            }

            $newRow[$newField] = $mappedValue

            if ($error) {
                $log += "Row $($i+1) Field '$($m.NewField)': $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }

        # Check for missing mandatory fields
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                foreach ($map in $mapLookup[$mandatoryField]) {
                    $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                    
                    $newRow[$newField] = ""
                    
                    $error = "Required field '$mandatoryField' is missing from input data"
                    $log += "Row $($i+1): $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    # Perform data validation after processing
    $validationResults = Validate-DataIntegrity $originalFingerprints $result
    $global:DataValidationResults = $validationResults
    
    # Add validation results to the log
    $summary = $validationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
    if ($summary) {
        $log += "=== DATA VALIDATION RESULTS ==="
        $log += "Data Integrity Status: $($summary.DataIntegrityStatus)"
        $log += "Original rows: $($summary.OriginalRowCount), Processed rows: $($summary.ProcessedRowCount)"
        
        if ($summary.MissingRowCount -gt 0) {
            $log += "WARNING: $($summary.MissingRowCount) rows from original data are missing in processed data"
        }
        
        if ($summary.UnmatchedRowCount -gt 0) {
            $log += "WARNING: $($summary.UnmatchedRowCount) rows in processed data don't clearly match original data"
        }
        
        if ($summary.DataIntegrityStatus -eq "PASSED") {
            $log += "All original data rows are preserved in processed output"
        }
    }
    
    return @{ Data = $result; Log = $log }
}

# UI Creation
Write-Host "Creating UI components..."

$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Transformation Tool with Validation"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)

$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($tableLayout)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($buttonPanel, 0, 0)

# Create buttons
$btnLoadFiles = New-Object System.Windows.Forms.Button
$btnLoadFiles.Text = "Load Files"
$btnLoadFiles.Width = 120
$btnLoadFiles.Location = New-Object System.Drawing.Point(10, 10)
$buttonPanel.Controls.Add($btnLoadFiles)

$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Text = "Show Errors Only"
$btnFilter.Width = 120
$btnFilter.Location = New-Object System.Drawing.Point(140, 10)
$buttonPanel.Controls.Add($btnFilter)

$btnMappedOnly = New-Object System.Windows.Forms.Button
$btnMappedOnly.Text = "Show Mapped Only"
$btnMappedOnly.Width = 120
$btnMappedOnly.Location = New-Object System.Drawing.Point(270, 10)
$buttonPanel.Controls.Add($btnMappedOnly)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Show All Data"
$btnReset.Width = 120
$btnReset.Location = New-Object System.Drawing.Point(400, 10)
$buttonPanel.Controls.Add($btnReset)

$btnSortByData = New-Object System.Windows.Forms.Button
$btnSortByData.Text = "Sort by Data"
$btnSortByData.Width = 120
$btnSortByData.Location = New-Object System.Drawing.Point(530, 10)
$buttonPanel.Controls.Add($btnSortByData)

$btnSortByMapping = New-Object System.Windows.Forms.Button
$btnSortByMapping.Text = "Sort by Mapping"
$btnSortByMapping.Width = 120
$btnSortByMapping.Location = New-Object System.Drawing.Point(660, 10)
$buttonPanel.Controls.Add($btnSortByMapping)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Width = 120
$btnExport.Location = New-Object System.Drawing.Point(790, 10)
$buttonPanel.Controls.Add($btnExport)

$btnExportErrors = New-Object System.Windows.Forms.Button
$btnExportErrors.Text = "Export Errors"
$btnExportErrors.Width = 120
$btnExportErrors.Location = New-Object System.Drawing.Point(920, 10)
$buttonPanel.Controls.Add($btnExportErrors)

$btnValidationReport = New-Object System.Windows.Forms.Button
$btnValidationReport.Text = "Validation Report"
$btnValidationReport.Width = 120
$btnValidationReport.Location = New-Object System.Drawing.Point(1050, 10)
$buttonPanel.Controls.Add($btnValidationReport)

# Create content panels
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($contentPanel, 0, 1)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 500
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$splitContainer.SplitterWidth = 5
$contentPanel.Controls.Add($splitContainer)

$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel1.Controls.Add($gridPanel)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Text = "Ready. Click 'Load Files' to begin."

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackColor = [System.Drawing.Color]::White
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.RowHeadersWidth = 60
$grid.RowHeadersVisible = $true
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 30
$grid.DefaultCellStyle.NullValue = ""
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true

# Event handlers
$grid.Add_DataError({
    param($sender, $e)
    if ($txtLog) {
        $txtLog.AppendText("`r`nGrid display error in cell [Row $($e.RowIndex), Column $($e.ColumnIndex)]: $($e.Exception.Message)")
    }
    $e.ThrowException = $false
})

$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.RowHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.RowHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridPanel.Controls.Add($grid)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel2.Controls.Add($logPanel)
$logPanel.Controls.Add($txtLog)

# Additional grid event handlers
$grid.Add_RowPostPaint({
    param($sender, $e)
    $rowNumber = ($e.RowIndex + 1).ToString()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($rowNumber, $grid.Font)
    
    $x = $e.RowBounds.Location.X + ($grid.RowHeadersWidth - $size.Width) / 2
    $y = $e.RowBounds.Location.Y + ($e.RowBounds.Height - $size.Height) / 2
    
    $location = New-Object System.Drawing.PointF($x, $y)
    $e.Graphics.DrawString($rowNumber, $grid.Font, $brush, $location)
})

$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    $hitTestInfo = $grid.HitTest($e.X, $e.Y)
    if ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
        $grid.ClearSelection()
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $grid.Rows[$i].Cells[$hitTestInfo.ColumnIndex].Selected = $true
        }
    }
})

$grid.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $grid.ClearSelection()
        $grid.Rows[$e.RowIndex].Selected = $true
        $txtLog.AppendText("`r`nSelected row $($e.RowIndex + 1) via row header")
    }
})

$grid.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    
    $rowIndex = $e.RowIndex
    $colName = if ($e.ColumnIndex -lt $grid.Columns.Count) { 
        $grid.Columns[$e.ColumnIndex].Name 
    } else { 
        return 
    }
    
    $key = "$rowIndex-$colName"
    if ($global:ErrorCells.ContainsKey($key)) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightPink
        $e.FormattingApplied = $true
    }
})

# Function definitions
function Load-Files {
    $global:ErrorCells.Clear()
    $global:ShowOnlyMappedFields = $false   
    $txtLog.Text = "Starting file loading process..."

    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    
    # Select Mapping CSV
    $openFileDialog.Title = "Select Mapping CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled mapping file selection."
        return
    }
    
    $mappingPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading mapping file: $mappingPath"
    
    $mapping = Load-Mapping $mappingPath
    if (-not $mapping -or $mapping.Count -eq 0) {
        $txtLog.Text += "`r`nMapping file is null or empty. Loading failed."
        return
    }
    
    $global:Mapping = $mapping
    $txtLog.Text += "`r`nLoaded mapping with $($mapping.Count) row(s)"
    
    # Build the source-to-new field map
    $global:SourceToNewFieldMap.Clear()
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
            if (-not $global:SourceToNewFieldMap.ContainsKey($m.SourceField)) {
                $global:SourceToNewFieldMap[$m.SourceField] = @()
            }
            $global:SourceToNewFieldMap[$m.SourceField] += $m.NewField
            $txtLog.Text += "`r`nField mapping: '$($m.SourceField)' => '$($m.NewField)'"
        }
    }
    
    $multiMappedFields = 0
    foreach ($sourceField in $global:SourceToNewFieldMap.Keys) {
        if ($global:SourceToNewFieldMap[$sourceField].Count -gt 1) {
            $multiMappedFields++
            $txtLog.Text += "`r`nMulti-mapped field: '$sourceField' => " + ($global:SourceToNewFieldMap[$sourceField] -join ", ")
        }
    }
    
    $txtLog.Text += "`r`nTotal source fields with mappings: $($global:SourceToNewFieldMap.Count) (including $multiMappedFields with multiple mappings)"
    $txtLog.Text += "`r`nOriginal column order from CSV: $($global:OriginalColumnOrder -join ', ')"
    
    # Extract column order from mapping
    $global:MappingColumnOrder = @()
    foreach ($m in $mapping) {
        $fieldName = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $m.SourceField }
        if (-not [string]::IsNullOrEmpty($fieldName) -and $global:MappingColumnOrder -notcontains $fieldName) {
            $global:MappingColumnOrder += $fieldName
        }
    }
    
    $txtLog.Text += "`r`nMapping column order: $($global:MappingColumnOrder -join ', ')"
    
    # Select Data CSV
    $openFileDialog.Title = "Select Data CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled data file selection."
        return
    }
    
    $dataPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading data file: $dataPath"
    
    $data = Load-Data $dataPath
    if (-not $data -or $data.Count -eq 0) {
        $txtLog.Text += "`r`nData file is null or empty."
        return
    }
    
    $txtLog.Text += "`r`nLoaded data with $($data.Count) row(s)"
    
    # Process data
    $txtLog.Text += "`r`nProcessing data..."
    try {
        $result = Process-Data $data $global:Mapping
        $global:ProcessedData = $result.Data
        $global:LogData = $result.Log
        
        if ($global:LogData.Count -gt 0) {
            $txtLog.Text += "`r`n--- Processing Results ---"
            foreach ($error in $global:LogData) {
                $txtLog.Text += "`r`n$error"
            }
        } else {
            $txtLog.Text += "`r`nNo validation errors found."
        }
    } catch {
        $txtLog.Text += "`r`nError during processing: $_"
        $global:ProcessedData = $data
    }
    
    # Display in grid
    if ($global:ProcessedData.Count -gt 0) {
        $firstRow = $global:ProcessedData[0]
        $columns = $firstRow.PSObject.Properties.Name
        
        $txtLog.Text += "`r`nDetected columns in processed data: $($columns -join ', ')"
        
        Show-MappedOnly
        
        $txtLog.Text += "`r`nData loaded successfully. Found $($global:ProcessedData.Count) records."
    } else {
        $txtLog.Text += "`r`nNo data to display"
    }
}

function Show-ErrorsOnly {
    $txtLog.Text += "`r`nFiltering to show only rows with errors..."
    
    if ($global:ErrorCells.Count -eq 0) {
        $txtLog.Text += "`r`nNo error cells detected. Nothing to filter."
        return
    }
    
    $errorRows = @{}
    foreach ($key in $global:ErrorCells.Keys) {
        $parts = $key -split "-"
        if ($parts.Count -eq 2) {
            $rowIndex = [int]$parts[0]
            $errorRows[$rowIndex] = $true
        }
    }
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $errorRows.ContainsKey($i)
    }
    
    $txtLog.Text += "`r`nShowing only rows with errors."
}

function Show-MappedOnly {
    $txtLog.Text += "`r`nFiltering to show only fields specified in the mapping..."
    
    if ($null -eq $global:Mapping -or $global:Mapping.Count -eq 0) {
        $txtLog.Text += "`r`nNo mapping loaded. Nothing to filter."
        return
    }
    
    $global:ShowOnlyMappedFields = $true
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing only mapped fields."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nNo data loaded to reprocess."
    }
}

function Show-AllData {
    $txtLog.Text += "`r`nShowing all rows..."
    
    $global:ShowOnlyMappedFields = $false
    
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $true
    }
    
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            Update-GridView
            
            $txtLog.Text += "`r`nShowing all fields and rows."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nAll rows are now visible."
    }
}

function Export-Data {
    if ($null -eq $global:ProcessedData -or $global:ProcessedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.")
        return
    }
    
    if ($global:ErrorCells.Count -gt 0) {
        $errorCount = $global:ErrorCells.Count
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            "Warning: Your data contains $errorCount error(s). Do you still want to proceed with the export?",
            "Export Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
            $txtLog.Text += "`r`nExport cancelled due to validation errors."
            return
        }
    }
    
    if ($global:DataValidationResults.Count -gt 0) {
        $summary = $global:DataValidationResults | Where-Object { $_.Type -eq "ValidationSummary" } | Select-Object -First 1
        if ($summary -and $summary.DataIntegrityStatus -ne "PASSED") {
            $validationWarning = [System.Windows.Forms.MessageBox]::Show(
                "Warning: Data validation detected potential issues:`n" +
                "- Original rows: $($summary.OriginalRowCount)`n" +
                "- Processed rows: $($summary.ProcessedRowCount)`n" +
                "- Missing rows: $($summary.MissingRowCount)`n" +
                "- Unmatched rows: $($summary.UnmatchedRowCount)`n`n" +
                "Do you still want to proceed with the export?",
                "Data Validation Warning",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($validationWarning -eq [System.Windows.Forms.DialogResult]::No) {
                $txtLog.Text += "`r`nExport cancelled due to data validation issues."
                return
            }
        }
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Processed Data"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:ProcessedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nData exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting data: $_"
        }
    }
}

function Export-Errors {
    if ($null -eq $global:LogData -or $global:LogData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No errors to export.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Error Log"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $errorObjects = @()
            foreach ($error in $global:LogData) {
                $errorObjects += [PSCustomObject]@{ ErrorMessage = $error }
            }
            
            $errorObjects | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Errors exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nErrors exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting errors: $_"
        }
    }
}

function Export-ValidationReport {
    if ($null -eq $global:DataValidationResults -or $global:DataValidationResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No validation results available. Please load and process data first.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Data Validation Report"
    $saveFileDialog.FileName = "DataValidationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $success = Export-DataValidationReport $global:DataValidationResults $saveFileDialog.FileName
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Validation report exported successfully to $($saveFileDialog.FileName)")
                $txtLog.Text += "`r`nValidation report exported successfully to $($saveFileDialog.FileName)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error exporting validation report.")
                $txtLog.Text += "`r`nError exporting validation report."
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting validation report: $_")
            $txtLog.Text += "`r`nError exporting validation report: $_"
        }
    }
}

function Update-GridView {
    param([bool]$useMappingOrder = $false)
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $grid.Columns.Clear()
        
        $firstRow = $global:ProcessedData[0]
        $availableColumns = $firstRow.PSObject.Properties.Name
        
        $orderedColumns = @()
        
        if ($useMappingOrder -and $global:MappingColumnOrder.Count -gt 0) {
            $txtLog.AppendText("`r`nUsing mapping file column order")
            
            foreach ($col in $global:MappingColumnOrder) {
                if ($availableColumns -contains $col) {
                    $orderedColumns += $col
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        } 
        else {
            $txtLog.AppendText("`r`nUsing data file column order with renamed fields")
            
            foreach ($origCol in $global:OriginalColumnOrder) {
                $targetCol = $origCol
                
                if ($global:SourceToNewFieldMap.ContainsKey($origCol)) {
                    $targetCol = $global:SourceToNewFieldMap[$origCol]
                    $txtLog.AppendText("`r`nTranslating column '$origCol' to '$targetCol'")
                }
                
                if ($availableColumns -contains $targetCol -and $orderedColumns -notcontains $targetCol) {
                    $orderedColumns += $targetCol
                }
            }
            
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        }
        
        if ($orderedColumns.Count -eq 0) {
            $orderedColumns = $availableColumns
        }
        
        $txtLog.AppendText("`r`nFinal column order: $($orderedColumns -join ', ')")
        
        # Create columns
        foreach ($colName in $orderedColumns) {
            $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $column.HeaderText = $colName
            $column.Name = $colName
            $column.Width = 120
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $grid.Columns.Add($column)
        }
        
        # Add rows
        for ($rowIdx = 0; $rowIdx -lt $global:ProcessedData.Count; $rowIdx++) {
            $rowObj = $global:ProcessedData[$rowIdx]
            
            $grid.Rows.Add() | Out-Null
            $currentRow = $grid.Rows[$grid.Rows.Count - 1]
            
            if ($rowIdx % 2 -eq 1) {
                $currentRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
            }
            
            for ($colIdx = 0; $colIdx -lt $orderedColumns.Count; $colIdx++) {
                $colName = $orderedColumns[$colIdx]
                
                if ($null -eq $rowObj.$colName) {
                    $currentRow.Cells[$colIdx].Value = [DBNull]::Value
                } else {
                    $currentRow.Cells[$colIdx].Value = "$($rowObj.$colName)"
                }
            }
        }
        
        # Mark error cells
        foreach ($key in $global:ErrorCells.Keys) {
            $parts = $key -split "-"
            if ($parts.Count -eq 2) {
                $rowIndex = [int]$parts[0]
                $colName = $parts[1]
                
                $colIndex = -1
                for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                    if ($grid.Columns[$i].Name -eq $colName) {
                        $colIndex = $i
                        break
                    }
                }
                
                if ($rowIndex -lt $grid.Rows.Count -and $colIndex -ge 0) {
                    $grid.Rows[$rowIndex].Cells[$colIndex].Style.BackColor = [System.Drawing.Color]::LightPink
                }
            }
        }
    }
}

function Sort-ByDataOrder {
    $txtLog.AppendText("`r`nSorting by original data file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $false
    }
}

function Sort-ByMappingOrder {
    $txtLog.AppendText("`r`nSorting by mapping file column order")
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $true
    }
}

# Button event handlers
$btnLoadFiles.Add_Click({ Load-Files })
$btnFilter.Add_Click({ Show-ErrorsOnly })
$btnMappedOnly.Add_Click({ Show-MappedOnly })
$btnReset.Add_Click({ Show-AllData })
$btnSortByData.Add_Click({ Sort-ByDataOrder })
$btnSortByMapping.Add_Click({ Sort-ByMappingOrder })
$btnExport.Add_Click({ Export-Data })
$btnExportErrors.Add_Click({ Export-Errors })
$btnValidationReport.Add_Click({ Export-ValidationReport })

$form.Add_Shown({
    $txtLog.AppendText("`r`nApplication ready. Please click 'Load Files' to begin.")
})

Write-Host "Starting application..."
[void]$form.ShowDialog()
