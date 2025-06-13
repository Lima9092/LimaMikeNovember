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

# ISBN Conversion Functions
function Convert-ISBN13ToISBN10 {
    param([string]$isbn13)
    
    if ([string]::IsNullOrWhiteSpace($isbn13)) {
        return $isbn13
    }
    
    # Remove any hyphens, spaces, or other non-digit characters
    $cleanISBN = $isbn13 -replace '[^0-9]', ''
    
    # Debug output
    Write-Host "Debug: Original ISBN: '$isbn13', Clean ISBN: '$cleanISBN', Length: $($cleanISBN.Length)"
    
    # Check if it's a valid 13-digit ISBN starting with 978 or 979
    if ($cleanISBN.Length -ne 13) {
        Write-Host "Debug: Not 13 digits, skipping"
        return $isbn13
    }
    
    if (-not ($cleanISBN.StartsWith('978') -or $cleanISBN.StartsWith('979'))) {
        Write-Host "Debug: Doesn't start with 978 or 979, skipping"
        return $isbn13
    }
    
    # Note: ISBN-10 can only be converted from ISBN-13 that starts with 978
    # ISBN-13 starting with 979 cannot be converted to ISBN-10 format
    if ($cleanISBN.StartsWith('979')) {
        Write-Host "Debug: ISBN-13 starts with 979 - cannot convert to ISBN-10 (no ISBN-10 equivalent exists)"
        return $isbn13
    }
    
    Write-Host "Debug: Valid ISBN-13 (978 prefix) detected, converting to ISBN-10..."
    
    # Extract the middle 9 digits (after 978 prefix, before check digit)
    $isbn10Base = $cleanISBN.Substring(3, 9)
    
    # Calculate ISBN-10 check digit
    $sum = 0
    for ($i = 0; $i -lt 9; $i++) {
        $digit = [int]::Parse($isbn10Base[$i].ToString())
        $sum += $digit * (10 - $i)
    }
    
    $checkDigit = (11 - ($sum % 11)) % 11
    $checkDigitStr = if ($checkDigit -eq 10) { 'X' } else { $checkDigit.ToString() }
    
    $result = $isbn10Base + $checkDigitStr
    Write-Host "Debug: Converted to ISBN-10: '$result'"
    
    return $result
}

function Test-IsISBN13 {
    param([string]$value)
    
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $false
    }
    
    $cleanValue = $value -replace '[^0-9]', ''
    $isISBN13 = ($cleanValue.Length -eq 13 -and ($cleanValue.StartsWith('978') -or $cleanValue.StartsWith('979')))
    
    Write-Host "Debug: Testing '$value' -> Clean: '$cleanValue' -> Length: $($cleanValue.Length) -> Starts with 978/979: $($cleanValue.StartsWith('978') -or $cleanValue.StartsWith('979')) -> Is ISBN-13: $isISBN13"
    
    return $isISBN13
}

# CSV Processor Tool
# This script provides a GUI to:
# 1. Import a CSV file
# 2. Select primary and secondary headers
# 3. Create a new column based on primary values, falling back to secondary values when primary is empty
# 4. Optionally convert ISBN-13 numbers to ISBN-10 format (ONLY in selected fields)
# 5. Preview data in a grid view
# 6. Save as a new CSV file

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV Processor Tool"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.MinimumSize = New-Object System.Drawing.Size(600, 550)
# Enable resizing and handle resize events
$form.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Add_Resize({
    # Adjust the spacing between controls when the form is resized
    # This ensures everything looks good at different sizes
    $dataGridView.Width = $form.ClientSize.Width - 40  # 20px padding on each side
})

# CSV Data container
$global:csvData = $null
$global:headers = @()
$global:fileName = ""

# Import CSV button
$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(20, 20)
$importButton.Size = New-Object System.Drawing.Size(150, 30)
$importButton.Text = "Import CSV"
$importButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select a CSV File"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        try {
            $global:fileName = $openFileDialog.FileName
            $global:csvData = Import-Csv -Path $global:fileName
            $global:headers = $global:csvData[0].PSObject.Properties.Name
            
            # Clear and populate dropdown lists
            $primaryDropdown.Items.Clear()
            $secondaryDropdown.Items.Clear()
            foreach ($header in $global:headers) {
                $primaryDropdown.Items.Add($header)
                $secondaryDropdown.Items.Add($header)
            }
            
            $fileNameLabel.Text = "File: $($openFileDialog.SafeFileName)"
            $statusLabel.Text = "Status: CSV imported successfully. Select headers and new column name."
            
            # Enable controls
            $primaryDropdown.Enabled = $true
            $secondaryDropdown.Enabled = $true
            $newColumnNameTextBox.Enabled = $true
            $isbnCheckbox.Enabled = $true
            $previewButton.Enabled = $true
        }
        catch {
            [System.Windows.MessageBox]::Show("Error importing CSV: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})
$form.Controls.Add($importButton)

# File name label
$fileNameLabel = New-Object System.Windows.Forms.Label
$fileNameLabel.Location = New-Object System.Drawing.Point(180, 25)
$fileNameLabel.Size = New-Object System.Drawing.Size(400, 20)
$fileNameLabel.Text = "File: No file selected"
$form.Controls.Add($fileNameLabel)

# Primary header selection
$primaryLabel = New-Object System.Windows.Forms.Label
$primaryLabel.Location = New-Object System.Drawing.Point(20, 70)
$primaryLabel.Size = New-Object System.Drawing.Size(150, 20)
$primaryLabel.Text = "Primary Header:"
$form.Controls.Add($primaryLabel)

$primaryDropdown = New-Object System.Windows.Forms.ComboBox
$primaryDropdown.Location = New-Object System.Drawing.Point(180, 70)
$primaryDropdown.Size = New-Object System.Drawing.Size(200, 20)
$primaryDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$primaryDropdown.Enabled = $false
$form.Controls.Add($primaryDropdown)

# Secondary header selection
$secondaryLabel = New-Object System.Windows.Forms.Label
$secondaryLabel.Location = New-Object System.Drawing.Point(20, 110)
$secondaryLabel.Size = New-Object System.Drawing.Size(150, 20)
$secondaryLabel.Text = "Secondary Header:"
$form.Controls.Add($secondaryLabel)

$secondaryDropdown = New-Object System.Windows.Forms.ComboBox
$secondaryDropdown.Location = New-Object System.Drawing.Point(180, 110)
$secondaryDropdown.Size = New-Object System.Drawing.Size(200, 20)
$secondaryDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$secondaryDropdown.Enabled = $false
$form.Controls.Add($secondaryDropdown)

# New column name
$newColumnLabel = New-Object System.Windows.Forms.Label
$newColumnLabel.Location = New-Object System.Drawing.Point(20, 150)
$newColumnLabel.Size = New-Object System.Drawing.Size(150, 20)
$newColumnLabel.Text = "New Column Name:"
$form.Controls.Add($newColumnLabel)

$newColumnNameTextBox = New-Object System.Windows.Forms.TextBox
$newColumnNameTextBox.Location = New-Object System.Drawing.Point(180, 150)
$newColumnNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
$newColumnNameTextBox.Enabled = $false
$form.Controls.Add($newColumnNameTextBox)

# ISBN Conversion checkbox
$isbnCheckbox = New-Object System.Windows.Forms.CheckBox
$isbnCheckbox.Location = New-Object System.Drawing.Point(400, 150)
$isbnCheckbox.Size = New-Object System.Drawing.Size(300, 20)
$isbnCheckbox.Text = "Convert ISBN-13 to ISBN-10 (selected fields only)"
$isbnCheckbox.Enabled = $false
$form.Controls.Add($isbnCheckbox)

# Preview button
$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Location = New-Object System.Drawing.Point(20, 190)
$previewButton.Size = New-Object System.Drawing.Size(150, 30)
$previewButton.Text = "Preview Data"
$previewButton.Enabled = $false
$previewButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$previewButton.Add_Click({
    if ([string]::IsNullOrEmpty($primaryDropdown.SelectedItem) -or [string]::IsNullOrEmpty($secondaryDropdown.SelectedItem) -or [string]::IsNullOrEmpty($newColumnNameTextBox.Text)) {
        [System.Windows.MessageBox]::Show("Please select primary and secondary headers and enter a new column name.", "Required Fields", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    try {
        $primaryHeader = $primaryDropdown.SelectedItem.ToString()
        $secondaryHeader = $secondaryDropdown.SelectedItem.ToString()
        $newColumnName = $newColumnNameTextBox.Text
        $convertISBN = $isbnCheckbox.Checked
        
        # Initialize summary tracking
        $summary = @{
            TotalRecords = 0
            PrimaryUsed = 0
            SecondaryUsed = 0
            EmptyPrimary = 0
            EmptySecondary = 0
            ISBN13Detected = 0
            ISBN13With978 = 0
            ISBN13With979 = 0
            ISBN13Converted = 0
            ConversionErrors = 0
            ColumnStats = @{}
        }
        
        # Initialize column stats for only the selected fields
        if ($convertISBN) {
            $summary.ColumnStats[$primaryHeader] = @{
                TotalValues = 0
                ISBN13Found = 0
                ISBN13Converted = 0
                Examples = @()
            }
            $summary.ColumnStats[$secondaryHeader] = @{
                TotalValues = 0
                ISBN13Found = 0
                ISBN13Converted = 0
                Examples = @()
            }
        }
        
        # Create a new collection with the additional column and apply ISBN conversion
        $processedData = @()
        
        foreach ($row in $global:csvData) {
            $summary.TotalRecords++
            
            # Create a new hashtable for the row
            $newRow = @{}
            
            # Copy all properties - only convert ISBNs in selected fields
            foreach ($property in $row.PSObject.Properties) {
                $value = $property.Value
                $propertyName = $property.Name
                
                # Only process ISBN conversion for selected primary and secondary fields
                if ($convertISBN -and ($propertyName -eq $primaryHeader -or $propertyName -eq $secondaryHeader)) {
                    if ($value -and $value.ToString().Length -gt 0) {
                        $summary.ColumnStats[$propertyName].TotalValues++
                        
                        $valueStr = $value.ToString()
                        Write-Host "Debug: Processing selected column '$propertyName' with value '$valueStr'"
                        
                        # Test if this looks like an ISBN-13
                        if (Test-IsISBN13 -value $valueStr) {
                            $summary.ISBN13Detected++
                            $summary.ColumnStats[$propertyName].ISBN13Found++
                            
                            $cleanValue = $valueStr -replace '[^0-9]', ''
                            if ($cleanValue.StartsWith('978')) {
                                $summary.ISBN13With978++
                            } elseif ($cleanValue.StartsWith('979')) {
                                $summary.ISBN13With979++
                            }
                            
                            Write-Host "Debug: Converting ISBN in selected column '$propertyName'"
                            try {
                                $convertedISBN = Convert-ISBN13ToISBN10 -isbn13 $valueStr
                                if ($convertedISBN -ne $valueStr) {
                                    $summary.ISBN13Converted++
                                    $summary.ColumnStats[$propertyName].ISBN13Converted++
                                    $summary.ColumnStats[$propertyName].Examples += "$valueStr -> $convertedISBN"
                                }
                                $newRow[$propertyName] = $convertedISBN
                                Write-Host "Debug: Set selected column '$propertyName' to '$convertedISBN'"
                            } catch {
                                $summary.ConversionErrors++
                                Write-Host "Debug: Error converting ISBN: $_"
                                $newRow[$propertyName] = $value
                            }
                        } else {
                            $newRow[$propertyName] = $value
                        }
                    } else {
                        $newRow[$propertyName] = $value
                    }
                } else {
                    # For all other columns, just copy the value without conversion
                    $newRow[$propertyName] = $value
                }
            }
            
            # Now get the primary and secondary values (after conversion if applicable)
            $primaryValue = $newRow[$primaryHeader]
            $secondaryValue = $newRow[$secondaryHeader]
            
            # Track empty values
            if ([string]::IsNullOrWhiteSpace($primaryValue)) {
                $summary.EmptyPrimary++
            }
            if ([string]::IsNullOrWhiteSpace($secondaryValue)) {
                $summary.EmptySecondary++
            }
            
            # Add the new column with the logic
            if ([string]::IsNullOrWhiteSpace($primaryValue)) {
                $newRow[$newColumnName] = $secondaryValue
                $summary.SecondaryUsed++
            } else {
                $newRow[$newColumnName] = $primaryValue
                $summary.PrimaryUsed++
            }
            
            # Convert hashtable back to PSCustomObject and add to array
            $processedData += [PSCustomObject]$newRow
        }
        
        # Display summary
        Write-Host "`n=== PROCESSING SUMMARY ==="
        Write-Host "Total Records Processed: $($summary.TotalRecords)"
        Write-Host "Primary Column Used: $($summary.PrimaryUsed) times"
        Write-Host "Secondary Column Used: $($summary.SecondaryUsed) times"
        Write-Host "Empty Primary Values: $($summary.EmptyPrimary)"
        Write-Host "Empty Secondary Values: $($summary.EmptySecondary)"
        
        if ($convertISBN) {
            Write-Host "`n=== ISBN CONVERSION SUMMARY (Selected Fields Only) ==="
            Write-Host "Selected Primary Field: $primaryHeader"
            Write-Host "Selected Secondary Field: $secondaryHeader"
            Write-Host "Total ISBN-13 Numbers Detected: $($summary.ISBN13Detected)"
            Write-Host "ISBN-13 with 978 prefix: $($summary.ISBN13With978)"
            Write-Host "ISBN-13 with 979 prefix: $($summary.ISBN13With979)"
            Write-Host "Successful ISBN-13 to ISBN-10 Conversions: $($summary.ISBN13Converted)"
            Write-Host "Conversion Errors: $($summary.ConversionErrors)"
            
            Write-Host "`n=== SELECTED FIELDS BREAKDOWN ==="
            foreach ($columnName in @($primaryHeader, $secondaryHeader)) {
                if ($summary.ColumnStats.ContainsKey($columnName)) {
                    $stats = $summary.ColumnStats[$columnName]
                    Write-Host "Column '$columnName':"
                    Write-Host "  - Total values: $($stats.TotalValues)"
                    Write-Host "  - ISBN-13 found: $($stats.ISBN13Found)"
                    Write-Host "  - ISBN-13 converted: $($stats.ISBN13Converted)"
                    if ($stats.Examples.Count -gt 0) {
                        Write-Host "  - Examples:"
                        $stats.Examples | ForEach-Object { Write-Host "    $_" }
                    }
                }
            }
            
            if ($summary.ISBN13Detected -eq 0) {
                Write-Host "WARNING: No ISBN-13 numbers were detected in the selected fields."
                Write-Host "Check that your selected fields contain 13-digit numbers starting with 978 or 979."
            }
        }
        Write-Host "========================`n"
        
        # Display in DataGridView
        $dataGridView.DataSource = $null
        $dataTable = New-Object System.Data.DataTable
        
        # Create columns
        $allHeaders = $processedData[0].PSObject.Properties.Name
        foreach ($header in $allHeaders) {
            $dataTable.Columns.Add($header)
        }
        
        # Add rows
        foreach ($item in $processedData) {
            $row = $dataTable.NewRow()
            foreach ($header in $allHeaders) {
                $row[$header] = $item.$header
            }
            $dataTable.Rows.Add($row)
        }
        
        $dataGridView.DataSource = $dataTable
        $saveButton.Enabled = $true
        $global:processedData = $processedData
        
        $statusMessage = "Status: Preview generated. Click 'Save CSV' to export the data."
        if ($convertISBN) {
            $statusMessage += " (ISBN-13 converted to ISBN-10 in selected fields only)"
        }
        $statusLabel.Text = $statusMessage
    }
    catch {
        [System.Windows.MessageBox]::Show("Error generating preview: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})
$form.Controls.Add($previewButton)

# DataGridView for preview
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 240)
$dataGridView.Size = New-Object System.Drawing.Size(740, 250)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($dataGridView)

# Save CSV button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Location = New-Object System.Drawing.Point(180, 190)
$saveButton.Size = New-Object System.Drawing.Size(150, 30)
$saveButton.Text = "Save CSV"
$saveButton.Enabled = $false
$saveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$saveButton.Add_Click({
    if ($null -eq $global:processedData) {
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV File"
    $saveFileDialog.DefaultExt = "csv"
    
    # Suggest a filename based on the original
    $fileInfo = [System.IO.FileInfo]$global:fileName
    $suggestedName = "$($fileInfo.Directory)\$($fileInfo.BaseName)_processed$($fileInfo.Extension)"
    $saveFileDialog.FileName = $suggestedName
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $global:processedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $statusLabel.Text = "Status: CSV saved successfully to $($saveFileDialog.FileName)"
        }
        catch {
            [System.Windows.MessageBox]::Show("Error saving CSV: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})
$form.Controls.Add($saveButton)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 510)
$statusLabel.Size = New-Object System.Drawing.Size(740, 40)
$statusLabel.Text = "Status: Import a CSV file to begin."
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)

# Global variable for processed data
$global:processedData = $null

# Show the form
$form.ShowDialog()
