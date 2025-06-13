# Media Type Classifier - GUI Application
# Classifies library items based on RCN and other field data

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuration
$MediaTypeFields = @('Media Type', 'ShelfMark')

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Media Type Classifier"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)

# Create button panel instead of toolbar
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Height = 50
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$buttonPanel.BackColor = [System.Drawing.SystemColors]::Control

$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = "Open CSV"
$openButton.Size = New-Object System.Drawing.Size(100, 30)
$openButton.Location = New-Object System.Drawing.Point(10, 10)
$openButton.UseVisualStyleBackColor = $true

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save CSV"
$saveButton.Size = New-Object System.Drawing.Size(100, 30)
$saveButton.Location = New-Object System.Drawing.Point(250, 10)
$saveButton.UseVisualStyleBackColor = $true

$processButton = New-Object System.Windows.Forms.Button
$processButton.Text = "Process Data"
$processButton.Size = New-Object System.Drawing.Size(120, 30)
$processButton.Location = New-Object System.Drawing.Point(120, 10)
$processButton.UseVisualStyleBackColor = $true

$buttonPanel.Controls.AddRange(@($openButton, $saveButton, $processButton))
$form.Controls.Add($buttonPanel)

# Create status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready - Load a CSV file to begin"
$statusBar.Items.Add($statusLabel)
$form.Controls.Add($statusBar)

# Create data grid view
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 70)
$dataGridView.Size = New-Object System.Drawing.Size(1160, 660)
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $false
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
$dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$form.Controls.Add($dataGridView)

# Global variables
$script:csvData = $null
$script:currentFilePath = ""

# Enhanced media type classification function
function ClassifyMediaType {
    param($row)
    
    # Get all column names for comprehensive searching
    $allColumns = $row.PSObject.Properties.Name
    
    # First, check for DVD in ANY column (highest priority)
    foreach ($col in $allColumns) {
        if ($row.PSObject.Properties[$col] -and $row.$col) {
            $colValue = $row.$col.ToString().ToLower()
            # Enhanced DVD detection - case insensitive, handles various formats
            if ($colValue -match '\bdvd\b|dvd[-\s]|blu[-\s]?ray|blue[-\s]?ray|blueray|video.*disc|digital.*video') {
                return 'DVD'
            }
        }
    }
    
    # Second, check for Fiction/Non-fiction in ANY column
    foreach ($col in $allColumns) {
        if ($row.PSObject.Properties[$col] -and $row.$col) {
            $colValue = $row.$col.ToString().ToLower()
            
            # Fiction detection - various forms
            if ($colValue -match '\bfiction\b|\bfic\b|novel|story|narrative|tale') {
                return 'AF'
            }
            
            # Non-fiction detection - various forms  
            if ($colValue -match '\bnon[-\s]?fiction\b|\bnonfiction\b|\bnon[-\s]?fic\b|\bnonfic\b') {
                return 'ANF'
            }
        }
    }
    
    # Third, check RCN column for barcode-based detection
    if ($row.PSObject.Properties['RCN'] -and $row.RCN) {
        $rcnValue = $row.RCN.ToString().Trim()
        
        # Media type detection for barcodes/ISBNs
        if ($rcnValue -match '^\d{13}$') {
            $prefix = $rcnValue.Substring(0, 3)
            
            # ISBN prefixes (skip - handled by separate ISBN lookup tool)
            if ($prefix -eq '978' -or $prefix -eq '979') {
                # Check if it's specifically an audiobook (non-book format)
                foreach ($col in $allColumns) {
                    if ($row.PSObject.Properties[$col] -and $row.$col) {
                        $colValue = $row.$col.ToString().ToLower()
                        if ($colValue -match 'audio.?book|cd.?book|spoken|sound.?recording|playaway') {
                            return 'AUD-ADULT'
                        }
                    }
                }
                # Skip other ISBN processing - let ISBN lookup tool handle books
                return $null
            }
            
            # Magazine/Periodical ISSN prefixes and patterns
            elseif ($prefix -match '^(977|976)' -or $rcnValue -match '^0\d{12}$') {
                return 'MAGAZ'
            }
            
            # DVD/Video EAN prefixes
            else {
                $videoTvPrefixes = @('503', '504', '505', '506', '507', '508', '509', '725', '750', '760', '761', '762', '763', '764', '765')
                if ($videoTvPrefixes -contains $prefix) {
                    return 'DVD'
                }
                
                # Major DVD distributor prefixes
                if ($rcnValue.StartsWith('883') -or $rcnValue.StartsWith('024') -or $rcnValue.StartsWith('031') -or $rcnValue.StartsWith('014')) {
                    return 'DVD'
                }
            }
        }
        
        # Check for 10-digit ISBN (skip - handled by ISBN lookup tool)
        elseif ($rcnValue -match '^\d{10}$') {
            foreach ($col in $allColumns) {
                if ($row.PSObject.Properties[$col] -and $row.$col) {
                    $colValue = $row.$col.ToString().ToLower()
                    if ($colValue -match 'audio.?book|cd.?book|spoken|sound.?recording|playaway') {
                        return 'AUD-ADULT'
                    }
                }
            }
            # Skip other 10-digit ISBN processing - let ISBN lookup tool handle books
            return $null
        }
        
        # Check for 8-digit ISSN (magazines/periodicals)
        elseif ($rcnValue -match '^\d{4}-?\d{4}$') {
            return 'MAGAZ'
        }
    }
    
    # Fourth, comprehensive content-based detection using ALL columns
    foreach ($col in $allColumns) {
        if ($row.PSObject.Properties[$col] -and $row.$col) {
            $colValue = $row.$col.ToString().ToLower()
            
            # Audiobook detection
            if ($colValue -match 'audio.?book|cd.?book|spoken|sound.?recording|playaway|unabridged|abridged') {
                return 'AUD-ADULT'
            }
            
            # CD/Music detection
            if ($colValue -match '\bcd\b|music|audio.?cd|compact.?disc') {
                return 'MCD'
            }
            
            # Magazine/Periodical detection
            if ($colValue -match 'magazine|periodical|journal|serial|issue|vol\.|volume|no\.|number') {
                return 'MAGAZ'
            }
			 
            # DVD Player detection
            if ($colValue -match 'dvd player') {
                return 'EQUIP'
            }
			
            # Board Game detection
            if ($colValue -match 'board game') {
                return 'BGAME'
            }
							 
            # Jigsaw detection
            if ($colValue -match 'jigsaw|puzzle') {
                return 'JIGSAW'
            }
			
            # Game detection
            if ($colValue -match 'game|ps2|ps3|ps4|ps5|xbox|nintendo|wii|wiiu|switch') {
                return 'GAME'
            }
        }
    }
    
    return $null
}

# Transform reference codes
function TransformReferenceCode {
    param($value)
    
    if (-not $value) { return $value }
    
    $trimmedValue = $value.ToString().Trim()
    switch ($trimmedValue) {
        'AREF' { return 'ANF' }
        'CREF' { return 'CNF' }
        default { return $trimmedValue }
    }
}

# Load CSV file
function LoadCSV {
    param($filePath)
    
    try {
        $statusLabel.Text = "Loading CSV file..."
        $script:csvData = Import-Csv -Path $filePath
        $script:currentFilePath = $filePath
        
        # Create DataTable for display
        $dataTable = New-Object System.Data.DataTable
        
        # Add columns
        if ($script:csvData.Count -gt 0) {
            $firstRow = $script:csvData[0]
            foreach ($property in $firstRow.PSObject.Properties) {
                $column = $dataTable.Columns.Add($property.Name, [string])
            }
        }
        
        # Add rows
        foreach ($row in $script:csvData) {
            $dataRow = $dataTable.NewRow()
            foreach ($property in $row.PSObject.Properties) {
                $dataRow[$property.Name] = if ($property.Value) { $property.Value.ToString() } else { "" }
            }
            $dataTable.Rows.Add($dataRow)
        }
        
        $dataGridView.DataSource = $dataTable
        $statusLabel.Text = "Loaded $($script:csvData.Count) records from $(Split-Path $filePath -Leaf)"
        
        # Highlight lms_category column if it exists
        if ($dataGridView.Columns.Contains("lms_category")) {
            $dataGridView.Columns["lms_category"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading CSV file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Error loading file"
    }
}

# Save CSV file
function SaveCSV {
    param($filePath)
    
    try {
        $statusLabel.Text = "Saving CSV file..."
        
        # Update csvData from DataGridView
        if ($dataGridView.DataSource -and $script:csvData) {
            $dataTable = $dataGridView.DataSource
            for ($i = 0; $i -lt $script:csvData.Count; $i++) {
                $row = $script:csvData[$i]
                $dataRow = $dataTable.Rows[$i]
                
                foreach ($property in $row.PSObject.Properties) {
                    $row.$($property.Name) = $dataRow[$property.Name]
                }
            }
        }
        
        $script:csvData | Export-Csv -Path $filePath -NoTypeInformation
        $statusLabel.Text = "Saved $($script:csvData.Count) records to $(Split-Path $filePath -Leaf)"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving CSV file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Error saving file"
    }
}

# Process data
function ProcessData {
    if (-not $script:csvData) {
        [System.Windows.Forms.MessageBox]::Show("Please load a CSV file first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Processing data..."
    $processedCount = 0
    
    # Check for existing category columns
    $hasLmsCategoryColumn = $script:csvData[0].PSObject.Properties.Name -contains "lms_category"
    $hasCategoryColumn = $script:csvData[0].PSObject.Properties.Name -contains "category"
    
    # Determine which column to use
    $categoryColumnName = ""
    if ($hasLmsCategoryColumn) {
        $categoryColumnName = "lms_category"
        $statusLabel.Text = "Processing data using 'lms_category' column..."
    } elseif ($hasCategoryColumn) {
        $categoryColumnName = "category"
        $statusLabel.Text = "Processing data using existing 'category' column (processed file detected)..."
    } else {
        # Add lms_category column if neither exists
        $categoryColumnName = "lms_category"
        foreach ($row in $script:csvData) {
            $row | Add-Member -MemberType NoteProperty -Name "lms_category" -Value ""
        }
        $statusLabel.Text = "Processing data with new 'lms_category' column..."
    }
    
    foreach ($row in $script:csvData) {
        $originalValue = ""
        if ($row.PSObject.Properties[$categoryColumnName] -and $row.$categoryColumnName) {
            $originalValue = $row.$categoryColumnName.ToString().Trim()
        }
        
        # Only process if the category field is blank or empty
        if (-not $originalValue -or $originalValue -eq "") {
            # Try to classify media type
            $mediaType = ClassifyMediaType $row
            if ($mediaType) {
                $row.$categoryColumnName = $mediaType
                $processedCount++
            }
        } else {
            # Apply reference code transformations to existing values if needed
            $transformedValue = TransformReferenceCode $originalValue
            if ($transformedValue -ne $originalValue) {
                $row.$categoryColumnName = $transformedValue
                $processedCount++
            }
        }
    }
    
    # Recreate the DataTable with updated data
    $dataTable = New-Object System.Data.DataTable
    
    # Add columns (including lms_category if it was added)
    if ($script:csvData.Count -gt 0) {
        $firstRow = $script:csvData[0]
        foreach ($property in $firstRow.PSObject.Properties) {
            $column = $dataTable.Columns.Add($property.Name, [string])
        }
    }
    
    # Add all rows with updated data
    foreach ($row in $script:csvData) {
        $dataRow = $dataTable.NewRow()
        foreach ($property in $row.PSObject.Properties) {
            $dataRow[$property.Name] = if ($property.Value) { $property.Value.ToString() } else { "" }
        }
        $dataTable.Rows.Add($dataRow)
    }
    
    # Update the DataGridView
    $dataGridView.DataSource = $dataTable
    
    # Highlight the category column being used
    if ($dataGridView.Columns.Contains($categoryColumnName)) {
        $dataGridView.Columns[$categoryColumnName].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
    }
    
    $statusLabel.Text = "Processing complete - $processedCount records updated in '$categoryColumnName' column"
    [System.Windows.Forms.MessageBox]::Show("Processing complete!`n`n$processedCount records were updated in the '$categoryColumnName' column.`n`nNote: Only blank/empty entries were processed to preserve existing data.", "Process Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Event handlers
$openButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file to open"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        LoadCSV $openFileDialog.FileName
    }
})

$saveButton.Add_Click({
    if (-not $script:csvData) {
        [System.Windows.Forms.MessageBox]::Show("No data to save.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save CSV file"
    
    # Fixed filename generation for compatibility with older PowerShell versions
    $saveFileDialog.FileName = if ($script:currentFilePath) { 
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:currentFilePath)
        "$baseName`_processed.csv" 
    } else { 
        "processed_data.csv" 
    }
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        SaveCSV $saveFileDialog.FileName
    }
})

$processButton.Add_Click({ ProcessData })

# Show form
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
