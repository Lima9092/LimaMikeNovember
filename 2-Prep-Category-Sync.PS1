<#
.SYNOPSIS
Category Syncing Tool - Sync categories between CSV files with matching column names

.DESCRIPTION
Simple tool to sync LMS category data from a source CSV to a destination CSV.
Uses barcode as reference and automatically adds missing columns.
#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Category Syncing Tool - CSV to CSV Transfer"
$form.Size = New-Object System.Drawing.Size(750, 700)
$form.StartPosition = 'CenterScreen'

# Source CSV controls
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = 'Source CSV (with categories):'
$lblSource.AutoSize = $true
$lblSource.Location = '10,15'

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = '10,35'
$txtSource.Width = 500
$txtSource.ReadOnly = $true

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = 'Browse...'
$btnBrowseSource.Location = '520,33'
$btnBrowseSource.Width = 75

# Destination CSV controls
$lblDestination = New-Object System.Windows.Forms.Label
$lblDestination.Text = 'Destination CSV (to update):'
$lblDestination.AutoSize = $true
$lblDestination.Location = '10,75'

$txtDestination = New-Object System.Windows.Forms.TextBox
$txtDestination.Location = '10,95'
$txtDestination.Width = 500
$txtDestination.ReadOnly = $true

$btnBrowseDestination = New-Object System.Windows.Forms.Button
$btnBrowseDestination.Text = 'Browse...'
$btnBrowseDestination.Location = '520,93'
$btnBrowseDestination.Width = 75

# Field selection (same for both files)
$lblBarcode = New-Object System.Windows.Forms.Label
$lblBarcode.Text = 'Barcode Field:'
$lblBarcode.AutoSize = $true
$lblBarcode.Location = '10,135'

$comboBarcodeField = New-Object System.Windows.Forms.ComboBox
$comboBarcodeField.Location = '100,132'
$comboBarcodeField.Width = 150
$comboBarcodeField.DropDownStyle = 'DropDownList'
$comboBarcodeField.Enabled = $false

$lblCategory = New-Object System.Windows.Forms.Label
$lblCategory.Text = 'Category Field:'
$lblCategory.AutoSize = $true
$lblCategory.Location = '270,135'

$comboCategoryField = New-Object System.Windows.Forms.ComboBox
$comboCategoryField.Location = '360,132'
$comboCategoryField.Width = 150
$comboCategoryField.DropDownStyle = 'DropDownList'
$comboCategoryField.Enabled = $false

# Additional fields to copy (optional)
$lblExtraFields = New-Object System.Windows.Forms.Label
$lblExtraFields.Text = 'Additional Fields to Copy:'
$lblExtraFields.AutoSize = $true
$lblExtraFields.Location = '10,170'

$clbExtraFields = New-Object System.Windows.Forms.CheckedListBox
$clbExtraFields.Location = '10,190'
$clbExtraFields.Size = New-Object System.Drawing.Size(500,80)
$clbExtraFields.Enabled = $false

# Sync options
$lblOptions = New-Object System.Windows.Forms.Label
$lblOptions.Text = 'Sync Options:'
$lblOptions.AutoSize = $true
$lblOptions.Location = '10,285'

$chkOverwrite = New-Object System.Windows.Forms.CheckBox
$chkOverwrite.Text = 'Overwrite existing data'
$chkOverwrite.Location = '100,285'
$chkOverwrite.AutoSize = $true
$chkOverwrite.Checked = $true

$chkEmptyOnly = New-Object System.Windows.Forms.CheckBox
$chkEmptyOnly.Text = 'Update empty fields only'
$chkEmptyOnly.Location = '280,285'
$chkEmptyOnly.AutoSize = $true

# Output controls
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = 'Output File:'
$lblOutput.AutoSize = $true
$lblOutput.Location = '10,320'

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = '10,340'
$txtOutput.Width = 500
$txtOutput.ReadOnly = $true

$btnSaveAs = New-Object System.Windows.Forms.Button
$btnSaveAs.Text = 'Save As...'
$btnSaveAs.Location = '520,338'
$btnSaveAs.Width = 75
$btnSaveAs.Enabled = $false

# Processing controls
$btnSync = New-Object System.Windows.Forms.Button
$btnSync.Text = 'Start Category Sync'
$btnSync.Location = '290,375'
$btnSync.Width = 150
$btnSync.Enabled = $false

# Status and progress
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Status: Ready'
$lblStatus.AutoSize = $true
$lblStatus.Location = '10,415'

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = 'Progress: 0 / 0'
$lblProgress.AutoSize = $true
$lblProgress.Location = '200,415'

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = '10,435'
$progressBar.Size = New-Object System.Drawing.Size(700,20)

# Results summary
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = 'Sync Summary:'
$lblSummary.AutoSize = $true
$lblSummary.Location = '10,465'

$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Location = '10,485'
$txtSummary.Size = New-Object System.Drawing.Size(700,20)
$txtSummary.ReadOnly = $true
$txtSummary.Text = 'Load source and destination files to see sync preview'

# Log box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = '10,515'
$txtLog.Size = New-Object System.Drawing.Size(700,100)
$txtLog.ReadOnly = $true

# Add controls to form
$controls = @(
    $lblSource, $txtSource, $btnBrowseSource,
    $lblDestination, $txtDestination, $btnBrowseDestination,
    $lblBarcode, $comboBarcodeField, $lblCategory, $comboCategoryField,
    $lblExtraFields, $clbExtraFields,
    $lblOptions, $chkOverwrite, $chkEmptyOnly,
    $lblOutput, $txtOutput, $btnSaveAs,
    $btnSync, $lblStatus, $lblProgress, $progressBar,
    $lblSummary, $txtSummary, $txtLog
)
foreach ($c in $controls) { $form.Controls.Add($c) }

# Global storage
$global:SourceData = $null
$global:DestinationData = $null
$global:CategoryLookup = @{}

# Function to ensure destination has all required columns
function Ensure-DestinationColumns {
    if (-not $global:SourceData -or -not $global:DestinationData) { return }
    
    # Get all columns from source data
    $sourceColumns = $global:SourceData[0].PSObject.Properties.Name
    $destColumns = $global:DestinationData[0].PSObject.Properties.Name
    
    $missingColumns = @()
    foreach ($sourceCol in $sourceColumns) {
        if ($destColumns -notcontains $sourceCol) {
            $missingColumns += $sourceCol
        }
    }
    
    if ($missingColumns.Count -gt 0) {
        $txtLog.AppendText("Adding missing columns to destination: $($missingColumns -join ', ')`r`n")
        
        # Add missing columns to each destination row
        foreach ($destRow in $global:DestinationData) {
            foreach ($missingCol in $missingColumns) {
                $destRow | Add-Member -NotePropertyName $missingCol -NotePropertyValue '' -Force
            }
        }
        
        $txtLog.AppendText("Missing columns added successfully.`r`n")
    }
}

# Function to populate field options
function Update-FieldOptions {
    if (-not $global:SourceData) { return }
    
    $comboBarcodeField.Items.Clear()
    $comboCategoryField.Items.Clear()
    $clbExtraFields.Items.Clear()
    
    $global:SourceData[0].PSObject.Properties.Name | ForEach-Object { 
        $comboBarcodeField.Items.Add($_)
        $comboCategoryField.Items.Add($_)
        $clbExtraFields.Items.Add($_)
    }
    
    $comboBarcodeField.Enabled = $true
    $comboCategoryField.Enabled = $true
    $clbExtraFields.Enabled = $true
    
    # Auto-select common field names
    if ($comboBarcodeField.Items.Contains('Barcode')) {
        $comboBarcodeField.SelectedItem = 'Barcode'
    } elseif ($comboBarcodeField.Items.Contains('barcode')) {
        $comboBarcodeField.SelectedItem = 'barcode'
    }
    
    if ($comboCategoryField.Items.Contains('lms_category')) {
        $comboCategoryField.SelectedItem = 'lms_category'
    } elseif ($comboCategoryField.Items.Contains('category')) {
        $comboCategoryField.SelectedItem = 'category'
    }
}

# Function to validate and preview sync
function Update-SyncPreview {
    if (-not $global:SourceData -or -not $global:DestinationData -or 
        -not $comboBarcodeField.SelectedItem -or -not $comboCategoryField.SelectedItem) {
        return
    }
    
    $txtSummary.Text = "Analyzing sync requirements..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Build category lookup from source
    $global:CategoryLookup = @{}
    $sourceValidCount = 0
    $sourceDuplicates = 0
    
    foreach ($sourceRow in $global:SourceData) {
        $barcode = $sourceRow.($comboBarcodeField.SelectedItem)
        
        if ($barcode -and $barcode.Trim() -ne '') {
            $cleanBarcode = $barcode.Trim()
            if ($global:CategoryLookup.ContainsKey($cleanBarcode)) {
                $sourceDuplicates++
            } else {
                $global:CategoryLookup[$cleanBarcode] = $sourceRow
                $sourceValidCount++
            }
        }
    }
    
    # Analyze destination requirements
    $destTotal = $global:DestinationData.Count
    $destMatches = 0
    $destUpdates = 0
    $destSkipped = 0
    $destEmpty = 0
    
    foreach ($destRow in $global:DestinationData) {
        $barcode = $destRow.($comboBarcodeField.SelectedItem)
        
        if (-not $barcode -or $barcode.Trim() -eq '') {
            $destEmpty++
            continue
        }
        
        $cleanBarcode = $barcode.Trim()
        if ($global:CategoryLookup.ContainsKey($cleanBarcode)) {
            $destMatches++
            $sourceRow = $global:CategoryLookup[$cleanBarcode]
            
            # Check if update is needed for main category field
            $currentCategory = $destRow.($comboCategoryField.SelectedItem)
            $newCategory = $sourceRow.($comboCategoryField.SelectedItem)
            
            $shouldUpdate = $false
            if ($chkOverwrite.Checked) {
                $shouldUpdate = $true
            } elseif ($chkEmptyOnly.Checked) {
                $shouldUpdate = (-not $currentCategory -or $currentCategory.Trim() -eq '')
            }
            
            if ($shouldUpdate) {
                $destUpdates++
            } else {
                $destSkipped++
            }
        }
    }
    
    $destNoMatch = $destTotal - $destMatches - $destEmpty
    
    $txtSummary.Text = "Source: $sourceValidCount records | Destination: $destTotal records | Matches: $destMatches | Will update: $destUpdates | Skip: $destSkipped"
    
    $txtLog.AppendText("Sync Analysis Complete:`r`n")
    $txtLog.AppendText("SOURCE FILE:`r`n")
    $txtLog.AppendText("- Total source records: $($global:SourceData.Count)`r`n")
    $txtLog.AppendText("- Valid barcodes: $sourceValidCount`r`n")
    $txtLog.AppendText("- Duplicate barcodes: $sourceDuplicates`r`n")
    $txtLog.AppendText("DESTINATION FILE:`r`n")
    $txtLog.AppendText("- Total destination records: $destTotal`r`n")
    $txtLog.AppendText("- Barcode matches found: $destMatches`r`n")
    $txtLog.AppendText("- No barcode match: $destNoMatch`r`n")
    $txtLog.AppendText("- Empty barcodes: $destEmpty`r`n")
    $txtLog.AppendText("SYNC PLAN:`r`n")
    $txtLog.AppendText("- Records to update: $destUpdates`r`n")
    $txtLog.AppendText("- Records to skip: $destSkipped`r`n")
    
    if ($destUpdates -gt 0) {
        $btnSync.Enabled = $true
        $txtLog.AppendText("Ready to sync! Click 'Start Category Sync' to proceed.`r`n")
    } else {
        $txtLog.AppendText("No updates needed with current settings.`r`n")
    }
}

# Function to perform the category sync
function Start-CategorySync {
    if (-not $txtOutput.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file location.")
        return
    }
    
    $btnSync.Enabled = $false
    $lblStatus.Text = "Syncing categories..."
    
    $totalRecords = $global:DestinationData.Count
    $progressBar.Maximum = $totalRecords
    $progressBar.Value = 0
    
    $updatedCount = 0
    $skippedCount = 0
    $noMatchCount = 0
    
    # Get selected extra fields to copy
    $extraFields = @()
    foreach ($item in $clbExtraFields.CheckedItems) {
        $extraFields += $item
    }
    
    $txtLog.AppendText("`r`nStarting category sync...`r`n")
    $txtLog.AppendText("Fields to sync: $($comboCategoryField.SelectedItem)")
    if ($extraFields.Count -gt 0) {
        $txtLog.AppendText(", $($extraFields -join ', ')")
    }
    $txtLog.AppendText("`r`n")
    
    # Process each destination record
    for ($i = 0; $i -lt $totalRecords; $i++) {
        $destRow = $global:DestinationData[$i]
        $barcode = $destRow.($comboBarcodeField.SelectedItem)
        
        if ($barcode -and $barcode.Trim() -ne '') {
            $cleanBarcode = $barcode.Trim()
            if ($global:CategoryLookup.ContainsKey($cleanBarcode)) {
                $sourceRow = $global:CategoryLookup[$cleanBarcode]
                
                # Check if update is needed for main category field
                $currentCategory = $destRow.($comboCategoryField.SelectedItem)
                $newCategory = $sourceRow.($comboCategoryField.SelectedItem)
                
                $shouldUpdate = $false
                if ($chkOverwrite.Checked) {
                    $shouldUpdate = $true
                } elseif ($chkEmptyOnly.Checked) {
                    $shouldUpdate = (-not $currentCategory -or $currentCategory.Trim() -eq '')
                }
                
                if ($shouldUpdate) {
                    # Update the main category field
                    $destRow.($comboCategoryField.SelectedItem) = $newCategory
                    
                    # Update any extra selected fields
                    foreach ($field in $extraFields) {
                        if ($field -ne $comboCategoryField.SelectedItem) {
                            $currentValue = $destRow.$field
                            $newValue = $sourceRow.$field
                            
                            $shouldUpdateField = $false
                            if ($chkOverwrite.Checked) {
                                $shouldUpdateField = $true
                            } elseif ($chkEmptyOnly.Checked) {
                                $shouldUpdateField = (-not $currentValue -or $currentValue.Trim() -eq '')
                            }
                            
                            if ($shouldUpdateField) {
                                $destRow.$field = $newValue
                            }
                        }
                    }
                    
                    $updatedCount++
                } else {
                    $skippedCount++
                }
            } else {
                $noMatchCount++
            }
        }
        
        # Update progress
        $progressBar.Value = $i + 1
        $lblProgress.Text = "Progress: $($i + 1) / $totalRecords"
        
        # Update UI every 100 records
        if (($i + 1) % 100 -eq 0) {
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    
    # Save the updated data
    try {
        $global:DestinationData | Export-Csv -Path $txtOutput.Text -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Sync completed successfully!"
        $txtLog.AppendText("Sync completed successfully!`r`n")
        $txtLog.AppendText("FINAL RESULTS:`r`n")
        $txtLog.AppendText("- Records updated: $updatedCount`r`n")
        $txtLog.AppendText("- Records skipped: $skippedCount`r`n")
        $txtLog.AppendText("- No match found: $noMatchCount`r`n")
        $txtLog.AppendText("- File saved to: $($txtOutput.Text)`r`n")
        
        [System.Windows.Forms.MessageBox]::Show("Category sync completed!`r`n`r`nRecords updated: $updatedCount`r`nRecords skipped: $skippedCount`r`nNo match found: $noMatchCount`r`n`r`nFile saved to:`r`n$($txtOutput.Text)")
        
    } catch {
        $lblStatus.Text = "Error saving file"
        $txtLog.AppendText("Error saving file: $_`r`n")
        [System.Windows.Forms.MessageBox]::Show("Error saving file: $_")
    }
    
    $btnSync.Enabled = $true
}

# Browse Source event
$btnBrowseSource.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV Files (*.csv)|*.csv'
    $ofd.Title = 'Select Source CSV (with categories)'
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtSource.Text = $ofd.FileName
        try {
            $global:SourceData = Import-Csv -Path $ofd.FileName
            if ($global:SourceData.Count -gt 0) {
                Update-FieldOptions
                $txtLog.AppendText("Loaded source file: $($global:SourceData.Count) records`r`n")
                Update-SyncPreview
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to load source CSV: $_")
        }
    }
})

# Browse Destination event
$btnBrowseDestination.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV Files (*.csv)|*.csv'
    $ofd.Title = 'Select Destination CSV (to update)'
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtDestination.Text = $ofd.FileName
        try {
            $global:DestinationData = Import-Csv -Path $ofd.FileName
            if ($global:DestinationData.Count -gt 0) {
                $btnSaveAs.Enabled = $true
                $txtLog.AppendText("Loaded destination file: $($global:DestinationData.Count) records`r`n")
                
                # Ensure destination has all columns from source
                Ensure-DestinationColumns
                
                Update-SyncPreview
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to load destination CSV: $_")
        }
    }
})

# Save As event
$btnSaveAs.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV Files (*.csv)|*.csv'
    $sfd.Title = 'Save Updated CSV As...'
    if ($sfd.ShowDialog() -eq 'OK') { 
        $txtOutput.Text = $sfd.FileName
        Update-SyncPreview
    }
})

# Field selection events
$comboBarcodeField.Add_SelectedIndexChanged({ Update-SyncPreview })
$comboCategoryField.Add_SelectedIndexChanged({ Update-SyncPreview })

# Option change events
$chkOverwrite.Add_CheckedChanged({ Update-SyncPreview })
$chkEmptyOnly.Add_CheckedChanged({ 
    if ($chkEmptyOnly.Checked) { $chkOverwrite.Checked = $false }
    Update-SyncPreview 
})
$chkOverwrite.Add_CheckedChanged({ 
    if ($chkOverwrite.Checked) { $chkEmptyOnly.Checked = $false }
    Update-SyncPreview 
})

# Sync button event
$btnSync.Add_Click({ Start-CategorySync })

# Show the form
[void]$form.ShowDialog()
