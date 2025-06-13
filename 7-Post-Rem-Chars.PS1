Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:csvData = @()
$script:originalData = @()
$script:headers = @()
$script:filePath = ""
$script:selectedColumns = @{}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV Data Cleaner"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.BackColor = [System.Drawing.SystemColors]::Control

# Create main layout panel
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$mainPanel.RowCount = 3
$mainPanel.ColumnCount = 1
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80))) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 140))) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$form.Controls.Add($mainPanel)

# ===== TOP PANEL - File Operations =====
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Fill"
$topPanel.BackColor = [System.Drawing.SystemColors]::Control
$mainPanel.Controls.Add($topPanel, 0, 0)

# Import CSV Button
$importBtn = New-Object System.Windows.Forms.Button
$importBtn.Text = "Import CSV File"
$importBtn.Location = New-Object System.Drawing.Point(10, 10)
$importBtn.Size = New-Object System.Drawing.Size(100, 30)
$importBtn.BackColor = [System.Drawing.SystemColors]::Control
$importBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$importBtn.FlatStyle = "Standard"
$importBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$topPanel.Controls.Add($importBtn)

# File info label
$fileInfoLabel = New-Object System.Windows.Forms.Label
$fileInfoLabel.Text = "No file loaded - Click 'Import CSV File' to begin"
$fileInfoLabel.Location = New-Object System.Drawing.Point(140, 20)
$fileInfoLabel.Size = New-Object System.Drawing.Size(400, 20)
$fileInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$fileInfoLabel.ForeColor = [System.Drawing.SystemColors]::GrayText
$topPanel.Controls.Add($fileInfoLabel)

# Action buttons panel
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(10, 45)
$actionPanel.Size = New-Object System.Drawing.Size(800, 30)
$topPanel.Controls.Add($actionPanel)

$saveAsBtn = New-Object System.Windows.Forms.Button
$saveAsBtn.Text = "Save As..."
$saveAsBtn.Location = New-Object System.Drawing.Point(0, 0)
$saveAsBtn.Size = New-Object System.Drawing.Size(100, 30)
$saveAsBtn.BackColor = [System.Drawing.SystemColors]::Control
$saveAsBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$saveAsBtn.FlatStyle = "Standard"
$saveAsBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$saveAsBtn.Enabled = $false
$actionPanel.Controls.Add($saveAsBtn)

$resetBtn = New-Object System.Windows.Forms.Button
$resetBtn.Text = "Reset Data"
$resetBtn.Location = New-Object System.Drawing.Point(110, 0)
$resetBtn.Size = New-Object System.Drawing.Size(100, 30)
$resetBtn.BackColor = [System.Drawing.SystemColors]::Control
$resetBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$resetBtn.FlatStyle = "Standard"
$resetBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$resetBtn.Enabled = $false
$actionPanel.Controls.Add($resetBtn)

# ===== MIDDLE PANEL - Column Selection and Actions =====
$middlePanel = New-Object System.Windows.Forms.Panel
$middlePanel.Dock = "Fill"
$middlePanel.BackColor = [System.Drawing.SystemColors]::Control
$mainPanel.Controls.Add($middlePanel, 0, 1)

# Column selection label
$columnsLabel = New-Object System.Windows.Forms.Label
$columnsLabel.Text = "Select columns to clean (click to toggle):"
$columnsLabel.Location = New-Object System.Drawing.Point(10, 5)
$columnsLabel.Size = New-Object System.Drawing.Size(250, 20)
$columnsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$middlePanel.Controls.Add($columnsLabel)

# Selection status
$selectionStatusLabel = New-Object System.Windows.Forms.Label
$selectionStatusLabel.Text = ""
$selectionStatusLabel.Location = New-Object System.Drawing.Point(270, 5)
$selectionStatusLabel.Size = New-Object System.Drawing.Size(200, 20)
$selectionStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$selectionStatusLabel.ForeColor = [System.Drawing.SystemColors]::GrayText
$middlePanel.Controls.Add($selectionStatusLabel)

# Column buttons panel
$columnButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$columnButtonsPanel.Location = New-Object System.Drawing.Point(10, 25)
$columnButtonsPanel.Size = New-Object System.Drawing.Size(1150, 80)
$columnButtonsPanel.FlowDirection = "LeftToRight"
$columnButtonsPanel.WrapContents = $true
$columnButtonsPanel.AutoScroll = $true
$middlePanel.Controls.Add($columnButtonsPanel)

# Action buttons group
$actionGroupLabel = New-Object System.Windows.Forms.Label
$actionGroupLabel.Text = "Cleaning Actions:"
$actionGroupLabel.Location = New-Object System.Drawing.Point(10, 110)
$actionGroupLabel.Size = New-Object System.Drawing.Size(120, 20)
$actionGroupLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$middlePanel.Controls.Add($actionGroupLabel)

$delSpaceBtn = New-Object System.Windows.Forms.Button
$delSpaceBtn.Text = "Del Space"
$delSpaceBtn.Location = New-Object System.Drawing.Point(140, 108)
$delSpaceBtn.Size = New-Object System.Drawing.Size(90, 25)
$delSpaceBtn.BackColor = [System.Drawing.SystemColors]::Control
$delSpaceBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$delSpaceBtn.FlatStyle = "Standard"
$delSpaceBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$delSpaceBtn.Enabled = $false
$middlePanel.Controls.Add($delSpaceBtn)

$remCharsBtn = New-Object System.Windows.Forms.Button
$remCharsBtn.Text = "Rem Chars"
$remCharsBtn.Location = New-Object System.Drawing.Point(240, 108)
$remCharsBtn.Size = New-Object System.Drawing.Size(90, 25)
$remCharsBtn.BackColor = [System.Drawing.SystemColors]::Control
$remCharsBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$remCharsBtn.FlatStyle = "Standard"
$remCharsBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$remCharsBtn.Enabled = $false
$middlePanel.Controls.Add($remCharsBtn)

$hyphenSpaceBtn = New-Object System.Windows.Forms.Button
$hyphenSpaceBtn.Text = "Hyphen Space"
$hyphenSpaceBtn.Location = New-Object System.Drawing.Point(340, 108)
$hyphenSpaceBtn.Size = New-Object System.Drawing.Size(100, 25)
$hyphenSpaceBtn.BackColor = [System.Drawing.SystemColors]::Control
$hyphenSpaceBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
$hyphenSpaceBtn.FlatStyle = "Standard"
$hyphenSpaceBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$hyphenSpaceBtn.Enabled = $false
$middlePanel.Controls.Add($hyphenSpaceBtn)

# ===== BOTTOM PANEL - Data View =====
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock = "Fill"
$bottomPanel.BackColor = [System.Drawing.SystemColors]::Control
$mainPanel.Controls.Add($bottomPanel, 0, 2)

# Data view label
$dataLabel = New-Object System.Windows.Forms.Label
$dataLabel.Text = "Data Preview:"
$dataLabel.Location = New-Object System.Drawing.Point(10, 5)
$dataLabel.Size = New-Object System.Drawing.Size(200, 20)
$dataLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$bottomPanel.Controls.Add($dataLabel)

# Create DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 25)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = "AllCells"
$dataGridView.SelectionMode = "FullRowSelect"
$dataGridView.MultiSelect = $false
$dataGridView.RowHeadersVisible = $true
$dataGridView.ColumnHeadersHeightSizeMode = "AutoSize"
$dataGridView.AllowUserToResizeColumns = $true
$dataGridView.AllowUserToResizeRows = $false
$dataGridView.BackgroundColor = [System.Drawing.SystemColors]::Window
$dataGridView.GridColor = [System.Drawing.Color]::LightGray
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
$bottomPanel.Controls.Add($dataGridView)

# Status bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready - Import a CSV file to begin"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Functions
function Load-CSVFile {
    param([string]$path)
    
    try {
        $script:originalData = Import-Csv -Path $path
        $script:csvData = $script:originalData | ForEach-Object { $_.PSObject.Copy() }
        $script:headers = $script:csvData[0].PSObject.Properties.Name
        $script:filePath = $path
        $script:selectedColumns = @{}
        
        Create-ColumnButtons
        Update-DataGrid
        Update-FileLabel
        Enable-Controls
        $statusLabel.Text = "File loaded: " + $script:csvData.Count + " rows, " + $script:headers.Count + " columns"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading file: " + $_.Exception.Message, "Error", "OK", "Error")
        $statusLabel.Text = "Error loading file"
    }
}

function Create-ColumnButtons {
    $columnButtonsPanel.Controls.Clear()
    $script:selectedColumns = @{}
    
    foreach ($header in $script:headers) {
        $columnBtn = New-Object System.Windows.Forms.Button
        $columnBtn.Text = $header
        $columnBtn.Size = New-Object System.Drawing.Size(120, 30)
        $columnBtn.BackColor = [System.Drawing.SystemColors]::Control
        $columnBtn.ForeColor = [System.Drawing.SystemColors]::ControlText
        $columnBtn.FlatStyle = "Standard"
        $columnBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $columnBtn.Margin = New-Object System.Windows.Forms.Padding(2)
        $columnBtn.Tag = $header
        $script:selectedColumns[$header] = $false
        
        # Add click event
        $columnBtn.Add_Click({
            param($sender, $e)
            $columnName = $sender.Tag
            $script:selectedColumns[$columnName] = -not $script:selectedColumns[$columnName]
            
            if ($script:selectedColumns[$columnName]) {
                $sender.BackColor = [System.Drawing.Color]::LightBlue
                $sender.ForeColor = [System.Drawing.Color]::Black
            } else {
                $sender.BackColor = [System.Drawing.SystemColors]::Control
                $sender.ForeColor = [System.Drawing.SystemColors]::ControlText
            }
            
            Update-SelectionStatus
        })
        
        $columnButtonsPanel.Controls.Add($columnBtn)
    }
}

function Update-SelectionStatus {
    $selectedCount = ($script:selectedColumns.Values | Where-Object { $_ -eq $true }).Count
    if ($selectedCount -eq 0) {
        $selectionStatusLabel.Text = "No columns selected"
        $selectionStatusLabel.ForeColor = [System.Drawing.SystemColors]::GrayText
        $delSpaceBtn.Enabled = $false
        $remCharsBtn.Enabled = $false
        $hyphenSpaceBtn.Enabled = $false
    } else {
        $selectionStatusLabel.Text = $selectedCount.ToString() + " column(s) selected"
        $selectionStatusLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $delSpaceBtn.Enabled = $true
        $remCharsBtn.Enabled = $true
        $hyphenSpaceBtn.Enabled = $true
    }
}

function Update-DataGrid {
    $dataGridView.DataSource = $null
    
    if ($script:csvData.Count -gt 0) {
        # Create DataTable
        $dataTable = New-Object System.Data.DataTable
        
        # Add columns
        foreach ($header in $script:headers) {
            $dataTable.Columns.Add($header) | Out-Null
        }
        
        # Add rows (limit to first 1000 for performance)
        $rowCount = [Math]::Min(1000, $script:csvData.Count)
        for ($i = 0; $i -lt $rowCount; $i++) {
            $row = $script:csvData[$i]
            $dataRow = $dataTable.NewRow()
            foreach ($header in $script:headers) {
                $dataRow[$header] = $row.$header
            }
            $dataTable.Rows.Add($dataRow)
        }
        
        $dataGridView.DataSource = $dataTable
        
        # Highlight selected columns
        Highlight-SelectedColumns
        
        if ($script:csvData.Count -gt 1000) {
            $dataLabel.Text = "Data Preview (showing first 1000 of " + $script:csvData.Count + " rows):"
        } else {
            $dataLabel.Text = "Data Preview (" + $script:csvData.Count + " rows):"
        }
    }
}

function Highlight-SelectedColumns {
    foreach ($column in $dataGridView.Columns) {
        if ($script:selectedColumns[$column.Name] -eq $true) {
            $column.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightBlue
            $column.HeaderCell.Style.BackColor = [System.Drawing.Color]::LightBlue
            $column.HeaderCell.Style.ForeColor = [System.Drawing.Color]::Black
        } else {
            $column.DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Window
            $column.HeaderCell.Style.BackColor = [System.Drawing.SystemColors]::Control
            $column.HeaderCell.Style.ForeColor = [System.Drawing.SystemColors]::ControlText
        }
    }
}

function Update-FileLabel {
    if ($script:filePath) {
        $fileName = [System.IO.Path]::GetFileName($script:filePath)
        $fileInfoLabel.Text = "File: " + $fileName
        $fileInfoLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
    } else {
        $fileInfoLabel.Text = "No file loaded - Click 'Import CSV File' to begin"
        $fileInfoLabel.ForeColor = [System.Drawing.SystemColors]::GrayText
    }
}

function Clean-DataInColumns {
    param([string]$cleanType)
    
    $selectedColumnNames = $script:selectedColumns.Keys | Where-Object { $script:selectedColumns[$_] -eq $true }
    $cleanedCount = 0
    
    foreach ($row in $script:csvData) {
        foreach ($column in $selectedColumnNames) {
            $originalValue = $row.$column
            if ($originalValue) {
                switch ($cleanType) {
                    "DelSpace" {
                        # Remove only spaces, keep other characters
                        $cleanedValue = $originalValue -replace '\s+', ''
                    }
                    "RemChars" {
                        # Remove special characters but keep spaces and alphanumeric
                        $cleanedValue = $originalValue -replace '[^a-zA-Z0-9\s]', ''
                    }
                    "HyphenSpace" {
                        # Replace spaces with hyphens, keep other characters
                        $cleanedValue = $originalValue -replace '\s+', '-'
                    }
                }
                $row.$column = $cleanedValue
                if ($originalValue -ne $cleanedValue) {
                    $cleanedCount++
                }
            }
        }
    }
    
    return $cleanedCount, $selectedColumnNames.Count
}

function Enable-Controls {
    $saveAsBtn.Enabled = $true
    $resetBtn.Enabled = $true
}

function Reset-Data {
    if ($script:originalData.Count -gt 0) {
        $script:csvData = $script:originalData | ForEach-Object { $_.PSObject.Copy() }
        Update-DataGrid
        $statusLabel.Text = "Data reset to original state"
    }
}

function Save-CSVFile {
    param([string]$path)
    
    try {
        $script:csvData | Export-Csv -Path $path -NoTypeInformation
        $statusLabel.Text = "File saved successfully to: " + (Split-Path $path -Leaf)
        [System.Windows.Forms.MessageBox]::Show("File saved successfully!", "Success", "OK", "Information")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving file: " + $_.Exception.Message, "Error", "OK", "Error")
        $statusLabel.Text = "Error saving file"
    }
}

# Event handlers
$importBtn.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV File to Import"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        Load-CSVFile -path $openFileDialog.FileName
    }
})

$saveAsBtn.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save CSV File As"
    
    if ($script:filePath) {
        $saveFileDialog.FileName = [System.IO.Path]::GetFileName($script:filePath)
    }
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        Save-CSVFile -path $saveFileDialog.FileName
        $script:filePath = $saveFileDialog.FileName
        Update-FileLabel
    }
})

$resetBtn.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will reset all data to the original state. Any changes will be lost.`n`nContinue?",
        "Confirm Reset",
        "YesNo",
        "Warning"
    )
    
    if ($result -eq "Yes") {
        Reset-Data
    }
})

# Event handlers for action buttons
$delSpaceBtn.Add_Click({
    $selectedColumnNames = $script:selectedColumns.Keys | Where-Object { $script:selectedColumns[$_] -eq $true }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will remove all spaces from the selected columns:`n`n" + ($selectedColumnNames -join ', ') + "`n`nDo you want to continue?",
        "Confirm Delete Spaces",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $cleanedCount, $columnCount = Clean-DataInColumns -cleanType "DelSpace"
        Update-DataGrid
        $statusLabel.Text = "Delete Spaces: " + $cleanedCount + " cell(s) cleaned in " + $columnCount + " column(s)"
    }
})

$remCharsBtn.Add_Click({
    $selectedColumnNames = $script:selectedColumns.Keys | Where-Object { $script:selectedColumns[$_] -eq $true }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will remove special characters (keeping letters, numbers, and spaces) from the selected columns:`n`n" + ($selectedColumnNames -join ', ') + "`n`nDo you want to continue?",
        "Confirm Remove Characters",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $cleanedCount, $columnCount = Clean-DataInColumns -cleanType "RemChars"
        Update-DataGrid
        $statusLabel.Text = "Remove Characters: " + $cleanedCount + " cell(s) cleaned in " + $columnCount + " column(s)"
    }
})

$hyphenSpaceBtn.Add_Click({
    $selectedColumnNames = $script:selectedColumns.Keys | Where-Object { $script:selectedColumns[$_] -eq $true }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will replace all spaces with hyphens in the selected columns:`n`n" + ($selectedColumnNames -join ', ') + "`n`nDo you want to continue?",
        "Confirm Hyphen Spaces",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $cleanedCount, $columnCount = Clean-DataInColumns -cleanType "HyphenSpace"
        Update-DataGrid
        $statusLabel.Text = "Hyphen Spaces: " + $cleanedCount + " cell(s) cleaned in " + $columnCount + " column(s)"
    }
})

# Handle form and control resizing
$form.Add_Resize({
    $panelWidth = $bottomPanel.Width - 20
    $panelHeight = $bottomPanel.Height - 35
    $dataGridView.Size = New-Object System.Drawing.Size($panelWidth, $panelHeight)
    
    $columnButtonsPanel.Size = New-Object System.Drawing.Size(($middlePanel.Width - 20), 80)
})

# Show form
$form.Add_Shown({
    $form.Activate()
    # Initial resize
    $panelWidth = $bottomPanel.Width - 20
    $panelHeight = $bottomPanel.Height - 35
    $dataGridView.Size = New-Object System.Drawing.Size($panelWidth, $panelHeight)
})

[System.Windows.Forms.Application]::Run($form)
