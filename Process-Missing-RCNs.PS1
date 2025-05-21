Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Process messages for missing RCNs"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)

# Create import button
$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(20, 20)
$importButton.Size = New-Object System.Drawing.Size(150, 30)
$importButton.Text = "Import Items CSV"
$importButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create save button (initially disabled)
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Location = New-Object System.Drawing.Point(180, 20)
$saveButton.Size = New-Object System.Drawing.Size(150, 30)
$saveButton.Text = "Save Output CSV"
$saveButton.Enabled = $false
$saveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(340, 25)
$statusLabel.Size = New-Object System.Drawing.Size(500, 25)
$statusLabel.Text = "Ready to import CSV file..."
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Create record count labels
$recordCountLabel = New-Object System.Windows.Forms.Label
$recordCountLabel.Location = New-Object System.Drawing.Point(20, 60)
$recordCountLabel.Size = New-Object System.Drawing.Size(250, 25)
$recordCountLabel.Text = "Total Records: 0"
$recordCountLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$emptyRcnCountLabel = New-Object System.Windows.Forms.Label
$emptyRcnCountLabel.Location = New-Object System.Drawing.Point(280, 60)
$emptyRcnCountLabel.Size = New-Object System.Drawing.Size(250, 25)
$emptyRcnCountLabel.Text = "Records with Empty RCN: 0"
$emptyRcnCountLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create a panel for the top section
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Location = New-Object System.Drawing.Point(20, 95)
$topPanel.Size = New-Object System.Drawing.Size(940, 275)
$topPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Create DataGridView for the data
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(0, 0)
$dataGridView.Size = New-Object System.Drawing.Size(940, 275)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill

# Add DataGridView to top panel
$topPanel.Controls.Add($dataGridView)

# Create label for input grid
$inputLabel = New-Object System.Windows.Forms.Label
$inputLabel.Location = New-Object System.Drawing.Point(20, 380)
$inputLabel.Size = New-Object System.Drawing.Size(940, 25)
$inputLabel.Text = "Records with Empty RCN to be Processed:"
$inputLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Create a panel for the bottom section
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Location = New-Object System.Drawing.Point(20, 410)
$bottomPanel.Size = New-Object System.Drawing.Size(940, 230)
$bottomPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Create a second DataGridView for the output data
$outputGridView = New-Object System.Windows.Forms.DataGridView
$outputGridView.Location = New-Object System.Drawing.Point(0, 0)
$outputGridView.Size = New-Object System.Drawing.Size(940, 230)
$outputGridView.AllowUserToAddRows = $false
$outputGridView.AllowUserToDeleteRows = $false
$outputGridView.ReadOnly = $true
$outputGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$outputGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$outputGridView.Dock = [System.Windows.Forms.DockStyle]::Fill

# Add output DataGridView to bottom panel
$bottomPanel.Controls.Add($outputGridView)

# Create output label
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(20, 650)
$outputLabel.Size = New-Object System.Drawing.Size(940, 25)
$outputLabel.Text = "Output will be saved as 'item_messages_remarks_authx.csv'"
$outputLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Variables to store data
$global:inputData = $null
$global:outputData = $null

# Import button click event
$importButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Items CSV File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Read CSV file
            $global:inputData = Import-Csv -Path $openFileDialog.FileName
            
            # Check if required headers exist
            $requiredHeaders = @("shelf","branch","category","parentFlag","location","barcode","rcn")
            $missingHeaders = $requiredHeaders | Where-Object { $global:inputData[0].PSObject.Properties.Name -notcontains $_ }
            
            if ($missingHeaders) {
                [System.Windows.Forms.MessageBox]::Show("The CSV is missing required headers: $($missingHeaders -join ', ')", "Missing Headers", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            
            # Populate the DataGridView
            $dataGridView.DataSource = [System.Collections.ArrayList]($global:inputData)
            
            # Count records
            $totalRecords = $global:inputData.Count
            $emptyRcnRecords = ($global:inputData | Where-Object { [string]::IsNullOrWhiteSpace($_.rcn) }).Count
            
            $recordCountLabel.Text = "Total Records: $totalRecords"
            $emptyRcnCountLabel.Text = "Records with Empty RCN: $emptyRcnRecords"
            
            # Create output data for records with empty RCN
            $global:outputData = $global:inputData | Where-Object { [string]::IsNullOrWhiteSpace($_.rcn) } | ForEach-Object {
                [PSCustomObject]@{
                    Barcode = $_.barcode
                    textflag = "2"
                    Text = "This item requires a unique code such as an ISBN/ISSN etc. Please scan or enter in the items ISBN to update the record as soon as possible."
                }
            }
            
            # Populate the output DataGridView
            $outputGridView.DataSource = [System.Collections.ArrayList]($global:outputData)
            
            # Enable save button if we have output data
            $saveButton.Enabled = $global:outputData.Count -gt 0
            
            $statusLabel.Text = "Imported $totalRecords records. Found $emptyRcnRecords with empty RCN."
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error importing CSV file: $_", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Save button click event
$saveButton.Add_Click({
    if ($global:outputData -eq $null -or $global:outputData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to save.", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Output CSV"
    $saveFileDialog.FileName = "item_messages_remarks_authx.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:outputData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $statusLabel.Text = "Successfully saved to $($saveFileDialog.FileName)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving CSV file: $_", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Form resize event to adjust vertical spacing
$form.Add_Resize({
    # Calculate the vertical split point - allocate 45% of space to top, 55% to bottom
    $totalAvailableHeight = $form.ClientSize.Height - 130 # Subtract space for buttons/labels
    $topGridHeight = [Math]::Floor($totalAvailableHeight * 0.45)
    $bottomGridHeight = $totalAvailableHeight - $topGridHeight - 50 # Allow for spacing
    
    # Set top panel height
    $topPanel.Height = $topGridHeight
    
    # Adjust input label position
    $inputLabel.Top = $topPanel.Bottom + 10
    
    # Adjust bottom panel position and height
    $bottomPanel.Top = $inputLabel.Bottom + 5
    $bottomPanel.Height = $bottomGridHeight
    
    # Adjust output label position
    $outputLabel.Top = $bottomPanel.Bottom + 10
    
    # Adjust widths
    $controlWidth = $form.ClientSize.Width - 40
    $topPanel.Width = $controlWidth
    $bottomPanel.Width = $controlWidth
    $inputLabel.Width = $controlWidth
    $outputLabel.Width = $controlWidth
})

# Add controls to the form
$form.Controls.Add($importButton)
$form.Controls.Add($saveButton)
$form.Controls.Add($statusLabel)
$form.Controls.Add($recordCountLabel)
$form.Controls.Add($emptyRcnCountLabel)
$form.Controls.Add($topPanel)
$form.Controls.Add($inputLabel)
$form.Controls.Add($bottomPanel)
$form.Controls.Add($outputLabel)

# Show the form
$form.ShowDialog()
