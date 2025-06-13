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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Process messages for missing RCNs"
$form.Size = New-Object System.Drawing.Size(1000, 750)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.MinimumSize = New-Object System.Drawing.Size(800, 650)

# Create import button
$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(20, 20)
$importButton.Size = New-Object System.Drawing.Size(150, 30)
$importButton.Text = "Import Items CSV"
$importButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create save messages button (initially disabled)
$saveMessagesButton = New-Object System.Windows.Forms.Button
$saveMessagesButton.Location = New-Object System.Drawing.Point(180, 20)
$saveMessagesButton.Size = New-Object System.Drawing.Size(120, 30)
$saveMessagesButton.Text = "Save Messages"
$saveMessagesButton.Enabled = $false
$saveMessagesButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create save items button (initially disabled)
$saveItemsButton = New-Object System.Windows.Forms.Button
$saveItemsButton.Location = New-Object System.Drawing.Point(310, 20)
$saveItemsButton.Size = New-Object System.Drawing.Size(120, 30)
$saveItemsButton.Text = "Save Items"
$saveItemsButton.Enabled = $false
$saveItemsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create Holding RCN label
$holdingRcnLabel = New-Object System.Windows.Forms.Label
$holdingRcnLabel.Location = New-Object System.Drawing.Point(440, 25)
$holdingRcnLabel.Size = New-Object System.Drawing.Size(95, 25)
$holdingRcnLabel.Text = "Holding RCN:"
$holdingRcnLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create Holding RCN text box
$holdingRcnTextBox = New-Object System.Windows.Forms.TextBox
$holdingRcnTextBox.Location = New-Object System.Drawing.Point(540, 22)
$holdingRcnTextBox.Size = New-Object System.Drawing.Size(135, 23)
$holdingRcnTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(685, 25)
$statusLabel.Size = New-Object System.Drawing.Size(275, 25)
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
$global:originalFileName = ""

# Import button click event
$importButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Items CSV File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Store the original filename for later use
            $global:originalFileName = $openFileDialog.FileName
            
            # Read CSV file
            $global:inputData = Import-Csv -Path $openFileDialog.FileName
            
            # Check if essential headers exist (only require barcode and rcn)
            $essentialHeaders = @("barcode","rcn")
            $availableHeaders = $global:inputData[0].PSObject.Properties.Name
            $missingEssentialHeaders = $essentialHeaders | Where-Object { $availableHeaders -notcontains $_ }
            
            if ($missingEssentialHeaders) {
                [System.Windows.Forms.MessageBox]::Show("The CSV is missing essential headers: $($missingEssentialHeaders -join ', ')", "Missing Essential Headers", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            
            # Apply holding RCN to records with blank RCN if holding RCN is provided
            $holdingRcnValue = $holdingRcnTextBox.Text.Trim()
            if (![string]::IsNullOrWhiteSpace($holdingRcnValue)) {
                foreach ($record in $global:inputData) {
                    if ([string]::IsNullOrWhiteSpace($record.rcn)) {
                        $record.rcn = $holdingRcnValue
                    }
                }
            }
            
            # Populate the DataGridView
            $dataGridView.DataSource = [System.Collections.ArrayList]($global:inputData)
            
            # Count records (before applying holding RCN for display purposes)
            $totalRecords = $global:inputData.Count
            $emptyRcnRecords = ($global:inputData | Where-Object { 
                [string]::IsNullOrWhiteSpace($_.rcn) -or 
                ($_.rcn -eq $holdingRcnValue -and ![string]::IsNullOrWhiteSpace($holdingRcnValue))
            }).Count
            
            $recordCountLabel.Text = "Total Records: $totalRecords"
            $emptyRcnCountLabel.Text = "Records with Empty/Holding RCN: $emptyRcnRecords"
            
            # Create output data for records with empty RCN or holding RCN
            $global:outputData = $global:inputData | Where-Object { 
                [string]::IsNullOrWhiteSpace($_.rcn) -or 
                ($_.rcn -eq $holdingRcnValue -and ![string]::IsNullOrWhiteSpace($holdingRcnValue))
            } | ForEach-Object {
                [PSCustomObject]@{
                    Barcode = $_.barcode
                    textflag = "2"
                    Text = "This item requires a unique code such as an ISBN/ISSN etc. Please scan or enter in the items ISBN to update the record as soon as possible."
                }
            }
            
            # Populate the output DataGridView
            $outputGridView.DataSource = [System.Collections.ArrayList]($global:outputData)
            
            # Enable save buttons if we have data
            $saveMessagesButton.Enabled = $global:outputData.Count -gt 0
            $saveItemsButton.Enabled = $global:inputData.Count -gt 0
            
            $holdingRcnMessage = if (![string]::IsNullOrWhiteSpace($holdingRcnValue)) { " Applied holding RCN '$holdingRcnValue' to empty records." } else { "" }
            $statusLabel.Text = "Imported $totalRecords records. Found $emptyRcnRecords needing processing.$holdingRcnMessage"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error importing CSV file: $_", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Save messages button click event
$saveMessagesButton.Add_Click({
    if ($global:outputData -eq $null -or $global:outputData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No message data to save.", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Messages CSV"
    $saveFileDialog.FileName = "item_messages_remarks_authx.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:outputData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $statusLabel.Text = "Successfully saved messages to $($saveFileDialog.FileName)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving messages CSV file: $_", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Save items button click event
$saveItemsButton.Add_Click({
    if ($global:inputData -eq $null -or $global:inputData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No item data to save.", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Generate default filename based on original file
    $defaultFileName = if ($global:originalFileName) {
        $directory = [System.IO.Path]::GetDirectoryName($global:originalFileName)
        $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($global:originalFileName)
        $extension = [System.IO.Path]::GetExtension($global:originalFileName)
        [System.IO.Path]::Combine($directory, "$fileNameWithoutExt-HoldingRCNs$extension")
    } else {
        "items-HoldingRCNs.csv"
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Updated Items CSV"
    $saveFileDialog.FileName = [System.IO.Path]::GetFileName($defaultFileName)
    if ($global:originalFileName) {
        $saveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($global:originalFileName)
    }
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:inputData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $statusLabel.Text = "Successfully saved items to $($saveFileDialog.FileName)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving items CSV file: $_", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
$form.Controls.Add($saveMessagesButton)
$form.Controls.Add($saveItemsButton)
$form.Controls.Add($holdingRcnLabel)
$form.Controls.Add($holdingRcnTextBox)
$form.Controls.Add($statusLabel)
$form.Controls.Add($recordCountLabel)
$form.Controls.Add($emptyRcnCountLabel)
$form.Controls.Add($topPanel)
$form.Controls.Add($inputLabel)
$form.Controls.Add($bottomPanel)
$form.Controls.Add($outputLabel)

# Show the form
$form.ShowDialog()
