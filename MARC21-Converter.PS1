# CSV to MARC-21 Converter GUI Tool (Items & Loans)
# Requires Windows PowerShell with .NET Framework

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# MARC-21 Helper Functions for Items
function Convert-ItemToMARC21 {
    param(
        [PSCustomObject]$Record
    )
    
    $marc = @()
    
    # Leader (24 characters) - basic structure for bibliographic record
    $leader = "00000nam a2200000 a 4500"
    $marc += "LDR $leader"
    
    # Control fields
    $marc += "001 $($Record.barcode)"  # Control number (barcode)
    
    # Get current date in YYMMDD format for 008 field
    $currentDate = Get-Date -Format "yyMMdd"
    $field008 = "${currentDate}s        xx            000 0 eng d"
    $marc += "008 $field008"
    
    # ISBN (020 field) - if RCN looks like ISBN
    if ($Record.rcn -and $Record.rcn -match '^\d{9}[\dX]$|^\d{13}$') {
        $marc += "020 `$a$($Record.rcn)"
    }
    
    # Location/Branch (040 field - Cataloging source)
    if ($Record.branch) {
        $marc += "040 `$a$($Record.branch)`$c$($Record.branch)"
    }
    
    # Call number based on category and shelf information
    $callNumber = ""
    if ($Record.shelf -and $Record.shelf -ne "") {
        if ($Record.shelf -match '^\d+\.?\d*$') {
            # Dewey Decimal Number
            $callNumber = $Record.shelf
            $marc += "082 04`$a$callNumber"
        } else {
            # Literary/subject classification
            $callNumber = $Record.shelf
            $marc += "050 00`$a$callNumber"
        }
    }
    
    # Title (245 field) - Generate based on category and barcode
    $title = "Item"
    if ($Record.category) {
        switch ($Record.category.ToUpper()) {
            "AF" { $title = "Adult Fiction Item" }
            "ANF" { $title = "Adult Non-Fiction Item" }
            "TF" { $title = "Teen Fiction Item" }
            "DVD" { $title = "DVD Item" }
            "MCD" { $title = "Music CD Item" }
            "MAGAZ" { $title = "Magazine Item" }
            default { $title = "$($Record.category) Item" }
        }
    }
    $title += " [$($Record.barcode)]"
    $marc += "245 10`$a$title"
    
    # Physical description (300 field)
    $physicalDesc = "1 item"
    if ($Record.category) {
        switch ($Record.category.ToUpper()) {
            "DVD" { $physicalDesc = "1 videodisc" }
            "MCD" { $physicalDesc = "1 audio disc" }
            "MAGAZ" { $physicalDesc = "1 magazine" }
            default { $physicalDesc = "1 volume" }
        }
    }
    $marc += "300 `$a$physicalDesc"
    
    # Local holdings (852 field)
    $holdingsField = "852 "
    if ($Record.branch) {
        $holdingsField += "`$a$($Record.branch)"
    }
    if ($callNumber) {
        $holdingsField += "`$h$callNumber"
    }
    if ($Record.barcode) {
        $holdingsField += "`$p$($Record.barcode)"
    }
    $marc += $holdingsField
    
    # Item information (876 field) - Local item data
    $itemField = "876 "
    if ($Record.barcode) {
        $itemField += "`$p$($Record.barcode)"
    }
    if ($Record.location) {
        $itemField += "`$l$($Record.location)"
    }
    if ($Record.shelf) {
        $itemField += "`$h$($Record.shelf)"
    }
    $marc += $itemField
    
    return $marc -join "`r`n"
}

# MARC-21 Helper Functions for Loans/Circulation
function Convert-LoanToMARC21 {
    param(
        [PSCustomObject]$Record
    )
    
    $marc = @()
    
    # Leader for circulation record
    $leader = "00000nas a2200000 n 4500"
    $marc += "LDR $leader"
    
    # Control fields
    $marc += "001 CIRC_$($Record.itembarcode)_$($Record.borrbarcode)"  # Unique circulation ID
    
    # Date/time of transaction (005 field)
    if ($Record.issuedate) {
        try {
            $issueDateTime = [DateTime]::ParseExact($Record.issuedate, "dd/MM/yyyy HH:mm:ss", $null)
            $marc005 = $issueDateTime.ToString("yyyyMMddHHmmss.0")
            $marc += "005 $marc005"
        } catch {
            $marc += "005 $(Get-Date -Format 'yyyyMMddHHmmss.0')"
        }
    }
    
    # Circulation data in custom 9XX fields (local use)
    
    # 901 - Item information
    $itemField = "901 "
    if ($Record.itembarcode) {
        $itemField += "`$a$($Record.itembarcode)"
    }
    if ($Record.issuebranch) {
        $itemField += "`$b$($Record.issuebranch)"
    }
    $marc += $itemField
    
    # 902 - Borrower information
    if ($Record.borrbarcode) {
        $marc += "902 `$a$($Record.borrbarcode)"
    }
    
    # 903 - Loan dates
    $loanField = "903 "
    if ($Record.issuedate) {
        $loanField += "`$a$($Record.issuedate)"
    }
    if ($Record.duedate) {
        $loanField += "`$b$($Record.duedate)"
    }
    $marc += $loanField
    
    # 904 - Transaction type and status
    $marc += "904 `$aLOAN`$bACTIVE"
    
    # 910 - Administrative note
    $adminNote = "Circulation record for item $($Record.itembarcode) borrowed by $($Record.borrbarcode)"
    if ($Record.issuedate -and $Record.duedate) {
        $adminNote += " from $($Record.issuedate) due $($Record.duedate)"
    }
    $marc += "910 `$a$adminNote"
    
    return $marc -join "`r`n"
}

function Export-MARCFile {
    param(
        [Array]$Records,
        [string]$OutputPath,
        [string]$DataType
    )
    
    $marcContent = @()
    $recordCount = 0
    
    foreach ($record in $Records) {
        $shouldProcess = $false
        $marcRecord = ""
        
        if ($DataType -eq "Items" -and $record.barcode) {
            $marcRecord = Convert-ItemToMARC21 -Record $record
            $shouldProcess = $true
        } elseif ($DataType -eq "Loans" -and $record.itembarcode -and $record.borrbarcode) {
            $marcRecord = Convert-LoanToMARC21 -Record $record
            $shouldProcess = $true
        }
        
        if ($shouldProcess) {
            $marcContent += "--- RECORD $($recordCount + 1) ---"
            $marcContent += $marcRecord
            $marcContent += ""
            $recordCount++
        }
    }
    
    # Add summary header
    $header = @(
        "MARC-21 Export File ($DataType Data)",
        "Generated: $(Get-Date)",
        "Total Records: $recordCount",
        "=" * 50,
        ""
    )
    
    $finalContent = $header + $marcContent
    $finalContent | Out-File -FilePath $OutputPath -Encoding UTF8
    
    return $recordCount
}

function Detect-CSVType {
    param(
        [Array]$CSVData
    )
    
    if ($CSVData.Count -eq 0) {
        return "Unknown"
    }
    
    $firstRecord = $CSVData[0]
    $properties = $firstRecord.PSObject.Properties.Name
    
    # Check for items data columns
    $itemsColumns = @("shelf", "branch", "category", "barcode", "rcn")
    $itemsMatch = ($itemsColumns | Where-Object { $_ -in $properties }).Count
    
    # Check for loans data columns
    $loansColumns = @("issuedate", "itembarcode", "duedate", "borrbarcode")
    $loansMatch = ($loansColumns | Where-Object { $_ -in $properties }).Count
    
    if ($loansMatch -ge 3) {
        return "Loans"
    } elseif ($itemsMatch -ge 3) {
        return "Items"
    } else {
        return "Unknown"
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV to MARC-21 Converter (Items & Loans)"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Create controls
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "CSV to MARC-21 Library Format Converter"
$lblTitle.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.Size = New-Object System.Drawing.Size(600, 25)
$lblTitle.TextAlign = "MiddleCenter"

$lblInstruction = New-Object System.Windows.Forms.Label
$lblInstruction.Text = "Select a CSV file containing library items or loans data to convert to MARC-21 format:"
$lblInstruction.Location = New-Object System.Drawing.Point(20, 60)
$lblInstruction.Size = New-Object System.Drawing.Size(600, 20)

# File selection group
$gbFileSelection = New-Object System.Windows.Forms.GroupBox
$gbFileSelection.Text = "File Selection"
$gbFileSelection.Location = New-Object System.Drawing.Point(20, 90)
$gbFileSelection.Size = New-Object System.Drawing.Size(600, 80)

$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Location = New-Object System.Drawing.Point(15, 25)
$txtFilePath.Size = New-Object System.Drawing.Size(450, 20)
$txtFilePath.ReadOnly = $true

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(475, 23)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 25)

# Data type detection
$lblDataType = New-Object System.Windows.Forms.Label
$lblDataType.Location = New-Object System.Drawing.Point(15, 50)
$lblDataType.Size = New-Object System.Drawing.Size(450, 20)
$lblDataType.Text = "Data type will be detected automatically"
$lblDataType.ForeColor = [System.Drawing.Color]::Blue

# Preview group
$gbPreview = New-Object System.Windows.Forms.GroupBox
$gbPreview.Text = "Data Preview"
$gbPreview.Location = New-Object System.Drawing.Point(20, 180)
$gbPreview.Size = New-Object System.Drawing.Size(600, 250)

$txtPreview = New-Object System.Windows.Forms.TextBox
$txtPreview.Location = New-Object System.Drawing.Point(15, 25)
$txtPreview.Size = New-Object System.Drawing.Size(570, 210)
$txtPreview.Multiline = $true
$txtPreview.ScrollBars = "Both"
$txtPreview.ReadOnly = $true
$txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)

# Convert button
$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Text = "Convert to MARC-21"
$btnConvert.Location = New-Object System.Drawing.Point(250, 450)
$btnConvert.Size = New-Object System.Drawing.Size(120, 30)
$btnConvert.BackColor = [System.Drawing.Color]::LightBlue
$btnConvert.Enabled = $false

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 490)
$lblStatus.Size = New-Object System.Drawing.Size(600, 20)
$lblStatus.Text = "Ready. Please select a CSV file."

# Add controls to form
$gbFileSelection.Controls.Add($txtFilePath)
$gbFileSelection.Controls.Add($btnBrowse)
$gbFileSelection.Controls.Add($lblDataType)
$gbPreview.Controls.Add($txtPreview)

$form.Controls.Add($lblTitle)
$form.Controls.Add($lblInstruction)
$form.Controls.Add($gbFileSelection)
$form.Controls.Add($gbPreview)
$form.Controls.Add($btnConvert)
$form.Controls.Add($lblStatus)

# Global variables
$script:csvData = $null
$script:dataType = "Unknown"

# Browse button event handler
$btnBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFilePath.Text = $openFileDialog.FileName
        $lblStatus.Text = "Loading CSV file..."
        
        try {
            # Import CSV data
            $script:csvData = Import-Csv -Path $openFileDialog.FileName
            
            # Detect data type
            $script:dataType = Detect-CSVType -CSVData $script:csvData
            $lblDataType.Text = "Detected data type: $($script:dataType)"
            
            if ($script:dataType -eq "Unknown") {
                $lblDataType.ForeColor = [System.Drawing.Color]::Red
                $lblDataType.Text += " - Please check your CSV format"
            } else {
                $lblDataType.ForeColor = [System.Drawing.Color]::Green
            }
            
            # Display preview
            $previewText = "Data Type: $($script:dataType)`r`n"
            $previewText += "Columns: " + ($script:csvData[0].PSObject.Properties.Name -join ", ") + "`r`n`r`n"
            
            if ($script:dataType -eq "Items") {
                $previewText += "Sample item records (first 5):`r`n"
            } elseif ($script:dataType -eq "Loans") {
                $previewText += "Sample loan records (first 5):`r`n"
            } else {
                $previewText += "Sample records (first 5):`r`n"
            }
            
            $previewText += "-" * 50 + "`r`n"
            
            $sampleCount = [Math]::Min(5, $script:csvData.Count)
            for ($i = 0; $i -lt $sampleCount; $i++) {
                $record = $script:csvData[$i]
                $previewText += "Record $($i + 1):`r`n"
                foreach ($prop in $record.PSObject.Properties) {
                    if ($prop.Value) {
                        $previewText += "  $($prop.Name): $($prop.Value)`r`n"
                    }
                }
                $previewText += "`r`n"
            }
            
            $txtPreview.Text = $previewText
            $btnConvert.Enabled = ($script:dataType -ne "Unknown")
            $lblStatus.Text = "CSV loaded successfully. $($script:csvData.Count) records found. Type: $($script:dataType)"
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading CSV file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Error loading CSV file."
        }
    }
})

# Convert button event handler
$btnConvert.Add_Click({
    if ($script:csvData -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select a CSV file first.", "No File Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($script:dataType -eq "Unknown") {
        [System.Windows.Forms.MessageBox]::Show("Cannot process unknown data type. Please check your CSV format.", "Unknown Data Type", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "MARC files (*.mrc)|*.mrc|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save MARC-21 File"
    
    if ($script:dataType -eq "Items") {
        $saveFileDialog.FileName = "library_items.mrc"
    } else {
        $saveFileDialog.FileName = "library_loans.mrc"
    }
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lblStatus.Text = "Converting to MARC-21 format..."
        
        try {
            $recordCount = Export-MARCFile -Records $script:csvData -OutputPath $saveFileDialog.FileName -DataType $script:dataType
            
            $message = "Conversion completed successfully!`r`n`r`n" +
                      "Data type: $($script:dataType)`r`n" +
                      "Records processed: $recordCount`r`n" +
                      "Output file: $($saveFileDialog.FileName)`r`n`r`n" +
                      "Would you like to open the output file?"
            
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Conversion Complete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process notepad.exe -ArgumentList $saveFileDialog.FileName
            }
            
            $lblStatus.Text = "Conversion completed. $recordCount $($script:dataType.ToLower()) records exported."
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error during conversion: $($_.Exception.Message)", "Conversion Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Error during conversion."
        }
    }
})

# Show the form
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
