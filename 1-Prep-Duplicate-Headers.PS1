Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV Duplicate Column Handler"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create the input file label and textbox
$inputLabel = New-Object System.Windows.Forms.Label
$inputLabel.Location = New-Object System.Drawing.Point(20, 20)
$inputLabel.Size = New-Object System.Drawing.Size(100, 20)
$inputLabel.Text = "Input CSV File:"
$form.Controls.Add($inputLabel)

$inputTextBox = New-Object System.Windows.Forms.TextBox
$inputTextBox.Location = New-Object System.Drawing.Point(130, 20)
$inputTextBox.Size = New-Object System.Drawing.Size(350, 20)
$inputTextBox.ReadOnly = $true
$form.Controls.Add($inputTextBox)

# Create the browse button for input file
$browseInputButton = New-Object System.Windows.Forms.Button
$browseInputButton.Location = New-Object System.Drawing.Point(490, 19)
$browseInputButton.Size = New-Object System.Drawing.Size(80, 23)
$browseInputButton.Text = "Browse..."
$form.Controls.Add($browseInputButton)

# Create the output file label and textbox
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(20, 60)
$outputLabel.Size = New-Object System.Drawing.Size(100, 20)
$outputLabel.Text = "Output CSV File:"
$form.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(130, 60)
$outputTextBox.Size = New-Object System.Drawing.Size(350, 20)
$outputTextBox.ReadOnly = $true
$form.Controls.Add($outputTextBox)

# Create the browse button for output file
$browseOutputButton = New-Object System.Windows.Forms.Button
$browseOutputButton.Location = New-Object System.Drawing.Point(490, 59)
$browseOutputButton.Size = New-Object System.Drawing.Size(80, 23)
$browseOutputButton.Text = "Browse..."
$form.Controls.Add($browseOutputButton)

# Create the log textbox
$logTextBox = New-Object System.Windows.Forms.RichTextBox
$logTextBox.Location = New-Object System.Drawing.Point(20, 100)
$logTextBox.Size = New-Object System.Drawing.Size(550, 300)
$logTextBox.ReadOnly = $true
$logTextBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($logTextBox)

# Create the process button
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(230, 420)
$processButton.Size = New-Object System.Drawing.Size(120, 30)
$processButton.Text = "Process CSV"
$processButton.Enabled = $false
$form.Controls.Add($processButton)

# Function to log messages
function Log-Message {
    param (
        [string]$Message,
        [string]$Color = "Black"
    )
    
    $logTextBox.SelectionColor = $Color
    $logTextBox.AppendText("$Message`r`n")
    $logTextBox.ScrollToCaret()
}

# Browse input file button click event
$browseInputButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Input CSV File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputTextBox.Text = $openFileDialog.FileName
        
        # Suggest default output filename
        $directory = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
        $filename = [System.IO.Path]::GetFileNameWithoutExtension($openFileDialog.FileName)
        $extension = [System.IO.Path]::GetExtension($openFileDialog.FileName)
        $outputTextBox.Text = [System.IO.Path]::Combine($directory, "$filename`_fixed$extension")
        
        # Enable process button if both input and output are set
        if ($inputTextBox.Text -ne "" -and $outputTextBox.Text -ne "") {
            $processButton.Enabled = $true
        }
    }
})

# Browse output file button click event
$browseOutputButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Select Output CSV File"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputTextBox.Text = $saveFileDialog.FileName
        
        # Enable process button if both input and output are set
        if ($inputTextBox.Text -ne "" -and $outputTextBox.Text -ne "") {
            $processButton.Enabled = $true
        }
    }
})

# Process button click event
$processButton.Add_Click({
    try {
        $inputFile = $inputTextBox.Text
        $outputFile = $outputTextBox.Text
        
        # Clear the log
        $logTextBox.Clear()
        
        # Check if input file exists
        if (-not (Test-Path $inputFile)) {
            Log-Message "Error: Input file does not exist." "Red"
            return
        }
        
        Log-Message "Processing CSV file..." "Blue"
        Log-Message "Input file: $inputFile" "Black"
        Log-Message "Output file: $outputFile" "Black"
        
        # Read the CSV file's first line to get headers
        $headerLine = Get-Content $inputFile -First 1
        $originalHeaders = $headerLine -split ','
        
        # Trim headers and handle quotes if present
        $cleanHeaders = @()
        foreach ($header in $originalHeaders) {
            $cleanHeader = $header.Trim()
            # Remove quotes if they exist
            if ($cleanHeader -match '^"(.*)"$') {
                $cleanHeader = $matches[1]
            }
            $cleanHeaders += $cleanHeader
        }
        
        Log-Message "Original headers: $($cleanHeaders -join ', ')" "Black"
        
        # Track duplicate headers and their counts
        $headerCount = @{}
        $newHeaders = @()
        $duplicatesFound = $false
        
        foreach ($header in $cleanHeaders) {
            if ($headerCount.ContainsKey($header)) {
                $duplicatesFound = $true
                $headerCount[$header]++
                $newHeader = "$header-$($headerCount[$header])"
                Log-Message "Duplicate found: '$header' renamed to '$newHeader'" "DarkOrange"
                $newHeaders += $newHeader
            } else {
                $headerCount[$header] = 0
                $newHeaders += $header
            }
        }
        
        if (-not $duplicatesFound) {
            Log-Message "No duplicate headers found. No changes needed." "Green"
            return
        }
        
        # Read content of the input file
        $allLines = Get-Content $inputFile
        
        # Replace the first line with new headers
        $newHeaderLine = $newHeaders -join ','
        $allLines[0] = $newHeaderLine
        
        # Write to output file
        $allLines | Set-Content $outputFile -Encoding UTF8
        
        Log-Message "CSV file processed successfully. Modified CSV saved to: $outputFile" "Green"
    }
    catch {
        Log-Message "Error: $_" "Red"
    }
})

# Show the form
[void]$form.ShowDialog()
