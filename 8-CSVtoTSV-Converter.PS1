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
# Custom naming convention reminder function
function Show-NamingConventionReminder {
    $customDialog = New-Object System.Windows.Forms.Form
    $customDialog.Text = "Naming Convention"
    $customDialog.Size = New-Object System.Drawing.Size(500, 400)
    $customDialog.StartPosition = "CenterScreen"
    $customDialog.MinimizeBox = $false
    $customDialog.MaximizeBox = $false
    $customDialog.FormBorderStyle = "FixedDialog"
    $customDialog.Icon = [System.Drawing.SystemIcons]::Information

    # Title label with bold font
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Naming Convention"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 25)
    $customDialog.Controls.Add($titleLabel)

    # Description text
    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = "The naming convention is as below with the x representing the authority number. The asterisk (*) indicates that the file is mandatory; the others may not be necessary if the prison system/spreadsheet does not allow statuses, messages, remarks or handle reservations."    
    $descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $descriptionLabel.Location = New-Object System.Drawing.Point(20, 45)
    $descriptionLabel.Size = New-Object System.Drawing.Size(460, 70)
    $descriptionLabel.AutoSize = $false
    $descriptionLabel.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $customDialog.Controls.Add($descriptionLabel)

    # Naming conventions list with proper bullet points
    $conventions = @(
        "- borrowers_authx *",
        "- borrower_status_authx",
        "- borrower_messages_remarks_authx",
        "- items_authx *",
        "- item_status_authx",
        "- item_messages_remarks_authx",
        "- loans_authx *",
        "- reservations_authx",
        "- loan_history_authx"
    )

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $listBox.Location = New-Object System.Drawing.Point(30, 115)
    $listBox.Size = New-Object System.Drawing.Size(440, 180)
    $listBox.BorderStyle = "None"
    $listBox.SelectionMode = "None"
    $listBox.BackColor = $customDialog.BackColor
    
    foreach ($convention in $conventions) {
        $listBox.Items.Add($convention)
    }
    
    $customDialog.Controls.Add($listBox)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(200, 320)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $customDialog.Controls.Add($okButton)
    $customDialog.AcceptButton = $okButton

    # Show the dialog
    $customDialog.ShowDialog() | Out-Null
}# CSV to Axiell Format Converter - GUI Tool
# 1. Minimizes the PowerShell window on launch.
# 2. Lets you load a CSV into a DataGridView.
# 3. Exports data in Axiell "field=value" format (tab-delimited, skipping blank fields).
# 4. Displays a 3×3 grid of recommended file-name conventions (with Copy buttons).
# 5. Uses Panel approach with fixed positioning for better alignment of text and buttons.
# 6. Pops up a Naming Convention message before saving.

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

# Minimize the PowerShell console
$consolePtr = [Win32]::GetConsoleWindow()
if ($consolePtr -ne [IntPtr]::Zero) {
    # 6 = Minimize
    [Win32]::ShowWindow($consolePtr, 6)
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$global:CsvData = $null
$global:CsvPath = $null

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV to Axiell Converter"
$form.Size = New-Object System.Drawing.Size(1200, 750)
$form.StartPosition = "CenterScreen"
$form.Topmost = $false

# TableLayout: 2 rows
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Absolute, 110)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($tableLayout)

# Top panel (row 1) for buttons & group box
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($topPanel, 0, 0)

# "Load CSV" button
$btnLoadCSV = New-Object System.Windows.Forms.Button
$btnLoadCSV.Text = "Load CSV"
$btnLoadCSV.Width = 120
$btnLoadCSV.Location = New-Object System.Drawing.Point(10, 10)
$topPanel.Controls.Add($btnLoadCSV)

# "Export Axiell" button
$btnExportAxiell = New-Object System.Windows.Forms.Button
$btnExportAxiell.Text = "Export Axiell"
$btnExportAxiell.Width = 120
$btnExportAxiell.Location = New-Object System.Drawing.Point(140, 10)
$topPanel.Controls.Add($btnExportAxiell)

# Group box for the 3×3 naming grid
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Recommended Filenames"
$groupBox.Location = New-Object System.Drawing.Point(280, 0)
$groupBox.Size = New-Object System.Drawing.Size(900, 110)  # Reduced height slightly
$groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$topPanel.Controls.Add($groupBox)

# Recommended naming conventions
$namingConventions = @(
    "borrowers_authx",
    "borrower_status_authx",
    "borrower_messages_remarks_authx",
    "items_authx",
    "item_status_authx",
    "item_messages_remarks_authx",
    "loans_authx",
    "reservations_authx",
    "loan_history_authx"
)

# Logging function
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("`r`n[$timestamp] $message")
}

# Create the 3x3 grid using fixed positions
# Define cell dimensions and spacing
$cellWidth = 280    # Adjusted width for better spacing
$cellHeight = 25    # Height of each cell
$cellMarginX = 10   # Increased horizontal margin for more equal spacing
$cellMarginY = 5    # Vertical spacing between cells
$buttonWidth = 60   # Width of copy button
$buttonHeight = 23  # Height of copy button
$startX = 15        # Increased start X for more even border spacing
$startY = 20        # Adjusted starting Y position

# Create and add labels and buttons for each naming convention
for ($i = 0; $i -lt $namingConventions.Count; $i++) {
    # Calculate row and column
    $row = [math]::Floor($i / 3)
    $col = $i % 3
    
    # Calculate position
    $x = $startX + ($col * ($cellWidth + $cellMarginX))
    $y = $startY + ($row * ($cellHeight + $cellMarginY))
    
    # Create label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $namingConventions[$i]
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size(($cellWidth - $buttonWidth - 5), $cellHeight)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $groupBox.Controls.Add($label)
    
    # Create copy button - positioned to the right of the label
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = "Copy"
    $btnCopy.Location = New-Object System.Drawing.Point(($x + $cellWidth - $buttonWidth), $y)
    $btnCopy.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $btnCopy.Tag = $namingConventions[$i]
    $btnCopy.Add_Click({
        param($sender, $args)
        [System.Windows.Forms.Clipboard]::SetText($sender.Tag)
        Write-Log("Copied '$($sender.Tag)' to clipboard.")
    })
    $groupBox.Controls.Add($btnCopy)
}

# SplitContainer for DataGrid (top) and Log (bottom)
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 400
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$tableLayout.Controls.Add($splitContainer, 0, 1)

$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Panel1.Controls.Add($gridPanel)

# DataGridView to preview CSV
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.AutoSizeColumnsMode = "Fill"
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ColumnHeadersHeight = 30
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.BackgroundColor = [System.Drawing.Color]::White
$gridPanel.Controls.Add($grid)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Panel2.Controls.Add($logPanel)

# TextBox for logs
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Text = "Ready."
$logPanel.Controls.Add($txtLog)

# Load CSV
function Load-CSVFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select a CSV file"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $openFileDialog.FileName
        Write-Log "Loading CSV file: $csvPath"
        try {
            $global:CsvData = Import-Csv $csvPath
            $global:CsvPath = $csvPath
            if ($global:CsvData.Count -gt 0) {
                Write-Log "Loaded $($global:CsvData.Count) rows."
                Update-GridView
            } else {
                Write-Log "CSV is empty or could not be parsed."
            }
        } catch {
            Write-Log "Error loading CSV: $_"
        }
    } else {
        Write-Log "User cancelled loading CSV."
    }
}

# Update DataGridView
function Update-GridView {
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    if (-not $global:CsvData) { return }
    $grid.AutoGenerateColumns = $true
    
    $list = New-Object System.Collections.Generic.List[Object]
    foreach ($row in $global:CsvData) {
        $list.Add($row)
    }
    $grid.DataSource = $list
}

# Export to Axiell format
function Export-AxiellFormat {
    if (-not $global:CsvData -or $global:CsvData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No CSV data loaded. Please load a CSV first.")
        return
    }

    # Show naming convention reminder before Save dialog
    Show-NamingConventionReminder

    # Proceed to save
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Title = "Save Axiell Format (TSV)"
    $saveFileDialog.Filter = "TSV Files (*.tsv)|*.tsv|All Files (*.*)|*.*"
    
    $dialogResult = $saveFileDialog.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $outPath = $saveFileDialog.FileName
        Write-Log "Exporting to Axiell format => $outPath"
        try {
            $lines = @()
            foreach ($row in $global:CsvData) {
                $fieldValuePairs = @()
                foreach ($property in $row.PSObject.Properties) {
                    $fieldName  = $property.Name
                    $fieldValue = $property.Value
                    # Skip blank fields
                    if (-not [string]::IsNullOrEmpty($fieldValue)) {
                        $fieldValuePairs += "$fieldName=$fieldValue"
                    }
                }
                $lines += ($fieldValuePairs -join "`t")
            }
            
            $lines | Out-File -FilePath $outPath -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show(
                "Exported successfully to $outPath",
                "Axiell Export"
            )
            Write-Log "Axiell export completed: $outPath"
        } catch {
            Write-Log "Error exporting: $_"
            [System.Windows.Forms.MessageBox]::Show("Error exporting to Axiell format: $_","Error")
        }
    } else {
        Write-Log "Export canceled by user."
    }
}

# Hook up button clicks
$btnLoadCSV.Add_Click({ Load-CSVFile })
$btnExportAxiell.Add_Click({ Export-AxiellFormat })

# On form show, log a message
$form.Add_Shown({
    Write-Log "Application started. Please click 'Load CSV' to begin."
})

[void] $form.ShowDialog()
