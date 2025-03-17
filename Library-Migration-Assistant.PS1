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
$consolePtr = [Win32]::GetConsoleWindow()
if ($consolePtr -ne [IntPtr]::Zero) {
    # 6 = Minimize the window
    [Win32]::ShowWindow($consolePtr, 6)
}

# Library Migration Assistant PowerShell Script
# This script helps map fields between library systems
# Usage: .\LibraryMigrationAssistant.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$global:SourceData = $null
$global:SourceFields = @()
$global:TargetFields = @()
$global:Mappings = @{}
$global:SuggestedMappings = @{}
$global:TransformFunctions = @{}
$global:LogOutput = ""

$global:ProcessScript = ".\Transform-Functions.ps1"
. $global:ProcessScript
$global:TransformationFunctions = @(Get-Command -CommandType Function | Where-Object { $_.Name -like "*Transform" -or $_.Name -like "Custom-Text*" } | Select-Object -ExpandProperty Name)

# Create main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Library Migration Assistant"
$mainForm.Size = New-Object System.Drawing.Size(1000, 700)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $true

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Location = New-Object System.Drawing.Point(0, 0)
$tabControl.Size = New-Object System.Drawing.Size(990, 650)

# Create tabs
$tabLoad = New-Object System.Windows.Forms.TabPage
$tabLoad.Text = "Load Files"

$tabAnalyze = New-Object System.Windows.Forms.TabPage
$tabAnalyze.Text = "Analyze"

$tabMap = New-Object System.Windows.Forms.TabPage
$tabMap.Text = "Map Fields"

$tabExport = New-Object System.Windows.Forms.TabPage
$tabExport.Text = "Export"

# Add tabs to the tab control
$tabControl.Controls.Add($tabLoad)
$tabControl.Controls.Add($tabAnalyze)
$tabControl.Controls.Add($tabMap)
$tabControl.Controls.Add($tabExport)

# Add tab control to the form
$mainForm.Controls.Add($tabControl)

# Initialize disabled tabs
$tabAnalyze.Enabled = $false
$tabMap.Enabled = $false
$tabExport.Enabled = $false

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBarLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBarLabel.Text = "Ready"
$statusBar.Items.Add($statusBarLabel)
$mainForm.Controls.Add($statusBar)

#--------------------------
# LOAD FILES TAB
#--------------------------

# Source data file section
$sourceGroupBox = New-Object System.Windows.Forms.GroupBox
$sourceGroupBox.Text = "Source System Data"
$sourceGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$sourceGroupBox.Size = New-Object System.Drawing.Size(950, 80)
$tabLoad.Controls.Add($sourceGroupBox)

$lblSourceFile = New-Object System.Windows.Forms.Label
$lblSourceFile.Text = "Source Data File (CSV):"
$lblSourceFile.Location = New-Object System.Drawing.Point(10, 30)
$lblSourceFile.Size = New-Object System.Drawing.Size(150, 20)
$sourceGroupBox.Controls.Add($lblSourceFile)

$txtSourceFile = New-Object System.Windows.Forms.TextBox
$txtSourceFile.Location = New-Object System.Drawing.Point(160, 28)
$txtSourceFile.Size = New-Object System.Drawing.Size(650, 20)
$sourceGroupBox.Controls.Add($txtSourceFile)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(820, 26)
$btnBrowseSource.Size = New-Object System.Drawing.Size(100, 24)
$sourceGroupBox.Controls.Add($btnBrowseSource)

# Target data file section
$targetGroupBox = New-Object System.Windows.Forms.GroupBox
$targetGroupBox.Text = "Target System Requirements"
$targetGroupBox.Location = New-Object System.Drawing.Point(10, 100)
$targetGroupBox.Size = New-Object System.Drawing.Size(950, 80)
$tabLoad.Controls.Add($targetGroupBox)

$lblTargetFile = New-Object System.Windows.Forms.Label
$lblTargetFile.Text = "Target Requirements (CSV):"
$lblTargetFile.Location = New-Object System.Drawing.Point(10, 30)
$lblTargetFile.Size = New-Object System.Drawing.Size(150, 20)
$targetGroupBox.Controls.Add($lblTargetFile)

$txtTargetFile = New-Object System.Windows.Forms.TextBox
$txtTargetFile.Location = New-Object System.Drawing.Point(160, 28)
$txtTargetFile.Size = New-Object System.Drawing.Size(650, 20)
$targetGroupBox.Controls.Add($txtTargetFile)

$btnBrowseTarget = New-Object System.Windows.Forms.Button
$btnBrowseTarget.Text = "Browse..."
$btnBrowseTarget.Location = New-Object System.Drawing.Point(820, 26)
$btnBrowseTarget.Size = New-Object System.Drawing.Size(100, 24)
$targetGroupBox.Controls.Add($btnBrowseTarget)

# Load button
$btnLoadData = New-Object System.Windows.Forms.Button
$btnLoadData.Text = "Load Data"
$btnLoadData.Location = New-Object System.Drawing.Point(400, 190)
$btnLoadData.Size = New-Object System.Drawing.Size(150, 30)
$tabLoad.Controls.Add($btnLoadData)

# Log textbox
$txtLoadLog = New-Object System.Windows.Forms.TextBox
$txtLoadLog.Location = New-Object System.Drawing.Point(10, 230)
$txtLoadLog.Size = New-Object System.Drawing.Size(950, 380)
$txtLoadLog.Multiline = $true
$txtLoadLog.ScrollBars = "Vertical"
$txtLoadLog.ReadOnly = $true
$tabLoad.Controls.Add($txtLoadLog)

#--------------------------
# ANALYZE TAB
#--------------------------

# Source fields listbox
$sourceListGroupBox = New-Object System.Windows.Forms.GroupBox
$sourceListGroupBox.Text = "Source Fields"
$sourceListGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$sourceListGroupBox.Size = New-Object System.Drawing.Size(250, 500)
$tabAnalyze.Controls.Add($sourceListGroupBox)

$lstSourceFields = New-Object System.Windows.Forms.ListBox
$lstSourceFields.Location = New-Object System.Drawing.Point(10, 20)
$lstSourceFields.Size = New-Object System.Drawing.Size(230, 470)
$sourceListGroupBox.Controls.Add($lstSourceFields)

# Target fields listbox
$targetListGroupBox = New-Object System.Windows.Forms.GroupBox
$targetListGroupBox.Text = "Target Fields"
$targetListGroupBox.Location = New-Object System.Drawing.Point(730, 10)
$targetListGroupBox.Size = New-Object System.Drawing.Size(250, 500)
$tabAnalyze.Controls.Add($targetListGroupBox)

$lstTargetFields = New-Object System.Windows.Forms.ListBox
$lstTargetFields.Location = New-Object System.Drawing.Point(10, 20)
$lstTargetFields.Size = New-Object System.Drawing.Size(230, 470)
$targetListGroupBox.Controls.Add($lstTargetFields)

# Preview panel
$previewGroupBox = New-Object System.Windows.Forms.GroupBox
$previewGroupBox.Text = "Data Preview"
$previewGroupBox.Location = New-Object System.Drawing.Point(270, 10)
$previewGroupBox.Size = New-Object System.Drawing.Size(450, 500)
$tabAnalyze.Controls.Add($previewGroupBox)

$txtDataPreview = New-Object System.Windows.Forms.TextBox
$txtDataPreview.Location = New-Object System.Drawing.Point(10, 20)
$txtDataPreview.Size = New-Object System.Drawing.Size(430, 470)
$txtDataPreview.Multiline = $true
$txtDataPreview.ScrollBars = "Vertical"
$txtDataPreview.ReadOnly = $true
$previewGroupBox.Controls.Add($txtDataPreview)

# Analyze button
$btnAnalyze = New-Object System.Windows.Forms.Button
$btnAnalyze.Text = "Analyze and Generate Mapping Suggestions"
$btnAnalyze.Location = New-Object System.Drawing.Point(350, 520)
$btnAnalyze.Size = New-Object System.Drawing.Size(300, 30)
$tabAnalyze.Controls.Add($btnAnalyze)

#--------------------------
# MAPPING TAB
#--------------------------

# Create DataGridView for mappings
$dgvMappings = New-Object System.Windows.Forms.DataGridView
$dgvMappings.Location = New-Object System.Drawing.Point(10, 10)
$dgvMappings.Size = New-Object System.Drawing.Size(950, 300)
$dgvMappings.AllowUserToAddRows = $false
$dgvMappings.AllowUserToDeleteRows = $false
$dgvMappings.ReadOnly = $true
$dgvMappings.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgvMappings.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvMappings.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dgvMappings.BackgroundColor = [System.Drawing.SystemColors]::Window
$dgvMappings.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dgvMappings.RowHeadersVisible = $false
$tabMap.Controls.Add($dgvMappings)

# Add columns to DataGridView (Target Field, Source Field, Transformation, Required)
$dgvMappings.ColumnCount = 4
$dgvMappings.Columns[0].Name = "Target Field"
$dgvMappings.Columns[1].Name = "Source Field"
$dgvMappings.Columns[2].Name = "Transformation"
$dgvMappings.Columns[3].Name = "Required"

# ---------------------------
# Mapping Edit Panel
# ---------------------------
# Declare layout variables as integers
[int]$leftX          = 20
[int]$leftXControl   = 130
[int]$rightX         = 420
[int]$rightXControl  = 520
[int]$startY         = 20
[int]$verticalSpace  = 40

$mapEditGroupBox = New-Object System.Windows.Forms.GroupBox
$mapEditGroupBox.Text = "Edit Mapping"
$mapEditGroupBox.Location = New-Object System.Drawing.Point(10, 320)
$mapEditGroupBox.Size = New-Object System.Drawing.Size(950, 320)
$tabMap.Controls.Add($mapEditGroupBox)

############################
# LEFT COLUMN
############################

# Row 1: Target Field
$lblTargetField = New-Object System.Windows.Forms.Label
$lblTargetField.Text = "Target Field:"
$lblTargetField.Location = New-Object System.Drawing.Point($leftX, $startY)
$lblTargetField.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblTargetField)

$txtTargetField = New-Object System.Windows.Forms.TextBox
$txtTargetField.Location = New-Object System.Drawing.Point($leftXControl, $startY)
$txtTargetField.Size = New-Object System.Drawing.Size(200, 20)
$txtTargetField.ReadOnly = $true
$txtTargetField.BackColor = [System.Drawing.SystemColors]::Control
$mapEditGroupBox.Controls.Add($txtTargetField)

# Row 2: Map From
$lblSourceField = New-Object System.Windows.Forms.Label
$lblSourceField.Text = "Map From:"
$lblSourceField.Location = New-Object System.Drawing.Point($leftX, ($startY + $verticalSpace))
$lblSourceField.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblSourceField)

$cboSourceField = New-Object System.Windows.Forms.ComboBox
$cboSourceField.Location = New-Object System.Drawing.Point($leftXControl, ($startY + $verticalSpace))
$cboSourceField.Size = New-Object System.Drawing.Size(200, 20)
$cboSourceField.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$mapEditGroupBox.Controls.Add($cboSourceField)

# Row 3: Transformation
$lblTransform = New-Object System.Windows.Forms.Label
$lblTransform.Text = "Transformation:"
$lblTransform.Location = New-Object System.Drawing.Point($leftX, ($startY + (2 * $verticalSpace)))
$lblTransform.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblTransform)

$cboTransform = New-Object System.Windows.Forms.ComboBox
$cboTransform.Location = New-Object System.Drawing.Point($leftXControl, ($startY + (2 * $verticalSpace)))
$cboTransform.Size = New-Object System.Drawing.Size(200, 20)
$cboTransform.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Populate the transformation combo box
$cboTransform.Items.Clear()
$cboTransform.Items.Add("None")
foreach ($func in $global:TransformationFunctions) {
    $cboTransform.Items.Add($func)
}
$cboTransform.Items.Add("New Custom Text...")
$cboTransform.SelectedIndex = 0

$mapEditGroupBox.Controls.Add($cboTransform)

# --------------------------------------------------
# EVENT: If "New Custom Text..." is selected
# --------------------------------------------------
$cboTransform.Add_SelectedIndexChanged({
    if ($cboTransform.SelectedItem -eq "New Custom Text...") {
        # Show popup form
        $popupForm = New-Object System.Windows.Forms.Form
        $popupForm.Text = "Enter Custom Text"
        $popupForm.Size = New-Object System.Drawing.Size(300, 150)
        $popupForm.StartPosition = "CenterParent"
        
        $lblCustom = New-Object System.Windows.Forms.Label
        $lblCustom.Text = "Custom Text:"
        $lblCustom.Location = New-Object System.Drawing.Point(10, 10)
        $lblCustom.Size = New-Object System.Drawing.Size(80, 20)
        $popupForm.Controls.Add($lblCustom)
        
        $txtCustom = New-Object System.Windows.Forms.TextBox
        $txtCustom.Location = New-Object System.Drawing.Point(100, 10)
        $txtCustom.Size = New-Object System.Drawing.Size(150, 20)
        $popupForm.Controls.Add($txtCustom)
        
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Location = New-Object System.Drawing.Point(100, 50)
        $btnOK.Size = New-Object System.Drawing.Size(80, 30)
        $popupForm.Controls.Add($btnOK)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Location = New-Object System.Drawing.Point(190, 50)
        $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
        $popupForm.Controls.Add($btnCancel)
        
        $btnOK.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($txtCustom.Text)) {
                $customText = $txtCustom.Text.Trim()

                # 1) Sanitize the function name
                $sanitized = ($customText -replace "\s+", "-") -replace "[^A-Za-z0-9\-]", ""
                if (-not $sanitized) { $sanitized = "Custom" }
                $newFuncName = "Custom-Text-" + $sanitized

                # 2) Build the function code to append
                $escapedText = $customText -replace "'", "''"
                $newFuncDef = @"
function $newFuncName(`$value) {
    return '$escapedText'
}
`$global:TransformFunctions["$newFuncName"] = `${function:$newFuncName}

"@ + "`r`n"

                # 3) Read the current file (create if needed), then just APPEND
                if (-not (Test-Path $global:ProcessScript)) {
                    New-Item -Path $global:ProcessScript -ItemType File -Force | Out-Null
                }
                Add-Content -Path $global:ProcessScript -Value $newFuncDef -Encoding UTF8

                # 4) Dot-source so it's available now
                . $global:ProcessScript

                # 5) Rebuild $global:TransformFunctions for the GUI
                $global:TransformFunctions = @{}
                $funcs = Get-Command -CommandType Function | Where-Object {
                    $_.Name -like "*Transform" -or $_.Name -like "Custom-Text*"
                }
                foreach ($f in $funcs) {
                    $global:TransformFunctions[$f.Name] = ${function:$($f.Name)}
                }

                # 6) Refresh the combo box
                $cboTransform.Items.Clear()
                $cboTransform.Items.Add("None")
                foreach ($key in $global:TransformFunctions.Keys) {
                    $cboTransform.Items.Add($key)
                }
                $cboTransform.Items.Add("New Custom Text...")

                # Select the newly created function
                $index = $cboTransform.Items.IndexOf($newFuncName)
                if ($index -ge 0) {
                    $cboTransform.SelectedIndex = $index
                }
                else {
                    $cboTransform.SelectedIndex = 0
                }
            }
            else {
                # No text => revert to "None"
                $cboTransform.SelectedIndex = 0
            }
            $popupForm.Close()
        })
        
        $btnCancel.Add_Click({
            $cboTransform.SelectedIndex = 0
            $popupForm.Close()
        })
        
        $popupForm.ShowDialog()
    }
})

# Row 4: Update Mapping button (aligned left)
$btnUpdateMapping = New-Object System.Windows.Forms.Button
$btnUpdateMapping.Text = "Update Mapping"
$btnUpdateMapping.Location = New-Object System.Drawing.Point($leftX, ($startY + (3 * $verticalSpace)))
$btnUpdateMapping.Size = New-Object System.Drawing.Size(120, 30)
$mapEditGroupBox.Controls.Add($btnUpdateMapping)

############################
# RIGHT COLUMN
############################

# Row 1: Notes
$lblNotesRight = New-Object System.Windows.Forms.Label
$lblNotesRight.Text = "Notes:"
$lblNotesRight.Location = New-Object System.Drawing.Point($rightX, $startY)
$lblNotesRight.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblNotesRight)

$txtNotesRight = New-Object System.Windows.Forms.TextBox
$txtNotesRight.Location = New-Object System.Drawing.Point($rightXControl, $startY)
$txtNotesRight.Size = New-Object System.Drawing.Size(350, 60)
$txtNotesRight.Multiline = $true
$txtNotesRight.ReadOnly = $true
$txtNotesRight.BackColor = [System.Drawing.SystemColors]::Control
$txtNotesRight.ScrollBars = "Vertical"
$mapEditGroupBox.Controls.Add($txtNotesRight)

# Row 2: Field Type (directly below the Notes box)
$lblFieldType = New-Object System.Windows.Forms.Label
$lblFieldType.Text = "Field Type:"
$lblFieldType.Location = New-Object System.Drawing.Point($rightX, ($startY + 70))
$lblFieldType.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblFieldType)

$txtFieldType = New-Object System.Windows.Forms.TextBox
$txtFieldType.Location = New-Object System.Drawing.Point($rightXControl, ($startY + 70))
$txtFieldType.Size = New-Object System.Drawing.Size(200, 20)
$txtFieldType.ReadOnly = $true
$txtFieldType.BackColor = [System.Drawing.SystemColors]::Control
$mapEditGroupBox.Controls.Add($txtFieldType)

# Row 3: Validation Rule (directly below Field Type)
$lblValidationRule = New-Object System.Windows.Forms.Label
$lblValidationRule.Text = "Validation Rule:"
$lblValidationRule.Location = New-Object System.Drawing.Point($rightX, ($startY + 70 + $verticalSpace - 10))
$lblValidationRule.Size = New-Object System.Drawing.Size(100, 20)
$mapEditGroupBox.Controls.Add($lblValidationRule)

$txtValidationRule = New-Object System.Windows.Forms.TextBox
$txtValidationRule.Location = New-Object System.Drawing.Point($rightXControl, ($startY + 70 + $verticalSpace - 10))
$txtValidationRule.Size = New-Object System.Drawing.Size(200, 20)
$txtValidationRule.ReadOnly = $true
$txtValidationRule.BackColor = [System.Drawing.SystemColors]::Control
$mapEditGroupBox.Controls.Add($txtValidationRule)

#--------------------------
# EXPORT TAB
#--------------------------

# Export preview
$lblExportPreview = New-Object System.Windows.Forms.Label
$lblExportPreview.Text = "Mapping File Preview:"
$lblExportPreview.Location = New-Object System.Drawing.Point(10, 10)
$lblExportPreview.Size = New-Object System.Drawing.Size(150, 20)
$tabExport.Controls.Add($lblExportPreview)

$txtExportPreview = New-Object System.Windows.Forms.TextBox
$txtExportPreview.Location = New-Object System.Drawing.Point(10, 35)
$txtExportPreview.Size = New-Object System.Drawing.Size(950, 400)
$txtExportPreview.Multiline = $true
$txtExportPreview.ScrollBars = "Vertical"
$txtExportPreview.ReadOnly = $true
$tabExport.Controls.Add($txtExportPreview)

# Export options
$exportOptionsGroupBox = New-Object System.Windows.Forms.GroupBox
$exportOptionsGroupBox.Text = "Export Options"
$exportOptionsGroupBox.Location = New-Object System.Drawing.Point(10, 450)
$exportOptionsGroupBox.Size = New-Object System.Drawing.Size(950, 100)
$tabExport.Controls.Add($exportOptionsGroupBox)

$lblExportFormat = New-Object System.Windows.Forms.Label
$lblExportFormat.Text = "Export Format:"
$lblExportFormat.Location = New-Object System.Drawing.Point(20, 30)
$lblExportFormat.Size = New-Object System.Drawing.Size(100, 20)
$exportOptionsGroupBox.Controls.Add($lblExportFormat)

$cboExportFormat = New-Object System.Windows.Forms.ComboBox
$cboExportFormat.Location = New-Object System.Drawing.Point(130, 28)
$cboExportFormat.Size = New-Object System.Drawing.Size(150, 20)
$cboExportFormat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboExportFormat.Items.Add("CSV")
$cboExportFormat.Items.Add("PowerShell Script")
$cboExportFormat.SelectedIndex = 0
$exportOptionsGroupBox.Controls.Add($cboExportFormat)

$chkIncludeValidation = New-Object System.Windows.Forms.CheckBox
$chkIncludeValidation.Text = "Include Validation Rules"
$chkIncludeValidation.Location = New-Object System.Drawing.Point(300, 28)
$chkIncludeValidation.Size = New-Object System.Drawing.Size(180, 20)
$chkIncludeValidation.Checked = $true
$exportOptionsGroupBox.Controls.Add($chkIncludeValidation)

$chkIncludeTransformations = New-Object System.Windows.Forms.CheckBox
$chkIncludeTransformations.Text = "Include Transformations"
$chkIncludeTransformations.Location = New-Object System.Drawing.Point(500, 28)
$chkIncludeTransformations.Size = New-Object System.Drawing.Size(180, 20)
$chkIncludeTransformations.Checked = $true
$exportOptionsGroupBox.Controls.Add($chkIncludeTransformations)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Mapping File"
$btnExport.Location = New-Object System.Drawing.Point(750, 60)
$btnExport.Size = New-Object System.Drawing.Size(180, 30)
$exportOptionsGroupBox.Controls.Add($btnExport)

#--------------------------
# EVENT HANDLERS
#--------------------------

# Browse for source file
$btnBrowseSource.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Source Data File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSourceFile.Text = $openFileDialog.FileName
    }
})

# Browse for target file
$btnBrowseTarget.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Target Requirements File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTargetFile.Text = $openFileDialog.FileName
    }
})

# Parse CSV with header
function Import-CsvWithHeader {
    param (
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "File not found: $FilePath"
    }
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    if ($extension -eq ".csv") {
        return Import-Csv -Path $FilePath
    }
    elseif ($extension -eq ".xlsx") {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "The ImportExcel module is required to read Excel files. Would you like to install it now?",
                "Module Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
                
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                try {
                    Install-Module -Name ImportExcel -Scope CurrentUser -Force
                    Import-Module ImportExcel
                }
                catch {
                    throw "Failed to install ImportExcel module: $_"
                }
            }
            else {
                throw "ImportExcel module is required but not installed."
            }
        }
        
        return Import-Excel -Path $FilePath
    }
    else {
        throw "Unsupported file format: $extension"
    }
}

# Load data
$btnLoadData.Add_Click({
    $txtLoadLog.Clear()
    $statusBarLabel.Text = "Loading data..."
    
    $sourceFile = $txtSourceFile.Text
    $targetFile = $txtTargetFile.Text
    
    if (-not $sourceFile -or -not (Test-Path $sourceFile)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid source data file.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    if (-not $targetFile -or -not (Test-Path $targetFile)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid target requirements file.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Load source data
    try {
        $global:SourceData = Import-CsvWithHeader -FilePath $sourceFile
        $sourceFields = $global:SourceData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        
        $txtLoadLog.AppendText("Loaded source data with $($sourceFields.Count) fields`r`n")
        $txtLoadLog.AppendText("Source data has $($global:SourceData.Count) rows`r`n`r`n")
        
        # Update source fields listbox
        $lstSourceFields.Items.Clear()
        foreach ($field in $sourceFields) {
            $lstSourceFields.Items.Add($field)
        }
        
        # Store source fields for later
        $global:SourceFields = $sourceFields
        
        # Update source dropdown in mapping tab
        $cboSourceField.Items.Clear()
        $cboSourceField.Items.Add("")  # Empty option
        foreach ($field in $sourceFields) {
            $cboSourceField.Items.Add($field)
        }
    }
    catch {
        $txtLoadLog.AppendText("Error loading source data: $_`r`n")
        [System.Windows.Forms.MessageBox]::Show("Failed to load source data: $_", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Load target requirements
    try {
        $targetData = Import-CsvWithHeader -FilePath $targetFile
        
        # Extract target field requirements from the new requirements CSV
        $global:TargetFields = @()
        foreach ($row in $targetData) {
            $fieldName = $row.FieldName
            if (-not [string]::IsNullOrWhiteSpace($fieldName)) {
                $fieldInfo = @{
                    FieldName       = $fieldName
                    DataType        = $row.FieldType
                    Mandatory       = ($row.Mandatory -eq "Yes")
                    ValidationRule  = $row.ValidationRule
                    Notes           = $row.Notes
                    Transformation  = $false
                    TransformFunction = ""
                    Type            = $row.FieldType
                }
                $global:TargetFields += $fieldInfo
            }
        }
        
        $txtLoadLog.AppendText("Loaded $($global:TargetFields.Count) target field requirements`r`n")
        
        # Update target fields listbox
        $lstTargetFields.Items.Clear()
        foreach ($field in $global:TargetFields) {
            $displayName = $field.FieldName
            if ($field.Mandatory) {
                $displayName += " [Required]"
            }
            $lstTargetFields.Items.Add($displayName)
        }
    }
    catch {
        $txtLoadLog.AppendText("Error loading target requirements: $_`r`n")
        [System.Windows.Forms.MessageBox]::Show("Failed to load target requirements: $_", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $txtLoadLog.AppendText("All data loaded successfully.`r`nPlease go to the Analyze tab to continue.`r`n")
    $statusBarLabel.Text = "Data loaded successfully"
    
    # Enable tabs
    $tabAnalyze.Enabled = $true
    $tabMap.Enabled = $true
    $tabExport.Enabled = $true
})

# Calculate similarity score between two strings
function Get-StringSimilarity {
    param (
        [string]$String1,
        [string]$String2
    )
    
    $str1 = $String1.ToLower()
    $str2 = $String2.ToLower()
    
    # Start with a direct match check
    if ($str1 -eq $str2) {
        return 1.0
    }
    
    # Check for substring matches
    if ($str1.Contains($str2) -or $str2.Contains($str1)) {
        $maxLength = [Math]::Max($str1.Length, $str2.Length)
        $minLength = [Math]::Min($str1.Length, $str2.Length)
        return $minLength / $maxLength * 0.9  # Slightly less than perfect match
    }
    
    # Split into words and check word overlap
    $words1 = $str1 -split '\W+' | Where-Object { $_ }
    $words2 = $str2 -split '\W+' | Where-Object { $_ }
    
    if ($words1.Count -gt 0 -and $words2.Count -gt 0) {
        $commonWords = $words1 | Where-Object { $words2 -contains $_ }
        $totalWords = ($words1 + $words2 | Select-Object -Unique).Count
        
        if ($totalWords -gt 0) {
            return $commonWords.Count / $totalWords * 0.8  # Word matching gets a good score
        }
    }
    
    # If no strong matches, calculate character-based similarity
    $setA = New-Object System.Collections.Generic.HashSet[char]
    foreach ($char in [char[]]$str1) { $setA.Add($char) | Out-Null }
    
    $setB = New-Object System.Collections.Generic.HashSet[char]
    foreach ($char in [char[]]$str2) { $setB.Add($char) | Out-Null }
    
    $intersection = 0
    foreach ($char in $setA) {
        if ($setB.Contains($char)) {
            $intersection++
        }
    }
    
    $union = $setA.Count + $setB.Count - $intersection
    
    if ($union -eq 0) {
        return 0.0
    }
    
    return $intersection / $union * 0.6  # Character-based similarity gets a lower weight
}

# Check if data is compatible with a given type
function Test-DataTypeCompatibility {
    param (
        [object[]]$Data,
        [string]$TargetType
    )
    
    # Skip if no data
    if (-not $Data -or $Data.Count -eq 0) {
        return $true
    }
    
    # Get a sample value (first non-null)
    $sample = $Data | Where-Object { $_ -ne $null -and $_ -ne "" } | Select-Object -First 1
    
    if (-not $sample) {
        return $true  # All nulls, assume compatible
    }
    
    switch ($TargetType.ToLower()) {
        "string" {
            # Everything can be a string
            return $true
        }
        "numeric" {
            # Check if can be converted to number
            try {
                [decimal]$sample | Out-Null
                return $true
            }
            catch {
                return $false
            }
        }
        "money" {
            # Check if can be converted to decimal
            try {
                [decimal]$sample | Out-Null
                return $true
            }
            catch {
                return $false
            }
        }
        "date" {
            # Check if can be parsed as date
            try {
                [datetime]$sample | Out-Null
                return $true
            }
            catch {
                # Check common date patterns
                $datePatterns = @(
                    '^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$',    # DD/MM/YYYY or MM/DD/YYYY
                    '^\d{4}[/-]\d{1,2}[/-]\d{1,2}$',       # YYYY/MM/DD
                    '^\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}$'  # DD Month YYYY
                )
                
                foreach ($pattern in $datePatterns) {
                    if ($sample -match $pattern) {
                        return $true
                    }
                }
                
                return $false
            }
        }
        "datetime" {
            # Check if can be parsed as datetime
            try {
                [datetime]$sample | Out-Null
                return $true
            }
            catch {
                return $false
            }
        }
        default {
            # Unknown type, assume compatible
            return $true
        }
    }
}

# Suggest a transformation function based on data type mismatch
function Get-TransformationSuggestion {
    param (
        [object[]]$Data,
        [string]$TargetType
    )
    
    if (-not $Data -or $Data.Count -eq 0) {
        return $null
    }
    
    # Get a few sample values
    $samples = $Data | Where-Object { $_ -ne $null -and $_ -ne "" } | Select-Object -First 5
    
    if (-not $samples -or $samples.Count -eq 0) {
        return $null
    }
    
    switch ($TargetType.ToLower()) {
        "date" {
            # If string that looks like date, suggest date parsing
            if ($samples[0] -is [string]) {
                $datePatterns = @(
                    '^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',    # DD/MM/YYYY or MM/DD/YYYY
                    '^\d{4}[/-]\d{1,2}[/-]\d{1,2}',       # YYYY/MM/DD
                    '^\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}'   # DD Month YYYY
                )
                
                foreach ($pattern in $datePatterns) {
                    if ($samples[0] -match $pattern) {
                        return "Date Format"
                    }
                }
            }
        }
        "datetime" {
            # Similar to date but for datetime
            if ($samples[0] -is [string]) {
                $datePatterns = @(
                    '^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',    # Any that starts with a date
                    '^\d{4}[/-]\d{1,2}[/-]\d{1,2}'       # Any that starts with a date
                )
                
                foreach ($pattern in $datePatterns) {
                    if ($samples[0] -match $pattern) {
                        return "Date Format"
                    }
                }
            }
        }
        "numeric" {
            # If string that contains non-numeric chars, suggest cleanup
            if ($samples[0] -is [string] -and $samples[0] -match '[^0-9.-]') {
                return "Numeric Only"
            }
        }
        "money" {
            # If string that contains non-numeric chars or currency symbols, suggest cleanup
            if ($samples[0] -is [string] -and $samples[0] -match '[^0-9.-]') {
                return "Numeric Only"
            }
        }
    }
    
    # No specific transformation needed
    return $null
}

# Source field selection event
$lstSourceFields.Add_SelectedIndexChanged({
    if ($lstSourceFields.SelectedItem) {
        $field = $lstSourceFields.SelectedItem
        $txtDataPreview.Clear()
        
        if ($global:SourceData) {
            # Get values for this field
            $values = $global:SourceData | Select-Object -ExpandProperty $field -ErrorAction SilentlyContinue
            
            if ($values) {
                # Get basic stats
                $totalCount = $values.Count
                $nonEmptyCount = ($values | Where-Object { $_ -ne $null -and $_ -ne "" }).Count
                $uniqueCount = ($values | Select-Object -Unique).Count
                
                # Display statistics
                $txtDataPreview.AppendText("Field: $field`r`n")
                $txtDataPreview.AppendText("Non-empty values: $nonEmptyCount of $totalCount ($([math]::Round(($nonEmptyCount/$totalCount)*100, 1))%)`r`n")
                $txtDataPreview.AppendText("Unique values: $uniqueCount`r`n`r`n")
                
                # Display sample values
                $txtDataPreview.AppendText("Sample values:`r`n")
                $values | Select-Object -First 10 | ForEach-Object {
                    $txtDataPreview.AppendText("- $_`r`n")
                }
            }
            else {
                $txtDataPreview.AppendText("No data available for field: $field")
            }
        }
    }
})

# Target field selection event
$lstTargetFields.Add_SelectedIndexChanged({
    if ($lstTargetFields.SelectedItem) {
        $selectedItem = $lstTargetFields.SelectedItem
        # Remove " [Required]" from the field name using the correct -replace syntax
        $fieldName = $selectedItem -replace ' $begin:math:display$Required$end:math:display$', ''
        
        $targetField = $global:TargetFields | Where-Object { $_.FieldName -eq $fieldName } | Select-Object -First 1
        
        if ($targetField) {
            $txtDataPreview.Clear()
            $txtDataPreview.AppendText("Target Field: $($targetField.FieldName)`r`n")
            $txtDataPreview.AppendText("Data Type: $($targetField.DataType)`r`n")
            
            # Compute the required field text before appending
            $req = if ($targetField.Mandatory) { 'Yes' } else { 'No' }
            $txtDataPreview.AppendText("Required: $req`r`n")
            
            if ($targetField.Validation) {
                $txtDataPreview.AppendText("Validation Rule: $($targetField.ValidationRule)`r`n")
            }
            
            if ($targetField.Transformation) {
                $txtDataPreview.AppendText("Default Transformation: $($targetField.TransformFunction)`r`n")
            }
            
            # Show current mapping if exists
            if ($global:Mappings.ContainsKey($targetField.FieldName)) {
                $mapping = $global:Mappings[$targetField.FieldName]
                $txtDataPreview.AppendText("`r`nCurrent Mapping:`r`n")
                $txtDataPreview.AppendText("Mapped from: $($mapping.SourceField)`r`n")
                
                if ($mapping.Transformation) {
                    $txtDataPreview.AppendText("Transformation: $($mapping.Transformation)`r`n")
                }
            }
            
            # Show mapping suggestions if available
            if ($global:SuggestedMappings.ContainsKey($targetField.FieldName)) {
                $suggestions = $global:SuggestedMappings[$targetField.FieldName]
                $txtDataPreview.AppendText("`r`nSuggested mappings:`r`n")
                
                foreach ($suggestion in $suggestions | Select-Object -First 5) {
                    $txtDataPreview.AppendText("- $($suggestion.SourceField) (score: $([math]::Round($suggestion.Score, 2)))`r`n")
                }
            }
        }
    }
})

$scriptContent += "`$mappings | Export-Csv -Path 'field_mappings.csv' -NoTypeInformation`r`n"

# In your Analyze code:
$btnAnalyze.Add_Click({
    if (-not $global:SourceFields -or -not $global:TargetFields) {
        [System.Windows.Forms.MessageBox]::Show("Please load source and target data first", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $statusBarLabel.Text = "Analyzing fields..."
    $global:SuggestedMappings = @{}
    
    foreach ($targetField in $global:TargetFields) {
        $fieldName = $targetField.FieldName.Trim()
        $suggestions = @()
        
        # Case-insensitive exact match
        $exactMatches = $global:SourceFields | Where-Object { $_.Trim().ToLower() -eq $fieldName.ToLower() }
        foreach ($m in $exactMatches) {
            $suggestions += @{ SourceField = $m; Score = 1.0; Reason = "Exact match" }
        }
        
        if (-not $suggestions) {
            foreach ($sourceField in $global:SourceFields) {
                $similarity = Get-StringSimilarity -String1 $fieldName -String2 $sourceField.Trim()
                if ($similarity -ge 0.5) {
                    $suggestions += @{ SourceField = $sourceField; Score = $similarity; Reason = "Similar name" }
                }
            }
        }
        
        $sorted = $suggestions | Sort-Object -Property Score -Descending
        $global:SuggestedMappings[$fieldName] = $sorted
        
        # If top suggestion's Score >= 0.99, we auto-assign
        if ($sorted.Count -gt 0 -and $sorted[0].Score -ge 0.99) {
            $bestMatch = $sorted[0].SourceField
            $global:Mappings[$fieldName] = @{
                SourceField    = $bestMatch
                Transformation = $null
                Notes          = $targetField.Notes
            }
        }
    }
    
    Update-MappingGrid
    $tabControl.SelectedTab = $tabMap
    $statusBarLabel.Text = "Analysis complete"
    [System.Windows.Forms.MessageBox]::Show("Field analysis complete. Please review and edit the suggested mappings.", "Analysis Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

function Update-MappingGrid {
    $dgvMappings.Rows.Clear()
    
    foreach ($targetField in $global:TargetFields) {
        $fieldName = $targetField.FieldName
        $required = if ($targetField.Mandatory) { "Yes" } else { "No" }
        
        if ($global:Mappings.ContainsKey($fieldName)) {
            $mapping = $global:Mappings[$fieldName]
            $sourceField = $mapping.SourceField
            $transform = $mapping.Transformation
        }
        else {
            $sourceField = ""
            $transform = ""
        }
        
        # Add a new row
        $rowIndex = $dgvMappings.Rows.Add($fieldName, $sourceField, $transform, $required)
        
        # If mandatory + missing source field => highlight pink
        if ($targetField.Mandatory -and -not $sourceField) {
            $dgvMappings.Rows[$rowIndex].Cells[1].Style.BackColor = [System.Drawing.Color]::LightPink
        }
        else {
            # Reset to white (or a default color)
            $dgvMappings.Rows[$rowIndex].Cells[1].Style.BackColor = [System.Drawing.Color]::White
        }
    }
    
    Update-ExportPreview
}

# Selection changed in mapping grid
$dgvMappings.Add_SelectionChanged({
    if ($dgvMappings.SelectedRows.Count -gt 0) {
        # Get the selected row and target field name from column 0
        $selectedRow = $dgvMappings.SelectedRows[0]
        $targetFieldName = $selectedRow.Cells[0].Value
        $txtTargetField.Text = $targetFieldName

        # Load mapping details if a mapping exists for this target field
        if ($global:Mappings.ContainsKey($targetFieldName)) {
            $mapping = $global:Mappings[$targetFieldName]
            $cboSourceField.Text = $mapping.SourceField
            if ($mapping.Transformation) {
                $cboTransform.Text = $mapping.Transformation
            }
            else {
                $cboTransform.Text = "None"
            }
        }
        else {
            $cboSourceField.Text = ""
            $cboTransform.SelectedIndex = 0  # "None" is at index 0
        }

        # Retrieve the target field from the requirements
        $targetField = $global:TargetFields | Where-Object { $_.FieldName -eq $targetFieldName } | Select-Object -First 1
        if ($targetField) {
            $txtNotesRight.Text = $targetField.Notes
            $txtFieldType.Text = $targetField.DataType
            $txtValidationRule.Text = $targetField.ValidationRule
        }
        else {
            $txtNotes.Text = ""
            $txtFieldType.Text = ""
            $txtValidationRule.Text = ""
        }
    }
})

# Update mapping
$btnUpdateMapping.Add_Click({
    $targetField = $txtTargetField.Text
    $sourceField = $cboSourceField.Text
    $transformChoice = $cboTransform.Text
    $notesValue = $txtNotesRight.Text  # adjust if your notes textbox is named differently

    if (-not $targetField) {
        [System.Windows.Forms.MessageBox]::Show("No target field selected", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # If the transformation selected is a custom text transform (its name starts with "Custom-Text-")
    if ($transformChoice -like "Custom-Text-*") {
        # For the grid, display "Transform" in the Source Field
        $global:Mappings[$targetField] = @{
            SourceField    = "Transform"
            Transformation = $transformChoice
            Notes          = $notesValue
        }
    }
    else {
        if ($sourceField) {
            $global:Mappings[$targetField] = @{
                SourceField    = $sourceField
                Transformation = if ($transformChoice -eq "None") { $null } else { $transformChoice }
                Notes          = $notesValue
            }
        }
        else {
            $global:Mappings.Remove($targetField)
        }
    }
    
    Update-MappingGrid
    $statusBarLabel.Text = "Updated mapping for $targetField"
})

# Update export preview
function Update-ExportPreview {
    $format = $cboExportFormat.Text
    $includeValidation = $chkIncludeValidation.Checked
    $includeTransformations = $chkIncludeTransformations.Checked
    
    $txtExportPreview.Clear()
    
    switch ($format) {
        "CSV" {
            # Generate CSV content with "Required" as the last column
            $csvContent = "SourceField,NewField,DataType,Mandatory,Validation,ValidationRule,Transformation,TransformFunction,Required`r`n"
            
            foreach ($targetField in $global:TargetFields) {
                $fieldName = $targetField.FieldName
                
                # Skip if no mapping or if mapping has neither a source field nor a transformation
                if (-not $global:Mappings.ContainsKey($fieldName)) {
                    continue
                }
                $mapping = $global:Mappings[$fieldName]
                if (-not $mapping.SourceField -and -not $mapping.Transformation) {
                    continue
                }
                
                # Now we know this field is matched in some way
                $dataType = $targetField.DataType
                $mandatory = if ($targetField.Mandatory) { "Y" } else { "N" }
                
                $sourceField = $mapping.SourceField
                $transform = $mapping.Transformation
                
                # If transform is a custom text transform, skip the source field
                if ($transform -like "Custom-Text-*") {
                    $sourceField = ""
                }
                
                $required = if ($targetField.Mandatory) { "Y" } else { "N" }
                
                # Validation
                $validation = "N"
                $validationRule = ""
                if ($includeValidation -and $targetField.ValidationRule) {
                    $validation = "Y"
                    $validationRule = $targetField.ValidationRule
                }
                
                # Transformation
                $transformation = "N"
                $transformFunc = ""
                if ($includeTransformations -and $transform) {
                    $transformation = "Y"
                    $transformFunc = $transform
                }
                
                # Add row to CSV
                $csvContent += "$sourceField,$fieldName,$dataType,$mandatory,$validation,`"$validationRule`",$transformation,`"$transformFunc`",$required`r`n"
            }
            
            $txtExportPreview.Text = $csvContent
        }
        "PowerShell Script" {
            # Similar logic for skipping unmapped fields or fields with neither source nor transform
            $scriptContent = "# Mapping file for Library Data Transformation Tool`r`n"
            $scriptContent += "# Created on $(Get-Date)`r`n`r`n"
            $scriptContent += "`$mappings = @(`r`n"
            
            foreach ($targetField in $global:TargetFields) {
                $fieldName = $targetField.FieldName
                if (-not $global:Mappings.ContainsKey($fieldName)) {
                    continue
                }
                $mapping = $global:Mappings[$fieldName]
                if (-not $mapping.SourceField -and -not $mapping.Transformation) {
                    continue
                }
                
                # We have a valid mapping
                $scriptContent += "    @{`r`n"
                $scriptContent += "        SourceField = '$($mapping.SourceField)'`r`n"
                $scriptContent += "        NewField = '$fieldName'`r`n"
                $scriptContent += "        DataType = '$($targetField.DataType)'`r`n"
                $scriptContent += "        Mandatory = '$(if($targetField.Mandatory){'Y'}else{'N'})'`r`n"
                
                if ($includeValidation -and $targetField.ValidationRule) {
                    $scriptContent += "        Validation = 'Y'`r`n"
                    $scriptContent += "        ValidationRule = '$($targetField.ValidationRule)'`r`n"
                }
                else {
                    $scriptContent += "        Validation = 'N'`r`n"
                    $scriptContent += "        ValidationRule = ''`r`n"
                }
                
                if ($includeTransformations -and $mapping.Transformation) {
                    $scriptContent += "        Transformation = 'Y'`r`n"
                    $scriptContent += "        TransformFunction = '$($mapping.Transformation)'`r`n"
                }
                else {
                    $scriptContent += "        Transformation = 'N'`r`n"
                    $scriptContent += "        TransformFunction = ''`r`n"
                }
                
                $scriptContent += "        Required = '$(if($targetField.Mandatory){'Y'}else{'N'})'`r`n"
                $scriptContent += "    },`r`n"
            }
            
            $scriptContent += ")`r`n`r`n"
            $scriptContent += "# Export to CSV`r`n"
            $scriptContent += "`$mappings | Export-Csv -Path 'field_mappings.csv' -NoTypeInformation`r`n"
            
            $txtExportPreview.Text = $scriptContent
        }
    }
}

# Export format changed
$cboExportFormat.Add_SelectedIndexChanged({
    Update-ExportPreview
})

# Include validation changed
$chkIncludeValidation.Add_CheckedChanged({
    Update-ExportPreview
})

# Include transformations changed
$chkIncludeTransformations.Add_CheckedChanged({
    Update-ExportPreview
})

# Export mapping
$btnExport.Add_Click({
    # 1) Check for any required fields that are unmapped or have neither a source field nor a transformation
    $unmappedMandatory = @()
    foreach ($targetField in $global:TargetFields) {
        if ($targetField.Mandatory) {
            if (-not $global:Mappings.ContainsKey($targetField.FieldName)) {
                $unmappedMandatory += $targetField.FieldName
            }
            else {
                $mapping = $global:Mappings[$targetField.FieldName]
                if (-not $mapping.SourceField -and -not $mapping.Transformation) {
                    $unmappedMandatory += $targetField.FieldName
                }
            }
        }
    }

    if ($unmappedMandatory.Count -gt 0) {
        $message = "The following required fields have no source or transformation:`r`n`r`n" + 
                   ($unmappedMandatory -join "`r`n") + 
                   "`r`n`r`nDo you want to continue exporting?"
        
        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Unmapped Required Fields",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::No) {
            return
        }
    }
    
    # 2) Check if we have any mappings at all
    if (-not $global:Mappings -or $global:Mappings.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No mappings have been created yet.", 
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # 3) Proceed with the export
    $format = $cboExportFormat.Text
    
    if ($format -eq "CSV") {
        $extension = ".csv"
        $filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    }
    else {
        $extension = ".ps1"
        $filter = "PowerShell Files (*.ps1)|*.ps1|All Files (*.*)|*.*"
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = $filter
    $saveFileDialog.Title = "Save Mapping File"
    $saveFileDialog.DefaultExt = $extension
    $saveFileDialog.FileName = "field_mappings$extension"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $txtExportPreview.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8
            
            $statusBarLabel.Text = "Mapping exported to $($saveFileDialog.FileName)"
            [System.Windows.Forms.MessageBox]::Show(
                "Mapping has been exported to $($saveFileDialog.FileName)", 
                "Export Successful", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to export mapping: $_", 
                "Export Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $statusBarLabel.Text = "Export failed"
        }
    }
})

# Show the form
$mainForm.ShowDialog()
