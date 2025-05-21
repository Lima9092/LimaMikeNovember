# PowerShell Data Anonymiser with GUI
# - Loads a CSV
# - Uses header->rule mappings defined at top
# - Anonymises fields based on configured rules
# - Displays data in Excel-style grid view
# - Toggles between real and anonymised data

# --- Configuration: map input headers to anonymisation types ---
$global:FieldMappings = @{
    forename       = 'Forename'
    surname        = 'Surname'
    barcode        = 'PRN'
    prison_id      = 'PRN'
    wing           = 'Wing'
    cell           = 'Cell'
    language       = 'Language'
    DOB            = 'DOB'
    ethnicity      = 'Ethnicity'
    release_date   = 'ReleaseDate'
    'Borrowed By'  = 'FullName'   # combines Forename and Surname
    Username       = 'PRN'
    borrbarcode    = 'PRN'
}

# Global variables for data storage
$global:OriginalData = $null
$global:AnonymisedData = $null
$global:IsShowingAnonymised = $false
$global:UseSuffix = $false

# --- Utility functions ---
function Get-AdultBirthDate {
    param(
        [int]$MinAge = 18,
        [int]$MaxAge = 80,
        [string]$Format = 'dd/MM/yyyy'
    )
    $t = Get-Date
    $earliest = $t.AddYears(-$MaxAge)
    $latest   = $t.AddYears(-$MinAge)
    $range    = ($latest - $earliest).Days
    if ($range -le 0) { $range = 365 * ($MaxAge - $MinAge) }
    $days = Get-Random -Minimum 0 -Maximum $range
    $dt = $earliest.AddDays($days)
    return $dt.ToString($Format)
}

function Get-RecentDate {
    param(
        [int]$MaxMonthsAgo = 6,
        [string]$Format = 'dd/MM/yyyy'
    )
    $t = Get-Date
    $earliest = $t.AddMonths(-$MaxMonthsAgo)
    $range    = ($t - $earliest).Days
    if ($range -le 0) { $range = 1 }
    $days = Get-Random -Minimum 0 -Maximum $range
    $dt = $earliest.AddDays($days)
    return $dt.ToString($Format)
}

function Get-PastDate {
    param(
        [int]$MinMonthsAgo = 6,
        [int]$MaxYearsAgo  = 5,
        [string]$Format    = 'dd/MM/yyyy'
    )
    $t        = Get-Date
    $latest   = $t.AddMonths(-$MinMonthsAgo)
    $earliest = $t.AddYears(-$MaxYearsAgo)
    $range    = ($latest - $earliest).Days
    if ($range -le 0) { $range = 365 * $MaxYearsAgo }
    $days = Get-Random -Minimum 0 -Maximum $range
    $dt   = $earliest.AddDays($days)
    return $dt.ToString($Format)
}

function Get-RandomString {
    param(
        [int]$Length   = 8,
        [string]$CharSet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    )
    $s    = ''
    $rand = New-Object System.Random
    for ($i=0; $i -lt $Length; $i++) { $s += $CharSet[$rand.Next(0, $CharSet.Length)] }
    return $s
}

# --- Value generator ---
function Get-GeneratedValue {
    param(
        [string]$fieldName,
        [string]$fieldType,
        [object]$validationRule,
        [PSObject]$rowData = $null  # Added parameter for row context
    )
    switch ($fieldType.ToLower()) {
        'prn'         { 
            $prnBase = 'T' + (Get-Random -Minimum 1000 -Maximum 9999) + (Get-Random -InputObject @('AA','AB','AC','ND','HD'))
            
            # Add branch suffix if enabled and branch column exists
            if ($global:UseSuffix -and $rowData -ne $null) {
                # Try to find a branch field in the row
                $branchField = $rowData.PSObject.Properties.Name | Where-Object { 
                    $_ -like "*branch*" -or $_ -eq "branch" 
                } | Select-Object -First 1
                
                if ($branchField -and ![string]::IsNullOrWhiteSpace($rowData.$branchField)) {
                    return "$prnBase-$($rowData.$branchField)"
                }
            }
            
            return $prnBase
        }
        'libraryno'   { (Get-Random -Minimum 100000000 -Maximum 999999999).ToString() }
        'forename'    { (Get-Random -InputObject @('John','James','Steve','Dave','Phil','Graham','Pete','Paul','Mohammed','George','Jeremy','Alex','William','Ryan','Daniel','Michael','Robert','Thomas','Chris','Matthew')) + '-T' }
        'surname'     { (Get-Random -InputObject @('Smith','Davies','Dobson','Collins','Clarke','Roberts','Welsh','Jones','Green','Brown','Mohammed','Wilson','Taylor','Anderson','Johnson','Thompson','White','Martin','Walker','Robinson')) + '-T' }
        'fullname'    {
            $first = Get-GeneratedValue -fieldName $fieldName -fieldType 'Forename' -validationRule $null
            $last  = Get-GeneratedValue -fieldName $fieldName -fieldType 'Surname'  -validationRule $null
            "$first $last"
        }
        'wing'        { Get-Random -InputObject @('Main','West','East','South','North') }
        'cell'        { Get-Random -InputObject @('E83','W27','N40','S28','E73','W34','N39','S93') }
        'language'    { Get-Random -InputObject @('English','Spanish','Arabic','Welsh') }
        'dob'         { Get-AdultBirthDate }
        'ethnicity'   { Get-Random -InputObject @('White British','White','Welsh','Asian British','Black British','Indian','Pakistani','Chinese','Caribbean') }
        'releasedate' { Get-PastDate -MinMonthsAgo 6 -MaxYearsAgo 5 -Format 'dd/MM/yyyy' }
        default       { Get-RandomString }
    }
}

# === GUI setup using WinForms ===
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CSV Data Anonymiser'
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = 'CenterScreen'
$form.Icon = [System.Drawing.SystemIcons]::Application

# View label
$lblViewMode = New-Object System.Windows.Forms.Label
$lblViewMode.Text = 'View: REAL DATA'
$lblViewMode.Location = New-Object System.Drawing.Point(510, 15)
$lblViewMode.Size = New-Object System.Drawing.Size(130, 20)
$lblViewMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblViewMode.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($lblViewMode)

# --- Excel-style DataGridView for CSV display ---
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Size = New-Object System.Drawing.Size(760, 450)
$dataGrid.Location = New-Object System.Drawing.Point(10, 50)
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.ReadOnly = $true
$dataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGrid.BackgroundColor = [System.Drawing.Color]::White
$dataGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dataGrid.RowHeadersVisible = $false
$dataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGrid.MultiSelect = $false
$dataGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$dataGrid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::White
$dataGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightBlue
$dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$dataGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($dataGrid)

# Function to load data into grid
function Load-DataToGrid {
    param([array]$data)
    
    # Clear existing data
    $dataGrid.DataSource = $null
    $dataGrid.Columns.Clear()
    
    if ($data -eq $null -or $data.Count -eq 0) {
        return
    }
    
    # Create DataTable for binding
    $dataTable = New-Object System.Data.DataTable
    
    # Add columns based on first row properties
    $columnNames = $data[0].PSObject.Properties.Name
    foreach ($column in $columnNames) {
        [void]$dataTable.Columns.Add($column)
    }
    
    # Add rows
    foreach ($item in $data) {
        $row = $dataTable.NewRow()
        foreach ($column in $columnNames) {
            $row[$column] = $item.$column
        }
        [void]$dataTable.Rows.Add($row)
    }
    
    # Set data source
    $dataGrid.DataSource = $dataTable
    
    # Format grid for Excel-like appearance
    foreach ($column in $dataGrid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    }
    
    # Update row count display
    $lblStatus.Text = "Loaded $($data.Count) records"
}

# --- Load CSV button ---
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Load CSV'
$btnLoad.Location = New-Object System.Drawing.Point(10, 10)
$btnLoad.Size = New-Object System.Drawing.Size(90, 30)
$btnLoad.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnLoad.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    if ($dlg.ShowDialog() -eq 'OK') {
        try {
            $lblStatus.Text = "Loading data from $($dlg.FileName)..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # Import CSV with error handling
            $global:OriginalData = Import-Csv $dlg.FileName
            $global:CsvData = $global:OriginalData
            $global:AnonymisedData = $null
            $global:IsShowingAnonymised = $false
            $lblViewMode.Text = 'View: REAL DATA'
            $lblViewMode.ForeColor = [System.Drawing.Color]::DarkBlue
            
            # Load data into grid
            Load-DataToGrid -data $global:CsvData
            
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblStatus.Text = "Loaded $($global:CsvData.Count) records from $($dlg.FileName)"
        }
        catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show(
                "Error loading CSV: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $lblStatus.Text = "Error loading CSV file."
        }
    }
})
$form.Controls.Add($btnLoad)

# --- Anonymise button ---
$btnAnon = New-Object System.Windows.Forms.Button
$btnAnon.Text = 'Anonymise'
$btnAnon.Location = New-Object System.Drawing.Point(110, 10)
$btnAnon.Size = New-Object System.Drawing.Size(90, 30)
$btnAnon.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnAnon.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnAnon.Add_Click({
    if (-not $global:OriginalData) {
        [System.Windows.Forms.MessageBox]::Show(
            'No CSV loaded. Please load a file first.',
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = "Anonymising data..."
        
        # Create a deep copy of the original data
        $global:AnonymisedData = @()
        foreach ($originalRow in $global:OriginalData) {
            $newRow = New-Object PSObject
            foreach ($prop in $originalRow.PSObject.Properties) {
                $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
            }
            $global:AnonymisedData += $newRow
        }
        
        # Build mapping of actual CSV columns to types
        $validMappings = @{}
        $missingKeys = @()
        foreach ($mapping in $global:FieldMappings.GetEnumerator()) {
            $mapKey = $mapping.Key
            $mapVal = $mapping.Value
            $match = $global:AnonymisedData[0].PSObject.Properties.Name |
                    Where-Object { $_.ToLower() -eq $mapKey.ToLower() }
            if ($match) {
                $validMappings[$match] = $mapVal
            } else {
                $missingKeys += $mapKey
            }
        }
        
        if ($validMappings.Count -eq 0) {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show(
                'No matching fields found to anonymise. Check your mappings.',
                'Error',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        # Apply anonymisation
        foreach ($row in $global:AnonymisedData) {
            foreach ($col in $validMappings.Keys) {
                try {
                    $row.$col = Get-GeneratedValue -fieldName $col -fieldType $validMappings[$col] -validationRule $null -rowData $row
                } catch {
                    # skip
                }
            }
        }
        
        # Switch to anonymised view
        $global:CsvData = $global:AnonymisedData
        $global:IsShowingAnonymised = $true
        $lblViewMode.Text = 'View: ANONYMISED'
        $lblViewMode.ForeColor = [System.Drawing.Color]::DarkRed
        
        # Refresh the data grid
        Load-DataToGrid -data $global:CsvData
        
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $msg = "Anonymisation complete for $($global:CsvData.Count) records."
        if ($missingKeys.Count -gt 0) {
            $msg += " Skipped mappings for: $($missingKeys -join ', ')"
        }
        $lblStatus.Text = $msg
        
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            'Anonymisation Complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.MessageBox]::Show(
            "Error during anonymisation: $_",
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})
$form.Controls.Add($btnAnon)

# --- Save CSV button ---
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Save CSV'
$btnSave.Location = New-Object System.Drawing.Point(210, 10)
$btnSave.Size = New-Object System.Drawing.Size(90, 30)
$btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnSave.Add_Click({
    if (-not $global:CsvData) {
        [System.Windows.Forms.MessageBox]::Show(
            'No data to save.',
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.DefaultExt = 'csv'
    $dlg.AddExtension = $true
    
    if ($dlg.ShowDialog() -eq 'OK') {
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $lblStatus.Text = "Saving data to $($dlg.FileName)..."
            
            $global:CsvData | Export-Csv -Path $dlg.FileName -NoTypeInformation
            
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblStatus.Text = "Data saved successfully to $($dlg.FileName)"
            
            [System.Windows.Forms.MessageBox]::Show(
                'CSV saved successfully',
                'Success',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show(
                "Error saving CSV: $_",
                'Error',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})
$form.Controls.Add($btnSave)

# --- Toggle View button ---
$btnToggleView = New-Object System.Windows.Forms.Button
$btnToggleView.Text = 'Toggle View'
$btnToggleView.Location = New-Object System.Drawing.Point(310, 10)
$btnToggleView.Size = New-Object System.Drawing.Size(90, 30)
$btnToggleView.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnToggleView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnToggleView.Add_Click({
    if (-not $global:OriginalData) {
        [System.Windows.Forms.MessageBox]::Show(
            'No data loaded. Please load a file first.',
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if (-not $global:AnonymisedData) {
        [System.Windows.Forms.MessageBox]::Show(
            'No anonymised data available. Please anonymise data first.',
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $global:IsShowingAnonymised = -not $global:IsShowingAnonymised
    
    if ($global:IsShowingAnonymised) {
        $global:CsvData = $global:AnonymisedData
        $lblViewMode.Text = 'View: ANONYMISED'
        $lblViewMode.ForeColor = [System.Drawing.Color]::DarkRed
    } else {
        $global:CsvData = $global:OriginalData
        $lblViewMode.Text = 'View: REAL DATA'
        $lblViewMode.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    
    # Refresh the data grid
    Load-DataToGrid -data $global:CsvData
})
$form.Controls.Add($btnToggleView)

# --- Suffix Toggle Button ---
$btnSuffix = New-Object System.Windows.Forms.Button
$btnSuffix.Text = 'Suffix: OFF'
$btnSuffix.Location = New-Object System.Drawing.Point(410, 10)
$btnSuffix.Size = New-Object System.Drawing.Size(90, 30)
$btnSuffix.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnSuffix.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnSuffix.BackColor = [System.Drawing.Color]::LightGray
$btnSuffix.Add_Click({
    $global:UseSuffix = -not $global:UseSuffix
    
    if ($global:UseSuffix) {
        $btnSuffix.Text = 'Suffix: ON'
        $btnSuffix.BackColor = [System.Drawing.Color]::LightGreen
        $lblStatus.Text = "Branch suffix will be added to PRN values during anonymization"
    } else {
        $btnSuffix.Text = 'Suffix: OFF'
        $btnSuffix.BackColor = [System.Drawing.Color]::LightGray
        $lblStatus.Text = "Branch suffix will NOT be added to PRN values"
    }
    
    # If we already have anonymized data, re-anonymize with new setting
    if ($global:AnonymisedData -and $global:AnonymisedData.Count -gt 0) {
        # Ask user if they want to re-anonymize with new suffix setting
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to re-anonymize the data with the new suffix setting?",
            "Re-Anonymize Data?",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq 'Yes') {
            # Trigger anonymize function by clicking the anonymize button
            $btnAnon.PerformClick()
        }
    }
})
$form.Controls.Add($btnSuffix)

# --- Add search/filter capability ---
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Filter:"
$lblSearch.Location = New-Object System.Drawing.Point(10, 530)
$lblSearch.Size = New-Object System.Drawing.Size(40, 20)
$lblSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(50, 530)
$txtSearch.Size = New-Object System.Drawing.Size(150, 20)
$txtSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$txtSearch.Add_TextChanged({
    if (-not $global:CsvData) { return }
    
    $searchText = $txtSearch.Text.ToLower()
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        # If search is empty, show all data
        Load-DataToGrid -data $global:CsvData
        return
    }
    
    # Filter data based on search text
    $filteredData = $global:CsvData | Where-Object {
        $row = $_
        $found = $false
        foreach ($prop in $row.PSObject.Properties) {
            if ($prop.Value -and $prop.Value.ToString().ToLower().Contains($searchText)) {
                $found = $true
                break
            }
        }
        $found
    }
    
    # Load filtered data
    Load-DataToGrid -data $filteredData
    $lblStatus.Text = "Found $($filteredData.Count) records containing '$searchText'"
})
$form.Controls.Add($txtSearch)

# --- Clear filter button ---
$btnClearFilter = New-Object System.Windows.Forms.Button
$btnClearFilter.Text = "Clear Filter"
$btnClearFilter.Location = New-Object System.Drawing.Point(210, 530)
$btnClearFilter.Size = New-Object System.Drawing.Size(90, 20)
$btnClearFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnClearFilter.Add_Click({
    $txtSearch.Clear()
    if ($global:CsvData) {
        Load-DataToGrid -data $global:CsvData
    }
})
$form.Controls.Add($btnClearFilter)

# --- Status label (moved to bottom) ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready. Please load a CSV file.'
$lblStatus.Location = New-Object System.Drawing.Point(320, 530)
$lblStatus.Size = New-Object System.Drawing.Size(450, 20)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lblStatus)

# Form resize event handler
$form.Add_Resize({
    # Make sure controls resize with the form
    $dataGrid.Width = $form.ClientSize.Width - 20
    $dataGrid.Height = $form.ClientSize.Height - 110
    
    # Ensure status label stays at bottom and scales properly
    $lblStatus.Width = $form.ClientSize.Width - 330
})

# Show the form
[void]$form.ShowDialog()
