# Convert an array of PSCustomObjects to a DataTable
function ConvertTo-DataTable {
    param([Parameter(ValueFromPipeline=$true)] $InputObject)
    $dt = New-Object System.Data.DataTable
    if (-not $InputObject -or $InputObject.Count -eq 0) { 
        # Make sure we at least have some columns
        $dt.Columns.Add("Message") | Out-Null
        $row = $dt.NewRow()
        $row["Message"] = "No data available"
        $dt.Rows.Add($row)
        return $dt 
    }
    $props = $InputObject[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    foreach ($prop in $props) { [void]$dt.Columns.Add($prop) }
    foreach ($obj in $InputObject) {
        $row = $dt.NewRow()
        foreach ($prop in $props) { $row[$prop] = $obj.$prop }
        $dt.Rows.Add($row)
    }
    return $dt
}# Global Variables
$global:KnownMatchesCSVFile = "Known-Matches.csv"  # Default file (current directory)
$global:FuzzyThreshold = 0.75  # Minimum fuzzy match score
$global:reportData = @()  # Added global variable to store report data

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Compute Levenshtein distance using a jagged array
function Get-LevenshteinDistance {
    param([string]$s, [string]$t)
    $s = if ([string]::IsNullOrEmpty($s)) { "" } else { $s }
    $t = if ([string]::IsNullOrEmpty($t)) { "" } else { $t }
    $n = $s.Length; $m = $t.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $d = @()
    for ($i = 0; $i -le $n; $i++) {
         $row = New-Object 'int[]' ($m+1)
         $d += ,$row
    }
    for ($i = 0; $i -le $n; $i++) { $d[$i][0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0][$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
         for ($j = 1; $j -le $m; $j++) {
             $cost = if ($s[$i-1] -eq $t[$j-1]) { 0 } else { 1 }
             $a = $d[$i-1][$j] + 1
             $b = $d[$i][$j-1] + 1
             $c = $d[$i-1][$j-1] + $cost
             $d[$i][$j] = [Math]::Min($a, [Math]::Min($b, $c))
         }
    }
    return $d[$n][$m]
}

# Return fuzzy match ratio (0 to 1)
function FuzzyMatch {
    param([string]$source, [string]$dest)
    $maxLength = [Math]::Max($source.Length, $dest.Length)
    if ($maxLength -eq 0) { return 1 }
    $distance = Get-LevenshteinDistance $source.ToLower() $dest.ToLower()
    return 1 - ($distance / $maxLength)
}

# Load CSV files from a folder
function Load-CsvFiles {
    param([string]$FolderPath)
    $csvFiles = Get-ChildItem -Path $FolderPath -Filter *.csv
    $tables = @{}
    foreach ($file in $csvFiles) {
        try {
            $data = Import-Csv -Path $file.FullName
            $tableName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $tables[$tableName] = $data
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading file: $($file.Name)`n$_")
        }
    }
    return $tables
}

# Extract header fields from source CSVs (assumes header row)
function Get-Fields {
    param([object]$TableData)
    if ($TableData.Count -gt 0) { return $TableData[0].PSObject.Properties.Name }
    return @()
}

# For requirements CSVs, assume the required field names are in a column named "FieldName"
function Get-RequirementFields {
    param([object]$TableData)
    $fields = @()
    foreach ($row in $TableData) {
        if ($row.FieldName) {
            $fields += $row.FieldName
        }
    }
    return $fields
}

# Load known matches from CSV; create file if not present
function Load-KnownMatches {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) {
        "Requirement,Known-Fields" | Out-File -Encoding utf8 $FilePath
    }
    $mapping = @{}
    try {
        $km = Import-Csv -Path $FilePath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading Known-Matches file.`n$_")
        return $mapping
    }
    foreach ($row in $km) {
        $req = $row.Requirement.ToLower().Trim()
        $knownRaw = $row.'Known-Fields'
        $knownFields = $knownRaw -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        $mapping[$req] = $knownFields
    }
    return $mapping
}

# Compare destination fields to source fields and generate mapping results
function Match-Fields {
    param(
        [hashtable]$DestFields,
        [hashtable]$SourceFields,
        [hashtable]$KnownMatches,
        [double]$Threshold = $global:FuzzyThreshold
    )
    $mappingResults = @{}
    foreach ($destTable in $DestFields.Keys) {
        $destFieldList = $DestFields[$destTable]
        $mappingResults[$destTable] = @()
        foreach ($dfield in $destFieldList) {
            $bestMatch = $null; $bestScore = 0; $candidateMatches = @(); $known = @()
            if ($KnownMatches.ContainsKey($dfield.ToLower())) { $known = $KnownMatches[$dfield.ToLower()] }
            foreach ($sourceTable in $SourceFields.Keys) {
                $sfields = $SourceFields[$sourceTable]
                foreach ($sfield in $sfields) {
                    $score = FuzzyMatch $sfield $dfield
                    if ($known.Count -gt 0 -and $known -contains $sfield.ToLower()) {
                        $score = [Math]::Max($score, 0.9)
                    }
                    if ($score -gt $bestScore) {
                        $bestScore = $score
                        $bestMatch = @{ SourceTable = $sourceTable; SourceField = $sfield; Score = $score }
                    }
                    if ($score -ge $Threshold) {
                        $candidateMatches += @{ SourceTable = $sourceTable; SourceField = $sfield; Score = $score }
                    }
                }
            }
            if ($bestScore -lt $global:FuzzyThreshold) { $bestMatch = $null }
            $mappingResults[$destTable] += @{
                DestField   = $dfield
                BestMatch   = $bestMatch
                Candidates  = $candidateMatches
            }
        }
    }
    return $mappingResults
}

# Generate a report (an array of custom objects) from mapping results for Requirements Audit
function Generate-CSVReport {
    param([hashtable]$MappingResults)
    $reportData = @()
    foreach ($destTable in $MappingResults.Keys) {
        foreach ($mapping in $MappingResults[$destTable]) {
            $destField = $mapping.DestField
            if ($mapping.BestMatch -ne $null) {
                $reportData += [PSCustomObject]@{
                    DestinationTable = $destTable
                    DestinationField = $destField
                    SourceTable      = $mapping.BestMatch.SourceTable
                    SourceField      = $mapping.BestMatch.SourceField
                    Score            = [math]::Round($mapping.BestMatch.Score,2)
                }
            } else {
                $reportData += [PSCustomObject]@{
                    DestinationTable = $destTable
                    DestinationField = $destField
                    SourceTable      = ""
                    SourceField      = "Not Matched"
                    Score            = ""
                }
            }
        }
    }
    return $reportData
}

# Generate a report from unmatched source fields for Source Audit.
function Generate-UnmatchedReport {
    param(
        [hashtable]$SourceFields,
        [hashtable]$MappingResults
    )
    $unmatched = @()
    foreach ($sourceTable in $SourceFields.Keys) {
        foreach ($field in $SourceFields[$sourceTable]) {
            $found = $false
            foreach ($destTable in $MappingResults.Keys) {
                foreach ($mapping in $MappingResults[$destTable]) {
                    if ($mapping.BestMatch -ne $null -and $mapping.BestMatch.SourceTable -eq $sourceTable -and $mapping.BestMatch.SourceField -eq $field) {
                        $found = $true
                        break
                    }
                }
                if ($found) { break }
            }
            if (-not $found) {
                $unmatched += [PSCustomObject]@{
                    SourceTable = $sourceTable
                    UnmatchedField = $field
                }
            }
        }
    }
    return $unmatched
}

# Function to apply color formatting to Requirements grid based on DestinationTable
function Apply-RequirementsFormatting {
    param([System.Windows.Forms.DataGridView]$Grid)
    
    # Create a dictionary to hold group colors for each DestinationTable.
    $currentDestTable = ""
    $useLightBlue = $true
    
    for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
        $row = $Grid.Rows[$i]
        if ($row.Cells.Count -lt 2 -or $row.IsNewRow) { continue } # Skip if row doesn't have enough cells
        
        $destTable = $row.Cells["DestinationTable"].Value
        if ([string]::IsNullOrEmpty($destTable)) {
            continue
        }
        
        # Change color scheme when destination table changes
        if ($destTable -ne $currentDestTable) {
            $currentDestTable = $destTable
            $useLightBlue = -not $useLightBlue
        }
        
        # Set background color based on alternating destination tables
        $row.DefaultCellStyle.BackColor = if ($useLightBlue) { 
            [System.Drawing.Color]::LightBlue 
        } else { 
            [System.Drawing.Color]::White 
        }
        
        # Always highlight "Not Matched" in the SourceField column with pink
        if ($row.Cells["SourceField"].Value -eq "Not Matched") {
            $row.Cells["SourceField"].Style.BackColor = [System.Drawing.Color]::LightPink
        }
    }
}

# Function to apply color formatting to Source grid based on SourceTable
function Apply-SourceFormatting {
    param([System.Windows.Forms.DataGridView]$Grid)
    
    # Apply alternating colors for Source Audit tab by SourceTable
    $currentSourceTable = ""
    $useSourceLightBlue = $true
    
    # First sort the grid by SourceTable if it's not already sorted
    $needToSort = $true
    foreach ($col in $Grid.Columns) {
        if ($col.HeaderText -eq "SourceTable" -and $col.SortMode -ne "NotSortable") {
            $needToSort = $false
            break
        }
    }
    
    # Apply formatting based on source table
    for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
        $row = $Grid.Rows[$i]
        if ($row.Cells.Count -lt 1 -or $row.IsNewRow) { continue } # Skip if row doesn't have enough cells
        
        $srcTable = $row.Cells["SourceTable"].Value
        if ([string]::IsNullOrEmpty($srcTable)) {
            continue
        }
        
        # Change color scheme when source table changes
        if ($srcTable -ne $currentSourceTable) {
            $currentSourceTable = $srcTable
            $useSourceLightBlue = -not $useSourceLightBlue
        }
        
        # Set background color based on alternating source tables
        $row.DefaultCellStyle.BackColor = if ($useSourceLightBlue) { 
            [System.Drawing.Color]::LightBlue 
        } else { 
            [System.Drawing.Color]::White 
        }
    }
}

# GUI Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Table Auditing Tool"
$form.Size = New-Object System.Drawing.Size(900,700)
$form.StartPosition = "CenterScreen"

# Create a TabControl with two tabs: Requirements and Source
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10,100)
$tabControl.Size = New-Object System.Drawing.Size(860,500)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$tabPageReq = New-Object System.Windows.Forms.TabPage
$tabPageReq.Text = "Requirements Audit"
$tabPageSource = New-Object System.Windows.Forms.TabPage
$tabPageSource.Text = "Source Audit"

$tabControl.TabPages.Add($tabPageReq)
$tabControl.TabPages.Add($tabPageSource)

# Add tab selection event to reapply formatting when switching tabs
$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabPageSource) {
        # Source Audit tab selected - make sure it's sorted and formatted
        foreach ($col in $dgvSource.Columns) {
            if ($col.Name -eq "SourceTable") {
                $dgvSource.Sort($col, [System.ComponentModel.ListSortDirection]::Ascending)
                break
            }
        }
        Apply-SourceFormatting -Grid $dgvSource
    }
    elseif ($tabControl.SelectedTab -eq $tabPageReq) {
        # Requirements Audit tab selected - make sure it's formatted
        Apply-RequirementsFormatting -Grid $dgvRequirements
    }
})

# DataGridView for Requirements Audit
$dgvRequirements = New-Object System.Windows.Forms.DataGridView
$dgvRequirements.Location = New-Object System.Drawing.Point(0,0)
$dgvRequirements.Size = $tabPageReq.ClientSize
$dgvRequirements.Dock = 'Fill'
$dgvRequirements.ReadOnly = $true
$dgvRequirements.AllowUserToAddRows = $false
$dgvRequirements.AutoGenerateColumns = $true
$dgvRequirements.AutoSizeColumnsMode = 'Fill'
$dgvRequirements.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    # Wait for sort to complete
    [System.Windows.Forms.Application]::DoEvents()
    # Apply formatting
    Apply-RequirementsFormatting -Grid $dgvRequirements
})
$tabPageReq.Controls.Add($dgvRequirements)

# DataGridView for Source Audit
$dgvSource = New-Object System.Windows.Forms.DataGridView
$dgvSource.Location = New-Object System.Drawing.Point(0,0)
$dgvSource.Size = $tabPageSource.ClientSize
$dgvSource.Dock = 'Fill'
$dgvSource.ReadOnly = $true
$dgvSource.AllowUserToAddRows = $false
$dgvSource.AutoGenerateColumns = $true
$dgvSource.AutoSizeColumnsMode = 'Fill'
$dgvSource.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    # Wait for sort to complete
    [System.Windows.Forms.Application]::DoEvents()
    # Apply formatting
    Apply-SourceFormatting -Grid $dgvSource
})
$tabPageSource.Controls.Add($dgvSource)

# Textboxes and Browse Buttons for folders
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Location = New-Object System.Drawing.Point(10,20)
$lblSource.Size = New-Object System.Drawing.Size(150,20)
$lblSource.Text = "Source CSV Folder:"
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(170,20)
$txtSource.Size = New-Object System.Drawing.Size(500,20)
$form.Controls.Add($txtSource)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Location = New-Object System.Drawing.Point(680,18)
$btnBrowseSource.Size = New-Object System.Drawing.Size(75,23)
$btnBrowseSource.Text = "Browse"
$btnBrowseSource.Add_Click({
    $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSource.Text = $folderDlg.SelectedPath
    }
})
$form.Controls.Add($btnBrowseSource)

$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Location = New-Object System.Drawing.Point(10,60)
$lblDest.Size = New-Object System.Drawing.Size(150,20)
$lblDest.Text = "Destination CSV Folder:"
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(170,60)
$txtDest.Size = New-Object System.Drawing.Size(500,20)
$form.Controls.Add($txtDest)

$btnBrowseDest = New-Object System.Windows.Forms.Button
$btnBrowseDest.Location = New-Object System.Drawing.Point(680,58)
$btnBrowseDest.Size = New-Object System.Drawing.Size(75,23)
$btnBrowseDest.Text = "Browse"
$btnBrowseDest.Add_Click({
    $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDest.Text = $folderDlg.SelectedPath
    }
})
$form.Controls.Add($btnBrowseDest)

# "Run Analysis" Button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(10,620)
$btnRun.Size = New-Object System.Drawing.Size(100,30)
$btnRun.Text = "Run Analysis"
$btnRun.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnRun.Add_Click({
    try {
        $sourceFolder = $txtSource.Text
        $destFolder = $txtDest.Text
        
        if (-not (Test-Path $sourceFolder)) {
            [System.Windows.Forms.MessageBox]::Show("Source folder not found.")
            return
        }
        if (-not (Test-Path $destFolder)) {
            [System.Windows.Forms.MessageBox]::Show("Destination folder not found.")
            return
        }
        
        # Load CSVs
        $sourceTables = Load-CsvFiles -FolderPath $sourceFolder
        if ($sourceTables.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No source CSV files found in the selected folder.")
            return
        }
        
        $destTables = Load-CsvFiles -FolderPath $destFolder
        if ($destTables.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No destination CSV files found in the selected folder.")
            return
        }
        
        # Extract fields from CSVs
        $sourceFields = @{}
        foreach ($tableName in $sourceTables.Keys) { 
            $sourceFields[$tableName] = Get-Fields -TableData $sourceTables[$tableName]
            Write-Host "Source table: $tableName, Fields: $($sourceFields[$tableName] -join ', ')"
        }
        
        $destFields = @{}
        foreach ($tableName in $destTables.Keys) { 
            # Use Get-RequirementFields (expects a column named "FieldName")
            $destFields[$tableName] = Get-RequirementFields -TableData $destTables[$tableName]
            Write-Host "Destination table: $tableName, Fields: $($destFields[$tableName] -join ', ')"
        }
        
        $knownMatches = Load-KnownMatches -FilePath $global:KnownMatchesCSVFile
        $mappingResults = Match-Fields -DestFields $destFields -SourceFields $sourceFields -KnownMatches $knownMatches -Threshold $global:FuzzyThreshold
        
        # Generate Requirements Audit report and sort by DestinationTable
        $global:reportData = Generate-CSVReport -MappingResults $mappingResults | Sort-Object DestinationTable
        
        # For debugging
        Write-Host "Generated report data count: $($global:reportData.Count)"
        if ($global:reportData.Count -gt 0) {
            Write-Host "Sample report data: $($global:reportData[0] | Out-String)"
        }
        
        # Create DataTable and a BindingSource for Requirements grid
        $dtReq = New-Object System.Data.DataTable
        foreach ($prop in @("DestinationTable", "DestinationField", "SourceTable", "SourceField", "Score")) {
            [void]$dtReq.Columns.Add($prop)
        }
        
        # Add rows to the DataTable
        foreach ($item in $global:reportData) {
            $row = $dtReq.NewRow()
            $row["DestinationTable"] = $item.DestinationTable
            $row["DestinationField"] = $item.DestinationField
            $row["SourceTable"] = $item.SourceTable
            $row["SourceField"] = $item.SourceField
            $row["Score"] = $item.Score
            $dtReq.Rows.Add($row)
        }
        
        # If no data, add a message row
        if ($dtReq.Rows.Count -eq 0) {
            $row = $dtReq.NewRow()
            $row["DestinationTable"] = "No data found"
            $dtReq.Rows.Add($row)
        }
        
        # Create a BindingSource
        $bs = New-Object System.Windows.Forms.BindingSource
        $bs.DataSource = $dtReq
        
        # Set the DataSource to the BindingSource
        $dgvRequirements.DataSource = $bs
        
        # Apply formatting to Requirements grid
        Apply-RequirementsFormatting -Grid $dgvRequirements

        # Generate Source Audit report (unmatched source fields)
        $unmatchedReport = Generate-UnmatchedReport -SourceFields $sourceFields -MappingResults $mappingResults | Sort-Object SourceTable
        
        # Create DataTable for Source audit
        $dtSource = New-Object System.Data.DataTable
        foreach ($prop in @("SourceTable", "UnmatchedField")) {
            [void]$dtSource.Columns.Add($prop)
        }
        
        # Add rows to the DataTable
        foreach ($item in $unmatchedReport) {
            $row = $dtSource.NewRow()
            $row["SourceTable"] = $item.SourceTable
            $row["UnmatchedField"] = $item.UnmatchedField
            $dtSource.Rows.Add($row)
        }
        
        # If no data, add a message row
        if ($dtSource.Rows.Count -eq 0) {
            $row = $dtSource.NewRow()
            $row["SourceTable"] = "No unmatched fields found"
            $dtSource.Rows.Add($row)
        }
        
        # Create a BindingSource
        $bsSource = New-Object System.Windows.Forms.BindingSource
        $bsSource.DataSource = $dtSource
        
        # Set the DataSource to the BindingSource
        $dgvSource.DataSource = $bsSource
        
        # Force proper formatting for Source Audit tab
        foreach ($col in $dgvSource.Columns) {
            if ($col.Name -eq "SourceTable") {
                $dgvSource.Sort($col, [System.ComponentModel.ListSortDirection]::Ascending)
                break
            }
        }
        Apply-SourceFormatting -Grid $dgvSource
        
        # Add a Refresh button for the Source Audit tab
        if (-not $form.Controls.Contains($btnRefreshSource)) {
            $btnRefreshSource = New-Object System.Windows.Forms.Button
            $btnRefreshSource.Location = New-Object System.Drawing.Point(330,620)
            $btnRefreshSource.Size = New-Object System.Drawing.Size(120,30)
            $btnRefreshSource.Text = "Refresh Source Tab"
            $btnRefreshSource.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
            $btnRefreshSource.Add_Click({
                foreach ($col in $dgvSource.Columns) {
                    if ($col.Name -eq "SourceTable") {
                        $dgvSource.Sort($col, [System.ComponentModel.ListSortDirection]::Ascending)
                        break
                    }
                }
                Apply-SourceFormatting -Grid $dgvSource
            })
            $form.Controls.Add($btnRefreshSource)
        }
        
        [System.Windows.Forms.MessageBox]::Show("Analysis complete. Requirements rows: " + $global:reportData.Count + "; Source unmatched rows: " + $unmatchedReport.Count)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error during analysis: $_`n$($_.ScriptStackTrace)")
    }
})
$form.Controls.Add($btnRun)

# "Save Report" Button – saves Requirements Audit data
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Location = New-Object System.Drawing.Point(120,620)
$btnSave.Size = New-Object System.Drawing.Size(100,30)
$btnSave.Text = "Save Report"
$btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnSave.Add_Click({
    try {
        $saveFileDlg = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDlg.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDlg.FileName = "RequirementsAudit.csv"
        
        if ($saveFileDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Check if we have data to save
            if ($null -eq $global:reportData -or $global:reportData.Count -eq 0) {
                # Try to get data from the grid instead
                $data = @()
                if ($dgvRequirements.Rows.Count -gt 0 -and $dgvRequirements.Columns.Count -gt 1) {
                    foreach ($row in $dgvRequirements.Rows) {
                        if ($row.IsNewRow) { continue }
                        
                        $obj = [PSCustomObject]@{
                            DestinationTable = $row.Cells["DestinationTable"].Value
                            DestinationField = $row.Cells["DestinationField"].Value
                            SourceTable = $row.Cells["SourceTable"].Value
                            SourceField = $row.Cells["SourceField"].Value
                            Score = $row.Cells["Score"].Value
                        }
                        $data += $obj
                    }
                    $data | Export-Csv -Path $saveFileDlg.FileName -NoTypeInformation -Force
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("No data available to save.")
                    return
                }
            }
            else {
                # Use the global data we already have
                $global:reportData | Export-Csv -Path $saveFileDlg.FileName -NoTypeInformation -Force
            }
            
            [System.Windows.Forms.MessageBox]::Show("Report saved to:`n$($saveFileDlg.FileName)")
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving report: $_")
    }
})
$form.Controls.Add($btnSave)

$form.Controls.Add($tabControl)
$form.ShowDialog() | Out-Null
