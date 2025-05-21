################################################################################
# PowerShell Data Anonymiser with GUI
# - Loads a CSV
# - Uses header->rule mappings defined at top
# - Anonymises fields based on configured rules
################################################################################

# --- Configuration: map input headers to anonymisation types ---
$global:FieldMappings = @{
    forename        = 'Forename'
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

# --- Utility functions ---
function Get-AdultBirthDate {
    param(
        [int]$MinAge = 18,
        [int]$MaxAge = 80,
        [string]$Format = 'dd/MM/yyyy'
    )
    $t = Get-Date
    $earliest = $t.AddYears(-$MaxYearsAgo)
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
        [object]$validationRule
    )
    switch ($fieldType.ToLower()) {
        'prn'         { 'A' + (Get-Random -Minimum 1000 -Maximum 9999) + (Get-Random -InputObject @('AA','AB','AC','ND','HD')) }
        'libraryno'   { (Get-Random -Minimum 100000000 -Maximum 999999999).ToString() }
        'forename'    { Get-Random -InputObject @('John','James','Steve','Dave','Phil','Graham','Pete','Paul','Mohammed','George','Jeremy') }
        'surname'     { Get-Random -InputObject @('Smith','Davies','Dobson','Collins','Clarke','Roberts','Welsh','Jones','Green','Brown','Mohammed') }
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

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CSV Data Anonymiser'
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = 'CenterScreen'

# Data grid to display CSV
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Size     = New-Object System.Drawing.Size(760,450)
$dataGrid.Location = New-Object System.Drawing.Point(10,50)
$form.Controls.Add($dataGrid)

# --- Load CSV button ---
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load CSV'
$btnLoad.Location = New-Object System.Drawing.Point(10,10)
$btnLoad.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    if ($dlg.ShowDialog() -eq 'OK') {
        $global:CsvData   = Import-Csv $dlg.FileName
        $dataGrid.DataSource = $global:CsvData
        [System.Windows.Forms.MessageBox]::Show(
            "Loaded $($global:CsvData.Count) records from $($dlg.FileName)",
            "Load Successful"
        )
    }
})
$form.Controls.Add($btnLoad)

# --- Anonymise button ---
$btnAnon = New-Object System.Windows.Forms.Button
$btnAnon.Text     = 'Anonymise'
$btnAnon.Location = New-Object System.Drawing.Point(110,10)
$btnAnon.Add_Click({
    if (-not $global:CsvData) {
        [System.Windows.Forms.MessageBox]::Show('No CSV loaded. Please load a file first.', 'Error')
        return
    }
    # Build mapping of actual CSV columns to types
    $validMappings = @{}
    $missingKeys   = @()
    foreach ($mapping in $global:FieldMappings.GetEnumerator()) {
        $mapKey = $mapping.Key
        $mapVal = $mapping.Value
        $match  = $global:CsvData[0].PSObject.Properties.Name |
                  Where-Object { $_.ToLower() -eq $mapKey.ToLower() }
        if ($match) {
            $validMappings[$match] = $mapVal
        } else {
            $missingKeys += $mapKey
        }
    }
    if ($validMappings.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No matching fields found to anonymise. Check your mappings.', 'Error')
        return
    }
    # Apply anonymisation
    foreach ($row in $global:CsvData) {
        foreach ($col in $validMappings.Keys) {
            try {
                $row.$col = Get-GeneratedValue -fieldName $col -fieldType $validMappings[$col] -validationRule $null
            } catch {
                # skip
            }
        }
    }
    $dataGrid.Refresh()
    $msg = "Anonymisation complete for $($global:CsvData.Count) records."
    if ($missingKeys.Count -gt 0) {
        $msg += " Skipped mappings for: $($missingKeys -join ', ')"
    }
    [System.Windows.Forms.MessageBox]::Show($msg, 'Anonymisation Complete')
})
$form.Controls.Add($btnAnon)

# --- Save CSV button ---
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Save CSV'
$btnSave.Location = New-Object System.Drawing.Point(210,10)
$btnSave.Add_Click({
    if (-not $global:CsvData) {
        [System.Windows.Forms.MessageBox]::Show('No data to save.', 'Error')
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    if ($dlg.ShowDialog() -eq 'OK') {
        $global:CsvData | Export-Csv -Path $dlg.FileName -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show('Anonymised CSV saved successfully', 'Success')
    }
})
$form.Controls.Add($btnSave)

# Show the form
[void]$form.ShowDialog()
