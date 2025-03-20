# Global variables
$global:FieldDefs = $null
$global:DataTable = $null

function Get-AdultBirthDate {
    param (
        [int]$MinAge = 18,
        [int]$MaxAge = 80,
        [string]$Format = "dd/MM/yyyy"
    )
    $today = Get-Date
    $earliestDate = $today.AddYears(-$MaxAge)
    $latestDate = $today.AddYears(-$MinAge)
    $daysRange = ($latestDate - $earliestDate).Days
    if ($daysRange -le 0) { $daysRange = 365 * ($MaxAge - $MinAge) }
    $randomDays = Get-Random -Minimum 0 -Maximum $daysRange
    $resultDate = $earliestDate.AddDays($randomDays)
    Write-Host "Generated birth date: $($resultDate.ToString($Format)) (age: $([Math]::Floor(($today - $resultDate).TotalDays / 365)) years)"
    return $resultDate.ToString($Format)
}

function Get-RecentDate {
    param (
        [int]$MaxMonthsAgo = 6,
        [string]$Format = "dd/MM/yyyy"
    )
    $today = Get-Date
    $earliestDate = $today.AddMonths(-$MaxMonthsAgo)
    $daysRange = ($today - $earliestDate).Days
    if ($daysRange -le 0) { $daysRange = 1 }
    $randomDays = Get-Random -Minimum 0 -Maximum $daysRange
    $resultDate = $earliestDate.AddDays($randomDays)
    Write-Host "Generated recent date: $($resultDate.ToString($Format))"
    return $resultDate.ToString($Format)
}

function Get-FutureDate {
    param (
        [int]$MinYearsAhead = 1,
        [int]$MaxYearsAhead = 3,
        [string]$Format = "dd/MM/yyyy"
    )
    $today = Get-Date
    $earliestFuture = $today.AddYears($MinYearsAhead)
    $latestFuture = $today.AddYears($MaxYearsAhead)
    $daysRange = ($latestFuture - $earliestFuture).Days
    if ($daysRange -le 0) { $daysRange = 365 * ($MaxYearsAhead - $MinYearsAhead) }
    $randomDays = Get-Random -Minimum 0 -Maximum $daysRange
    $futureDate = $earliestFuture.AddDays($randomDays)
    return $futureDate.ToString($Format)
}

# Load required assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Generator Functions
function Get-RandomString {
    param (
        [int]$Length = 10,
        [string]$CharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    )
    $result = ""
    $random = New-Object System.Random
    for ($i = 0; $i -lt $Length; $i++) {
        $result += $CharSet[$random.Next(0, $CharSet.Length)]
    }
    return $result
}

function Get-RandomDate {
    param (
        [int]$MinYear = 2010,
        [int]$MaxYear = 2023
    )
    $year = Get-Random -Minimum $MinYear -Maximum ($MaxYear + 1)
    $month = Get-Random -Minimum 1 -Maximum 13
    $day = Get-Random -Minimum 1 -Maximum 29
    return (Get-Date -Year $year -Month $month -Day $day).ToString("yyyy-MM-dd")
}

function Get-RandomDateTime {
    param (
        [int]$MinYear = 2010,
        [int]$MaxYear = 2023
    )
    $year = Get-Random -Minimum $MinYear -Maximum ($MaxYear + 1)
    $month = Get-Random -Minimum 1 -Maximum 13
    $day = Get-Random -Minimum 1 -Maximum 29
    $hour = Get-Random -Minimum 0 -Maximum 24
    $minute = Get-Random -Minimum 0 -Maximum 60
    $second = Get-Random -Minimum 0 -Maximum 60
    return (Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $minute -Second $second).ToString("yyyy-MM-dd HH:mm:ss")
}

function Get-RandomNumeric {
    param (
        [int]$Min = 1,
        [int]$Max = 1000
    )
    return Get-Random -Minimum $Min -Maximum $Max
}

function Get-GeneratedValue {
    param(
        $fieldName,
        $fieldType,
        $validationRule
    )
    
    switch ($fieldName.ToLower()) {
        "barcode" { $value = "C" + (Get-Random -Minimum 100000000 -Maximum 999999999).ToString() }
        "type" { $value = (Get-Random -InputObject @(0, 6)) }
        "category" { $value = (Get-Random -InputObject @("ADULT", "CHILD", "SENIOR", "YOUNG ADULT")) }
        "category2" { 
            $categories = @("0-4", "5-7", "8-12", "13-17", "18-24", "25-29", "30-40", "41-50", "51-60", "61-70")
            $value = (Get-Random -InputObject $categories)
        }
        "branch" { $value = (Get-Random -InputObject @("NCLNC", "MAIN", "EAST", "WEST", "NORTH", "SOUTH")) }
        "location" { $value = (Get-Random -InputObject @("GEN", "FIC", "REF")) }
        "awaybranch" { $value = (Get-Random -InputObject @("NCLNC", "MAIN", "EAST", "WEST", "NORTH", "SOUTH")) }
        "rcn" { $value = "BX" + (Get-Random -Minimum 10000000 -Maximum 100000000).ToString() }
        "holdback" { 
            $date = Get-RandomDate -MinYear 2020 -MaxYear 2023
            $dateObj = [datetime]::ParseExact($date, "yyyy-MM-dd", $null)
            $value = $dateObj.ToString("dd/MM/yyyy")
        }
        "totalissues" { $value = (Get-Random -Minimum 0 -Maximum 10) }
        "added" { 
            $date = Get-RandomDate -MinYear 1995 -MaxYear 2023
            $dateObj = [datetime]::ParseExact($date, "yyyy-MM-dd", $null)
            $value = $dateObj.ToString("dd/MM/yyyy")
        }
        "expirydate" { $value = Get-FutureDate -MinYearsAhead 1 -MaxYearsAhead 3 }
        "seen" { 
            $date = Get-RecentDate -MaxMonthsAgo 6
            $time = (Get-Date -Hour (Get-Random -Minimum 8 -Maximum 22) -Minute (Get-Random -Minimum 0 -Maximum 60) -Second (Get-Random -Minimum 0 -Maximum 60)).ToString("HH:mm:ss")
            $value = "$date $time" 
        }
        "seen2" { 
            $date = Get-RecentDate -MaxMonthsAgo 6
            $time = (Get-Date -Hour (Get-Random -Minimum 8 -Maximum 22) -Minute (Get-Random -Minimum 0 -Maximum 60) -Second (Get-Random -Minimum 0 -Maximum 60)).ToString("HH:mm:ss")
            $value = "$date $time" 
        }
        "activitycount" { $value = (Get-Random -Minimum 1 -Maximum 1000) }
        "title" { $value = (Get-Random -InputObject @("Mr", "Mrs", "Ms", "Miss")) }
        "forename" { $value = (Get-Random -InputObject @("John", "Jane", "Michael", "Sarah", "David", "Emma", "Robert", "Lisa", "Thomas", "Elizabeth")) }
        "initials" { $value = (Get-Random -InputObject @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")) }
        "surname" { $value = (Get-Random -InputObject @("Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia", "Wilson", "Taylor")) }
        "birth" { $value = Get-AdultBirthDate -MinAge 18 -MaxAge 80 -Format "dd/MM/yyyy" }
        "catbirth" { $value = Get-AdultBirthDate -MinAge 18 -MaxAge 80 -Format "dd/MM/yyyy" }
        "tickets" { $value = (Get-Random -Minimum 1 -Maximum 5) }
        "overduesmailcode" { $value = (Get-Random -Minimum 1 -Maximum 10) }
        "itemavailablemailcode" { $value = (Get-Random -Minimum 1 -Maximum 10) }
        "preoverduesmailcode" { $value = (Get-Random -Minimum 1 -Maximum 10) }
        "homedelivery" { $value = (Get-Random -InputObject @(0, 1)) }
        "firstaddress_subdwelling" { 
            $prefix = (Get-Random -InputObject @("Flat", "Apartment", "Unit", "Suite"))
            $number = (Get-Random -Minimum 1 -Maximum 20)
            $value = "$prefix $number" 
        }
        "firstaddress_streetnumber" { $value = (Get-Random -Minimum 1 -Maximum 200) }
        "firstaddress_street" { 
            $streetTypes = @("Street", "Road", "Avenue", "Drive", "Lane", "Way", "Close", "Crescent")
            $streetNames = @("Main", "High", "Church", "Park", "Oak", "Pine", "Maple", "Cedar", "Elm", "Some")
            $streetType = (Get-Random -InputObject $streetTypes)
            $streetName = (Get-Random -InputObject $streetNames)
            $value = "$streetName $streetType" 
        }
        "firstaddress_locality" { $value = (Get-Random -InputObject @("Locality", "A Locality", "District", "Area", "Village", "Suburb")) }
        "firstaddress_town" { $value = (Get-Random -InputObject @("A Town", "Nottingham", "Birmingham", "Manchester", "London", "Liverpool", "Edinburgh")) }
        "firstaddress_postcode" { 
            $first = (Get-Random -InputObject @("AB", "CD", "EF", "GH", "NG", "ST", "PST"))
            $second = (Get-Random -Minimum 1 -Maximum 99)
            $third = (Get-Random -InputObject @("CDE", "XYZ", "ABC", "DEF", "GHI"))
            $value = "$first$second $third" 
        }
        "firstaddress_addressflag" { $value = 1 }
        "firstaddress_active" { $value = 1 }
        "firsttelephone_1" { 
            $areaCode = (Get-Random -Minimum 1000 -Maximum 10000)
            $phoneNumber = (Get-Random -Minimum 100000 -Maximum 1000000)
            $value = "0$areaCode$phoneNumber" 
        }
        "email" { 
            $domains = @("hotmail.com", "gmail.com", "yahoo.com", "outlook.com", "btinternet.com")
            $domain = (Get-Random -InputObject $domains)
            if ($global:DataTable -and $global:DataTable.Rows.Count -gt 0) {
                $row = $global:DataTable.Rows[0]
                $fn = if ($row["forename"]) { $row["forename"].ToLower() } else { "user" }
                $sn = if ($row["surname"]) { $row["surname"].ToLower() } else { (Get-Random -Minimum 100 -Maximum 999) }
                $value = "$fn.$sn@$domain"
            } else {
                $value = "user" + (Get-Random -Minimum 100 -Maximum 999) + "@$domain"
            }
        }
        "mobiletelephone" { $value = "07" + (Get-Random -Minimum 100000000 -Maximum 1000000000).ToString() }
        "dpflag" { $value = (Get-Random -InputObject @("Yes", "No")) }
        "dpdate" { $value = Get-RandomDate -MinYear 2005 -MaxYear 2023 }
        "cost" { $value = [Math]::Round((Get-Random -Minimum 499 -Maximum 3999) / 100, 2) }
        "shelf" { $value = (Get-Random -Minimum 1 -Maximum 11) }
        "parentflag" { $value = (Get-Random -InputObject @(0, 1, 2)) }
        "parentitem" { 
            $value = "C" + (Get-Random -Minimum 100000000 -Maximum 999999999).ToString() 
        }
        "issuesequence" { $value = (Get-Random -Minimum 1 -Maximum 10) }
        "supplier" { 
            $publishers = @(
                "Penguin Random House", 
                "HarperCollins", 
                "Simon & Schuster", 
                "Hachette Book Group", 
                "Macmillan Publishers"
            )
            $value = (Get-Random -InputObject $publishers)
        }
        "size" { 
            $bookSizes = @(
                "A-format (178 x 111mm)",
                "B-format (198 x 129mm)",
                "C-format (216 x 135mm)",
                "Trade paperback (235 x 155mm)",
                "Royal (234 x 156mm)"
            )
            $value = (Get-Random -InputObject $bookSizes)
        }
        "filling" { $value = (Get-Random -Minimum 1 -Maximum 101) }
        "sequence" { $value = (Get-Random -Minimum 1 -Maximum 11) }
        
        # Circulation/loan record fields
        "borrbarcode" { 
            $prefix = (Get-Random -InputObject @("D0", "D7"))
            $value = $prefix + (Get-Random -Minimum 10000000 -Maximum 100000000).ToString() 
        }
        "itembarcode" { $value = "C" + (Get-Random -Minimum 100000000 -Maximum 999999999).ToString() }
        "issuebranch" { 
            $value = (Get-Random -InputObject @("NCLNC", "NCLBU", "NCLAS", "MAIN", "EAST", "WEST", "NORTH", "SOUTH")) 
        }
        
        # --- CHANGED to date‐only:
        "issuedate" {
            $today = Get-Date
            $daysBack = Get-Random -Minimum 14 -Maximum 60
            $issueDate = $today.AddDays(-$daysBack)
            $value = $issueDate.ToString("dd/MM/yyyy")
        }
        "duedate" { 
            $today = Get-Date
            $daysToAdd = Get-Random -Minimum 14 -Maximum 28
            $dueDate = $today.AddDays($daysToAdd)
            $value = $dueDate.ToString("dd/MM/yyyy")
        }
        "renewdate" { 
            if ((Get-Random -Minimum 0 -Maximum 10) -lt 4) {
                # Sometimes empty
                $value = ""
            } else {
                # Otherwise a recent date
                $today = Get-Date
                $daysBack = Get-Random -Minimum 1 -Maximum 14
                $renewDate = $today.AddDays(-$daysBack)
                $value = $renewDate.ToString("dd/MM/yyyy")
            }
        }
        
        "renewcount" { 
            $value = (Get-Random -Minimum 0 -Maximum 4).ToString()
        }
        "unseenrenewcount" { 
            $value = (Get-Random -Minimum 0 -Maximum 3).ToString()
        }

        # --- CHANGED: use same list as "branch"
        "renewbranch" {
            $value = (Get-Random -InputObject @("NCLNC","MAIN","EAST","WEST","NORTH","SOUTH"))
        }
        
        "parentissueflag" { 
            $value = (Get-Random -Minimum 0 -Maximum 3).ToString()
        }

        # This second "parentitem" we assume is needed for circulation
        "parentitem" {
            if ((Get-Random -Minimum 0 -Maximum 10) -lt 3) {
                $value = "C" + (Get-Random -Minimum 100000000 -Maximum 999999999).ToString()
            } else {
                $value = ""
            }
        }
        "idcheck" { 
            $idTypes = @("Utilities Bill", "Driving Licence", "Passport", "Bank Statement", "Council Tax Bill", "Parents ticket")
            $numIDs = Get-Random -Minimum 1 -Maximum 3
            $selectedIDs = @()
            for ($i = 0; $i -lt $numIDs; $i++) {
                $selectedIDs += (Get-Random -InputObject $idTypes)
            }
            $value = [string]::Join(",", $selectedIDs)
        }
        "marketing" { $value = (Get-Random -InputObject @("Yes", "No")) }
        "pin" { $value = (Get-Random -Minimum 1000 -Maximum 10000).ToString() }
        "gender" {
            if ($validationRule -eq "^(?:Male|Female)$") {
                $value = (Get-Random -InputObject @("Male", "Female"))
            } else {
                $value = (Get-Random -InputObject @(1, 2))
            }
        }
        "ethnic" { $value = (Get-Random -InputObject @("White British", "Asian", "Black", "Mixed", "Other")) }
        "language" { $value = (Get-Random -InputObject @("English", "French", "Spanish", "German", "Polish", "Urdu", "Arabic")) }
        "resident" { $value = 1 }
        "wing" { $value = "Wing " + (Get-Random -InputObject @("A", "B", "C", "D", "E", "F")) }
        "cell" { $value = "Cell " + (Get-Random -Minimum 100 -Maximum 500).ToString() }
        "prison_id" { 
            $letter1 = "A"
            $number = (Get-Random -Minimum 100 -Maximum 1000).ToString()
            $letter2 = (Get-Random -InputObject @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
            $letter3 = (Get-Random -InputObject @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
            $value = "$letter1$number$letter2$letter3"
        }
        "release_date" { 
            try {
                $value = Get-FutureDate -MinYearsAhead 1 -MaxYearsAhead 5
            }
            catch {
                $today = Get-Date
                $futureDate = $today.AddYears(1).AddDays((Get-Random -Minimum 0 -Maximum 365))
                $value = $futureDate.ToString("dd/MM/yyyy")
                Write-Host "Used fallback for release_date: $value"
            }
        }
        default { $value = $null }
    }
    
    # If nothing was generated but a validation rule or type is present:
    if ($null -eq $value) {
        if ($validationRule -and $validationRule.Trim() -ne "") {
            if ($validationRule -eq "^(?:Male|Female)$") {
                $value = (Get-Random -InputObject @("Male", "Female"))
            }
            elseif ($validationRule -match "^\^\[A-Z0-9\]") {
                $length = 8
                if ($validationRule -match "\{(\d+),(\d+)\}") {
                    $minLength = [int]$Matches[1]
                    $maxLength = [int]$Matches[2]
                    $length = Get-Random -Minimum $minLength -Maximum ($maxLength + 1)
                }
                elseif ($validationRule -match "\{(\d+)\}") {
                    $length = [int]$Matches[1]
                }
                $value = Get-RandomString -Length $length -CharSet "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            }
            elseif ($validationRule -match "^\^\.\{(\d+),(\d+)\}\$") {
                $minLength = [int]$Matches[1]
                $maxLength = [int]$Matches[2]
                $length = Get-Random -Minimum $minLength -Maximum ($maxLength + 1)
                $value = Get-RandomString -Length $length
            }
        } else {
            if ($fieldType -like "*String*") {
                $maxLength = 10
                if ($fieldType -match "\((\d+)\)") {
                    $maxLength = [int]$Matches[1]
                }
                $value = Get-RandomString -Length ([Math]::Min(5, $maxLength))
            }
            elseif ($fieldType -eq "Numeric") {
                $value = (Get-Random -Minimum 1 -Maximum 100)
            }
            elseif ($fieldType -eq "Date") {
                $value = Get-RandomDate
            }
            elseif ($fieldType -eq "DateTime") {
                $value = Get-RandomDateTime
            }
            else {
                $value = "$fieldName-" + (Get-Random -Minimum 1000 -Maximum 9999).ToString()
            }
        }
    }
    
    # Check validation
    if ($validationRule -and $validationRule.Trim() -ne "" -and $value) {
        try {
            if (-not ($value -match $validationRule)) {
                Write-Host "Warning: Generated value '$value' for field '$fieldName' doesn't match validation rule: $validationRule"
            }
        } catch {
            Write-Host "Warning: Invalid regex pattern for field '$fieldName': $validationRule"
        }
    }
    
    return $value
}

# Form creation
$form = New-Object System.Windows.Forms.Form
$form.Text = "Simplified Test Data Generator"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# CSV File Selection
$lblReq = New-Object System.Windows.Forms.Label
$lblReq.Text = "Requirements CSV File:"
$lblReq.Location = New-Object System.Drawing.Point(10, 10)
$lblReq.AutoSize = $true
$form.Controls.Add($lblReq)

$txtReq = New-Object System.Windows.Forms.TextBox
$txtReq.Location = New-Object System.Drawing.Point(150, 10)
$txtReq.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($txtReq)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(660, 8)
$btnBrowse.Size = New-Object System.Drawing.Size(75, 25)
$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title = "Select Requirements CSV File"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtReq.Text = $ofd.FileName
        try {
            $global:FieldDefs = Import-Csv -Path $txtReq.Text
            $txtLog.AppendText("Loaded $($global:FieldDefs.Count) field definitions.`r`n")
        }
        catch {
            $errorMsg = $_.Exception.Message
            $txtLog.AppendText("Error loading CSV: $errorMsg`r`n")
        }
    }
})
$form.Controls.Add($btnBrowse)

# Number of entries
$lblEntries = New-Object System.Windows.Forms.Label
$lblEntries.Text = "Number of Entries:"
$lblEntries.Location = New-Object System.Drawing.Point(10, 45)
$lblEntries.AutoSize = $true
$form.Controls.Add($lblEntries)

$txtEntries = New-Object System.Windows.Forms.TextBox
$txtEntries.Location = New-Object System.Drawing.Point(150, 42)
$txtEntries.Size = New-Object System.Drawing.Size(100, 20)
$txtEntries.Text = "50"
$form.Controls.Add($txtEntries)

# DataGridView
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(5, 80)
$grid.Size = New-Object System.Drawing.Size(990, 400)
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grid)

# Log area
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(5, 490)
$txtLog.Size = New-Object System.Drawing.Size(990, 80)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($txtLog)

# Generate button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Data"
$btnGenerate.Location = New-Object System.Drawing.Point(10, 580)
$btnGenerate.Size = New-Object System.Drawing.Size(120, 30)
$btnGenerate.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnGenerate.Add_Click({
    if (-not $global:FieldDefs) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid Requirements CSV file", "Error")
        return
    }
    try {
        $entryCount = [int]$txtEntries.Text
        $txtLog.AppendText("Generating $entryCount rows...`r`n")
        $dataTable = New-Object System.Data.DataTable
        foreach ($field in $global:FieldDefs) {
            if ($field.FieldName -and $field.FieldName.Trim() -ne "") {
                $dataTable.Columns.Add($field.FieldName.Trim()) | Out-Null
                $txtLog.AppendText("Added column: $($field.FieldName.Trim())`r`n")
            }
        }
        for ($i = 0; $i -lt $entryCount; $i++) {
            $row = $dataTable.NewRow()
            foreach ($field in $global:FieldDefs) {
                if ($field.FieldName -and $field.FieldName.Trim() -ne "") {
                    $fieldName = $field.FieldName.Trim()
                    $fieldType = $field.FieldType
                    try {
                        $validationRule = $field.ValidationRule
                        $generatedValue = Get-GeneratedValue -fieldName $fieldName -fieldType $fieldType -validationRule $validationRule
                        $row[$fieldName] = $generatedValue
                        
                        # If still empty but field is Mandatory, try one more time
                        if ([string]::IsNullOrEmpty($row[$fieldName]) -and $field.Mandatory -eq "Yes") {
                            $txtLog.AppendText("Warning: Empty value for mandatory field '$fieldName', regenerating...`r`n")
                            $row[$fieldName] = Get-GeneratedValue -fieldName $fieldName -fieldType $fieldType -validationRule $validationRule
                        }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        $txtLog.AppendText("Error generating value for field '$fieldName': $errorMsg`r`n")
                        $row[$fieldName] = "ERROR"
                    }
                }
            }
            $dataTable.Rows.Add($row)
        }
        $global:DataTable = $dataTable
        $grid.DataSource = $dataTable
        $txtLog.AppendText("Generated $entryCount rows successfully`r`n")
    }
    catch {
        $errorMsg = $_.Exception.Message
        $txtLog.AppendText("Error: $errorMsg`r`n")
    }
})
$form.Controls.Add($btnGenerate)

# Export button
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Location = New-Object System.Drawing.Point(140, 580)
$btnExport.Size = New-Object System.Drawing.Size(120, 30)
$btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnExport.Add_Click({
    if (-not $global:DataTable -or $global:DataTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please generate data first.", "Error")
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $sfd.Title = "Save Generated Data As"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportPath = $sfd.FileName
        try {
            $global:DataTable | Export-Csv -Path $exportPath -NoTypeInformation
            $txtLog.AppendText("Data exported to $exportPath`r`n")
            [System.Windows.Forms.MessageBox]::Show("Data exported to $exportPath", "Export Successful")
        }
        catch {
            $errorMsg = $_.Exception.Message
            $txtLog.AppendText("Error exporting data: $errorMsg`r`n")
        }
    }
})
$form.Controls.Add($btnExport)

# Clear Log button
$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text = "Clear Log"
$btnClearLog.Location = New-Object System.Drawing.Point(270, 580)
$btnClearLog.Size = New-Object System.Drawing.Size(80, 30)
$btnClearLog.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnClearLog.Add_Click({
    $txtLog.Clear()
})
$form.Controls.Add($btnClearLog)

# Resize event
$form.Add_Resize({
    $leftMargin = 5
    $rightMargin = 5
    $grid.Width = $form.ClientSize.Width - $leftMargin - $rightMargin
    $grid.Left = $leftMargin
    $grid.Height = $form.ClientSize.Height - $grid.Top - $txtLog.Height - $btnGenerate.Height - 20
    $txtLog.Top = $grid.Bottom + 10
    $txtLog.Width = $grid.Width
    $txtLog.Left = $leftMargin
    $btnGenerate.Top = $txtLog.Bottom + 10
    $btnExport.Top = $txtLog.Bottom + 10
    $btnClearLog.Top = $txtLog.Bottom + 10
})

[void]$form.ShowDialog()
