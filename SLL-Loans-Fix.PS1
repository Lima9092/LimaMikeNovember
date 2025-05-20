Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-LoanUserMapper {
    [void][System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Library Loan/User Correlator'
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.StartPosition = 'CenterScreen'

    $lblLoans = New-Object System.Windows.Forms.Label -Property @{Text='Current Loans CSV:';Location='10,20';AutoSize=$true}
    $txtLoans = New-Object System.Windows.Forms.TextBox -Property @{Location='150,18';Width=320}
    $btnLoans = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,16';Width=75}
    $btnLoans.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtLoans.Text = $dlg.FileName }
    })

    $lblUsers = New-Object System.Windows.Forms.Label -Property @{Text='Library Users CSV:';Location='10,60';AutoSize=$true}
    $txtUsers = New-Object System.Windows.Forms.TextBox -Property @{Location='150,58';Width=320}
    $btnUsers = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,56';Width=75}
    $btnUsers.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtUsers.Text = $dlg.FileName }
    })

    $lblOut = New-Object System.Windows.Forms.Label -Property @{Text='Output File:';Location='10,100';AutoSize=$true}
    $txtOut = New-Object System.Windows.Forms.TextBox -Property @{Location='150,98';Width=320}
    $btnOut = New-Object System.Windows.Forms.Button -Property @{Text='Save As';Location='480,96';Width=75}
    $btnOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.FileName }
    })

    $btnRun = New-Object System.Windows.Forms.Button -Property @{Text='Run';Location='250,140';Width=100;Height=30}
    $btnRun.Add_Click({
        try {
            $lblStatus.Text = 'Status: Processing...'
            $form.Refresh()

            if (-not (Test-Path $txtLoans.Text)) { throw "Loans file not found: $($txtLoans.Text)" }
            if (-not (Test-Path $txtUsers.Text)) { throw "Users file not found: $($txtUsers.Text)" }

            $users = Import-Csv -Path $txtUsers.Text
            $userLookup = @{}
            $rawNameToUsername = @{}

            foreach ($u in $users) {
                $fullName = if ($u.Fullname) { $u.Fullname.Trim() } else { '' }
                $username = $u.Username

                if ($fullName -and $username) {
                    $rawNameToUsername[$fullName.ToLower()] = $username
                }

                if ($fullName -match '^([^,]+),\s*(.+)$') {
                    $surname = $matches[1].Trim()
                    $firstname = ($matches[2] -split '\s+')[0]
                    $formatted = "$firstname $surname"

                    if ($formatted -and $username) {
                        $userLookup[$formatted.ToLower()] = $username
                    }
                }
            }

            $csvLines = Get-Content -Path $txtLoans.Text
            if ($csvLines.Count -le 1) { throw "Loans CSV file appears to be empty or contains only headers" }

            $headers = $csvLines[0].Split(',') | Where-Object { $_.Trim('"') -ne '' } | ForEach-Object { $_.Trim('"') }
            $fixedHeaders = @(); $headerCount = @{}

            foreach ($header in $headers) {
                if (-not $headerCount.ContainsKey($header)) {
                    $headerCount[$header] = 0
                    $fixedHeaders += $header
                } else {
                    $headerCount[$header]++
                    $fixedHeaders += "$($header)_$($headerCount[$header])"
                }
            }

            $cleanedLoans = @()

            for ($i = 1; $i -lt $csvLines.Count; $i++) {
                $line = $csvLines[$i]
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                $values = $line -split ',(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)' | ForEach-Object { $_.Trim('"') }
                $rowData = @{}
                for ($j = 0; $j -lt $fixedHeaders.Count -and $j -lt $values.Count; $j++) {
                    $rowData[$fixedHeaders[$j]] = $values[$j]
                }

                $borrowedKey = $fixedHeaders | Where-Object { $_ -eq 'Borrowed' } | Select-Object -First 1
                $borrowedDateKey = $fixedHeaders | Where-Object { $_ -eq 'Borrowed' } | Select-Object -Skip 1 -First 1
                $borrowed = if ($borrowedKey) { $rowData[$borrowedKey] } else { '' }
                $borrowedDate = if ($borrowedDateKey) { $rowData[$borrowedDateKey] } else { '' }

                $barcode = ''; $borrow
