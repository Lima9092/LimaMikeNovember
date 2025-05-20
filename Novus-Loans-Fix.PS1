Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-LoanUserMapper {
    [void][System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Library Loan/User Correlator'
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.StartPosition = 'CenterScreen'

    # Input: Current Loans
    $lblLoans = New-Object System.Windows.Forms.Label -Property @{Text='Current Loans CSV:';Location='10,20';AutoSize=$true}
    $txtLoans = New-Object System.Windows.Forms.TextBox -Property @{Location='150,18';Width=320}
    $btnLoans = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,16';Width=75}
    $btnLoans.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtLoans.Text = $dlg.FileName }
    })

    # Input: Library Users
    $lblUsers = New-Object System.Windows.Forms.Label -Property @{Text='Library Users CSV:';Location='10,60';AutoSize=$true}
    $txtUsers = New-Object System.Windows.Forms.TextBox -Property @{Location='150,58';Width=320}
    $btnUsers = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,56';Width=75}
    $btnUsers.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtUsers.Text = $dlg.FileName }
    })

    # Output: Save As
    $lblOut = New-Object System.Windows.Forms.Label -Property @{Text='Output File:';Location='10,100';AutoSize=$true}
    $txtOut = New-Object System.Windows.Forms.TextBox -Property @{Location='150,98';Width=320}
    $btnOut = New-Object System.Windows.Forms.Button -Property @{Text='Save As';Location='480,96';Width=75}
    $btnOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.FileName }
    })

    # Run Button
    $btnRun = New-Object System.Windows.Forms.Button -Property @{Text='Run';Location='250,140';Width=100;Height=30}
    $btnRun.Add_Click({
        try {
            $lblStatus.Text = 'Status: Processing...'
            $form.Refresh()
            
            if (-not (Test-Path $txtLoans.Text)) {
                throw "Loans file not found: $($txtLoans.Text)"
            }
            if (-not (Test-Path $txtUsers.Text)) {
                throw "Users file not found: $($txtUsers.Text)"
            }
            
            # Import user data
            $users = Import-Csv -Path $txtUsers.Text

            # Build strict and fallback user lookup dictionaries
            $userLookup = @{}
            $rawNameToCard = @{}

            foreach ($u in $users) {
                $fullName = $u.'Full Name'.Trim()
                $cardNumber = $u.'Card Number'

                # Raw fallback match
                $rawNameToCard[$fullName.ToLower()] = $cardNumber

                # Strict match reformat if comma present
                if ($fullName -match '^([^,]+),\s*(.+)$') {
                    $surname = $matches[1].Trim()
                    $firstname = ($matches[2] -split '\s+')[0]
                    $formatted = "$firstname $surname"
                    $userLookup[$formatted.ToLower()] = $cardNumber
                }
            }

            # Handle duplicate headers in loan CSV
            $csvLines = Get-Content -Path $txtLoans.Text
            if ($csvLines.Count -le 1) {
                throw "Loans CSV file appears to be empty or contains only headers"
            }

            $headers = $csvLines[0].Split(',') | ForEach-Object { $_.Trim('"') }

            $fixedHeaders = @()
            $headerCount = @{}

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

                $values = $line -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)' | ForEach-Object { $_.Trim('"') }

                $rowData = @{}
                for ($j = 0; $j -lt $fixedHeaders.Count -and $j -lt $values.Count; $j++) {
                    $rowData[$fixedHeaders[$j]] = $values[$j]
                }

                $borrowed = $rowData['borrowed']
                if ($null -eq $borrowed) {
                    $borrowedKey = $rowData.Keys | Where-Object { $_ -like 'Borrowed*' } | Select-Object -First 1
                    $borrowed = if ($borrowedKey) { $rowData[$borrowedKey] } else { "" }
                }

                $barcode = ''
                $borrowedTitle = $borrowed
                if ($borrowed -match '^(.*?)\s*\((\d+)\)') {
                    $borrowedTitle = $matches[1].Trim()
                    $barcode = $matches[2]
                }

                $borrowedByKey = $rowData.Keys | Where-Object { $_ -like 'Borrowed By*' } | Select-Object -First 1
                $originalBorrower = if ($borrowedByKey) { $rowData[$borrowedByKey] } else { "" }

                $borrowedByFormatted = $originalBorrower
                $userBarcode = ''

                if ($originalBorrower -match '^(\S+)\s+(.+)$') {
                    $firstname = $matches[1]
                    $surname = $matches[2]
                    $borrowedByFormatted = "$surname, $firstname"
                    $lookupKey = "$firstname $surname"

                    if ($userLookup.ContainsKey($lookupKey.ToLower())) {
                        $userBarcode = $userLookup[$lookupKey.ToLower()]
                    } elseif ($rawNameToCard.ContainsKey($originalBorrower.ToLower())) {
                        $userBarcode = $rawNameToCard[$originalBorrower.ToLower()]
                    }
                }

                $borrowedDateKey = $fixedHeaders | Where-Object { $_ -like 'Borrowed.1 Date*' } | Select-Object -First 1
                $returnDueKey = $fixedHeaders | Where-Object { $_ -like 'Return Due*' } | Select-Object -First 1
                $shelfmarkKey = $fixedHeaders | Where-Object { $_ -like 'Shelfmark*' } | Select-Object -First 1
                $renewalsKey = $fixedHeaders | Where-Object { $_ -like 'Repeat Renewals*' } | Select-Object -First 1
                $deptKey = $fixedHeaders | Where-Object { $_ -like 'Department*' } | Select-Object -First 1
                $resKey = $fixedHeaders | Where-Object { $_ -like 'Reservations*' } | Select-Object -First 1

                $newLoan = [PSCustomObject]@{
                    'Borrowed (Raw)'  = $borrowed
                    'Borrowed Title'  = $borrowedTitle
                    'Item Barcode'         = $barcode
                    'Borrowed By'     = $borrowedByFormatted
                    'User Barcode'    = $userBarcode
                    'Borrowed.1 Date' = if ($borrowedDateKey) { $rowData[$borrowedDateKey] } else { "" }
                    'Return Due'      = if ($returnDueKey) { $rowData[$returnDueKey] } else { "" }
                    'Shelfmark'       = if ($shelfmarkKey) { $rowData[$shelfmarkKey] } else { "" }
                    'Repeat Renewals' = if ($renewalsKey) { $rowData[$renewalsKey] } else { "" }
                    'Department'      = if ($deptKey) { $rowData[$deptKey] } else { "" }
                    'Reservations'    = if ($resKey) { $rowData[$resKey] } else { "" }
                }

                $cleanedLoans += $newLoan
            }

            $cleanedLoans | Export-Csv -Path $txtOut.Text -NoTypeInformation
            $lblStatus.Text = "Status: Done! Processed $($cleanedLoans.Count) loans and saved to $($txtOut.Text)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            $lblStatus.Text = "Status: Error - $($_.Exception.Message)"
        }
    })

    # Status
    $lblStatus = New-Object System.Windows.Forms.Label -Property @{Text='Status: Ready';Location='10,190';Width=560;Height=80;AutoSize=$false}

    # Add controls
    $form.Controls.AddRange(@(
        $lblLoans, $txtLoans, $btnLoans,
        $lblUsers, $txtUsers, $btnUsers,
        $lblOut, $txtOut, $btnOut,
        $btnRun, $lblStatus
    ))

    [void]$form.ShowDialog()
}

# Launch GUI
Show-LoanUserMapper
