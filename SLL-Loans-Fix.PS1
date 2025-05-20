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

            $headers = $csvLines[0].Split(',') | ForEach-Object { $_.Trim('"') }
            $fixedHeaders = @(); $headerCount = @{}

            foreach ($header in $headers) {
                if ($header -ne '') {
                    if (-not $headerCount.ContainsKey($header)) {
                        $headerCount[$header] = 0
                        $fixedHeaders += $header
                    } else {
                        $headerCount[$header]++
                        $fixedHeaders += "$($header)_$($headerCount[$header])"
                    }
                }
            }

            # Log header information to status
            $lblStatus.Text = "Status: Processing... Found headers: $($fixedHeaders -join ', ')"
            $form.Refresh()

            # Simpler date conversion function
            function Convert-DateFormat {
                param(
                    [string]$Date
                )
                
                if ([string]::IsNullOrEmpty($Date)) {
                    return $Date
                }
                
                try {
                    # Try parsing with standard date formats
                    $parsedDate = [DateTime]::Parse($Date)
                    return $parsedDate.ToString("dd/MM/yyyy")
                }
                catch {
                    # If parsing fails, return the original string
                    return $Date
                }
            }

            $cleanedLoans = @()

            # Find key indices to ensure correct handling
            $borrowedTitleIndex = $fixedHeaders.IndexOf(($fixedHeaders | Where-Object { $_ -eq 'Borrowed' } | Select-Object -First 1))
            $borrowedDateIndex = $fixedHeaders.IndexOf(($fixedHeaders | Where-Object { $_ -eq 'Borrowed_1' } | Select-Object -First 1))
            $borrowedByIndex = $fixedHeaders.IndexOf(($fixedHeaders | Where-Object { $_ -like 'Borrowed By*' } | Select-Object -First 1))
            $returnDueDateIndex = $fixedHeaders.IndexOf(($fixedHeaders | Where-Object { $_ -like '*Date Return Due*' } | Select-Object -First 1))
            
            # Get the actual headers by index to ensure we're using the correct names
            $borrowedTitleKey = if ($borrowedTitleIndex -ge 0) { $fixedHeaders[$borrowedTitleIndex] } else { $null }
            $borrowedDateKey = if ($borrowedDateIndex -ge 0) { $fixedHeaders[$borrowedDateIndex] } else { $null }
            $borrowedByKey = if ($borrowedByIndex -ge 0) { $fixedHeaders[$borrowedByIndex] } else { $null }
            $returnDueDateKey = if ($returnDueDateIndex -ge 0) { $fixedHeaders[$returnDueDateIndex] } else { $null }
            
            # Update status with key information for debugging
            $lblStatus.Text = "Status: Processing... Using keys: Title=$borrowedTitleKey, Date=$borrowedDateKey, By=$borrowedByKey"
            $form.Refresh()

            for ($i = 1; $i -lt $csvLines.Count; $i++) {
                $line = $csvLines[$i]
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                $values = $line -split ',(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)' | ForEach-Object { $_.Trim('"') }
                $rowData = @{}
                for ($j = 0; $j -lt $fixedHeaders.Count -and $j -lt $values.Count; $j++) {
                    $rowData[$fixedHeaders[$j]] = $values[$j]
                }

                # Extract values
                $borrowed = if ($borrowedTitleKey -and $rowData.ContainsKey($borrowedTitleKey)) { $rowData[$borrowedTitleKey] } else { '' }
                $borrowedDate = if ($borrowedDateKey -and $rowData.ContainsKey($borrowedDateKey)) { $rowData[$borrowedDateKey] } else { '' }
                $originalBorrower = if ($borrowedByKey -and $rowData.ContainsKey($borrowedByKey)) { $rowData[$borrowedByKey] } else { '' }
                $returnDueDate = if ($returnDueDateKey -and $rowData.ContainsKey($returnDueDateKey)) { $rowData[$returnDueDateKey] } else { '' }

                # Convert dates
                $borrowedDate = Convert-DateFormat $borrowedDate
                $returnDueDate = Convert-DateFormat $returnDueDate

                # Parse item information
                $barcode = ''; $borrowedTitle = $borrowed
                if ($borrowed -match '^(.*?)\s*\((\d+)\)') {
                    $borrowedTitle = $matches[1].Trim()
                    $barcode = $matches[2]
                }

                # Process borrower information
                $borrowedByFormatted = $originalBorrower
                $userBarcode = ''

                if ($originalBorrower -match '^(\S+)\s+(.+)$') {
                    $firstname = $matches[1]
                    $surname = $matches[2]
                    $borrowedByFormatted = "$surname, $firstname"
                    $lookupKey = "$firstname $surname"

                    if ($userLookup.ContainsKey($lookupKey.ToLower())) {
                        $userBarcode = $userLookup[$lookupKey.ToLower()]
                    } elseif ($rawNameToUsername.ContainsKey($originalBorrower.ToLower())) {
                        $userBarcode = $rawNameToUsername[$originalBorrower.ToLower()]
                    }
                }

                # Create new loan record with explicitly named fields
                $newLoan = [ordered]@{
                    'Borrowed (Raw)' = $borrowed
                    'Borrowed Title' = $borrowedTitle
                    'Item Barcode' = $barcode
                    'Borrowed Date' = $borrowedDate
                    'Borrowed By' = $borrowedByFormatted
                    'Username' = $userBarcode
                }

                # Add any other fields (except the ones we've already processed)
                foreach ($k in $fixedHeaders) {
                    if ($k -eq $borrowedTitleKey -or $k -eq $borrowedDateKey -or $k -eq $borrowedByKey) {
                        continue # Skip fields we've already processed
                    }
                    
                    if ($rowData.ContainsKey($k)) {
                        $value = $rowData[$k]
                        # Convert any date fields
                        if ($k -like '*Date*') {
                            $value = Convert-DateFormat $value
                        }
                        $newLoan[$k] = $value
                    }
                }

                $cleanedLoans += [PSCustomObject]$newLoan
            }

            $cleanedLoans | Export-Csv -Path $txtOut.Text -NoTypeInformation
            $lblStatus.Text = "Status: Done! Processed $($cleanedLoans.Count) loans and saved to $($txtOut.Text)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            $lblStatus.Text = "Status: Error - $($_.Exception.Message)"
        }
    })

    $lblStatus = New-Object System.Windows.Forms.Label -Property @{Text='Status: Ready';Location='10,190';Width=560;Height=80;AutoSize=$false}

    $form.Controls.AddRange(@(
        $lblLoans, $txtLoans, $btnLoans,
        $lblUsers, $txtUsers, $btnUsers,
        $lblOut, $txtOut, $btnOut,
        $btnRun, $lblStatus
    ))

    [void]$form.ShowDialog()
}

Show-LoanUserMapper
