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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuration variables that can be modified as needed
$USERS_USERNAME_COLUMN = "Barcode" # Default username column
$USERS_FALLBACK_USERNAME_COLUMN = "Username" # Fallback column if the selected column is empty
$USERS_FULLNAME_COLUMN = "Fullname" # Default fullname column

function Show-LoanUserMapper {
    [void][System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Library Loan/User Correlator'
    $form.Size = New-Object System.Drawing.Size(600,500)
    $form.StartPosition = 'CenterScreen'

    # Add configuration input fields
    $lblUsernameCol = New-Object System.Windows.Forms.Label -Property @{Text='Primary Username Column:';Location='10,20';AutoSize=$true}
    $txtUsernameCol = New-Object System.Windows.Forms.TextBox -Property @{Location='170,18';Width=300;Text=$USERS_USERNAME_COLUMN}

    $lblFallbackCol = New-Object System.Windows.Forms.Label -Property @{Text='Fallback Username Column:';Location='10,60';AutoSize=$true}
    $txtFallbackCol = New-Object System.Windows.Forms.TextBox -Property @{Location='170,58';Width=300;Text=$USERS_FALLBACK_USERNAME_COLUMN}
    
    $chkUseFallback = New-Object System.Windows.Forms.CheckBox -Property @{
        Text='Use fallback if primary is empty';
        Location='170,88';
        Width=300;
        Checked=$true
    }

    $lblFullnameCol = New-Object System.Windows.Forms.Label -Property @{Text='Fullname Column:';Location='10,120';AutoSize=$true}
    $txtFullnameCol = New-Object System.Windows.Forms.TextBox -Property @{Location='170,118';Width=300;Text=$USERS_FULLNAME_COLUMN}

    $lblLoans = New-Object System.Windows.Forms.Label -Property @{Text='Current Loans CSV:';Location='10,160';AutoSize=$true}
    $txtLoans = New-Object System.Windows.Forms.TextBox -Property @{Location='170,158';Width=300}
    $btnLoans = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,156';Width=75}
    $btnLoans.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtLoans.Text = $dlg.FileName }
    })

    $lblUsers = New-Object System.Windows.Forms.Label -Property @{Text='Library Users CSV:';Location='10,200';AutoSize=$true}
    $txtUsers = New-Object System.Windows.Forms.TextBox -Property @{Location='170,198';Width=300}
    $btnUsers = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='480,196';Width=75}
    $btnUsers.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtUsers.Text = $dlg.FileName }
    })

    $lblOut = New-Object System.Windows.Forms.Label -Property @{Text='Output File:';Location='10,240';AutoSize=$true}
    $txtOut = New-Object System.Windows.Forms.TextBox -Property @{Location='170,238';Width=300}
    $btnOut = New-Object System.Windows.Forms.Button -Property @{Text='Save As';Location='480,236';Width=75}
    $btnOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.FileName }
    })

    $btnRun = New-Object System.Windows.Forms.Button -Property @{Text='Run';Location='250,280';Width=100;Height=30}
    $btnRun.Add_Click({
        try {
            $lblStatus.Text = 'Status: Processing...'
            $form.Refresh()

            # Get column names from the form
            $primaryUsernameColumn = $txtUsernameCol.Text.Trim()
            $fallbackUsernameColumn = $txtFallbackCol.Text.Trim()
            $useFallback = $chkUseFallback.Checked
            $fullnameColumn = $txtFullnameCol.Text.Trim()
            
            if ([string]::IsNullOrWhiteSpace($primaryUsernameColumn)) {
                throw "Primary username column name cannot be empty"
            }
            
            if ($useFallback -and [string]::IsNullOrWhiteSpace($fallbackUsernameColumn)) {
                throw "Fallback username column name cannot be empty when fallback is enabled"
            }
            
            if ([string]::IsNullOrWhiteSpace($fullnameColumn)) {
                throw "Fullname column name cannot be empty"
            }

            if (-not (Test-Path $txtLoans.Text)) { throw "Loans file not found: $($txtLoans.Text)" }
            if (-not (Test-Path $txtUsers.Text)) { throw "Users file not found: $($txtUsers.Text)" }

            $users = Import-Csv -Path $txtUsers.Text
            $userLookup = @{}
            $rawNameToUsername = @{}

            # Verify that the specified columns exist in the users CSV
            $firstUser = $users | Select-Object -First 1
            $userColumns = $firstUser.PSObject.Properties.Name
            
            if (-not ($userColumns -contains $primaryUsernameColumn)) {
                throw "Column '$primaryUsernameColumn' not found in users CSV. Available columns: $($userColumns -join ', ')"
            }
            
            if ($useFallback -and -not ($userColumns -contains $fallbackUsernameColumn)) {
                throw "Column '$fallbackUsernameColumn' not found in users CSV. Available columns: $($userColumns -join ', ')"
            }
            
            if (-not ($userColumns -contains $fullnameColumn)) {
                throw "Column '$fullnameColumn' not found in users CSV. Available columns: $($userColumns -join ', ')"
            }

            foreach ($u in $users) {
                $fullName = if ($u.$fullnameColumn) { $u.$fullnameColumn.Trim() } else { '' }
                
                # Get username with fallback logic
                $username = $u.$primaryUsernameColumn
                if ($useFallback -and [string]::IsNullOrWhiteSpace($username) -and $primaryUsernameColumn -ne $fallbackUsernameColumn) {
                    $username = $u.$fallbackUsernameColumn
                }

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

            function Convert-DateFormat {
                param(
                    [string]$Date
                )

                if ([string]::IsNullOrWhiteSpace($Date)) {
                    return $Date
                }

                try {
                    # Parse into a DateTime
                    $parsed = [DateTime]::Parse($Date)

                    # Detect if the original had a time component (simple colon check)
                    $hasTime = $Date -match '\d{1,2}:\d{2}'

                    if (-not $hasTime) {
                        # No time in input → set to 17:00:00
                        $parsed = $parsed.Date.AddHours(17)
                    }
                    else {
                        # Time was present; ensure seconds are populated (ToString round-trip)
                        $parsed = [DateTime]::ParseExact(
                            $parsed.ToString("yyyy-MM-dd HH:mm:ss"),
                            "yyyy-MM-dd HH:mm:ss",
                            $null
                        )
                    }

                    # Output in desired format
                    return $parsed.ToString("dd/MM/yyyy HH:mm:ss")
                }
                catch {
                    # If parsing fails, just return the original
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

                # Parse item information - UPDATED VERSION
				$barcode = ''; $borrowedTitle = $borrowed
				if ($borrowed -match '^(.*?)\s*\(([^)]+)\)') {
					$borrowedTitle = $matches[1].Trim()
					$barcodeRaw = $matches[2]
					
					# Extract only the part before the first space (if any)
					if ($barcodeRaw -match '^(\S+)') {
						$barcode = $matches[1]
					} else {
						$barcode = $barcodeRaw
					}
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
            
            # Add information about fallback usage
            if ($useFallback) {
                $fallbackCount = ($users | Where-Object { 
                    [string]::IsNullOrWhiteSpace($_.$primaryUsernameColumn) -and 
                    -not [string]::IsNullOrWhiteSpace($_.$fallbackUsernameColumn)
                }).Count
                
                if ($fallbackCount -gt 0) {
                    $lblStatus.Text += "`nFallback column '$fallbackUsernameColumn' was used for $fallbackCount users."
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            $lblStatus.Text = "Status: Error - $($_.Exception.Message)"
        }
    })

    $lblStatus = New-Object System.Windows.Forms.Label -Property @{Text='Status: Ready';Location='10,330';Width=560;Height=100;AutoSize=$false}

    $form.Controls.AddRange(@(
        $lblUsernameCol, $txtUsernameCol,
        $lblFallbackCol, $txtFallbackCol, $chkUseFallback,
        $lblFullnameCol, $txtFullnameCol,
        $lblLoans, $txtLoans, $btnLoans,
        $lblUsers, $txtUsers, $btnUsers,
        $lblOut, $txtOut, $btnOut,
        $btnRun, $lblStatus
    ))

    [void]$form.ShowDialog()
}

Show-LoanUserMapper
