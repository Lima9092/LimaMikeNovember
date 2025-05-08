Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-Form {
    [void][System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'CSV Correlation Tool'
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.StartPosition = 'CenterScreen'

    # Input CSV selector
    $lblIn = New-Object System.Windows.Forms.Label -Property @{Text='Input CSV:';Location = '10,20';AutoSize=$true}
    $txtIn = New-Object System.Windows.Forms.TextBox -Property @{Location='100,18';Width=360}
    $btnIn = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='470,16';Width=75}
    $btnIn.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtIn.Text = $dlg.FileName }
    })

    # Lookup CSV selector
    $lblLookup = New-Object System.Windows.Forms.Label -Property @{Text='Lookup CSV:';Location = '10,60';AutoSize=$true}
    $txtLookup = New-Object System.Windows.Forms.TextBox -Property @{Location='100,58';Width=360}
    $btnLookup = New-Object System.Windows.Forms.Button -Property @{Text='Browse';Location='470,56';Width=75}
    $btnLookup.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtLookup.Text = $dlg.FileName }
    })

    # Regex for key extraction
    $lblRegex = New-Object System.Windows.Forms.Label -Property @{Text='Key Regex:';Location='10,100';AutoSize=$true}
    $txtRegex = New-Object System.Windows.Forms.TextBox -Property @{Location='100,98';Width=360;Text='\((\d+)\)'}

    # Lookup key column in lookup CSV
    $lblKeyCol = New-Object System.Windows.Forms.Label -Property @{Text='Lookup Key Column:';Location='10,140';AutoSize=$true}
    $txtKeyCol = New-Object System.Windows.Forms.TextBox -Property @{Location='140,138';Width=150;Text='Barcode'}

    # Return column
    $lblRetCol = New-Object System.Windows.Forms.Label -Property @{Text='Return Column(s):';Location='10,180';AutoSize=$true}
    $txtRetCol = New-Object System.Windows.Forms.TextBox -Property @{Location='140,178';Width=150;Text='ISBN'}

    # Output file selector
    $lblOut = New-Object System.Windows.Forms.Label -Property @{Text='Output CSV:';Location='10,220';AutoSize=$true}
    $txtOut = New-Object System.Windows.Forms.TextBox -Property @{Location='100,218';Width=360}
    $btnOut = New-Object System.Windows.Forms.Button -Property @{Text='Save As';Location='470,216';Width=75}
    $btnOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'CSV Files (*.csv)|*.csv'
        if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.FileName }
    })

    # Run button
    $btnRun = New-Object System.Windows.Forms.Button -Property @{Text='Run';Location='250,260';Width=100;Height=30}
    $btnRun.Add_Click({ Process-Data })

    # Help button
    $btnHelp = New-Object System.Windows.Forms.Button -Property @{Text='Help';Location='360,260';Width=100;Height=30}
    $btnHelp.Add_Click({
        $helpText = @"
This tool helps you merge data from two CSV files in a few simple steps. It extracts a key value from each row of your main file using a regular expression, finds matching rows in another CSV based on that key, and appends the selected fields before exporting a combined CSV.

CSV Correlation Tool Help:

1. Input CSV: Select the primary CSV file containing the Title field.
2. Lookup CSV: Select the file with the lookup data.
3. Key Regex: .NET regex with one capture group to extract the key from the Title column.
4. Lookup Key Column: Column name in lookup CSV that matches extracted key.
5. Return Column(s): Comma-separated list of column names to append to the input.
6. Output CSV: Path for the resulting joined file.

Example Regex Patterns:
- Barcode in parentheses: \((\d+)\)
- Date YYYY-MM-DD: (\d{4}-\d{2}-\d{2})
- Email username: ^([^@]+)
- SKU code: SKU-(\w+)
- Invoice number: INV(\d{6})
"@
        [System.Windows.Forms.MessageBox]::Show($helpText, 'Help', 'OK', 'Information')
    })

    # Status label
    $lblStatus = New-Object System.Windows.Forms.Label -Property @{Text='Status: Ready';Location='10,310';AutoSize=$true}

    # Add controls to form
    $form.Controls.AddRange(@($lblIn,$txtIn,$btnIn,$lblLookup,$txtLookup,$btnLookup,
                              $lblRegex,$txtRegex,$lblKeyCol,$txtKeyCol,$lblRetCol,$txtRetCol,
                              $lblOut,$txtOut,$btnOut,$btnRun,$btnHelp,$lblStatus))

    # Data processing function
    function Process-Data {
        try {
            $lblStatus.Text = 'Status: Loading files...'
            $inFile = $txtIn.Text; $lookupFile = $txtLookup.Text; $outFile = $txtOut.Text
            if (-not (Test-Path $inFile) -or -not (Test-Path $lookupFile)) {
                throw 'Please select valid CSV files.'
            }
            $regex = [regex]$txtRegex.Text
            $keyCol = $txtKeyCol.Text; $retCols = $txtRetCol.Text -split ','
            $inData = Import-Csv $inFile
            $lookupData = Import-Csv $lookupFile | Group-Object -Property $keyCol -AsHashTable -AsString

            $lblStatus.Text = 'Status: Correlating...'
            foreach ($row in $inData) {
                $match = $regex.Match($row.Title)
                $key = if ($match.Success) { $match.Groups[1].Value } else { '' }
                if ($lookupData.ContainsKey($key)) {
                    $found = $lookupData[$key] | Select-Object -First 1
                    foreach ($col in $retCols) {
                        $row | Add-Member -NotePropertyName $col -NotePropertyValue $found.$col -Force
                    }
                } else {
                    foreach ($col in $retCols) { $row | Add-Member -NotePropertyName $col -NotePropertyValue '' -Force }
                }
            }
            $lblStatus.Text = 'Status: Saving output...'
            $inData | Export-Csv -Path $outFile -NoTypeInformation
            $lblStatus.Text = "Status: Done! Exported to $outFile"
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
            $lblStatus.Text = 'Status: Error'
        }
    }

    [void]$form.ShowDialog()
}

# Launch the form
Show-Form
