# Ensure $global:TransformFunctions is a hashtable
if (-not $global:TransformFunctions) {
    $global:TransformFunctions = @{}
}

# Functions

function Date-DDMMMYYYYToDDMMYYYY {
    param($value)
    if (-not $value) { return '' }
    
    # Define month mapping
    $monthMap = @{
        'JAN' = '01'; 'FEB' = '02'; 'MAR' = '03'; 'APR' = '04'
        'MAY' = '05'; 'JUN' = '06'; 'JUL' = '07'; 'AUG' = '08'
        'SEP' = '09'; 'OCT' = '10'; 'NOV' = '11'; 'DEC' = '12'
    }
    
    # Trim and process input
    $dateString = $value.Trim()
    
    # Try to parse the date format DD MMM YYYY
    if ($dateString -match '(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})') {
        $day = $matches[1].PadLeft(2, '0')
        $month = $monthMap[$matches[2].ToUpper()]
        $year = $matches[3]
        
        # Return DD/MM/YYYY format
        return "$day/$month/$year"
    }
    else {
        # Return original value if format doesn't match
        return $dateString
    }
}
$global:TransformFunctions['Date-DDMMMYYYYToDDMMYYYY'] = ${function:Date-DDMMMYYYYToDDMMYYYY}

function FName-FirstTransform {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    if ($n -match ',') {
        # "First, Last" or "First, Last-Last"
        return ($n -split '\s*,\s*')[0]
    }
    else {
        # "First Last" or "First Last-Last"
        return ($n -split '\s+')[0]
    }
}
$global:TransformFunctions['FName-FirstTransform'] = ${function:FName-FirstTransform}

function SName-LastTransform {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    if ($n -match ',') {
        # "First, Last" or "First, Last-Last"
        return ($n -split '\s*,\s*')[1]
    }
    else {
        # "First Last" or "First Last-Last"
        return ($n -split '\s+')[-1]
    }
}
$global:TransformFunctions['SName-LastTransform'] = ${function:SName-LastTransform}

function FName-LastTransform {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    if ($n -match ',') {
        # "Last, First" or "Last-Last, First"
        return ($n -split '\s*,\s*')[1]
    }
    else {
        # "Last First" or "Last-Last First"
        return ($n -split '\s+')[-1]
    }
}
$global:TransformFunctions['FName-LastTransform'] = ${function:FName-LastTransform}

function SName-FirstTransform {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    if ($n -match ',') {
        # "Last, First" or "Last-Last, First"
        return ($n -split '\s*,\s*')[0]
    }
    else {
        # "Last First" or "Last-Last First"
        return ($n -split '\s+')[0]
    }
}
$global:TransformFunctions['SName-FirstTransform'] = ${function:SName-FirstTransform}

function AdultFictionTransform($value) {
    # Check for ISBN-10 pattern and validate
    # Remove any hyphens or spaces for processing
    $cleanValue = $value -replace '[-\s]', ''
    
    # ISBN-10 regex: exactly 10 characters, digits 0-9 and possibly X as last character
    if ($cleanValue -match '^(\d{9}[\dX])$') {
        $isbn10 = $cleanValue.ToUpper()
        
        # Validate ISBN-10 using check digit algorithm inline
        $sum = 0
        for ($i = 0; $i -lt 9; $i++) {
            $digit = [int]$isbn10.Substring($i, 1)
            $sum += $digit * ($i + 1)
        }
        
        $remainder = $sum % 11
        $checkDigit = $isbn10.Substring(9, 1)
        
        # Check digit is X if remainder is 10, otherwise it is the remainder itself
        $isValid = if ($remainder -eq 10) {
            $checkDigit -eq 'X'
        } else {
            $checkDigit -eq $remainder.ToString()
        }
        
        if ($isValid) {
            return 'AF'
        }
    }
    
    # If not a valid ISBN-10 return blank
    return ''
}
$global:TransformFunctions['AdultFictionTransform'] = ${function:AdultFictionTransform}

function RemoveSlashSuffixTransform {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    if ($n -match '/') {
        # Remove forward slash and everything after it
        return ($n -split '/')[0].Trim()
    }
    else {
        # No slash found, return original trimmed value
        return $n
    }
}
$global:TransformFunctions['RemoveSlashSuffixTransform'] = ${function:RemoveSlashSuffixTransform}

function ExcludeNonPRN {
    param($value)
    if (-not $value) { return '' }
    $n = $value.Trim()
    # Only allow values matching ANNNNAA where A = alphanumeric, N = digit
    if ($n -match '^[A-Za-z0-9][0-9]{4}[A-Za-z0-9]{2}$') {
        return $n
    }
    else {
        return ''
    }
}

$global:TransformFunctions['ExcludeNonPRN'] = ${function:ExcludeNonPRN}


