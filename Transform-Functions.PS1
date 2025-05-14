# Ensure $global:TransformFunctions is a hashtable
if (-not $global:TransformFunctions) {
    $global:TransformFunctions = @{}
}

# Functions

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
    # If the incoming value starts with 489 or 499, return “Adult-fiction”
    if ($value -match '^(489|499)') {
        return 'Adult-fiction'
    }
    # otherwise return a blank string
    return ''
}

# Register it in your in‐memory lookup
$global:TransformFunctions['AdultFictionTransform'] = ${function:AdultFictionTransform}
