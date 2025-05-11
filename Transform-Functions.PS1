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


$global:TransformFunctions["BorrowerTypeTransform"]    = ${function:BorrowerTypeTransform}


function Adult-fictionTransform($value) {
    return 'Adult fiction'
}
$global:TransformFunctions["Adult-fictionTransform"] = ${function:Adult-fictionTransform}


function parentFlag0Transform($value) {
    return '0'
}
$global:TransformFunctions["parentFlag0Transform"] = ${function:parentFlag0Transform}
