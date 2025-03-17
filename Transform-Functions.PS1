# Ensure $global:TransformFunctions is a hashtable
if (-not $global:TransformFunctions) {
    $global:TransformFunctions = @{}
}

# Functions

function GenderTransform($value) {
    if ($value -match "Mr") {
        return "Male"
    }
    elseif ($value -match "Mrs" -or $value -match "Miss") {
        return "Female"
    }
    else {
        return ""
    }
}

$global:TransformFunctions["GenderTransform"]          = ${function:GenderTransform}

function BorrowerTypeTransform($value) {
    if ($value -match "^[A-Za-z][0-9]{4}[A-Za-z]{2}$") {
        return "Student"
    }
    else {
        return "Teacher"
    }
}

$global:TransformFunctions["BorrowerTypeTransform"]    = ${function:BorrowerTypeTransform}
function Custom-Text-Site-A($value) {
    return 'Site A'
}
$global:TransformFunctions["Custom-Text-Site-A"] = ${function:Custom-Text-Site-A}


function Custom-Text-1111($value) {
    return '1111'
}
$global:TransformFunctions["Custom-Text-1111"] = ${function:Custom-Text-1111}


