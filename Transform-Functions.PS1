# Ensure $global:TransformFunctions is a hashtable
if (-not $global:TransformFunctions) {
    $global:TransformFunctions = @{}
}

# Functions

function GenderTransform($value) {
    if ($value -match "\b(Mrs|Miss|Ms)\b") {
        return "Female"
    }
    elseif ($value -match "\bMr\b") {
        return "Male"
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




