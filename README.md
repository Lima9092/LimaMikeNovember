# Library Migration Assistant

A comprehensive PowerShell toolkit for library system migrations, helping librarians and system administrators map fields between different library systems and transform data during migration.

## Overview

The Library Migration Assistant consists of two main tools:

1. **Library Migration Assistant** - A field mapping tool for planning data migrations
2. **Library Data Transformation Tool** - A data processing utility that applies mappings to actual data

These tools work together to simplify the complex process of migrating from one library system to another while ensuring data integrity and consistency.

## Features

### Library Migration Assistant
- Load source data and target system requirements
- Analyze field compatibility between systems
- Generate intelligent mapping suggestions
- Create and edit field mappings with transformations
- Export mapping configurations for later use

### Library Data Transformation Tool
- Apply field mappings to real data
- Transform data according to target system requirements
- Validate data against business rules
- Identify and highlight data errors
- Export transformed data ready for import into the target system

## Requirements

- Windows OS with PowerShell 5.1 or higher
- .NET Framework 4.5 or higher
- For Excel file support: PowerShell ImportExcel module (optional)

## Installation

1. Clone this repository
2. Ensure PowerShell execution policy allows script execution
3. Run the scripts directly from PowerShell

```powershell
# Clone the repository
git clone https://github.com/yourusername/library-migration-assistant.git

# Navigate to the directory
cd library-migration-assistant

# Run the migration assistant
.\LibraryMigrationAssistant.ps1

# Run the transformation tool
.\LibraryDataTransformationTool.ps1
```

## Usage Guide

### Library Migration Assistant

1. **Load Files Tab**
   - Browse for your source data CSV/Excel file containing actual library data
   - Browse for your target requirements CSV file describing the destination system fields
   - Click "Load Data" to analyze the files

2. **Analyze Tab**
   - Review source and target fields
   - View data samples to understand field contents
   - Click "Analyze and Generate Mapping Suggestions" to get intelligent mapping recommendations

3. **Map Fields Tab**
   - Review and edit the suggested mappings
   - Select source fields to map to target fields
   - Apply transformations where needed (e.g., format changes, data normalization)
   - Update mappings as needed

4. **Export Tab**
   - Preview the mapping configuration
   - Choose export format (CSV or PowerShell script)
   - Include validation rules and transformations as needed
   - Export the mapping file for use with the Data Transformation Tool

### Library Data Transformation Tool

1. **Load Files**
   - Load the mapping file created with the Migration Assistant
   - Load the source data CSV to be transformed

2. **Data Exploration**
   - View the data with error highlighting
   - Toggle between original and mapped fields
   - Show only records with errors or only mapped fields

3. **Data Transformation**
   - Data is transformed according to mapping rules
   - Error cells are highlighted in pink
   - Review the log for details on transformations and errors

4. **Export**
   - Export the transformed data ready for import into the new system
   - Export error logs for further analysis and cleanup

## Transformation Functions

The system uses a collection of transformation functions to modify data during migration. These functions are defined in the `Transform-Functions.ps1` file.

### Built-in Transformations

```powershell
# Gender detection based on title
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

# Borrower type detection based on ID pattern
function BorrowerTypeTransform($value) {
    if ($value -match "^[A-Za-z][0-9]{4}[A-Za-z]{2}$") {
        return "Student"
    }
    else {
        return "Teacher"
    }
}
```

### Creating Custom Transformations

You can create and add custom transformation functions directly within the Migration Assistant UI or by editing the `Transform-Functions.ps1` file.

Each function should:
1. Accept a single input value
2. Return the transformed value
3. Be registered in the `$global:TransformFunctions` hashtable

Example custom function for ISBN formatting:

```powershell
function FormatISBN($value) {
    # Remove all non-numeric characters
    $cleanISBN = $value -replace "[^0-9X]", ""
    
    # Check if valid ISBN-10 or ISBN-13
    if ($cleanISBN.Length -eq 10 -or $cleanISBN.Length -eq 13) {
        return $cleanISBN
    }
    else {
        return $value  # Return original if not valid
    }
}
$global:TransformFunctions["FormatISBN"] = ${function:FormatISBN}
```

## Target System Requirements File Format

The target requirements CSV file should contain the following columns:

- **FieldName**: The name of the field in the target system
- **FieldType**: Data type (string, numeric, date, etc.)
- **Mandatory**: Whether the field is required ("Yes" or "No")
- **ValidationRule**: Regular expression or rule for validation
- **Notes**: Additional information about the field

Example:

```csv
FieldName,FieldType,Mandatory,ValidationRule,Notes
PatronID,string,Yes,^[A-Z][0-9]{6}$,Must start with letter followed by 6 digits
FirstName,string,Yes,,
LastName,string,Yes,,
BorrowerType,string,Yes,^(Student|Teacher|Staff)$,Must be one of the allowed values
EmailAddress,string,No,^[\w.%+-]+@[\w.-]+\.[a-zA-Z]{2,}$,Must be valid email format
```

## Field Mapping File Format

The mapping file (created by the Migration Assistant) defines how fields are mapped and transformed:

- **SourceField**: Field name from the source system
- **NewField**: Field name in the target system
- **DataType**: Expected data type
- **Mandatory**: Whether the field is required
- **Validation**: Whether validation is enabled
- **ValidationRule**: Rule to check data against
- **Transformation**: Whether transformation is enabled
- **TransformFunction**: Name of the function to apply
- **Required**: Whether the field is required

## Troubleshooting

### Common Issues

1. **Missing Fields**: If target fields are marked as required but not found in source data, they will be highlighted in pink in the Data Transformation Tool.

2. **Transformation Errors**: If a transformation function fails, the error will be logged and the cell highlighted.

3. **Excel Files**: If working with Excel files, ensure the ImportExcel module is installed (`Install-Module ImportExcel`).

### Getting Help

For additional help or to report issues, please file a ticket on the GitHub repository.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
