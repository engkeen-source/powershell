
# Set the path to the folder containing the xlsx files
$folderPath = "C:\Users\tanen\OneDrive\Script\PowerShell\CyberArk\CompareFolderPermission\WorkingDirectory\Testing\"

# Get all the xlsx files in the folder
$xlsxFiles = Get-ChildItem $folderPath -Filter *.xlsx

# Load the ImportExcel module
Import-Module ImportExcel

# Create an empty list to store the combined data
$data = New-Object System.Collections.Generic.List[System.Object]

# Loop through each xlsx file and append its data to the $data list
foreach ($file in $xlsxFiles) {
    # Get the hostname from the filename
    $hostname = $file.Name.Replace(".xlsx", "")

    # Load the xlsx file into a PowerShell object
    $xlsxData = Import-Excel $file.FullName

    # Add a "Hostname" column with the filename as the value
    $xlsxData | Add-Member -NotePropertyName "Hostname" -NotePropertyValue $hostname

    # Add a "Comparison Result" column with empty value
    $xlsxData | Add-Member -NotePropertyName "Comparison Result" -NotePropertyValue ""

    # Add a "Baseline Permission" column with empty value
    $xlsxData | Add-Member -NotePropertyName "Baseline Permission" -NotePropertyValue ""

    # Append the xlsx data to the $data list
    $data.AddRange($xlsxData)
}

# Rename the header
    # Get the current header names and rename the first three
    $headerNames = $data[0].psobject.Properties.Name
    $headerNames[0] = "FilePath"
    $headerNames[1] = "UserOrGroup"
    $headerNames[2] = "Permissions"

    # Create a new CSV with the updated header names
    $newData = $data | ForEach-Object {
        $row = $_
        $newRow = New-Object -TypeName psobject
        for ($i = 0; $i -lt $headerNames.Count; $i++) {
            $newRow | Add-Member -MemberType NoteProperty -Name $headerNames[$i] -Value $row.($row.psobject.Properties.Name[$i])
        }
        $newRow
    }

# Export the combined data to a new xlsx file
$outputFileName = $component + "_combined.csv"
$newData | Export-Csv -Path "$folderPath\$outputFileName" -NoTypeInformation
