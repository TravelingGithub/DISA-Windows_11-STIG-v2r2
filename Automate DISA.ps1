# Exit immediately without executing the rest of the script
#Write-Host "Script execution prevented"
#exit

#Requires -RunAsAdministrator test

<#
.SYNOPSIS
Reads registry entries from a CSV file and creates the specified keys and values.

.DESCRIPTION
This script reads data from a CSV file located in the same directory.
The CSV file must have the headers: ID, Registry_Path, Value_Name, Value_Type, Value_Data.
For each row in the CSV, the script:
- Creates the full registry key path if it doesn't exist, without deleting or overwriting existing keys.
- Adds or updates the specified registry value within the final key.
- Handles MultiString values correctly, expecting them to be semicolon-delimited in the CSV's Value_Data field.

.PARAMETER CsvFileName
The name of the CSV file to process. Defaults to 'registry_entries.csv'.

.EXAMPLE
.\CreateRegistryFromCsv.ps1
# Looks for 'registry_entries.csv' in the script's directory and processes it.

.EXAMPLE
.\CreateRegistryFromCsv.ps1 -CsvFileName "my_registry_settings.csv"
# Looks for 'my_registry_settings.csv' in the script's directory and processes it.

.NOTES
- Run this script with Administrator privileges as it modifies the registry (especially HKLM).
- Ensure the CSV file is correctly formatted with the required headers: ID,Registry_Path,Value_Name,Value_Type,Value_Data
- Supported Value_Type options in the CSV should match PowerShell's expectations:
  String, ExpandString, Binary, DWord, QWord, MultiString.
- For MultiString Value_Type, the Value_Data should be a single string with individual lines separated by semicolons (e.g., "String1;String2;Another String").

#>

param(
    [string]$CsvFileName = "registry_entries.csv"
)

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot

# Construct the full path to the CSV file
$csvFilePath = Join-Path -Path $scriptDirectory -ChildPath $CsvFileName

# Check if the CSV file exists
if (-not (Test-Path -Path $csvFilePath -PathType Leaf)) {
    Write-Error "CSV file not found at '$csvFilePath'. Please ensure the file exists in the same directory as the script and the name is correct."
    exit 1 # Exit the script if the CSV is not found
}

# Import data from the CSV file
try {
    $registryEntries = Import-Csv -Path $csvFilePath
    Write-Host "Successfully imported data from '$csvFilePath'."
}
catch {
    Write-Error "Failed to import CSV file '$csvFilePath'. Error: $($_.Exception.Message)"
    exit 1 # Exit if CSV import fails
}


# --- Loop through each entry defined in the CSV file ---
foreach ($entry in $registryEntries) {

    # Extract data for the current entry from the CSV row
    # Trim whitespace from values to avoid potential issues
    $registryPath = $entry.Registry_Path.Trim()
    $value = $entry.Value_Name.Trim()
    $valueData = $entry.Value_Data.Trim()
    $valueType = $entry.Value_Type.Trim()
    $entryId = $entry.ID.Trim() # Also trim the ID for cleaner logging

    # Basic validation for the current entry
    if ([string]::IsNullOrWhiteSpace($registryPath) -or [string]::IsNullOrWhiteSpace($value) -or [string]::IsNullOrWhiteSpace($valueType)) {
        Write-Warning "Skipping entry with ID '$entryId' due to missing required data (Registry_Path, Value_Name, or Value_Type)."
        continue # Skip to the next entry in the CSV
    }

    Write-Host "Processing Entry ID '$entryId': Path='$registryPath', Name='$value', Type='$valueType', Data='$valueData'"

    # --- Core logic to create path and set value ---

    # Split the path into individual components
    $pathParts = $registryPath -split '\\'

    # Check if the path starts with a valid registry hive prefix (e.g., HKLM:, HKCU:)
    if ($pathParts[0] -notmatch '^(HKLM|HKCU|HKCR|HKU|HKCC):$') {
         Write-Warning "Skipping entry ID '$entryId'. Invalid registry hive specified in path: '$registryPath'. Path must start with HKLM:, HKCU:, etc."
         continue # Skip to the next entry
    }

    # Build the path incrementally, creating keys as needed
    $currentPath = $pathParts[0]
    $pathCreationFailed = $false

    # Loop through each part of the path starting from the second element (the first subkey)
    for ($i = 1; $i -lt $pathParts.Count; $i++) {
        # Construct the next part of the path
        # Use Join-Path for robustness, especially if part names have tricky characters
        $currentPath = Join-Path -Path $currentPath -ChildPath $pathParts[$i]

        # Check if the key exists, create only if it doesn't
        if (-not (Test-Path -Path $currentPath)) {
            try {
                # Create the key
                New-Item -Path $currentPath -Force -ErrorAction Stop | Out-Null
                Write-Host "  Created registry key: $currentPath"
            }
            catch {
                 Write-Error "  Failed to create registry key: '$currentPath'. Error: $($_.Exception.Message). Check permissions."
                 $pathCreationFailed = $true # Mark that path creation failed
                 break # Stop trying to create further subkeys for this entry
            }
        } else {
            Write-Host "  Registry key already exists: $currentPath"
        }
    } # End for loop ($i)

    # Proceed only if the full path was successfully created or already existed
    if (-not $pathCreationFailed) {
        # Now handle the registry value at the final $registryPath

        # Variable to hold the data in the correct format for registry cmdlets
        $processedValueData = $valueData # Default to the original string data

        # --- Adjust ValueData type if necessary ---
        if ($valueType -eq 'MultiString') {
            # Split the semicolon-delimited string from CSV into a string array
            $processedValueData = $valueData -split ';'
            # Optional: Trim whitespace from each resulting string in the array
            $processedValueData = $processedValueData | ForEach-Object { $_.Trim() }
            Write-Host "  Note: Processed '$valueData' as MultiString array."
        }
        elseif ($valueType -eq 'Binary') {
             try {
                 # Example: Convert hex string (e.g., "01 AB FF") to byte array
                 $hexString = $valueData -replace '[^0-9A-Fa-f]' # Remove spaces, commas etc.
                 if (($hexString.Length % 2) -ne 0) { throw "Invalid hex string length."}
                 $bytes = [byte[]]::new($hexString.Length / 2)
                 for ($j = 0; $j -lt $bytes.Length; $j++) {
                     $bytes[$j] = [System.Convert]::ToByte($hexString.Substring($j * 2, 2), 16)
                 }
                 $processedValueData = $bytes
                 Write-Host "  Note: Processed '$valueData' as Binary data."
             } catch {
                 Write-Warning "  Skipping value set for '$value'. Invalid Binary data format in CSV for ID '$entryId'. Error: $($_.Exception.Message)"
                 continue # Skip setting value for this entry
             }
        }
        elseif ($valueType -eq 'DWord') {
           try {
               $processedValueData = [int32]$valueData # Use specific type cast
               Write-Host "  Note: Processed '$valueData' as DWord (Int32)."
           } catch {
               Write-Warning "  Skipping value set for '$value'. Invalid DWord data '$valueData' for entry ID '$entryId'. Error: $($_.Exception.Message)"
               continue # Skip setting value for this entry
           }
        }
         elseif ($valueType -eq 'QWord') {
           try {
               $processedValueData = [int64]$valueData # Use specific type cast
               Write-Host "  Note: Processed '$valueData' as QWord (Int64)."
           } catch {
               Write-Warning "  Skipping value set for '$value'. Invalid QWord data '$valueData' for entry ID '$entryId'. Error: $($_.Exception.Message)"
               continue # Skip setting value for this entry
           }
        }
        # String and ExpandString usually don't need special processing unless trimming was missed

        try {
            # Check if the value already exists
            $existingValue = Get-ItemProperty -Path $registryPath -Name $value -ErrorAction SilentlyContinue

            if ($existingValue -eq $null) {
                # Add the registry value to the final key if it doesn't exist
                New-ItemProperty -Path $registryPath -Name $value -Value $processedValueData -PropertyType $valueType -Force -ErrorAction Stop | Out-Null
                Write-Host "  Added registry value '$value' (Type: $valueType) to $registryPath"
            } else {
                # Check if the existing value needs updating
                $needsUpdate = $false
                $existingDataType = $existingValue.PSObject.Properties[$value].TypeNameOfValue
                $existingData = $existingValue.$value

                 # Check Type First
                 if ($existingDataType -ne $valueType) {
                    # Use mapping if PowerShell type names differ slightly from registry type names if needed
                    # Example: Might see Int32 instead of DWord sometimes depending on how value was read
                    # This simple check assumes exact match for now. Add more complex mapping if required.
                    Write-Host "  Note: Value '$value' exists but type differs (Existing: '$existingDataType', Required: '$valueType'). Updating."
                    $needsUpdate = $true
                 } else {
                     # Types match, now compare data
                     if ($valueType -eq 'MultiString' -or $valueType -eq 'Binary') {
                         # Compare arrays (order matters for MultiString, byte-by-byte for Binary)
                         if ($null -eq $existingData -or (Compare-Object -ReferenceObject $existingData -DifferenceObject $processedValueData -SyncWindow 0)) {
                             # Compare-Object returns differences if they exist, so if it returns anything, they are different.
                             # Also check if existing data is null when processed data is not.
                             $needsUpdate = $true
                         }
                     } elseif ($existingData -ne $processedValueData) {
                         # Simple comparison for other types (String, DWord, QWord, ExpandString)
                         $needsUpdate = $true
                     }
                 }


                if ($needsUpdate) {
                     # Overwrite the existing value with the new value/type
                    Set-ItemProperty -Path $registryPath -Name $value -Value $processedValueData -Type $valueType -Force -ErrorAction Stop | Out-Null
                    Write-Host "  Updated registry value '$value' with new data/type (Type: $valueType) at $registryPath"
                } else {
                    Write-Host "  Registry value '$value' already exists with the correct data/type at $registryPath. No changes needed."
                }
            }
        }
        catch {
            Write-Error "  Failed to set registry value '$value' at '$registryPath'. Error: $($_.Exception.Message). Check permissions or data/type compatibility."
        }
    } else {
         Write-Warning "Skipping value set for entry ID '$entryId' because path creation failed earlier."
    }

    Write-Host "Finished processing Entry ID '$entryId'."
    Write-Host "-----------------------------------------"

} # --- End foreach loop ($entry) ---

Write-Host "Script finished processing all entries in '$csvFilePath'."