param(
    [Parameter(Mandatory = $true)]
    [string]$sourceFile,
    [switch]$ChangeSlideMaster
)

# Extended Color Schemes
$colorSchemes = [ordered]@{
    "Green"  = [ordered]@{
        "244739" = "Green"
        "1B5744" = "Green"
        "247554" = "Green"
        "349E5F" = "Green"
        "37CC5C" = "Green"
        "4EE760" = "Green"
        "7BFC76" = "Green"
        "9FFF99" = "Green"
        "BFFFB7" = "Green"
        "E4FFD6" = "Green"
    }
    "Purple" = [ordered]@{
        "2A145A" = "Purple"
        "500A96" = "Purple"
        "612CB0" = "Purple"
        "743DD4" = "Purple"
        "8E5CEF" = "Purple"
        "9F78FC" = "Purple"
        "AA96FA" = "Purple"
        "BEAFFF" = "Purple"
        "C8C0FF" = "Purple"
        "DEDDFF" = "Purple"
    }
    "Blue"   = [ordered]@{
        "0D2155" = "Blue"
        "00227F" = "Blue"
        "0C32A4" = "Blue"
        "1D56C0" = "Blue"
        "0672CB" = "Blue"
        "58A5E6" = "Blue"
        "80C7FB" = "Blue"
        "9FDDFF" = "Blue"
        "CBEEFF" = "Blue"
        "E5F8FF" = "Blue"
    }
    "Red"    = [ordered]@{
        "4A193A" = "Red"
        "691D3F" = "Red"
        "85133F" = "Red"
        "B30B37" = "Red"
        "D2333D" = "Red"
        "E1633F" = "Red"
        "E17F3F" = "Red"
        "F4BB5E" = "Red"
        "F9D674" = "Red"
        "FBEECE" = "Red"
    }
}

# Function to get relevant color files based on file type
function Get-ColorFiles {
    param (
        [string]$dirPath,
        [string]$fileType,
        [switch]$ChangeSlideMaster
    )

    # Initialize an array to hold the files
    $files = @()

    # Determine the search pattern based on file type
    if ($fileType -match 'pptx|potx') {
        # Paths for PPT files
        $mediaPath = Join-Path -Path $dirPath -ChildPath "ppt\media"
        $slidesPath = Join-Path -Path $dirPath -ChildPath "ppt\slides"
        $chartsPath = Join-Path -Path $dirPath -ChildPath "ppt\charts"
        # Include slide layouts if the switch is set
        if ($ChangeSlideMaster) {
            $layoutsPath = Join-Path -Path $dirPath -ChildPath "ppt\slideLayouts"
        }

        # Get the files
        $files += Get-ChildItem -Path $mediaPath -Recurse -Include *.svg -ErrorAction SilentlyContinue
        $files += Get-ChildItem -Path $slidesPath -Recurse -Include *.xml -ErrorAction SilentlyContinue
        $files += Get-ChildItem -Path $chartsPath -Recurse -Include *.xml -ErrorAction SilentlyContinue
        if ($ChangeSlideMaster) {
            $files += Get-ChildItem -Path $layoutsPath -Recurse -Include *.xml -ErrorAction SilentlyContinue
        }

    } elseif ($fileType -match 'docx|dotx') {
        # Paths for DOC files
        $mediaPath = Join-Path -Path $dirPath -ChildPath "word\media"
        $docPath = Join-Path -Path $dirPath -ChildPath "word"

        # Get the files
        $files += Get-ChildItem -Path $mediaPath -Recurse -Include *.svg -ErrorAction SilentlyContinue
        $files += Get-ChildItem -Path $docPath -File -Filter *.xml -ErrorAction SilentlyContinue
    } else {
        # For other file types
        $files += Get-ChildItem -Path $dirPath -Recurse -Include *.xml, *.svg -ErrorAction SilentlyContinue
    }

    # Exclude theme1.xml etc to avoid overwriting themes colors
    $files = $files | Where-Object { $_.Name -notmatch '^theme\d+\.xml$' }

    return $files
}

# Function to detect the current color scheme from the file contents
function Detect-ColorScheme {
    param (
        [string]$dirPath,
        [string]$fileType
    )

    $files = Get-ColorFiles -dirPath $dirPath -fileType $fileType -ChangeSlideMaster:$false

    foreach ($file in $files) {
        $content = Get-Content -LiteralPath $file.FullName -Raw

        foreach ($scheme in $colorSchemes.Keys) {
            foreach ($color in $colorSchemes[$scheme].Keys) {
                if ($content -like "*$color*") {
                    return $scheme
                }
            }
        }
    }

    return $null
}

# Function to replace colors based on target scheme
function Replace-Colors {
    param (
        [string]$dirPath,
        [string]$sourceScheme,
        [string]$targetScheme,
        [string]$fileType,
        [switch]$ChangeSlideMaster
    )

    $sourceColors = @($colorSchemes[$sourceScheme].Keys)
    $targetColors = @($colorSchemes[$targetScheme].Keys)

    $files = Get-ColorFiles -dirPath $dirPath -fileType $fileType -ChangeSlideMaster:$ChangeSlideMaster

    foreach ($file in $files) {
        $content = Get-Content -LiteralPath $file.FullName

        # Replace each color code from the source scheme to the target scheme
        for ($i = 0; $i -lt $sourceColors.Count; $i++) {
            $sourceColor = $sourceColors[$i]
            $targetColor = $targetColors[$i]

            # Check if sourceColor and targetColor are not null or empty
            if (![string]::IsNullOrEmpty($sourceColor) -and ![string]::IsNullOrEmpty($targetColor)) {
                $content = $content -replace [Regex]::Escape($sourceColor), $targetColor
            } else {
                Write-Host "Skipping replacement for invalid source or target color." -ForegroundColor Yellow
            }
        }

        # Write the content back to the file
        Set-Content -LiteralPath $file.FullName -Value $content
    }
}


# Function to handle the unzipping, color replacement, and zipping process
function Process-Template {
    param (
        [string]$sourceFile,
        [string]$outputFile,
        [string]$sourceScheme,
        [string]$targetScheme,
        [string]$fileType
    )

    # Define temp directory for extraction
    $tempDir = Join-Path -Path $env:TEMP -ChildPath (New-Guid)

    # Ensure the temp directory exists
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    if ($fileType -eq ".svg") {
        # If the input file is an SVG, just copy it to the temp directory
        $svgFile = Join-Path -Path $tempDir -ChildPath "template.svg"
        Copy-Item -Path $sourceFile -Destination $svgFile

        # Replace colors directly in the SVG
        Replace-Colors -dirPath $tempDir -sourceScheme $sourceScheme -targetScheme $targetScheme -fileType $fileType

        # Copy the result to the output file location
        Copy-Item -Path $svgFile -Destination $outputFile
    } else {
        # If not an SVG, handle as a ZIP archive (e.g., .docx, .dotx, .pptx, .potx)
        $zipFile = Join-Path -Path $tempDir -ChildPath "template.zip"
        
        try {
            Copy-Item -Path $sourceFile -Destination $zipFile
            Expand-Archive -Path $zipFile -DestinationPath $tempDir

            # Remove the template.zip after extraction
            Remove-Item -Path $zipFile -Force
        } catch {
            Write-Host "Failed to copy or extract the template file." -ForegroundColor Red
            exit
        }

        # Replace colors
        Replace-Colors -dirPath $tempDir -sourceScheme $sourceScheme -targetScheme $targetScheme -fileType $sourceExtension

        # Compress the files again
        $tempZip = Join-Path -Path $tempDir -ChildPath "result.zip"
        
        try {
            Compress-Archive -Path (Join-Path $tempDir '*') -DestinationPath $tempZip
            Copy-Item -Path $tempZip -Destination $outputFile
        } catch {
            Write-Host "Failed to compress or copy the result zip." -ForegroundColor Red
            exit
        }
    }

    # Clean up temp directory
    Remove-Item -Path $tempDir -Recurse -Force
}

# Main script

# Ensure the source file exists
if (-not (Test-Path $sourceFile)) {
    Write-Host "The source file $sourceFile does not exist or was not provided." -ForegroundColor Red
    exit
}

# Get the directory, file name, and extension from the source file
$sourceDir = Split-Path -Parent $sourceFile
$sourceName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile)
$sourceExtension = [System.IO.Path]::GetExtension($sourceFile)

# Define supported extensions
$supportedExtensions = @(".pptx", ".potx", ".docx", ".dotx", ".svg")

# Validate the file extension
if ($sourceExtension -notin $supportedExtensions) {
    Write-Host "Unsupported file type: $sourceExtension. Supported types are: $supportedExtensions" -ForegroundColor Red
    exit
}

# Proceed with the script
Write-Host "Processing file: $sourceFile" -ForegroundColor Green

$version = $null
$detectedScheme = $null
$date = $null

# Regex pattern to find version info (- vX.Y) without including extra parts
$versionRegex = ' - v(\d+\.\d+)'

# Create a case-insensitive regex pattern from the color scheme keys
$colorSchemeRegex = "(?i) - ($($colorSchemes.Keys -join '|'))"

# Create a regex pattern for date
$dateRegex = '(\d{4}[.-]\d{2}[.-]\d{2})'  # Date pattern (yyyy-mm-dd or yyyy.mm.dd)


# Extract Version
if ($sourceName -match $versionRegex) {
    $version = $matches[1]
    $sourceName = $sourceName -replace $versionRegex, ""  # Remove version from source name
}

# Extract Color Scheme
if ($sourceName -match $colorSchemeRegex) {
    $detectedScheme = $matches[1]
    $sourceName = $sourceName -replace $colorSchemeRegex, ""  # Remove color scheme from source name
}

# Extract Date
if ($sourceName -match $dateRegex) {
    $date = $matches[1] -replace '[-]', '.'  # Replace hyphens with dots for output
    $sourceName = $sourceName -replace $dateRegex, ""  # Remove date from source name
}

# The remaining source name is the Document Name
$documentName = $sourceName.Trim()  # Trim leading/trailing whitespace

# Remove trailing hyphen if it exists
if ($documentName.EndsWith('-')) {
    $documentName = $documentName.TrimEnd('-').Trim()  # Trim hyphen and whitespace
}



# Remove any detected color scheme and clean up the name
if ($detectedScheme) {
    Write-Host "Detected color scheme from filename: $detectedScheme"
} else {
    Write-Host "No color scheme found in filename. Detecting from file content..."
    $tempDir = Join-Path -Path $env:TEMP -ChildPath (New-Guid)
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    if ($sourceExtension -eq ".svg") {
        # Handle SVG files differently (no zipping/unzipping needed)
        Copy-Item -Path $sourceFile -Destination $tempDir
    } else {
        # Detect color scheme from non-SVG file content (requires unzipping)
        $zipFile = Join-Path -Path $tempDir -ChildPath "template.zip"
        Copy-Item -Path $sourceFile -Destination $zipFile
        Expand-Archive -Path $zipFile -DestinationPath $tempDir
    }
    $detectedScheme = Detect-ColorScheme -dirPath $tempDir -fileType $sourceExtension
    Remove-Item -Path $tempDir -Recurse -Force

    if (-not $detectedScheme) {
        Write-Host "No matching color scheme found in the template." -ForegroundColor Red
        exit
    }

    Write-Host "Detected color scheme from content: $detectedScheme"
}

# Process for the remaining color schemes
$remainingSchemes = $colorSchemes.Keys | Where-Object { $_ -ne $detectedScheme }

foreach ($scheme in $remainingSchemes) {
    # Construct the output filename
    $outputFileBase = $documentName

    if ($version) {
        $outputFileBase += " - v$version"
    }
    if ($date) {
        $outputFileBase += " - $date"
    }
    # Always append the new color scheme
    $outputFileBase += " - $Scheme"

    $outputFile = Join-Path -Path $sourceDir -ChildPath "$outputFileBase$sourceExtension"
    
    # Output for debugging
    Write-Host "Processing: $outputFile"

    # Call the template processing function
    Process-Template -sourceFile $sourceFile -outputFile $outputFile -sourceScheme $detectedScheme -targetScheme $scheme -fileType $sourceExtension
}

Write-Host "Process completed. The new files are saved in the same directory as the source file."