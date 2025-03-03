# PowerShell Color Replacement Script

This PowerShell script automates the process of color replacement in template files. It supports various file types, including Microsoft Office formats (`.docx`, `.dotx`, `.pptx`, `.potx`) and `.SVG` files. The script detects the current color scheme from the filename or the content of the file and generates new files with specified color replacements. It also includes the ability to modify slide masters in PowerPoint templates for advanced customization.

## Features

- **Supported File Types**:
  - Microsoft Word: `.docx`, `.dotx`
  - Microsoft PowerPoint: `.pptx`, `.potx`
  - SVG: `.svg`

- **Color Scheme Detection**: Automatically detects the current color scheme from the filename or file content.

- **Color Replacement**: Replaces color schemes based on predefined mappings.

- **Slide Master Customization**: Optional parameter to update the slide master in PowerPoint files.

- **Output Generation**: Produces new files named according to the original document name, version, date, and new color scheme.

## Prerequisites

- Windows operating system with PowerShell 5.1 or later.
- The `Microsoft .NET Framework` for PowerShell.
- Access to the file system for reading and writing files.

## Usage

1. **Clone the Repository** or download the script file.

    ```bash
    git clone <repository-url>
    ```

2. **Run the Script**: Execute the script in PowerShell with the required parameter.

    ```powershell
    .\ColorReplacementScript.ps1 -sourceFile "path\to\your\template.file" [-ChangeSlideMaster]
    ```

### Parameters

- `-sourceFile`: (Mandatory) The path to the source file (template) you want to process.
- `-ChangeSlideMaster`: (Optional) Use this parameter to enable color replacement in the slide master for PowerPoint files (`.pptx`, `.potx`). This ensures consistency in themes and master slide layouts (USE WITH CAUTION!).

## Color Schemes
The script includes the following extended color schemes:

- **Green**
- **Purple**
- **Blue**
- **Red**

### Color Replacement Table

| Green Scheme                                                      | Purple Scheme                                                     | Blue Scheme                                                       | Red Scheme                                                        |
| ----------------------------------------------------------------- | ----------------------------------------------------------------- | ----------------------------------------------------------------- | ----------------------------------------------------------------- |
| ![#244739](https://dummyimage.com/20/244739/000000.png&text=+) `#244739` | ![#2A145A](https://dummyimage.com/20/2A145A/000000.png&text=+) `#2A145A` | ![#0D2155](https://dummyimage.com/20/0D2155/000000.png&text=+) `#0D2155` | ![#4A193A](https://dummyimage.com/20/4A193A/000000.png&text=+) `#4A193A` |
| ![#1B5744](https://dummyimage.com/20/1B5744/000000.png&text=+) `#1B5744` | ![#500A96](https://dummyimage.com/20/500A96/000000.png&text=+) `#500A96` | ![#00227F](https://dummyimage.com/20/00227F/000000.png&text=+) `#00227F` | ![#691D3F](https://dummyimage.com/20/691D3F/000000.png&text=+) `#691D3F` |
| ![#247554](https://dummyimage.com/20/247554/000000.png&text=+) `#247554` | ![#612CB0](https://dummyimage.com/20/612CB0/000000.png&text=+) `#612CB0` | ![#0C32A4](https://dummyimage.com/20/0C32A4/000000.png&text=+) `#0C32A4` | ![#85133F](https://dummyimage.com/20/85133F/000000.png&text=+) `#85133F` |
| ![#349E5F](https://dummyimage.com/20/349E5F/000000.png&text=+) `#349E5F` | ![#743DD4](https://dummyimage.com/20/743DD4/000000.png&text=+) `#743DD4` | ![#1D56C0](https://dummyimage.com/20/1D56C0/000000.png&text=+) `#1D56C0` | ![#B30B37](https://dummyimage.com/20/B30B37/000000.png&text=+) `#B30B37` |
| ![#37CC5C](https://dummyimage.com/20/37CC5C/000000.png&text=+) `#37CC5C` | ![#8E5CEF](https://dummyimage.com/20/8E5CEF/000000.png&text=+) `#8E5CEF` | ![#0672CB](https://dummyimage.com/20/0672CB/000000.png&text=+) `#0672CB` | ![#D2333D](https://dummyimage.com/20/D2333D/000000.png&text=+) `#D2333D` |
| ![#4EE760](https://dummyimage.com/20/4EE760/000000.png&text=+) `#4EE760` | ![#9F78FC](https://dummyimage.com/20/9F78FC/000000.png&text=+) `#9F78FC` | ![#58A5E6](https://dummyimage.com/20/58A5E6/000000.png&text=+) `#58A5E6` | ![#E1633F](https://dummyimage.com/20/E1633F/000000.png&text=+) `#E1633F` |
| ![#7BFC76](https://dummyimage.com/20/7BFC76/000000.png&text=+) `#7BFC76` | ![#AA96FA](https://dummyimage.com/20/AA96FA/000000.png&text=+) `#AA96FA` | ![#80C7FB](https://dummyimage.com/20/80C7FB/000000.png&text=+) `#80C7FB` | ![#E17F3F](https://dummyimage.com/20/E17F3F/000000.png&text=+) `#E17F3F` |
| ![#9FFF99](https://dummyimage.com/20/9FFF99/000000.png&text=+) `#9FFF99` | ![#BEAFFF](https://dummyimage.com/20/BEAFFF/000000.png&text=+) `#BEAFFF` | ![#9FDDFF](https://dummyimage.com/20/9FDDFF/000000.png&text=+) `#9FDDFF` | ![#F4BB5E](https://dummyimage.com/20/F4BB5E/000000.png&text=+) `#F4BB5E` |
| ![#BFFFB7](https://dummyimage.com/20/BFFFB7/000000.png&text=+) `#BFFFB7` | ![#C8C0FF](https://dummyimage.com/20/C8C0FF/000000.png&text=+) `#C8C0FF` | ![#CBEEFF](https://dummyimage.com/20/CBEEFF/000000.png&text=+) `#CBEEFF` | ![#F9D674](https://dummyimage.com/20/F9D674/000000.png&text=+) `#F9D674` |
| ![#E4FFD6](https://dummyimage.com/20/E4FFD6/000000.png&text=+) `#E4FFD6` | ![#DEDDFF](https://dummyimage.com/20/DEDDFF/000000.png&text=+) `#DEDDFF` | ![#E5F8FF](https://dummyimage.com/20/E5F8FF/000000.png&text=+) `#E5F8FF` | ![#FBEECE](https://dummyimage.com/20/FBEECE/000000.png&text=+) `#FBEECE` |

## How It Works

The script follows a series of steps to perform color replacement:

### 1. File Analysis

- **Filename Parsing**: The script extracts components from the source filename:
  - **Document Name**: The main name of the document.
  - **Version**: Identified by a pattern like `- vX.Y`.
  - **Date**: Recognized formats like `yyyy-mm-dd` or `yyyy.mm.dd`.
  - **Color Scheme**: Extracted from the filename if present.

- **Regular Expressions**: The script uses regex patterns to identify and extract these components.

### 2. Color Scheme Detection

- **From Filename**: If a color scheme is found in the filename, it's used as the source scheme.
- **From File Content**:
  - If no color scheme is detected in the filename, the script examines the file content.
  - **Extraction**: The script unpacks the file (if it's a zipped format like `.pptx` or `.docx`) and searches for known color codes within the files.
  - **Matching**: It compares found color codes against predefined color schemes to determine the source scheme.

### 3. Color Replacement

- **Target Schemes**: The script identifies all color schemes different from the source scheme as target schemes.
- **Processing Each Scheme**:
  - **File Extraction**: For each target scheme, the script extracts the contents of the source file into a temporary directory.
  - **Color Files Retrieval**: It gathers all relevant files that may contain color codes (e.g., `.xml`, `.svg`).
  - **Replacement Logic**:
    - The script replaces color codes from the source scheme with those from the target scheme.
    - It uses the order of colors defined in the color scheme mappings to ensure accurate replacement.
    - Optionally modifies the slide master when `-ChangeSlideMaster` is specified.
  - **File Reassembly**: After replacement, the script reassembles the files back into the original format.

### 4. Slide Master Update (`-ChangeSlideMaster`)

For PowerPoint files, when `-ChangeSlideMaster` is enabled, the script updates the slide master with the new color scheme. This ensures consistent appearance across all slides.

### 5. Output Generation

- **Filename Construction**: The script constructs the output filenames using the extracted components in the following order:
  - `DocumentName - Version - Date - ColorScheme`
  - Components not present in the original filename are omitted.
- **File Saving**: The new files with the replaced color schemes are saved in the same directory as the source file.

### 6. Cleanup

- **Temporary Files**: The script deletes any temporary files and directories created during the process to ensure no unnecessary files are left behind.

## Example

Suppose you have a PowerPoint template named:

```
Company Presentation - v1.0 - Blue - 2023-08-15.pptx
```

Running the script:

```powershell
.\ColorReplacementScript.ps1 -sourceFile "C:\Templates\Company Presentation - v1.0 - Blue - 2023-08-15.pptx" -ChangeSlideMaster
```

The script will:

- Detect that the current color scheme is **Blue**.
- Update the slide master and content for each target scheme.
- Generate new files with the same document name, version, and date but with other color schemes:

  - `Company Presentation - v1.0 - 2023.08.15 - Green.pptx`
  - `Company Presentation - v1.0 - 2023.08.15 - Purple.pptx`
  - `Company Presentation - v1.0 - 2023.08.15 - Red.pptx`

Each new file will have the colors replaced according to the target scheme.

## Detailed Explanation of the Script Components

### Functions

- **`Get-ColorFiles`**:
  - Retrieves all relevant files that may contain color codes based on the file type (e.g., `.pptx`, `.docx`, `.svg`).
  - Excludes certain files (like theme files) to avoid unintended replacements.
  - Will include PowerPoint Slide Master layouts when `-ChangeSlideMaster` is given.

- **`Detect-ColorScheme`**:
  - Uses `Get-ColorFiles` to obtain files.
  - Reads the content of these files to detect which color scheme is present by matching color codes.

- **`Replace-Colors`**:
  - Also uses `Get-ColorFiles` to obtain files.
  - Replaces each color code from the source scheme with the corresponding color code from the target scheme.
  - Writes the updated content back to the files.

- **`Process-Template`**:
  - Manages the overall process for each target color scheme.
  - Handles file extraction and reassembly.
  - Calls `Replace-Colors` for the actual replacement.

### Variables and Data Structures

- **`$colorSchemes`**:
  - An ordered hashtable containing the color schemes and their respective color codes.
  - Defines the mapping between different schemes.

- **Regex Patterns**:
  - **Version**: Identifies version information in the filename.
  - **Color Scheme**: Dynamically generated from the keys of `$colorSchemes`.
  - **Date**: Matches date formats to extract date information.

## Cleanup

After processing, the script automatically cleans up temporary files and directories created during execution to maintain a clean working environment.


## License

This project is licensed under the MIT License. See the LICENSE file for more details.