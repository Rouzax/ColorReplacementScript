# PowerShell Color Replacement Script

This PowerShell script is designed to automate the process of color replacement in template files. It supports various file types, including Microsoft Office formats (`.docx`, `.dotx`, `.pptx`, `.potx`) and SVG files. The script detects the current color scheme from the filename or the content of the file and generates new files with the specified color replacements.

## Features

- **Supported File Types**: 
  - Microsoft Word: `.docx`, `.dotx`
  - Microsoft PowerPoint: `.pptx`, `.potx`
  - SVG: `.svg`

- **Color Scheme Detection**: Automatically detects the current color scheme from the filename or file content.

- **Color Replacement**: Allows for the replacement of color schemes based on predefined mappings.

- **Output Generation**: Produces new files named according to the original document name, version, date, and new color scheme.

## Prerequisites

- Windows operating system with PowerShell 5.1 or later.
- The `Microsoft .NET Framework` for PowerShell.
- The script requires access to the file system for reading and writing files.

## Usage

1. **Clone the Repository** or download the script file.

    ```bash
    git clone <repository-url>
    ```

2. **Run the Script**: Execute the script in PowerShell with the required parameters.

    ```powershell
    .\ColorReplacementScript.ps1 -sourceFile "path\to\your\template.file"
    ```

### Parameters

- `-sourceFile`: (Mandatory) The path to the source file (template) you want to process.

## Color Schemes

### Color Replacement Table

| Green Scheme | Purple Scheme | Blue Scheme | Red Scheme |
|--------------|---------------|-------------|------------|
| ![#244739](https://dummyimage.com/20/244739/000000?text=+) `#244739` | ![#2A145A](https://dummyimage.com/20/2A145A/000000?text=+) `#2A145A` | ![#0D2155](https://dummyimage.com/20/0D2155/000000?text=+) `#0D2155` | ![#4A193A](https://dummyimage.com/20/4A193A/000000?text=+) `#4A193A` |
| ![#1B5744](https://dummyimage.com/20/1B5744/000000?text=+) `#1B5744` | ![#500A96](https://dummyimage.com/20/500A96/000000?text=+) `#500A96` | ![#00227F](https://dummyimage.com/20/00227F/000000?text=+) `#00227F` | ![#691D3F](https://dummyimage.com/20/691D3F/000000?text=+) `#691D3F` |
| ![#247554](https://dummyimage.com/20/247554/000000?text=+) `#247554` | ![#612CB0](https://dummyimage.com/20/612CB0/000000?text=+) `#612CB0` | ![#0C32A4](https://dummyimage.com/20/0C32A4/000000?text=+) `#0C32A4` | ![#85133F](https://dummyimage.com/20/85133F/000000?text=+) `#85133F` |
| ![#349E5F](https://dummyimage.com/20/349E5F/000000?text=+) `#349E5F` | ![#743DD4](https://dummyimage.com/20/743DD4/000000?text=+) `#743DD4` | ![#1D56C0](https://dummyimage.com/20/1D56C0/000000?text=+) `#1D56C0` | ![#B30B37](https://dummyimage.com/20/B30B37/000000?text=+) `#B30B37` |
| ![#37CC5C](https://dummyimage.com/20/37CC5C/000000?text=+) `#37CC5C` | ![#8E5CEF](https://dummyimage.com/20/8E5CEF/000000?text=+) `#8E5CEF` | ![#0672CB](https://dummyimage.com/20/0672CB/000000?text=+) `#0672CB` | ![#D2333D](https://dummyimage.com/20/D2333D/000000?text=+) `#D2333D` |
| ![#4EE760](https://dummyimage.com/20/4EE760/000000?text=+) `#4EE760` | ![#9F78FC](https://dummyimage.com/20/9F78FC/000000?text=+) `#9F78FC` | ![#58A5E6](https://dummyimage.com/20/58A5E6/000000?text=+) `#58A5E6` | ![#E1633F](https://dummyimage.com/20/E1633F/000000?text=+) `#E1633F` |
| ![#7BFC76](https://dummyimage.com/20/7BFC76/000000?text=+) `#7BFC76` | ![#AA96FA](https://dummyimage.com/20/AA96FA/000000?text=+) `#AA96FA` | ![#80C7FB](https://dummyimage.com/20/80C7FB/000000?text=+) `#80C7FB` | ![#E17F3F](https://dummyimage.com/20/E17F3F/000000?text=+) `#E17F3F` |
| ![#9FFF99](https://dummyimage.com/20/9FFF99/000000?text=+) `#9FFF99` | ![#BEAFFF](https://dummyimage.com/20/BEAFFF/000000?text=+) `#BEAFFF` | ![#9FDDFF](https://dummyimage.com/20/9FDDFF/000000?text=+) `#9FDDFF` | ![#F4BB5E](https://dummyimage.com/20/F4BB5E/000000?text=+) `#F4BB5E` |
| ![#BFFFB7](https://dummyimage.com/20/BFFFB7/000000?text=+) `#BFFFB7` | ![#C8C0FF](https://dummyimage.com/20/C8C0FF/000000?text=+) `#C8C0FF` | ![#CBEEFF](https://dummyimage.com/20/CBEEFF/000000?text=+) `#CBEEFF` | ![#F9D674](https://dummyimage.com/20/F9D674/000000?text=+) `#F9D674` |
| ![#E4FFD6](https://dummyimage.com/20/E4FFD6/000000?text=+) `#E4FFD6` | ![#DEDDFF](https://dummyimage.com/20/DEDDFF/000000?text=+) `#DEDDFF` | ![#E5F8FF](https://dummyimage.com/20/E5F8FF/000000?text=+) `#E5F8FF` | ![#FBEECE](https://dummyimage.com/20/FBEECE/000000?text=+) `#FBEECE` |


## How It Works

1. **File Analysis**: The script first checks the filename for version, color scheme, and date information.
2. **Color Scheme Detection**: If no color scheme is found in the filename, the script attempts to detect it from the file content.
3. **Color Replacement**: The specified source color scheme is replaced with the target color scheme across all applicable files.
4. **Output**: New files are created in the same directory as the source file, named according to the specified format.

## Example

To replace colors in a PowerPoint template file:

```powershell
.\ColorReplacementScript.ps1 -sourceFile "C:\path\to\template.pptx"
```