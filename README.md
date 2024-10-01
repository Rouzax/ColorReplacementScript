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

The script includes the following extended color schemes:

| Color Scheme | Hex Code | Fill Color         | Text Color  |
|--------------|----------|---------------------|-------------|
| Green        | `#244739`| ![#244739](https://dummyimage.com/15/244739/000000?text=+) | Black       |
|              | `#1B5744`| ![#1B5744](https://dummyimage.com/15/1B5744/000000?text=+) | Black       |
|              | `#247554`| ![#247554](https://dummyimage.com/15/247554/000000?text=+) | Black       |
|              | `#349E5F`| ![#349E5F](https://dummyimage.com/15/349E5F/000000?text=+) | Black       |
|              | `#37CC5C`| ![#37CC5C](https://dummyimage.com/15/37CC5C/000000?text=+) | Black       |
|              | `#4EE760`| ![#4EE760](https://dummyimage.com/15/4EE760/000000?text=+) | Black       |
|              | `#7BFC76`| ![#7BFC76](https://dummyimage.com/15/7BFC76/000000?text=+) | Black       |
|              | `#9FFF99`| ![#9FFF99](https://dummyimage.com/15/9FFF99/000000?text=+) | Black       |
|              | `#BFFFB7`| ![#BFFFB7](https://dummyimage.com/15/BFFFB7/000000?text=+) | Black       |
|              | `#E4FFD6`| ![#E4FFD6](https://dummyimage.com/15/E4FFD6/000000?text=+) | Black       |
| Purple       | `#2A145A`| ![#2A145A](https://dummyimage.com/15/2A145A/FFFFFF?text=+) | White       |
|              | `#500A96`| ![#500A96](https://dummyimage.com/15/500A96/FFFFFF?text=+) | White       |
|              | `#612CB0`| ![#612CB0](https://dummyimage.com/15/612CB0/FFFFFF?text=+) | White       |
|              | `#743DD4`| ![#743DD4](https://dummyimage.com/15/743DD4/FFFFFF?text=+) | White       |
|              | `#8E5CEF`| ![#8E5CEF](https://dummyimage.com/15/8E5CEF/FFFFFF?text=+) | White       |
|              | `#9F78FC`| ![#9F78FC](https://dummyimage.com/15/9F78FC/FFFFFF?text=+) | White       |
|              | `#AA96FA`| ![#AA96FA](https://dummyimage.com/15/AA96FA/FFFFFF?text=+) | White       |
|              | `#BEAFFF`| ![#BEAFFF](https://dummyimage.com/15/BEAFFF/000000?text=+) | Black       |
|              | `#C8C0FF`| ![#C8C0FF](https://dummyimage.com/15/C8C0FF/000000?text=+) | Black       |
|              | `#DEDDFF`| ![#DEDDFF](https://dummyimage.com/15/DEDDFF/000000?text=+) | Black       |
| Blue         | `#0D2155`| ![#0D2155](https://dummyimage.com/15/0D2155/FFFFFF?text=+) | White       |
|              | `#00227F`| ![#00227F](https://dummyimage.com/15/00227F/FFFFFF?text=+) | White       |
|              | `#0C32A4`| ![#0C32A4](https://dummyimage.com/15/0C32A4/FFFFFF?text=+) | White       |
|              | `#1D56C0`| ![#1D56C0](https://dummyimage.com/15/1D56C0/FFFFFF?text=+) | White       |
|              | `#0672CB`| ![#0672CB](https://dummyimage.com/15/0672CB/FFFFFF?text=+) | White       |
|              | `#58A5E6`| ![#58A5E6](https://dummyimage.com/15/58A5E6/000000?text=+) | Black       |
|              | `#80C7FB`| ![#80C7FB](https://dummyimage.com/15/80C7FB/000000?text=+) | Black       |
|              | `#9FDDFF`| ![#9FDDFF](https://dummyimage.com/15/9FDDFF/000000?text=+) | Black       |
|              | `#CBEEFF`| ![#CBEEFF](https://dummyimage.com/15/CBEEFF/000000?text=+) | Black       |
|              | `#E5F8FF`| ![#E5F8FF](https://dummyimage.com/15/E5F8FF/000000?text=+) | Black       |
| Red          | `#4A193A`| ![#4A193A](https://dummyimage.com/15/4A193A/FFFFFF?text=+) | White       |
|              | `#691D3F`| ![#691D3F](https://dummyimage.com/15/691D3F/FFFFFF?text=+) | White       |
|              | `#85133F`| ![#85133F](https://dummyimage.com/15/85133F/FFFFFF?text=+) | White       |
|              | `#B30B37`| ![#B30B37](https://dummyimage.com/15/B30B37/FFFFFF?text=+) | White       |
|              | `#D2333D`| ![#D2333D](https://dummyimage.com/15/D2333D/FFFFFF?text=+) | White       |
|              | `#E1633F`| ![#E1633F](https://dummyimage.com/15/E1633F/FFFFFF?text=+) | White       |
|              | `#E17F3F`| ![#E17F3F](https://dummyimage.com/15/E17F3F/FFFFFF?text=+) | White       |
|              | `#F4BB5E`| ![#F4BB5E](https://dummyimage.com/15/F4BB5E/000000?text=+) | Black       |
|              | `#F9D674`| ![#F9D674](https://dummyimage.com/15/F9D674/000000?text=+) | Black       |
|              | `#FBEECE`| ![#FBEECE](https://dummyimage.com/15/FBEECE/000000?text=+) | Black       |

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