# Certificate Generator

A powerful and user-friendly application for generating professional certificates from Excel data. This tool allows you to create customized certificates in bulk with dynamic field placement and styling.

## Features

- **Dynamic Field Handling**
  - Import data from Excel files
  - Automatic field detection from headers
  - Customizable field visibility
  - Individual field styling

- **Certificate Customization**
  - Drag-and-drop field placement
  - Adjustable font sizes
  - Custom color selection
  - Real-time preview
  - First field auto-centering

- **Batch Processing**
  - Generate multiple certificates at once
  - Progress tracking
  - PDF output format
  - Automatic file naming

- **Project Management**
  - Save certificate layouts
  - Load previous projects
  - Preserve all settings and positions

## Installation

### Option 1: Windows Executable (Recommended for Windows Users)
1. Download the latest release from the [Releases](https://github.com/yourusername/certificate-generator/releases) page
2. Extract the downloaded ZIP file
3. Run `certificate-generator.exe`

### Option 2: From Source
1. Clone the repository:
```bash
git clone https://github.com/yourusername/certificate-generator.git
cd certificate-generator
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Requirements

### For Windows Executable
- Windows 10 or later
- No additional software required

### For Source Installation
- Python 3.x
- Required packages:
  - tkinter
  - Pillow (PIL)
  - openpyxl
  - fpdf

## Usage

1. **Prepare Your Data**
   - Create an Excel file with headers for each field
   - Add student data in rows below the headers
   - Save the file in .xlsx format

2. **Load Template and Data**
   - Click "Load Template" to select your certificate template (PNG format)
   - Click "Load Excel" to import your student data

3. **Customize Fields**
   - Use checkboxes to toggle field visibility
   - Adjust font sizes using the spinners
   - Click "Color" to change text color
   - Drag and drop fields to position them
   - First field will be automatically centered

4. **Preview and Generate**
   - Click "Preview" to see how the certificate will look
   - Click "Generate" to create PDF certificates
   - Select output directory for the generated files


## Saving and Loading Projects

- Click "Project" > "Save Project" to save your current layout
- Click "Project" > "Load Project" to restore a previous layout
- Project files (.certproj) contain:
  - Template path
  - Field positions
  - Font settings
  - Excel data path

## Known Issues

- None reported

## Future Improvements

- Template customization options
- Additional font support
- More export formats
- Batch processing improvements
- Template library
- Custom field validation
- Additional platform executables (macOS, Linux)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please open an issue in the GitHub repository or contact the maintainers. 