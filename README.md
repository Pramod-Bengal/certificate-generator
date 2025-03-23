# Certificate Generator

A proprietary certificate generation application developed for EyeQDotnet Pvt. The application generates personalized certificates in both vertical and horizontal orientations from Excel data.

## Features

- Generate certificates in vertical and/or horizontal orientations
- Read data from Excel files
- User-friendly GUI for orientation selection
- Automatic output organization in separate folders
- Sanitized filenames for compatibility

## Prerequisites

- Python 3.6 or higher
- Required font files:
  - `arial.ttf`

## Required Files

Ensure you have the following files in your project directory:
- `generate.py` (main script)
- `cert_vertical.pdf` (vertical certificate template)
- `cert_horizontal.pdf` (horizontal certificate template)
- Excel file containing certificate data with columns:
  - "Name"
  - "Certificate_Number"

## Installation

1. Download the project files
2. Install required Python packages:
```bash
pip install pandas PyPDF2 reportlab
```

## Usage

1. Place your certificate templates in the project directory:
   - `cert_vertical.pdf` for vertical certificates
   - `cert_horizontal.pdf` for horizontal certificates

2. Prepare your Excel file with the following columns:
   - "Name": The name to appear on the certificate
   - "Certificate_Number": The unique certificate number

3. Run the script:
```bash
python generate.py
```

4. Follow the GUI prompts:
   - Select your Excel file when prompted
   - Choose orientation(s) using checkboxes
   - Click "Generate" to create certificates

5. Generated certificates will be saved in:
   - `output_vertical/` for vertical certificates
   - `output_horizontal/` for horizontal certificates

## Creating Windows Executable

To create a standalone Windows executable:

1. Install PyInstaller:
```bash
pip install pyinstaller
```

2. Create the executable:
```bash
pyinstaller --onefile generate.py
```

3. The executable will be created in the `dist` directory

### Important Notes for Executable:
- Copy your certificate templates (`cert_vertical.pdf` and `cert_horizontal.pdf`) to the same directory as the executable
- Font files will be included automatically in the executable
- Make sure your Excel file is accessible when running the executable

## Troubleshooting

1. **Font Errors**:
   - Ensure `arial.ttf` and `arialbd.ttf` are in the project directory
   - On Windows, you can copy these from `C:\Windows\Fonts\`

2. **File Not Found Errors**:
   - Verify certificate templates are named correctly
   - Check that templates are in the same directory as the script/executable

3. **Excel File Issues**:
   - Confirm column names are exactly "Name" and "Certificate_Number"
   - Make sure the Excel file is not open in another program

## Repository Access and Security

### Access Control
- This is a private repository accessible only to authorized EyeQDotnet Pvt. team members
- Access requests must be approved by repository administrators
- Team members must have 2FA enabled on their GitHub accounts

### Security Guidelines
1. Never commit sensitive data directly to the repository
2. Use environment variables for sensitive configurations
3. Keep your access tokens and credentials secure
4. Report any security concerns to the repository administrators immediately

### Getting Access
To request access to this repository:
1. Ensure you have a GitHub account with 2FA enabled
2. Contact the repository administrator with your GitHub username
3. Wait for access confirmation
4. Set up your local development environment following the installation instructions

## Copyright

© 2024 EyeQDotnet Pvt. All rights reserved.

This software is proprietary and confidential. Unauthorized copying, modification, distribution, or use of this software, via any medium, is strictly prohibited. 