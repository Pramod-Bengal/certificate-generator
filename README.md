# Certificate Generator

[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/shahil-sk/certificate-generator)

A Python-based application for generating certificates from templates with customizable text placement, fonts, and colors.

## Repository

This project is hosted on GitHub: [https://github.com/shahil-sk/certificate-generator](https://github.com/shahil-sk/certificate-generator)

## Features

- Load PNG template images
- Import student data from Excel files
- Drag-and-drop text placement
- Customizable font sizes and colors
- Preview certificates before generation
- Save and load placeholder positions
- Generate PDF certificates in bulk
- Progress tracking for batch generation

## Requirements

- Python 3.6 or higher
- Pillow (PIL)
- openpyxl
- fpdf2

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/shahil-sk/certificate-generator.git
   cd certificate-generator
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python script.py
   ```

2. Load a template image (PNG format)
3. Load student data from an Excel file
4. Drag the placeholders to desired positions
5. Customize font sizes and colors
6. Preview the certificate
7. Generate certificates for all students

## Excel File Format

The Excel file should have the following columns:
- Name
- ID
- Start Date
- End Date

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. 