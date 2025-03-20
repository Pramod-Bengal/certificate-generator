import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import io
from tkinter import filedialog

# Load Excel file
excel_file = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not excel_file:
    print("No Excel file selected. Exiting...")
    exit()
df = pd.read_excel(excel_file)

# Load PDF template
template_path = filedialog.askopenfilename(title="Select certificate template", filetypes=[("PDF files", "*.pdf")])
if not template_path:
    print("No template selected. Exiting...")
    exit()
pdf_reader = PdfReader(template_path)

# Font settings
pdfmetrics.registerFont(TTFont("Arial", "arial.ttf"))

# Page dimensions
page_width, page_height = landscape(A4)

# Output certificates
for index, row in df.iterrows():
    name = row["Name"]
    cert_number = row["Certificate_Number"]
    
    # Create a new PDF with text
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=landscape(A4))
    can.setFont("Arial", 18)
    
    # Add Name & Certificate Number at adjusted positions
    can.setFillColor("#FF0000")  # Set color to red for name
    can.drawString(405, 275, name)  # f-string not needed for single variable
    can.setFillColor("#000000")  # Set color to black for certificate number (white would be invisible)
    can.drawString(100, 555, f"Cert No: {cert_number}")
    can.save()
    
    # Merge with template
    packet.seek(0)
    new_pdf = PdfReader(packet)
    output = PdfWriter()
    page = pdf_reader.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    
    # Save output file
    output_filename = f"Certificate_{cert_number}-{name}.pdf"
    with open(output_filename, "wb") as outputStream:
        output.write(outputStream)
    
    print(f"Generated: {output_filename}")

print("All certificates generated successfully!")