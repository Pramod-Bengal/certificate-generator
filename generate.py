import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4, portrait
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import CMYKColor, Color
import io
from tkinter import filedialog, messagebox, Tk, Checkbutton, Button, Label, IntVar
import re
import os

# Create output directories if they don't exist
def ensure_output_dirs():
    base_dirs = ['output_vertical', 'output_horizontal']
    color_spaces = ['CMYK', 'RGB']
    
    for base_dir in base_dirs:
        for color_space in color_spaces:
            dir_name = os.path.join(base_dir, color_space)
            if not os.path.exists(dir_name):
                os.makedirs(dir_name)
                print(f"Created directory: {dir_name}")

# Load Excel file
excel_file = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not excel_file:
    print("No Excel file selected. Exiting...")
    exit()
df = pd.read_excel(excel_file)

# Create selection window
root = Tk()
root.title("Certificate Generation Settings")
root.geometry("300x150")

# Variables to store checkbox states
vertical_var = IntVar()
horizontal_var = IntVar()

# Create and pack widgets
Label(root, text="Select orientations to generate:", pady=10).pack()
Checkbutton(root, text="Vertical", variable=vertical_var).pack()
Checkbutton(root, text="Horizontal", variable=horizontal_var).pack()

# Function to handle selection
def confirm_selection():
    orientations = []
    if vertical_var.get():
        orientations.append('vertical')
    if horizontal_var.get():
        orientations.append('horizontal')
    
    if not orientations:
        messagebox.showwarning("Warning", "Please select at least one orientation!")
        return
    
    root.orientations = orientations
    root.destroy()

# Create and pack confirm button
Button(root, text="Generate", command=confirm_selection).pack(pady=10)

# Start the main loop
root.mainloop()

try:
    # Get the selected orientations
    orientations = getattr(root, 'orientations', [])
    
    if not orientations:
        print("No orientations selected. Exiting...")
        exit()

    # Create output directories
    ensure_output_dirs()

    # Function to sanitize filename
    def sanitize_filename(filename):
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    # Function to get colors based on color space
    def get_colors(color_space):
        if color_space == 'cmyk':
            # CMYK colors for printing
            name_color = CMYKColor(0, 1, 1, 0)  # Red in CMYK
            cert_color = CMYKColor(0, 0, 0, 0)  # White in CMYK
        else:
            # RGB colors for digital
            name_color = Color(1, 0, 0)  # Red in RGB
            cert_color = Color(1, 1, 1)  # White in RGB
        return name_color, cert_color

    # Function to generate certificates for a specific orientation and color space
    def generate_certificates(orientation_type, color_space):
        if orientation_type == 'vertical':
            template_path = "cert_vertical.pdf"
            page_size = portrait(A4)
            # Coordinates for vertical orientation
            name_x, name_y = 280,470
            cert_x, cert_y = 425, 820
            base_dir = "output_vertical"
        else:
            template_path = "cert_horizontal.pdf"
            page_size = landscape(A4)
            # Coordinates for horizontal orientation
            name_x, name_y = 405, 275
            cert_x, cert_y = 5, 583
            base_dir = "output_horizontal"

        # Set output directory based on color space
        output_dir = os.path.join(base_dir, color_space.upper())

        # Load PDF template
        pdf_reader = PdfReader(template_path)

        # Font settings
        try:
            pdfmetrics.registerFont(TTFont("Arial", "arial.ttf"))
            pdfmetrics.registerFont(TTFont("Arial-Bold", "arialbd.ttf"))
        except Exception as e:
            print(f"Error loading fonts: {str(e)}")
            print("Please make sure arial.ttf and arialbd.ttf are in the same directory as the script.")
            messagebox.showerror("Font Error", "Could not load required fonts. Please make sure arial.ttf and arialbd.ttf are in the same directory.")
            return

        # Get colors based on color space
        name_color, cert_color = get_colors(color_space)

        # Output certificates
        for index, row in df.iterrows():
            name = row["Name"]
            cert_number = row["Certificate_Number"]
            
            # Create a new PDF with text
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=page_size)
            
            # Add Name & Certificate Number at adjusted positions
            can.setFont("Arial-Bold", 22)
            can.setFillColor(name_color)
            can.drawString(name_x, name_y, name)
            
            can.setFont("Arial-Bold", 17)
            can.setFillColor(cert_color)
            can.drawString(cert_x, cert_y, f"{cert_number}")
            can.save()
            
            # Merge with template
            packet.seek(0)
            new_pdf = PdfReader(packet)
            output = PdfWriter()
            
            # Create a new PDF reader for each certificate to get a fresh template
            template_reader = PdfReader(template_path)
            template_page = template_reader.pages[0]
            
            # Get the text page
            text_page = new_pdf.pages[0]
            
            # Merge the pages (text on top of template)
            template_page.merge_page(text_page)
            
            # Add the merged page to the output
            output.add_page(template_page)
            
            # Generate output file
            generate_output(output, output_dir, cert_number, name)
            
            # Clean up resources
            packet.close()
            del can
            del new_pdf
            del output
            del template_reader
            
    # Save output file with orientation prefix and sanitized filename        
    def generate_output(output, output_dir, cert_number, name):
        output_filename = sanitize_filename(f"{cert_number}-{name}.pdf")
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, "wb") as outputStream:
            output.write(outputStream)
        print(f"Generated: {output_path}")
    
    # Generate certificates for each selected orientation and color space
    for orientation_type in orientations:
        print(f"\nGenerating {orientation_type} certificates...")
        # Generate CMYK version
        print(f"Generating CMYK version...")
        generate_certificates(orientation_type, 'cmyk')
        # Generate RGB version
        print(f"Generating RGB version...")
        generate_certificates(orientation_type, 'rgb')

    print("\nAll certificates generated successfully!")

except Exception as e:
    print(f"An error occurred: {str(e)}")
    messagebox.showerror("Error", f"An error occurred while generating certificates: {str(e)}")
finally:
    try:
        root.destroy()
    except:
        pass