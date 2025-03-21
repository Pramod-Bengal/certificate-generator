import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4, portrait
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import io
from tkinter import filedialog, messagebox, Tk, Checkbutton, Button, Label, IntVar
import re
import os


# Load Excel file
excel_file = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not excel_file:
    print("No Excel file selected. Exiting...")
    exit()
df = pd.read_excel(excel_file)

# Create orientation selection window
root = Tk()
root.title("Select Certificate Orientations")
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

    # Function to sanitize filename
    def sanitize_filename(filename):
        # Replace invalid characters with underscores
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    # Function to generate certificates for a specific orientation
    def generate_certificates(orientation_type):
        if orientation_type == 'vertical':
            template_path = "cert_vertical.pdf"
            page_size = portrait(A4)
            # Coordinates for vertical orientation
            name_x, name_y = 280,470
            cert_x, cert_y = 425, 820
            output_prefix = "Vertical_"
        else:
            template_path = "cert_horizontal.pdf"
            page_size = landscape(A4)
            # Coordinates for horizontal orientation
            name_x, name_y = 405, 275
            cert_x, cert_y = 5, 583
            output_prefix = "Horizontal_"

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

        # Output certificates
        for index, row in df.iterrows():
            name = row["Name"]
            cert_number = row["Certificate_Number"]
            
            # Create a new PDF with text
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=page_size)
            
            # Add Name & Certificate Number at adjusted positions
            can.setFont("Arial-Bold", 22)  # Using Arial Bold for name
            can.setFillColor("#FF0000")  # Set color to red for name
            can.drawString(name_x, name_y, name)  # f-string not needed for single variable
            
            can.setFont("Arial-Bold", 17)  # Using Arial Bold for certificate number
            can.setFillColor("white")  # Set color to white for certificate number
            can.drawString(cert_x, cert_y, f"{cert_number}")
            can.save()
            
            # Merge with template
            packet.seek(0)
            new_pdf = PdfReader(packet)
            output = PdfWriter()
            page = pdf_reader.pages[0]
            page.merge_page(new_pdf.pages[0])
            output.add_page(page)
            
            # Save output file with orientation prefix and sanitized filename
            output_filename = sanitize_filename(f"{output_prefix}Certificate_{cert_number}-{name}.pdf")
            with open(output_filename, "wb") as outputStream:
                output.write(outputStream)
            
            print(f"Generated: {output_filename}")

    # Generate certificates for each selected orientation
    for orientation_type in orientations:
        print(f"\nGenerating {orientation_type} certificates...")
        generate_certificates(orientation_type)

    print("\nAll certificates generated successfully!")

except Exception as e:
    print(f"An error occurred: {str(e)}")
    messagebox.showerror("Error", f"An error occurred while generating certificates: {str(e)}")
finally:
    try:
        root.destroy()
    except:
        pass