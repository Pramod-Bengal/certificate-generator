import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from openpyxl import load_workbook
from fpdf import FPDF
from PIL import ImageFont, ImageDraw
from tkinter import filedialog, messagebox
import re



class CertificateApp:
    
    def __init__(self, root):

        self.master = root
        self.root = root
        self.root.title("Certificate Generator")

        self.original_image = None
        self.display_image = None
        self.scale_x = 1
        self.scale_y = 1

        self.canvas = tk.Canvas(root, width=1000, height=700, bg="gray")
        self.canvas.pack()

        self.load_btn = tk.Button(root, text="Load Template", command=self.load_template)
        self.load_btn.pack(pady=10)

        self.placeholders = {}  # Store placeholder references

        self.excel_data = []  # Will store list of student dictionaries

        self.excel_btn = tk.Button(root, text="Load Excel", command=self.load_excel)
        self.excel_btn.pack(pady=5)

        self.include_name = tk.BooleanVar(value=True)
        self.include_id = tk.BooleanVar(value=True)
        self.include_start = tk.BooleanVar(value=True)
        self.include_end = tk.BooleanVar(value=True)


        # Section label
        tk.Label(self.master, text="Include Fields in Certificate:", font=("Arial", 12, "bold")).pack(pady=(10, 0))

        # Create a frame to group the checkbuttons
        checkbox_frame = tk.Frame(self.master)
        checkbox_frame.pack(pady=(0, 10))

        # Individual Checkbuttons
        tk.Checkbutton(checkbox_frame, text="Name", variable=self.include_name).grid(row=0, column=0, sticky="w", padx=5)
        tk.Checkbutton(checkbox_frame, text="ID", variable=self.include_id).grid(row=0, column=1, sticky="w", padx=5)
        tk.Checkbutton(checkbox_frame, text="Start Date", variable=self.include_start).grid(row=1, column=0, sticky="w", padx=5)
        tk.Checkbutton(checkbox_frame, text="End Date", variable=self.include_end).grid(row=1, column=1, sticky="w", padx=5)
        
        self.generate_btn = tk.Button(root, text="Generate Certificates", command=self.generate_certificates)
        self.generate_btn.pack(pady=5)



    def load_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png")])
        if not file_path:
            return

        self.original_image = Image.open(file_path)
        original_width, original_height = self.original_image.size

        max_width, max_height = 1000, 700
        ratio = min(max_width / original_width, max_height / original_height)
        new_size = (int(original_width * ratio), int(original_height * ratio))

        self.scale_x = original_width / new_size[0]
        self.scale_y = original_height / new_size[1]

        resized_img = self.original_image.resize(new_size)
        self.display_image = ImageTk.PhotoImage(resized_img)

        self.canvas.config(width=new_size[0], height=new_size[1])
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.display_image, anchor="nw")

        # Create draggable placeholders
        self.create_placeholder("Name")
        self.create_placeholder("ID")
        self.create_placeholder("Start Date")
        self.create_placeholder("End Date")

    def create_placeholder(self, label):
        widget = tk.Label(self.canvas, text=label, bg="yellow", fg="black")
        item = self.canvas.create_window(50, 50, window=widget, anchor="nw")

        def start_drag(event, canvas_item=item):
            self._drag_data = {"item": canvas_item, "x": event.x, "y": event.y}

        def do_drag(event):
            dx = event.x - self._drag_data["x"]
            dy = event.y - self._drag_data["y"]
            self.canvas.move(self._drag_data["item"], dx, dy)
            self._drag_data["x"] = event.x
            self._drag_data["y"] = event.y

        widget.bind("<Button-1>", start_drag)
        widget.bind("<B1-Motion>", do_drag)

        self.placeholders[label] = item  # Save it

    def get_placeholder_positions(self):
        """Get scaled coordinates for actual certificate."""
        coords = {}
        for label, item in self.placeholders.items():
            x, y = self.canvas.coords(item)
            coords[label] = (x * self.scale_x, y * self.scale_y)
        return coords
    
    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        wb = load_workbook(file_path)
        sheet = wb.active

        self.excel_data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header
            name, student_id, start_date, end_date = row
            self.excel_data.append({
                "Name": str(name),
                "ID": str(student_id),
                "Start Date": str(start_date),
                "End Date": str(end_date)
            })

        print(f"Loaded {len(self.excel_data)} students from Excel.")
        print("Include Name:", self.include_name.get())
        print("Include ID:", self.include_id.get())
        print("Include Start:", self.include_start.get())
        print("Include End:", self.include_end.get())

    def generate_certificates(self):
        if not self.excel_data:
            messagebox.showwarning("Warning", "No student data loaded!")
            return

        if not self.original_image:
            messagebox.showwarning("Warning", "No template image loaded!")
            return

        # Let user choose folder to save generated certificates
        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
            return  # User cancelled

        font_path = "arial.ttf"
        generated_count = 0  # Make sure to have this font file in your directory

        for student in self.excel_data:
            # Create a PDF object
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            # Use the already loaded original image for the template
            original_img = self.original_image.copy()

            # Create ImageDraw object
            draw = ImageDraw.Draw(original_img)

            # Get scaled placeholder positions
            placeholder_positions = self.get_placeholder_positions()
            # Name
            if self.include_name.get():
                name_x, name_y = placeholder_positions["Name"]
                font = ImageFont.truetype(font_path, 48)
                draw.text((name_x, name_y), student["Name"], font=font, fill="black")

            # ID
            if self.include_id.get():
                id_x, id_y = placeholder_positions["ID"]
                font = ImageFont.truetype(font_path, 32)
                draw.text((id_x, id_y), student["ID"], font=font, fill="black")

            # Start Date
            if self.include_start.get():
                start_x, start_y = placeholder_positions["Start Date"]
                font = ImageFont.truetype(font_path, 32)
                draw.text((start_x, start_y), student["Start Date"], font=font, fill="black")

            # End Date
            if self.include_end.get():
                end_x, end_y = placeholder_positions["End Date"]
                font = ImageFont.truetype(font_path, 32)
                draw.text((end_x, end_y), student["End Date"], font=font, fill="black")

            # Example: Add Name (larger font size)
            name_x, name_y = placeholder_positions["Name"]
            font_size = 48  # Larger size for Name
            try:
                font = ImageFont.truetype(font_path, font_size)
            except IOError:
                font = ImageFont.load_default()  # Fallback to default if font file is missing
            draw.text((name_x, name_y), student["Name"], font=font, fill="black")

            # Save the image temporarily as a PNG file
            temp_img_path = "temp_certificate.png"
            original_img.save(temp_img_path)

            # Convert to PDF
            pdf_image = Image.open(temp_img_path)
            pdf.image(temp_img_path, x=10, y=10, w=pdf_image.width / 10, h=pdf_image.height / 10)  # Adjust image size to fit PDF
            
            # # Save final PDF for this student in the generated folder
            # pdf_output_path = os.path.join(output_dir, f"{student['Name']}_certificate.pdf")
            # pdf.output(pdf_output_path)

            # # Save final PDF for this student
            # pdf_output_path = f"{student['Name']}_certificate.pdf"
            # pdf.output(pdf_output_path)

            # Clean file name (no special characters)
            safe_name = re.sub(r'[^\w\-_. ]', '', student['Name']).strip()
            pdf_output_path = os.path.join(output_dir, f"{safe_name}_certificate.pdf")
            pdf.output(pdf_output_path)
            generated_count += 1

            # Clean up temp PNG
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
        
        messagebox.showinfo("Done", f"{generated_count} certificate(s) generated successfully!")



if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateApp(root)
    root.mainloop()