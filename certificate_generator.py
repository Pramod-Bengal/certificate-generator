import os
import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from PIL import Image, ImageTk
import fitz  # PyMuPDF for PDF preview

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generator")
        
        self.template_path = None
        self.excel_path = None
        self.output_folder = None
        self.name_position = [100, 200]
        self.cert_position = [100, 300]
        self.name_color = "#0000FF"
        self.cert_color = "#FF0000"
        self.name_size = 40
        self.cert_size = 30
        self.selected_field = "Name"
        self.locked = {"Name": False, "Certificate Number": False}
        
        # UI Setup
        tk.Button(root, text="Select Template (PDF)", command=self.load_template).pack()
        tk.Button(root, text="Select Excel File", command=self.load_excel).pack()
        tk.Button(root, text="Select Output Folder", command=self.select_output_folder).pack()
        
        self.canvas = tk.Canvas(root, width=500, height=300, bg="white")
        self.canvas.pack()
        self.canvas.bind("<B1-Motion>", self.move_text)
        
        # Ensure text stays on top
        self.name_text = self.canvas.create_text(self.name_position, text="Name", font=("Arial", self.name_size), fill=self.name_color, tags="name_text")
        self.cert_text = self.canvas.create_text(self.cert_position, text="CERT1234", font=("Arial", self.cert_size), fill=self.cert_color, tags="cert_text")
        
        # Controls
        self.field_selector = tk.StringVar(value="Name")
        tk.Radiobutton(root, text="Name", variable=self.field_selector, value="Name", command=self.set_selected_field).pack()
        tk.Radiobutton(root, text="Certificate Number", variable=self.field_selector, value="Certificate Number", command=self.set_selected_field).pack()
        
        tk.Button(root, text="Lock/Unlock Position", command=self.toggle_lock).pack()
        tk.Button(root, text="Set Font Color", command=self.choose_color).pack()
        
        self.size_slider = tk.Scale(root, from_=10, to=100, orient=tk.HORIZONTAL, command=self.update_font_size)
        self.size_slider.pack()
        
        tk.Button(root, text="Generate Certificates", command=self.generate_certificates).pack()
    
    def load_template(self):
        self.template_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.template_path:
            self.display_template()
    
    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    
    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
    
    def set_selected_field(self):
        self.selected_field = self.field_selector.get()
    
    def move_text(self, event):
        if self.locked[self.selected_field]:
            return
        
        if self.selected_field == "Name":
            self.name_position = [event.x, event.y]
            self.canvas.coords("name_text", event.x, event.y)
            self.canvas.tag_raise("name_text")
        else:
            self.cert_position = [event.x, event.y]
            self.canvas.coords("cert_text", event.x, event.y)
            self.canvas.tag_raise("cert_text")
    
    def update_font_size(self, value):
        size = int(float(value))
        if self.selected_field == "Name":
            self.name_size = size
            self.canvas.itemconfig("name_text", font=("Arial", size))
        else:
            self.cert_size = size
            self.canvas.itemconfig("cert_text", font=("Arial", size))
    
    def choose_color(self):
        color = colorchooser.askcolor(title="Choose Font Color")[1]
        if self.selected_field == "Name":
            self.name_color = color
            self.canvas.itemconfig("name_text", fill=color)
        else:
            self.cert_color = color
            self.canvas.itemconfig("cert_text", fill=color)
    
    def toggle_lock(self):
        self.locked[self.selected_field] = not self.locked[self.selected_field]
        state = "Locked" if self.locked[self.selected_field] else "Unlocked"
        messagebox.showinfo("Position Lock", f"{self.selected_field} position is now {state}")
    
    def display_template(self):
        doc = fitz.open(self.template_path)
        page = doc[0]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((500, 300))
        self.template_img = ImageTk.PhotoImage(img)
        self.canvas.create_image(250, 150, image=self.template_img, tags="template")
        self.canvas.tag_lower("template")
    
    def generate_certificates(self):
        if not all([self.template_path, self.excel_path, self.output_folder]):
            messagebox.showerror("Error", "Please select all required files and folders.")
            return
        
        df = pd.read_excel(self.excel_path)
        pdf_reader = PdfReader(self.template_path)
        
        for idx, row in df.iterrows():
            name = row['Name']
            cert_num = row['Certificate_Number']
            output_pdf = os.path.join(self.output_folder, f"{name}_{cert_num}.pdf")
            pdf_writer = PdfWriter()
            
            page = pdf_reader.pages[0]
            temp_pdf = "temp_overlay.pdf"
            c = canvas.Canvas(temp_pdf, pagesize=(float(page.mediabox[2]), float(page.mediabox[3])))
            
            c.drawString(self.name_position[0], float(page.mediabox[3]) - self.name_position[1], name)
            
            c.setFont("Helvetica-Bold", self.cert_size)
            c.setFillColor(self.cert_color)
            c.drawString(self.cert_position[0], float(page.mediabox[3]) - self.cert_position[1], cert_num)
            c.save()
            
            with open(temp_pdf, "rb") as f:
                temp_pdf_reader = PdfReader(f)
                page.merge_page(temp_pdf_reader.pages[0])
                pdf_writer.add_page(page)
            
            with open(output_pdf, "wb") as f:
                pdf_writer.write(f)
            
            os.remove(temp_pdf)
        
        messagebox.showinfo("Success", "Certificates generated successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateGenerator(root)
    root.mainloop()