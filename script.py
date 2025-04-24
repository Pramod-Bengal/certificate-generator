import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageFont, ImageDraw
from openpyxl import load_workbook
from fpdf import FPDF
import re
from tkinter import colorchooser
import json
import threading
import sys
import platform



class CertificateApp:
    
    def __init__(self, root):

        self.master = root
        self.root = root
        self.root.title("Certificate Generator")
        self.set_icon()  # Set the icon after initializing the window

        self.original_image = None
        self.display_image = None
        self.scale_x = 1
        self.scale_y = 1
        self.template_path = None  # Add template_path attribute
        self.excel_path = None  # Add excel_path attribute

        self.placeholders = {}  # Store placeholder references

        self.excel_data = []  # Will store list of student dictionaries

        self.include_name = tk.BooleanVar(value=True)
        self.include_id = tk.BooleanVar(value=True)
        self.include_start = tk.BooleanVar(value=True)
        self.include_end = tk.BooleanVar(value=True)


        self.font_settings = {
            "Name": {"size": tk.IntVar(value=48), "color": tk.StringVar(value="#000000")},
            "ID": {"size": tk.IntVar(value=32), "color": tk.StringVar(value="#000000")},
            "Start Date": {"size": tk.IntVar(value=32), "color": tk.StringVar(value="#000000")},
            "End Date": {"size": tk.IntVar(value=32), "color": tk.StringVar(value="#000000")}
        }
        self.setup_ui()
        
    def setup_ui(self):
        self.master.title("Certificate Generator")
        self.master.configure(padx=20, pady=20)

        # ---- Top Navigation Bar ----
        nav_frame = tk.Frame(self.master, bg="#f0f0f0", height=50)
        nav_frame.pack(fill="x", pady=(0, 10))
        
        # Right side of nav (Controls)
        controls_frame = tk.Frame(nav_frame, bg="#f0f0f0")
        controls_frame.pack(side="left", padx=10)

        # File Operations
        file_menu = tk.Menubutton(controls_frame, text="Project", bg="#ffffff", relief="flat")
        file_menu.pack(side="left", padx=5)
        file_menu.menu = tk.Menu(file_menu, tearoff=0)
        file_menu["menu"] = file_menu.menu
        file_menu.menu.add_command(label="Save Project", command=self.save_project)
        file_menu.menu.add_command(label="Load Project", command=self.load_project)

        # ---- Main Content Area ----
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill="both", expand=True)

        # ---- Left Panel (Settings) ----
        settings_frame = tk.Frame(main_frame, width=250)
        settings_frame.pack(side="left", fill="y", padx=(0, 20))

        # ---- Load Data ----

        data_frame = tk.Frame(settings_frame, width=250)
        data_frame.pack(fill="x", padx=(0.10))

        # load template button
        load_template_btn = tk.Button(data_frame, text="Load Template", command=self.load_template, fg="black", bg="grey", relief="flat", padx=10)
        load_template_btn.pack(side="left", padx=5)

        # load execl Button
        load_execl_btn = tk.Button(data_frame, text="Load Excel", command=self.load_excel, fg="black", bg="grey", relief="flat", padx=10)
        load_execl_btn.pack(side="left", padx=5)

        # ---- Placeholder Toggles ----
        toggle_frame = tk.LabelFrame(settings_frame, text="Attributes", padx=10, pady=10)
        toggle_frame.pack(fill="x", pady=(0, 10))
        

        # Create checkbuttons with toggle commands
        name_cb = tk.Checkbutton(toggle_frame, text="Name", variable=self.include_name, 
                               command=lambda: self.toggle_placeholder("Name"))
        name_cb.pack(anchor="w", pady=2)
        
        id_cb = tk.Checkbutton(toggle_frame, text="ID", variable=self.include_id,
                             command=lambda: self.toggle_placeholder("ID"))
        id_cb.pack(anchor="w", pady=2)
        
        start_cb = tk.Checkbutton(toggle_frame, text="Start Date", variable=self.include_start,
                                command=lambda: self.toggle_placeholder("Start Date"))
        start_cb.pack(anchor="w", pady=2)
        
        end_cb = tk.Checkbutton(toggle_frame, text="End Date", variable=self.include_end,
                              command=lambda: self.toggle_placeholder("End Date"))
        end_cb.pack(anchor="w", pady=2)

        # ---- Font Settings ----
        font_frame = tk.LabelFrame(settings_frame, text="Font Settings", padx=10, pady=10)
        font_frame.pack(fill="x", pady=(10, 10))

        for i, field in enumerate(self.font_settings):
            field_frame = tk.Frame(font_frame)
            field_frame.pack(fill="x", pady=2)
            
            tk.Label(field_frame, text=field, width=10).pack(side="left")
            
            size_frame = tk.Frame(field_frame)
            size_frame.pack(side="left", padx=5)
            tk.Label(size_frame, text="Size:").pack(side="left")
            tk.Spinbox(size_frame, from_=10, to=100, textvariable=self.font_settings[field]["size"], 
                      width=5).pack(side="left", padx=2)
            
            color_btn = tk.Button(field_frame, text="Color", command=lambda f=field: self.choose_color(f),
                                relief="flat", bg="#f0f0f0")
            color_btn.pack(side="right")

        # ---- Action buttons ----

        action_frame = tk.LabelFrame(settings_frame, padx=10, pady=10)
        action_frame.pack(fill="x", pady=(10, 10))

         # Preview Button
        preview_btn = tk.Button(action_frame, text="Preview", command=self.preview_certificate, 
                              bg="#4CAF50", fg="white", relief="flat", padx=10)
        preview_btn.pack(side="left", padx=5)

        # Generate Button
        generate_btn = tk.Button(action_frame, text="Generate", command=self.generate_certificates,
                               bg="#2196F3", fg="white", relief="flat", padx=10)
        generate_btn.pack(side="left", padx=5)

        # ---- Center Canvas Area ----
        center_panel = tk.Frame(main_frame)
        center_panel.pack(side="left", fill="both", expand=True)

        self.canvas_frame = tk.Frame(center_panel, relief="sunken", borderwidth=2)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="white")
        self.canvas.pack(fill="both", expand=True)

        # ---- Progress Bar ----
        progress_frame = tk.Frame(self.master)
        progress_frame.pack(fill="x", pady=20)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(side="left", padx=10, expand=True)

    def set_icon(self):
        """Set the application icon based on the operating system."""
        try:
            if getattr(sys, 'frozen', False):
                # Running as a bundle (PyInstaller)
                base_path = sys._MEIPASS
            else:
                # Running as a script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            # Try different icon formats
            icon_paths = [
                os.path.join(base_path, 'icon.ico'),
                os.path.join(base_path, 'logo.png'),
                os.path.join(base_path, 'icon.png')
            ]
            
            icon_set = False
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    try:
                        if platform.system() == 'Windows':
                            self.root.iconbitmap(icon_path)
                            icon_set = True
                            break
                        elif platform.system() == 'Linux':
                            img = Image.open(icon_path)
                            photo = ImageTk.PhotoImage(img)
                            self.root.tk.call('wm', 'iconphoto', self.root._w, photo)
                            icon_set = True
                            break
                        elif platform.system() == 'Darwin':  # macOS
                            self.root.iconbitmap(icon_path)
                            icon_set = True
                            break
                    except Exception as e:
                        print(f"Error setting icon from {icon_path}: {e}")
                        continue
            
            if not icon_set:
                print("No suitable icon file found or could not be loaded")
        except Exception as e:
            print(f"Error setting icon: {e}")

    def choose_color(self, field):
        color = colorchooser.askcolor(title=f"Choose color for {field}")
        if color[1]:
            self.font_settings[field]["color"].set(color[1])

    def hex_to_rgb(self, hex_color):
        # Get the string value from StringVar if it's a StringVar
        if isinstance(hex_color, tk.StringVar):
            hex_color = hex_color.get()
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def load_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png")])
        if not file_path:
            return

        self.template_path = file_path  # Store the template path
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
        def render_placeholder():
            try:
                # Get font size with proper validation
                try:
                    font_size = int(self.font_settings[label]["size"].get())
                    if font_size < 1:
                        font_size = 32  # Default size if invalid
                except (ValueError, TypeError, ttk.TclError):
                    font_size = 32  # Default size if conversion fails

                # Get color with validation
                try:
                    color = self.font_settings[label]["color"].get()
                    if not color:
                        color = "#000000"  # Default color if empty
                except (ValueError, TypeError, ttk.TclError):
                    color = "#000000"  # Default color if conversion fails

                sample_value = {
                    "Name": "John Doe",
                    "ID": "ID12345",
                    "Start Date": "01-01-2024",
                    "End Date": "01-06-2024"
                }.get(label, label)
            
                font_path = "arial.ttf"
                try:
                    scaled_font_size = max(10, int(font_size / self.scale_y))
                    font = ImageFont.truetype(font_path, scaled_font_size)
                except IOError:
                    font = ImageFont.load_default()
            
                img = Image.new("RGBA", (500, 100), (255, 255, 255, 0))
                draw = ImageDraw.Draw(img)
                draw.text((0, 0), sample_value, font=font, fill=self.hex_to_rgb(color))
                bbox = img.getbbox()
                cropped_img = img.crop(bbox)
            
                return ImageTk.PhotoImage(cropped_img)
            except Exception as e:
                print(f"Error rendering placeholder: {e}")
                return None


        # Initial render
        preview_img = render_placeholder()
        img_label = tk.Label(self.canvas, image=preview_img, bg="white")
        img_label.image = preview_img  # keep a reference
        item = self.canvas.create_window(50, 50, window=img_label, anchor="nw")

        def start_drag(event, canvas_item=item):
            self._drag_data = {
                "item": canvas_item,
                "x": self.canvas.canvasx(event.x_root - self.canvas.winfo_rootx()),
                "y": self.canvas.canvasy(event.y_root - self.canvas.winfo_rooty())
            }

        def do_drag(event):
            new_x = self.canvas.canvasx(event.x_root - self.canvas.winfo_rootx())
            new_y = self.canvas.canvasy(event.y_root - self.canvas.winfo_rooty())
            dx = new_x - self._drag_data["x"]
            dy = new_y - self._drag_data["y"]
            self.canvas.move(self._drag_data["item"], dx, dy)
            self._drag_data["x"] = new_x
            self._drag_data["y"] = new_y


        img_label.bind("<Button-1>", start_drag)
        img_label.bind("<B1-Motion>", do_drag)

        # Update preview on font size change
        def update_preview(*args):
            new_img = render_placeholder()
            img_label.configure(image=new_img)
            img_label.image = new_img

        self.font_settings[label]["size"].trace("w", update_preview)
        self.font_settings[label]["color"].trace("w", update_preview)

        self.placeholders[label] = item

    def toggle_placeholder(self, label):
        """Show or hide a placeholder based on its toggle state."""
        if label in self.placeholders:
            var = {
                "Name": self.include_name,
                "ID": self.include_id,
                "Start Date": self.include_start,
                "End Date": self.include_end
            }[label]
            
            if var.get():
                self.canvas.itemconfig(self.placeholders[label], state="normal")
            else:
                self.canvas.itemconfig(self.placeholders[label], state="hidden")

    def get_placeholder_positions(self):
        """Get scaled coordinates for actual certificate."""
        coords = {}
        for label, item in self.placeholders.items():
            x, y = self.canvas.coords(item)
            coords[label] = (x * self.scale_x, y * self.scale_y)
        return coords
    
    def load_excel(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        self.excel_path = file_path

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
    
    def preview_certificate(self):
        if not self.original_image:
            messagebox.showwarning("Warning", "Template not loaded!")
            return

        if not self.excel_data:
            messagebox.showwarning("Warning", "No student data loaded!")
            return

        # Automatically select the first student
        student = self.excel_data[0]  # Always preview the first student

        # Create a copy of the image to work on
        img = self.original_image.copy()
        draw = ImageDraw.Draw(img)
        placeholder_positions = self.get_placeholder_positions()
        font_path = "arial.ttf"  # Make sure this font file is available

        for field, include_var in [
            ("Name", self.include_name),
            ("ID", self.include_id),
            ("Start Date", self.include_start),
            ("End Date", self.include_end)
        ]:
            if include_var.get():
                x, y = placeholder_positions[field]
                size = self.font_settings[field]["size"].get()
                color = self.font_settings[field]["color"].get()
                try:
                    font = ImageFont.truetype(font_path, size)
                except IOError:
                    font = ImageFont.load_default()
                
                # Calculate the width of the text to adjust the position
                text_width = draw.textlength(student[field], font=font)

                # If it's the Name field, adjust the x to center the text
                if field == "Name":
                    x = (img.width - text_width) // 2  # Center the text horizontally

                # Apply the text to the image
                draw.text((x, y), student[field], font=font, fill=self.hex_to_rgb(color))

        # Show preview in a new window
        preview_win = tk.Toplevel(self.root)
        preview_win.title("Certificate Preview")

        # Resize image for display
        preview_img = img.resize((900, int(img.height * (900 / img.width))), Image.LANCZOS)
        preview_photo = ImageTk.PhotoImage(preview_img)

        # Label to display the image in the preview window
        label = tk.Label(preview_win, image=preview_photo)
        label.image = preview_photo  # Keep a reference to the image
        label.pack()

        preview_win.mainloop()  # Start the Tkinter event loop for the preview window


    def generate_certificates(self):
        if not self.excel_data:
            messagebox.showwarning("Warning", "No student data loaded!")
            return
    
        if not self.original_image:
            messagebox.showwarning("Warning", "No template image loaded!")
            return
    
        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
            return
    
        font_path = "arial.ttf"
        generated_count = 0
    
        # Convert px to mm (1 px = 0.264583 mm)
        def px_to_mm(px): return px * 0.264583
    
        img_width_px, img_height_px = self.original_image.size
        pdf_width = px_to_mm(img_width_px)
        pdf_height = px_to_mm(img_height_px)
    
        # Update the progress bar
        def update_progressbar(current, total):
            progress = (current / total) * 100
            self.progress['value'] = progress
            self.master.update_idletasks()
    
        # Certificate generation function
        def generate_certificates_in_thread():
            nonlocal generated_count
            total_students = len(self.excel_data)
    
            for i, student in enumerate(self.excel_data):
                pdf = FPDF(unit="mm", format=(pdf_width, pdf_height))
            pdf.add_page()
    
            original_img = self.original_image.copy()
            draw = ImageDraw.Draw(original_img)

            placeholder_positions = self.get_placeholder_positions()
        
                # Add text fields
            for field, include_var in [
                    ("Name", self.include_name),
                    ("ID", self.include_id),
                    ("Start Date", self.include_start),
                    ("End Date", self.include_end)
                ]:
                    if include_var.get():
                        x, y = placeholder_positions[field]
                        size = self.font_settings[field]["size"].get()
                        color = self.font_settings[field]["color"].get()
                        try:
                            font = ImageFont.truetype(font_path, size)
                        except IOError:
                            font = ImageFont.load_default()
                        
                        # Calculate the width of the text to adjust the position
                        text_width = draw.textlength(student[field], font=font)
    
                        # If it's the Name field, adjust the x to center the text
                        if field == "Name":
                            x = (original_img.width - text_width) // 2  # Center the text horizontally
    
                        # Apply the text to the image
                        draw.text((x, y), student[field], font=font, fill=self.hex_to_rgb(color))
    
            temp_img_path = "temp_certificate.png"
            original_img.save(temp_img_path)

            pdf.image(temp_img_path, x=0, y=0, w=pdf_width, h=pdf_height)

            safe_name = re.sub(r'[^\w\-_. ]', '', student['Name']).strip()
            pdf_output_path = os.path.join(output_dir, f"{safe_name}_certificate.pdf")
            pdf.output(pdf_output_path)
            generated_count += 1
    
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
    
                # Update progress after each certificate is generated
                update_progressbar(i + 1, total_students)
    
            messagebox.showinfo("Done", f"{generated_count} certificate(s) generated successfully!")
    
        # Start the certificate generation in a separate thread
        threading.Thread(target=generate_certificates_in_thread, daemon=True).start()

    def save_project(self):
        if not self.original_image:
            messagebox.showwarning("Warning", "No template loaded!")
            return
    
        try:
            # Ensure template_path is set
            if not hasattr(self, 'template_path') or not self.template_path:
                messagebox.showwarning("Warning", "Template path not set!")
                return

            project_data = {
                "template_path": self.template_path,
                "positions": self.get_placeholder_positions(),
                "font_settings": {
                    field: {
                        "size": self.font_settings[field]["size"].get(),
                        "color": self.font_settings[field]["color"].get()
                    } for field in self.font_settings
                },
                "attributes": {
                    "Name": self.include_name.get(),
                    "ID": self.include_id.get(),
                    "Start Date": self.include_start.get(),
                    "End Date": self.include_end.get()
                },
                "excel_path": self.excel_path if hasattr(self, 'excel_path') else None
            }
    
            file_path = filedialog.asksaveasfilename(defaultextension=".certproj", filetypes=[("Certificate Project", "*.certproj")])
            if file_path:
                with open(file_path, "w") as f:
                    json.dump(project_data, f, indent=4)
                messagebox.showinfo("Saved", "Project saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save project: {e}")
    
    def load_project(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Certificate Project", "*.certproj")])
            if not file_path:
                return

            with open(file_path, "r") as f:
                project_data = json.load(f)

            # Clear existing UI elements
            self.canvas.delete("all")
            self.placeholders.clear()

            # Load template first
            if project_data.get("template_path") and os.path.exists(project_data["template_path"]):
                self.template_path = project_data["template_path"]
                self.original_image = Image.open(self.template_path)
                original_width, original_height = self.original_image.size

                # Calculate scaling
                max_width, max_height = 1000, 700
                ratio = min(max_width / original_width, max_height / original_height)
                new_size = (int(original_width * ratio), int(original_height * ratio))

                self.scale_x = original_width / new_size[0]
                self.scale_y = original_height / new_size[1]

                # Resize and display template
                resized_img = self.original_image.resize(new_size)
                self.display_image = ImageTk.PhotoImage(resized_img)
                self.canvas.config(width=new_size[0], height=new_size[1])
                self.canvas.create_image(0, 0, image=self.display_image, anchor="nw")

            # Load font settings
            if "font_settings" in project_data:
                for field in project_data["font_settings"]:
                    if field in self.font_settings:
                        self.font_settings[field]["size"].set(project_data["font_settings"][field]["size"])
                        self.font_settings[field]["color"].set(project_data["font_settings"][field]["color"])

            # Load attributes
            if "attributes" in project_data:
                self.include_name.set(project_data["attributes"].get("Name", True))
                self.include_id.set(project_data["attributes"].get("ID", True))
                self.include_start.set(project_data["attributes"].get("Start Date", True))
                self.include_end.set(project_data["attributes"].get("End Date", True))

            # Create placeholders with updated settings
            for label in ["Name", "ID", "Start Date", "End Date"]:
                self.create_placeholder(label)

            # Load positions after placeholders are created
            if "positions" in project_data:
                for label, (x, y) in project_data["positions"].items():
                    if label in self.placeholders:
                        self.canvas.coords(self.placeholders[label], x / self.scale_x, y / self.scale_y)

            # Load Excel data last
            if project_data.get("excel_path") and os.path.exists(project_data["excel_path"]):
                self.excel_path = project_data["excel_path"]
                self.load_excel(self.excel_path)

            messagebox.showinfo("Loaded", "Project loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project: {e}")



if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateApp(root)
    root.mainloop()