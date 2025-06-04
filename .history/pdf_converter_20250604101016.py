import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import PyPDF2
from docx import Document
import arabic_reshaper
from bidi.algorithm import get_display
from PIL import Image, ImageTk
import time

class PDFConverterApp:
    def __init__(self):
        # Set up the main window
        self.window = tk.Tk()
        self.window.title("PDF to Word Converter")
        self.window.geometry("800x600")
        self.window.configure(bg="#2C3E50")
        
        # Configure grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)
        
        # Create main frame
        self.main_frame = tk.Frame(self.window, bg="#2C3E50")
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Create animated background
        self.create_animated_background()
        
        # Title with custom font
        self.title_label = tk.Label(
            self.main_frame,
            text="PDF to Word Converter",
            font=("Helvetica", 32, "bold"),
            fg="#FFFFFF",
            bg="#2C3E50"
        )
        self.title_label.grid(row=0, column=0, pady=(40, 20))
        
        # Description
        self.desc_label = tk.Label(
            self.main_frame,
            text="Convert your PDF to Word with layout and font preservation\nSupports English, Arabic, and Urdu",
            font=("Helvetica", 14),
            fg="#FFFFFF",
            bg="#2C3E50"
        )
        self.desc_label.grid(row=1, column=0, pady=(0, 40))
        
        # File selection frame
        self.file_frame = tk.Frame(self.main_frame, bg="#2C3E50")
        self.file_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)
        
        self.file_label = tk.Label(
            self.file_frame,
            text="No file selected",
            font=("Helvetica", 12),
            fg="#FFFFFF",
            bg="#2C3E50"
        )
        self.file_label.grid(row=0, column=0, padx=10, pady=10)
        
        self.select_button = tk.Button(
            self.file_frame,
            text="Select PDF",
            command=self.select_file,
            font=("Helvetica", 14, "bold"),
            bg="#E74C3C",
            fg="#FFFFFF",
            activebackground="#C0392B",
            activeforeground="#FFFFFF",
            width=15,
            height=2
        )
        self.select_button.grid(row=0, column=1, padx=10, pady=10)
        
        # Convert button
        self.convert_button = tk.Button(
            self.main_frame,
            text="Convert to Word",
            command=self.start_conversion,
            state="disabled",
            font=("Helvetica", 14, "bold"),
            bg="#E74C3C",
            fg="#FFFFFF",
            activebackground="#C0392B",
            activeforeground="#FFFFFF",
            width=15,
            height=2
        )
        self.convert_button.grid(row=3, column=0, pady=20)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient="horizontal",
            length=300,
            mode="determinate"
        )
        self.progress_bar.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        
        # Status label
        self.status_label = tk.Label(
            self.main_frame,
            text="Ready",
            font=("Helvetica", 12),
            fg="#FFFFFF",
            bg="#2C3E50"
        )
        self.status_label.grid(row=5, column=0, pady=10)
        
        self.selected_file = None
        self.animation_running = True
        self.start_background_animation()

    def create_animated_background(self):
        # Create a canvas for the animated background
        self.canvas = tk.Canvas(self.window, highlightthickness=0, bg="#2C3E50")
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Create gradient circles
        self.circles = []
        for _ in range(5):
            x = self.window.winfo_width() * 0.5
            y = self.window.winfo_height() * 0.5
            circle = self.canvas.create_oval(x-100, y-100, x+100, y+100, 
                                          fill="#E74C3C", outline="")
            self.circles.append({"id": circle, "dx": 2, "dy": 2})

    def start_background_animation(self):
        def animate():
            if not self.animation_running:
                return
            
            for circle in self.circles:
                # Move circle
                self.canvas.move(circle["id"], circle["dx"], circle["dy"])
                
                # Get current position
                pos = self.canvas.coords(circle["id"])
                
                # Bounce off walls
                if pos[0] <= 0 or pos[2] >= self.window.winfo_width():
                    circle["dx"] *= -1
                if pos[1] <= 0 or pos[3] >= self.window.winfo_height():
                    circle["dy"] *= -1
                
                # Change color gradually
                color = self.canvas.itemcget(circle["id"], "fill")
                if color == "#E74C3C":
                    self.canvas.itemconfig(circle["id"], fill="#C0392B")
                else:
                    self.canvas.itemconfig(circle["id"], fill="#E74C3C")
            
            self.window.after(50, animate)
        
        animate()

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.selected_file = file_path
            self.file_label.configure(text=os.path.basename(file_path))
            self.convert_button.configure(state="normal")

    def convert_pdf_to_word(self):
        try:
            self.status_label.configure(text="Converting...")
            self.progress_bar["value"] = 20
            
            # Create output filename
            output_file = os.path.splitext(self.selected_file)[0] + ".docx"
            
            # Create a new Word document
            doc = Document()
            
            # Open the PDF file
            with open(self.selected_file, 'rb') as file:
                # Create PDF reader object
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Get number of pages
                num_pages = len(pdf_reader.pages)
                
                # Process each page
                for page_num in range(num_pages):
                    # Update progress
                    progress = 20 + (80 * (page_num + 1) / num_pages)
                    self.progress_bar["value"] = progress
                    
                    # Get the page
                    page = pdf_reader.pages[page_num]
                    
                    # Extract text
                    text = page.extract_text()
                    
                    # Handle Arabic/Urdu text
                    if any(ord(c) > 127 for c in text):
                        text = get_display(arabic_reshaper.reshape(text))
                    
                    # Add text to document
                    doc.add_paragraph(text)
                    
                    # Add page break if not the last page
                    if page_num < num_pages - 1:
                        doc.add_page_break()
            
            # Save the document
            doc.save(output_file)
            
            self.progress_bar["value"] = 100
            self.status_label.configure(text="Conversion completed!")
            messagebox.showinfo("Success", "PDF has been converted to Word successfully!")
            
        except Exception as e:
            self.status_label.configure(text="Error during conversion")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.progress_bar["value"] = 0
            self.convert_button.configure(state="normal")

    def start_conversion(self):
        self.convert_button.configure(state="disabled")
        # Start conversion in a separate thread
        thread = threading.Thread(target=self.convert_pdf_to_word)
        thread.daemon = True
        thread.start()

    def run(self):
        self.window.mainloop()
        self.animation_running = False

if __name__ == "__main__":
    app = PDFConverterApp()
    app.run() 