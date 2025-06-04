import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from pdf2docx import Converter
import os
import arabic_reshaper
from bidi.algorithm import get_display
import threading

class PDFConverterApp:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("PDF to Word Converter")
        self.window.geometry("600x400")
        self.window.configure(fg_color="#f0f0f0")

        # Configure grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        # Create main frame
        self.main_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="PDF to Word Converter",
            font=("Helvetica", 24, "bold")
        )
        self.title_label.grid(row=0, column=0, pady=20)

        # File selection frame
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)

        self.file_label = ctk.CTkLabel(
            self.file_frame,
            text="No file selected",
            font=("Helvetica", 12)
        )
        self.file_label.grid(row=0, column=0, padx=10, pady=10)

        self.select_button = ctk.CTkButton(
            self.file_frame,
            text="Select PDF",
            command=self.select_file
        )
        self.select_button.grid(row=0, column=1, padx=10, pady=10)

        # Convert button
        self.convert_button = ctk.CTkButton(
            self.main_frame,
            text="Convert to Word",
            command=self.start_conversion,
            state="disabled"
        )
        self.convert_button.grid(row=2, column=0, pady=20)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)

        # Status label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="Ready",
            font=("Helvetica", 12)
        )
        self.status_label.grid(row=4, column=0, pady=10)

        self.selected_file = None

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
            self.progress_bar.set(0.2)

            # Create output filename
            output_file = os.path.splitext(self.selected_file)[0] + ".docx"

            # Convert PDF to Word
            cv = Converter(self.selected_file)
            cv.convert(output_file, start=0, end=None)
            cv.close()

            self.progress_bar.set(1.0)
            self.status_label.configure(text="Conversion completed!")
            messagebox.showinfo("Success", "PDF has been converted to Word successfully!")
            
        except Exception as e:
            self.status_label.configure(text="Error during conversion")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.progress_bar.set(0)
            self.convert_button.configure(state="normal")

    def start_conversion(self):
        self.convert_button.configure(state="disabled")
        # Start conversion in a separate thread
        thread = threading.Thread(target=self.convert_pdf_to_word)
        thread.daemon = True
        thread.start()

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = PDFConverterApp()
    app.run() 