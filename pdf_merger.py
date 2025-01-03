import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os

class DraggableListbox(tk.Listbox):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind('<Button-1>', self.on_click)
        self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', self.on_release)
        self.drag_data = {'y': 0, 'item': None, 'index': None}
        
    def on_click(self, event):
        index = self.nearest(event.y)
        if index >= 0:
            self.drag_data['y'] = event.y
            self.drag_data['item'] = self.get(index)
            self.drag_data['index'] = index
            
    def on_drag(self, event):
        if self.drag_data['item']:
            new_index = self.nearest(event.y)
            if new_index != self.drag_data['index']:
                self.delete(self.drag_data['index'])
                self.insert(new_index, self.drag_data['item'])
                self.drag_data['index'] = new_index
                self.selection_clear(0, tk.END)
                self.selection_set(new_index)
                
    def on_release(self, event):
        if self.drag_data['item']:
            self.event_generate('<<ListboxReordered>>')
        self.drag_data = {'y': 0, 'item': None, 'index': None}

class PDFMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")
        self.root.geometry("1200x800")
        
        self.pdf_files = []  # List of (path, doc) tuples
        self.current_preview_page = 1
        self.current_pdf_index = -1
        self.create_gui()
        
    def create_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left side - Controls
        controls_frame = ttk.LabelFrame(main_frame, text="Controls", padding="5")
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        # Add PDF button
        ttk.Button(controls_frame, text="Add PDFs", command=self.add_pdfs).grid(row=0, column=0, pady=5)
        
        # PDF List
        ttk.Label(controls_frame, text="PDF Files (drag to reorder):").grid(row=1, column=0, pady=2)
        self.files_listbox = DraggableListbox(controls_frame, height=20, width=40)
        self.files_listbox.grid(row=2, column=0, pady=5)
        self.files_listbox.bind('<<ListboxSelect>>', self.on_select_pdf)
        self.files_listbox.bind('<<ListboxReordered>>', self.on_files_reordered)
        
        # Remove PDF button
        ttk.Button(controls_frame, text="Remove Selected PDF", command=self.remove_pdf).grid(row=3, column=0, pady=5)
        
        # Merge button
        ttk.Button(controls_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=4, column=0, pady=20)
        
        # Right side - Preview
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(preview_frame, width=600, height=600, bg='white')
        self.preview_canvas.grid(row=0, column=0, columnspan=3)
        
        # Navigation
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        ttk.Button(nav_frame, text="Previous", command=self.prev_page).grid(row=0, column=0, padx=5)
        self.page_label = ttk.Label(nav_frame, text="Page: 0/0")
        self.page_label.grid(row=0, column=1, padx=5)
        ttk.Button(nav_frame, text="Next", command=self.next_page).grid(row=0, column=2, padx=5)

    def add_pdfs(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for path in file_paths:
            try:
                doc = fitz.open(path)
                filename = os.path.basename(path)
                self.pdf_files.append((path, doc))
                self.files_listbox.insert(tk.END, f"{filename} ({len(doc)} pages)")
                
                # Select first PDF if this is the first one added
                if len(self.pdf_files) == 1:
                    self.files_listbox.selection_set(0)
                    self.current_pdf_index = 0
                    self.current_preview_page = 1
                    self.update_preview()
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load {path}: {str(e)}")

    def on_files_reordered(self, event):
        new_pdf_files = []
        for i in range(self.files_listbox.size()):
            filename = self.files_listbox.get(i).split(" (")[0]
            for path, doc in self.pdf_files:
                if os.path.basename(path) == filename:
                    new_pdf_files.append((path, doc))
                    break
        self.pdf_files = new_pdf_files

    def on_select_pdf(self, event):
        selection = self.files_listbox.curselection()
        if selection:
            self.current_pdf_index = selection[0]
            self.current_preview_page = 1
            self.update_preview()

    def remove_pdf(self):
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.files_listbox.delete(index)
            path, doc = self.pdf_files.pop(index)
            doc.close()
            
            # Update preview if necessary
            if not self.pdf_files:
                self.current_pdf_index = -1
                self.current_preview_page = 1
                self.preview_canvas.delete("all")
                self.page_label.config(text="Page: 0/0")
            elif index == self.current_pdf_index:
                self.current_pdf_index = min(index, len(self.pdf_files) - 1)
                self.current_preview_page = 1
                self.update_preview()

    def update_preview(self):
        if self.current_pdf_index < 0 or not self.pdf_files:
            return
            
        try:
            _, doc = self.pdf_files[self.current_pdf_index]
            page = doc[self.current_preview_page - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            canvas_width = 600
            canvas_height = 600
            
            scale = min(canvas_width/img.width, canvas_height/img.height)
            new_width = int(img.width * scale)
            new_height = int(img.height * scale)
            
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.preview_image = ImageTk.PhotoImage(img)
            
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(canvas_width//2, canvas_height//2, 
                                          image=self.preview_image, anchor="center")
            
            total_pages = len(doc)
            self.page_label.config(text=f"Page: {self.current_preview_page}/{total_pages}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update preview: {str(e)}")

    def prev_page(self):
        if self.current_pdf_index >= 0 and self.current_preview_page > 1:
            self.current_preview_page -= 1
            self.update_preview()

    def next_page(self):
        if self.current_pdf_index >= 0:
            _, doc = self.pdf_files[self.current_pdf_index]
            if self.current_preview_page < len(doc):
                self.current_preview_page += 1
                self.update_preview()

    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF files to merge")
            return
            
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                  filetypes=[("PDF files", "*.pdf")])
        if output_path:
            try:
                pdf_writer = PdfWriter()
                
                for path, _ in self.pdf_files:
                    pdf_reader = PdfReader(path)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
                
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                    
                messagebox.showinfo("Success", "PDFs merged successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def __del__(self):
        # Clean up by closing all PDF documents
        for _, doc in self.pdf_files:
            doc.close()

def main():
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()