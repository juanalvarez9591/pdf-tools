import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image, ImageTk

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

class PDFSplitterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter and Merger")
        self.root.geometry("1200x800")
        
        self.current_pdf_path = None
        self.page_ranges = []
        self.total_pages = 0
        self.current_preview_page = 1
        self.doc = None
        self.goto_page_var = tk.StringVar()
        
        self.create_gui()
        
    def create_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        controls_frame = ttk.LabelFrame(main_frame, text="Controls", padding="5")
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        ttk.Button(controls_frame, text="Load PDF", command=self.load_pdf).grid(row=0, column=0, pady=5)
        
        range_frame = ttk.Frame(controls_frame)
        range_frame.grid(row=1, column=0, pady=5)
        
        ttk.Label(range_frame, text="From page:").grid(row=0, column=0)
        self.start_page_var = tk.StringVar()
        self.start_page_entry = ttk.Entry(range_frame, textvariable=self.start_page_var, width=5)
        self.start_page_entry.grid(row=0, column=1, padx=2)
        
        ttk.Label(range_frame, text="To page:").grid(row=0, column=2)
        self.end_page_var = tk.StringVar()
        self.end_page_entry = ttk.Entry(range_frame, textvariable=self.end_page_var, width=5)
        self.end_page_entry.grid(row=0, column=3, padx=2)
        
        ttk.Button(controls_frame, text="Add Range", command=self.add_range).grid(row=2, column=0, pady=5)
        
        ttk.Label(controls_frame, text="Selected Ranges (drag to reorder):").grid(row=3, column=0, pady=2)
        self.ranges_listbox = DraggableListbox(controls_frame, height=10, width=30)
        self.ranges_listbox.grid(row=4, column=0, pady=5)
        self.ranges_listbox.bind('<<ListboxReordered>>', self.on_ranges_reordered)
        
        ttk.Button(controls_frame, text="Remove Selected Range", command=self.remove_range).grid(row=5, column=0, pady=5)
        
        ttk.Button(controls_frame, text="Export PDF", command=self.export_pdf).grid(row=6, column=0, pady=20)
        
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        self.preview_canvas = tk.Canvas(preview_frame, width=600, height=600, bg='white')
        self.preview_canvas.grid(row=0, column=0, columnspan=4)
        
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.grid(row=1, column=0, columnspan=4, pady=5)
        
        ttk.Button(nav_frame, text="Previous", command=self.prev_page).grid(row=0, column=0, padx=5)
        self.page_label = ttk.Label(nav_frame, text="Page: 0/0")
        self.page_label.grid(row=0, column=1, padx=5)
        ttk.Button(nav_frame, text="Next", command=self.next_page).grid(row=0, column=2, padx=5)
        
        ttk.Label(nav_frame, text="Go to page:").grid(row=0, column=3, padx=5)
        goto_entry = ttk.Entry(nav_frame, textvariable=self.goto_page_var, width=5)
        goto_entry.grid(row=0, column=4, padx=2)
        ttk.Button(nav_frame, text="Go", command=self.goto_page).grid(row=0, column=5, padx=5)
        
        goto_entry.bind('<Return>', lambda e: self.goto_page())

    def on_ranges_reordered(self, event):
        new_ranges = []
        for i in range(self.ranges_listbox.size()):
            item = self.ranges_listbox.get(i)
            parts = item.split()
            start = int(parts[1])
            end = int(parts[3])
            new_ranges.append((start, end))
        self.page_ranges = new_ranges
        
    def goto_page(self):
        try:
            page_num = int(self.goto_page_var.get())
            if 1 <= page_num <= self.total_pages:
                self.current_preview_page = page_num
                self.update_preview()
            else:
                messagebox.showerror("Error", f"Please enter a page number between 1 and {self.total_pages}")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid page number")
        self.goto_page_var.set("")
        
    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            try:
                self.current_pdf_path = file_path
                self.doc = fitz.open(file_path)
                self.total_pages = len(self.doc)
                self.current_preview_page = 1
                self.page_ranges = []
                self.ranges_listbox.delete(0, tk.END)
                self.update_preview()
                messagebox.showinfo("Success", f"Loaded PDF with {self.total_pages} pages")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
            
    def update_preview(self):
        if not self.doc:
            return
            
        try:
            page = self.doc[self.current_preview_page - 1]
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
            
            self.page_label.config(text=f"Page: {self.current_preview_page}/{self.total_pages}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update preview: {str(e)}")
            
    def prev_page(self):
        if self.current_preview_page > 1:
            self.current_preview_page -= 1
            self.update_preview()
            
    def next_page(self):
        if self.current_preview_page < self.total_pages:
            self.current_preview_page += 1
            self.update_preview()
            
    def add_range(self):
        try:
            start = int(self.start_page_var.get())
            end = int(self.end_page_var.get())
            
            if start < 1 or end > self.total_pages or start > end:
                raise ValueError
                
            range_str = f"Pages {start} to {end}"
            self.page_ranges.append((start, end))
            self.ranges_listbox.insert(tk.END, range_str)
            
            self.start_page_var.set("")
            self.end_page_var.set("")
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid page numbers")
            
    def remove_range(self):
        selection = self.ranges_listbox.curselection()
        if selection:
            index = selection[0]
            self.ranges_listbox.delete(index)
            self.page_ranges.pop(index)
            
    def export_pdf(self):
        if not self.page_ranges:
            messagebox.showerror("Error", "No page ranges selected")
            return
            
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                  filetypes=[("PDF files", "*.pdf")])
        if output_path:
            try:
                pdf_reader = PdfReader(self.current_pdf_path)
                pdf_writer = PdfWriter()
                
                for start, end in self.page_ranges:
                    for page_num in range(start - 1, end):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                    
                messagebox.showinfo("Success", "PDF exported successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    root = tk.Tk()
    app = PDFSplitterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()