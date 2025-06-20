import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
import hashlib
from typing import List, Tuple, Dict, Set
from dataclasses import dataclass
from enum import Enum

class FileStatus(Enum):
    NORMAL = "normal"
    DUPLICATE = "duplicate"

@dataclass
class PDFFile:
    """Represents a PDF file with metadata"""
    path: str
    doc: fitz.Document
    hash: str
    status: FileStatus = FileStatus.NORMAL
    
    @property
    def filename(self) -> str:
        return os.path.basename(self.path)
    
    @property
    def page_count(self) -> int:
        return len(self.doc)
    
    def close(self):
        """Close the PDF document"""
        self.doc.close()

class DraggableListbox(tk.Listbox):
    """Enhanced listbox with drag-and-drop reordering"""
    
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind('<Button-1>', self._on_click)
        self.bind('<B1-Motion>', self._on_drag)
        self.bind('<ButtonRelease-1>', self._on_release)
        self._drag_data = {'y': 0, 'item': None, 'index': None}
        
    def _on_click(self, event):
        index = self.nearest(event.y)
        if index >= 0:
            self._drag_data['y'] = event.y
            self._drag_data['item'] = self.get(index)
            self._drag_data['index'] = index
            
    def _on_drag(self, event):
        if self._drag_data['item']:
            new_index = self.nearest(event.y)
            if new_index != self._drag_data['index']:
                self.delete(self._drag_data['index'])
                self.insert(new_index, self._drag_data['item'])
                self._drag_data['index'] = new_index
                self.selection_clear(0, tk.END)
                self.selection_set(new_index)
                
    def _on_release(self, event):
        if self._drag_data['item']:
            self.event_generate('<<ListboxReordered>>')
        self._drag_data = {'y': 0, 'item': None, 'index': None}

class PDFFileManager:
    """Manages PDF files and duplicate detection"""
    
    def __init__(self):
        self.files: List[PDFFile] = []
        self._hash_counts: Dict[str, int] = {}
    
    def add_file(self, file_path: str) -> Tuple[bool, str]:
        """
        Add a PDF file to the manager
        
        Returns:
            Tuple[bool, str]: (success, error_message or file_hash)
        """
        try:
            # Calculate file hash
            file_hash = self._calculate_hash(file_path)
            
            # Check if file already exists
            if self._file_exists(file_hash):
                return False, f"File already exists: {os.path.basename(file_path)}"
            
            # Open PDF document
            doc = fitz.open(file_path)
            pdf_file = PDFFile(path=file_path, doc=doc, hash=file_hash)
            
            # Add to files list
            self.files.append(pdf_file)
            
            # Update hash counts
            self._hash_counts[file_hash] = self._hash_counts.get(file_hash, 0) + 1
            
            # Update file statuses
            self._update_duplicate_statuses()
            
            return True, file_hash
            
        except Exception as e:
            return False, str(e)
    
    def remove_file(self, index: int) -> bool:
        """Remove file at given index"""
        if 0 <= index < len(self.files):
            pdf_file = self.files.pop(index)
            pdf_file.close()
            
            # Update hash counts
            self._hash_counts[pdf_file.hash] -= 1
            if self._hash_counts[pdf_file.hash] == 0:
                del self._hash_counts[pdf_file.hash]
            
            # Update file statuses
            self._update_duplicate_statuses()
            return True
        return False
    
    def reorder_files(self, new_order: List[str]) -> bool:
        """Reorder files based on filename list"""
        try:
            new_files = []
            for filename in new_order:
                for pdf_file in self.files:
                    if pdf_file.filename == filename:
                        new_files.append(pdf_file)
                        break
            
            if len(new_files) == len(self.files):
                self.files = new_files
                return True
            return False
        except Exception:
            return False
    
    def get_duplicate_files(self) -> List[PDFFile]:
        """Get list of files marked as duplicates"""
        return [f for f in self.files if f.status == FileStatus.DUPLICATE]
    
    def remove_duplicates(self) -> int:
        """Remove duplicate files, keeping first occurrence of each hash"""
        seen_hashes = set()
        files_to_keep = []
        removed_count = 0
        
        for pdf_file in self.files:
            if pdf_file.hash not in seen_hashes:
                seen_hashes.add(pdf_file.hash)
                files_to_keep.append(pdf_file)
            else:
                pdf_file.close()
                removed_count += 1
        
        self.files = files_to_keep
        self._hash_counts = {}
        for pdf_file in self.files:
            self._hash_counts[pdf_file.hash] = self._hash_counts.get(pdf_file.hash, 0) + 1
        
        self._update_duplicate_statuses()
        return removed_count
    
    def _calculate_hash(self, file_path: str) -> str:
        """Calculate MD5 hash of file"""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    
    def _file_exists(self, file_hash: str) -> bool:
        """Check if file with given hash already exists"""
        return file_hash in self._hash_counts
    
    def _update_duplicate_statuses(self):
        """Update status of all files based on hash counts"""
        for pdf_file in self.files:
            if self._hash_counts[pdf_file.hash] > 1:
                pdf_file.status = FileStatus.DUPLICATE
            else:
                pdf_file.status = FileStatus.NORMAL
    
    def close_all(self):
        """Close all PDF documents"""
        for pdf_file in self.files:
            pdf_file.close()

class PDFMergerGUI:
    """Main GUI for PDF merger application"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")
        self.root.geometry("1200x800")
        
        # Initialize components
        self.file_manager = PDFFileManager()
        self.current_preview_index = -1
        self.current_preview_page = 1
        
        # Create GUI
        self._create_gui()
        
    def _create_gui(self):
        """Create the main GUI layout"""
        # Configure grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        
        # Left panel - Controls
        self._create_control_panel(main_frame)
        
        # Right panel - Preview
        self._create_preview_panel(main_frame)
    
    def _create_control_panel(self, parent):
        """Create the left control panel"""
        control_frame = ttk.LabelFrame(parent, text="Controls", padding="5")
        control_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # Add PDF button
        ttk.Button(control_frame, text="Add PDFs", command=self._add_pdfs).grid(
            row=0, column=0, pady=5, sticky="ew")
        
        # File list
        ttk.Label(control_frame, text="PDF Files (drag to reorder):").grid(
            row=1, column=0, pady=(10, 2), sticky="w")
        
        self.files_listbox = DraggableListbox(control_frame, height=20, width=40)
        self.files_listbox.grid(row=2, column=0, pady=5, sticky="ew")
        self.files_listbox.bind('<<ListboxSelect>>', self._on_file_select)
        self.files_listbox.bind('<<ListboxReordered>>', self._on_files_reordered)
        
        # Remove button
        ttk.Button(control_frame, text="Remove Selected", command=self._remove_selected).grid(
            row=3, column=0, pady=5, sticky="ew")
        
        # Remove duplicates button
        self.remove_duplicates_btn = ttk.Button(
            control_frame, text="Remove Duplicates", 
            command=self._remove_duplicates, state="disabled")
        self.remove_duplicates_btn.grid(row=4, column=0, pady=5, sticky="ew")
        
        # Merge button
        ttk.Button(control_frame, text="Merge PDFs", command=self._merge_pdfs).grid(
            row=5, column=0, pady=20, sticky="ew")
    
    def _create_preview_panel(self, parent):
        """Create the right preview panel"""
        preview_frame = ttk.LabelFrame(parent, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(preview_frame, width=600, height=600, bg='white')
        self.preview_canvas.grid(row=0, column=0, columnspan=3, sticky="nsew")
        
        # Navigation frame
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        ttk.Button(nav_frame, text="Previous", command=self._prev_page).grid(row=0, column=0, padx=5)
        self.page_label = ttk.Label(nav_frame, text="Page: 0/0")
        self.page_label.grid(row=0, column=1, padx=5)
        ttk.Button(nav_frame, text="Next", command=self._next_page).grid(row=0, column=2, padx=5)
    
    def _add_pdfs(self):
        """Add PDF files to the list"""
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        
        for path in file_paths:
            success, result = self.file_manager.add_file(path)
            
            if not success:
                if "already exists" in result:
                    # Ask user if they want to add duplicate
                    filename = os.path.basename(path)
                    if messagebox.askyesno("Duplicate File", 
                                         f"File '{filename}' already exists.\n\nAdd it anyway?"):
                        # Force add by temporarily removing the hash check
                        self._force_add_file(path)
                else:
                    messagebox.showerror("Error", f"Failed to load {path}: {result}")
            else:
                self._refresh_file_list()
    
    def _force_add_file(self, path: str):
        """Force add a file even if it's a duplicate"""
        try:
            doc = fitz.open(path)
            file_hash = self.file_manager._calculate_hash(path)
            pdf_file = PDFFile(path=path, doc=doc, hash=file_hash)
            
            self.file_manager.files.append(pdf_file)
            self.file_manager._hash_counts[file_hash] = self.file_manager._hash_counts.get(file_hash, 0) + 1
            self.file_manager._update_duplicate_statuses()
            
            self._refresh_file_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load {path}: {str(e)}")
    
    def _refresh_file_list(self):
        """Refresh the file listbox display"""
        self.files_listbox.delete(0, tk.END)
        
        for i, pdf_file in enumerate(self.file_manager.files):
            display_text = f"{pdf_file.filename} ({pdf_file.page_count} pages)"
            self.files_listbox.insert(tk.END, display_text)
            
            # Color duplicate files
            if pdf_file.status == FileStatus.DUPLICATE:
                self.files_listbox.itemconfig(i, {'bg': '#ffcccc'})
        
        # Update remove duplicates button
        duplicate_count = len(self.file_manager.get_duplicate_files())
        if duplicate_count > 0:
            self.remove_duplicates_btn.config(state="normal")
        else:
            self.remove_duplicates_btn.config(state="disabled")
    
    def _remove_selected(self):
        """Remove the selected file"""
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            if self.file_manager.remove_file(index):
                self._refresh_file_list()
                
                # Update preview if necessary
                if not self.file_manager.files:
                    self._clear_preview()
                elif index == self.current_preview_index:
                    self.current_preview_index = min(index, len(self.file_manager.files) - 1)
                    self.current_preview_page = 1
                    self._update_preview()
    
    def _remove_duplicates(self):
        """Remove all duplicate files"""
        duplicate_count = len(self.file_manager.get_duplicate_files())
        if duplicate_count == 0:
            return
        
        if messagebox.askyesno("Remove Duplicates", 
                              f"Remove {duplicate_count} duplicate files?\n"
                              f"This will keep only the first occurrence of each unique file."):
            removed = self.file_manager.remove_duplicates()
            self._refresh_file_list()
            
            # Update preview
            if self.current_preview_index >= len(self.file_manager.files):
                self.current_preview_index = len(self.file_manager.files) - 1 if self.file_manager.files else -1
            
            if self.file_manager.files:
                self.files_listbox.selection_clear(0, tk.END)
                self.files_listbox.selection_set(self.current_preview_index)
                self.current_preview_page = 1
                self._update_preview()
            else:
                self._clear_preview()
    
    def _on_file_select(self, event):
        """Handle file selection"""
        selection = self.files_listbox.curselection()
        if selection:
            self.current_preview_index = selection[0]
            self.current_preview_page = 1
            self._update_preview()
    
    def _on_files_reordered(self, event):
        """Handle file reordering"""
        new_order = []
        for i in range(self.files_listbox.size()):
            filename = self.files_listbox.get(i).split(" (")[0]
            new_order.append(filename)
        
        if self.file_manager.reorder_files(new_order):
            self._refresh_file_list()
    
    def _update_preview(self):
        """Update the preview display"""
        if (self.current_preview_index < 0 or 
            self.current_preview_index >= len(self.file_manager.files)):
            return
        
        try:
            pdf_file = self.file_manager.files[self.current_preview_index]
            page = pdf_file.doc[self.current_preview_page - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Scale image to fit canvas
            canvas_width = 600
            canvas_height = 600
            scale = min(canvas_width/img.width, canvas_height/img.height)
            new_width = int(img.width * scale)
            new_height = int(img.height * scale)
            
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.preview_image = ImageTk.PhotoImage(img)
            
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(
                canvas_width//2, canvas_height//2, 
                image=self.preview_image, anchor="center")
            
            self.page_label.config(text=f"Page: {self.current_preview_page}/{pdf_file.page_count}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update preview: {str(e)}")
    
    def _clear_preview(self):
        """Clear the preview display"""
        self.current_preview_index = -1
        self.current_preview_page = 1
        self.preview_canvas.delete("all")
        self.page_label.config(text="Page: 0/0")
    
    def _prev_page(self):
        """Go to previous page"""
        if (self.current_preview_index >= 0 and 
            self.current_preview_page > 1):
            self.current_preview_page -= 1
            self._update_preview()
    
    def _next_page(self):
        """Go to next page"""
        if self.current_preview_index >= 0:
            pdf_file = self.file_manager.files[self.current_preview_index]
            if self.current_preview_page < pdf_file.page_count:
                self.current_preview_page += 1
                self._update_preview()
    
    def _merge_pdfs(self):
        """Merge all PDF files"""
        if not self.file_manager.files:
            messagebox.showerror("Error", "No PDF files to merge")
            return
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")])
        
        if output_path:
            try:
                pdf_writer = PdfWriter()
                
                for pdf_file in self.file_manager.files:
                    pdf_reader = PdfReader(pdf_file.path)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
                
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                messagebox.showinfo("Success", "PDFs merged successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def __del__(self):
        """Cleanup when object is destroyed"""
        self.file_manager.close_all()

def main():
    """Main entry point"""
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()