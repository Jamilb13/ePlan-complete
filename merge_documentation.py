import os
import sys
import subprocess
import shutil
import time
import logging
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pypdf
import win32com.client
import ezdxf
from ezdxf.addons.drawing import Frontend, RenderContext, pymupdf, layout
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy

# Configuration
REPLACE_PLACEHOLDER = True  # If True, replace the placeholder page containing the link. If False, insert after it.

# Logger setup
logger = logging.getLogger("merge_doc")
logger.setLevel(logging.INFO)

def get_base_path():
    """Resolves the directory containing the running script or compiled executable."""
    if getattr(sys, 'frozen', False):
        # Running as compiled PyInstaller executable
        return os.path.dirname(sys.executable)
    else:
        # Running as standard Python script
        return os.path.dirname(os.path.abspath(__file__))

# Cache of already converted files to avoid redundant COM calls
# Maps external_file_path -> converted_pdf_path
CONVERSION_CACHE = {}

def get_office_pdf_path(src_path):
    """Generates target PDF path in the same folder."""
    base, _ = os.path.splitext(src_path)
    return base + "_converted.pdf"

def convert_docx_to_pdf(docx_path):
    """Converts a Word document to PDF using MS Word COM interface."""
    docx_path = os.path.abspath(docx_path)
    pdf_path = get_office_pdf_path(docx_path)
    
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for Word file: {pdf_path}")
        return pdf_path

    logger.info(f"Converting Word document: {docx_path} -> {pdf_path}")
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        doc.SaveAs(pdf_path, FileFormat=17) # wdFormatPDF = 17
        logger.info("Word conversion successful.")
        return pdf_path
    except Exception as e:
        logger.error(f"Failed to convert Word document {docx_path}: {e}")
        return None
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass

def convert_xlsx_to_pdf(xlsx_path):
    """Converts an Excel sheet to PDF using MS Excel COM interface."""
    xlsx_path = os.path.abspath(xlsx_path)
    pdf_path = get_office_pdf_path(xlsx_path)
    
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for Excel file: {pdf_path}")
        return pdf_path

    logger.info(f"Converting Excel sheet: {xlsx_path} -> {pdf_path}")
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(xlsx_path, ReadOnly=True)
        wb.ExportAsFixedFormat(0, pdf_path) # xlTypePDF = 0
        logger.info("Excel conversion successful.")
        return pdf_path
    except Exception as e:
        logger.error(f"Failed to convert Excel sheet {xlsx_path}: {e}")
        return None
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

def convert_dwg_to_pdf(dwg_path):
    """Converts a DWG drawing to DXF via ODAFileConverter and then renders it to PDF using ezdxf + PyMuPDF."""
    dwg_path = os.path.abspath(dwg_path)
    base, _ = os.path.splitext(dwg_path)
    
    # 1. Fallback: Check if a pre-converted PDF exists
    pre_converted_pdf = base + ".pdf"
    if os.path.exists(pre_converted_pdf):
        logger.info(f"Found pre-converted PDF for DWG file: {pre_converted_pdf}")
        return pre_converted_pdf

    # Check if a previously converted temp file exists
    pdf_path = base + "_converted.pdf"
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for DWG file: {pdf_path}")
        return pdf_path

    # Define intermediate DXF path
    dxf_path = base + ".dxf"

    # Find ODAFileConverter (check extracted temp directory first for portable packaging, then next to script/exe)
    oda_converter = None
    if getattr(sys, 'frozen', False):
        # PyInstaller extracts bundled files to sys._MEIPASS
        temp_path = os.path.join(sys._MEIPASS, "bin", "ODAFileConverter.exe")
        if os.path.exists(temp_path):
            oda_converter = temp_path
            
    if not oda_converter:
        oda_converter = os.path.join(get_base_path(), "bin", "ODAFileConverter.exe")
        
    if not os.path.exists(oda_converter):
        logger.error(f"ODAFileConverter.exe not found at: {oda_converter}")
        return None

    logger.info(f"Attempting to convert DWG drawing to DXF: {dwg_path}")
    
    # Use temp directory to avoid unicode directory path encoding issues in ODA
    with tempfile.TemporaryDirectory(prefix="oda_conv_") as tmp_dir:
        temp_dwg = os.path.join(tmp_dir, "drawing.dwg")
        shutil.copy2(dwg_path, temp_dwg)
        
        # ODA command line arguments
        cmd = [
            oda_converter,
            tmp_dir,
            tmp_dir,
            "ACAD2018",
            "DXF",
            "0",
            "0",
            "drawing.dwg"
        ]
        
        try:
            logger.info("Running ODAFileConverter...")
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if res.returncode != 0:
                logger.warning(f"ODAFileConverter returned exit code {res.returncode}. Stderr: {res.stderr}")
        except subprocess.TimeoutExpired:
            logger.error("ODAFileConverter conversion timed out (30s).")
            return None
        except Exception as e:
            logger.error(f"Error running ODAFileConverter: {e}")
            return None

        temp_dxf = os.path.join(tmp_dir, "drawing.dxf")
        if not os.path.exists(temp_dxf):
            logger.error("ODAFileConverter did not generate DXF file.")
            return None

        logger.info("Saving generated DXF file to target folder...")
        shutil.copy2(temp_dxf, dxf_path)
        logger.info(f"DXF saved: {dxf_path}")

        # Render DXF to PDF in temp folder using ezdxf + pymupdf
        temp_pdf = os.path.join(tmp_dir, "drawing.pdf")
        try:
            logger.info("Rendering DXF to PDF via Python (ezdxf + PyMuPDF)...")
            doc = ezdxf.readfile(temp_dxf)
            
            # Find non-empty paperspace layouts
            paperspace_layouts = []
            for name in doc.layout_names_in_taborder():
                lay = doc.layout(name)
                if lay.is_any_paperspace:
                    # Check if layout contains non-viewport graphical entities OR a viewport looking into modelspace (ID > 1)
                    non_vps = [e for e in lay if e.dxftype() != 'VIEWPORT']
                    viewports = [e for e in lay if e.dxftype() == 'VIEWPORT']
                    has_content = len(non_vps) > 0 or any(vp.dxf.id > 1 for vp in viewports)
                    if has_content:
                        paperspace_layouts.append(lay)
                        
            config = Configuration(background_policy=BackgroundPolicy.WHITE)
            
            if paperspace_layouts:
                logger.info(f"Found {len(paperspace_layouts)} non-empty paperspace layout(s). Rendering layouts.")
                import io
                writer = pypdf.PdfWriter()
                for lay in paperspace_layouts:
                    logger.info(f"Rendering paperspace layout: {lay.name}")
                    backend = pymupdf.PyMuPdfBackend()
                    frontend = Frontend(RenderContext(doc), backend, config=config)
                    frontend.draw_layout(lay)
                    
                    page = layout.Page(0, 0) # auto-fit layout
                    pdf_bytes = backend.get_pdf_bytes(page)
                    
                    pdf_file = io.BytesIO(pdf_bytes)
                    reader = pypdf.PdfReader(pdf_file)
                    for page_obj in reader.pages:
                        writer.add_page(page_obj)
                        
                with open(temp_pdf, "wb") as f:
                    writer.write(f)
            else:
                logger.info("No non-empty paperspace layouts found. Rendering modelspace instead.")
                msp = doc.modelspace()
                backend = pymupdf.PyMuPdfBackend()
                frontend = Frontend(RenderContext(doc), backend, config=config)
                frontend.draw_layout(msp)
                
                page = layout.Page(0, 0) # auto-fit layout
                pdf_bytes = backend.get_pdf_bytes(page)
                
                with open(temp_pdf, "wb") as f:
                    f.write(pdf_bytes)
            logger.info("Python DXF to PDF rendering successful.")
        except Exception as e:
            logger.error(f"Failed to render DXF to PDF: {e}")
            return None

        if os.path.exists(temp_pdf):
            logger.info("Saving generated PDF to target folder...")
            shutil.copy2(temp_pdf, pdf_path)
            logger.info(f"PDF saved: {pdf_path}")
            return pdf_path

    return None

def find_target_file(filename, pdf_dir, source_dir):
    """Searches for the target file in the PDF directory, or under the selected source directory."""
    # 1. Search in the same folder as the PDF
    path = os.path.join(pdf_dir, filename)
    if os.path.exists(path):
        return path
        
    # 2. Search recursively in the source directory
    for root, dirs, files in os.walk(source_dir):
        if filename in files:
            full_path = os.path.join(root, filename)
            logger.info(f"Found target file in search fallback: {full_path}")
            return full_path
            
    return None

def process_file_conversion(filename, pdf_dir, source_dir):
    """Resolves target file path, performs conversion to PDF, and returns the converted PDF path."""
    target_path = find_target_file(filename, pdf_dir, source_dir)
    if not target_path:
        logger.error(f"Target file not found anywhere in source directory: '{filename}'")
        return None

    # Check cache first
    if target_path in CONVERSION_CACHE:
        return CONVERSION_CACHE[target_path]

    _, ext = os.path.splitext(target_path.lower())
    converted_pdf = None
    
    if ext in (".docx", ".doc"):
        converted_pdf = convert_docx_to_pdf(target_path)
    elif ext in (".xlsx", ".xls"):
        converted_pdf = convert_xlsx_to_pdf(target_path)
    elif ext in (".dwg",):
        converted_pdf = convert_dwg_to_pdf(target_path)
    else:
        logger.warning(f"Unsupported file format for conversion: {ext} ('{target_path}')")
        
    if converted_pdf:
        CONVERSION_CACHE[target_path] = converted_pdf
        
    return converted_pdf

def scan_and_merge_pdf(pdf_path, source_dir, target_dir=None):
    """Scans a single PDF file, converts external links, and merges them into a completed PDF."""
    pdf_path = os.path.abspath(pdf_path)
    pdf_dir = os.path.dirname(pdf_path)
    filename_only = os.path.basename(pdf_path)
    
    # Avoid processing already completed files
    if filename_only.endswith("_complete.pdf") or "_converted" in filename_only:
        return

    logger.info(f"==================================================")
    logger.info(f"Processing PDF: {filename_only}")
    logger.info(f"==================================================")

    reader = pypdf.PdfReader(pdf_path)
    writer = pypdf.PdfWriter()
    
    has_merged_files = False
    
    # Iterate through pages
    for page_idx, page in enumerate(reader.pages):
        annotations = page.get("/Annots")
        
        # Check if page has a Launch action link
        launch_filename = None
        if annotations:
            if isinstance(annotations, pypdf.generic.IndirectObject):
                annotations = annotations.get_object()
                
            for annot in annotations:
                if isinstance(annot, pypdf.generic.IndirectObject):
                    annot = annot.get_object()
                
                action = annot.get("/A")
                if action:
                    if isinstance(action, pypdf.generic.IndirectObject):
                        action = action.get_object()
                    
                    action_type = action.get("/S")
                    if action_type == "/Launch":
                        f_spec = action.get("/F")
                        if f_spec:
                            if isinstance(f_spec, pypdf.generic.IndirectObject):
                                f_spec = f_spec.get_object()
                            
                            if isinstance(f_spec, dict):
                                launch_filename = f_spec.get("/F") or f_spec.get("/UF")
                            else:
                                launch_filename = f_spec
                                
                            if launch_filename:
                                if not isinstance(launch_filename, str):
                                    launch_filename = str(launch_filename)
                                launch_filename = launch_filename.strip()
                                break # Use the first launch link found on this page

        if launch_filename:
            logger.info(f"Page {page_idx + 1}: Found link to external file: '{launch_filename}'")
            
            # Convert target file to PDF
            converted_pdf_path = process_file_conversion(launch_filename, pdf_dir, source_dir)
            
            if converted_pdf_path and os.path.exists(converted_pdf_path):
                logger.info(f"Merging converted PDF into output...")
                
                # If we are inserting after, keep the original page
                if not REPLACE_PLACEHOLDER:
                    writer.add_page(page)
                    
                # Append pages of converted document
                ext_reader = pypdf.PdfReader(converted_pdf_path)
                for ext_page in ext_reader.pages:
                    writer.add_page(ext_page)
                    
                has_merged_files = True
                logger.info(f"Merged {len(ext_reader.pages)} pages from '{launch_filename}' successfully.")
            else:
                logger.warning(f"Could not merge '{launch_filename}' (file missing or conversion failed). Keeping original placeholder page.")
                writer.add_page(page)
        else:
            # Standard page, copy directly
            writer.add_page(page)

    if has_merged_files:
        out_folder = target_dir if target_dir else pdf_dir
        output_path = os.path.join(out_folder, filename_only.replace(".pdf", "_complete.pdf"))
        logger.info(f"Saving merged document to: {output_path}")
        with open(output_path, "wb") as out_f:
            writer.write(out_f)
        logger.info(f"Completed processing for: {filename_only}")
    else:
        logger.info(f"No changes made to {filename_only} (no launch links found or successfully merged).")

def main_process(source_dir, target_dir=None, progress_callback=None):
    """Core process that scans files and merges them."""
    logger.info("Starting documentation merging process...")
    logger.info(f"Source Directory: {source_dir}")
    if target_dir:
        logger.info(f"Target Directory: {target_dir}")
    else:
        logger.info("Target Directory: (Same folder as source PDF files)")

    # Find PDF files in the selected source directory recursively
    pdf_files = []
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith(".pdf") and not file.lower().endswith("_complete.pdf") and "_converted" not in file.lower():
                pdf_files.append(os.path.join(root, file))

    if not pdf_files:
        logger.warning("No target PDF files found in the source directory.")
        if progress_callback:
            progress_callback("Finished (No PDFs found)")
        return

    logger.info(f"Found {len(pdf_files)} PDF files to process.")
    for idx, pdf_file in enumerate(pdf_files):
        try:
            scan_and_merge_pdf(pdf_file, source_dir, target_dir)
        except Exception as e:
            logger.error(f"Error processing PDF '{pdf_file}': {e}", exc_info=True)
            
    logger.info("Keeping temporary converted files for verification.")
    logger.info("Documentation merging process finished.")
    
    if progress_callback:
        progress_callback("Finished successfully!")

# Tkinter GUI log handler
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')
            self.text_widget.configure(state='disabled')
        self.text_widget.after(0, append)

# GUI class
class MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ePlan Documentation Merger")
        self.root.geometry("750x550")
        self.root.minsize(600, 450)
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use('vista') # Use standard clean theme
        
        # Frame
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_label = ttk.Label(main_frame, text="ePlan Documentation Merger", font=("Segoe UI", 16, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 15))
        
        # Fields Frame
        fields_frame = ttk.LabelFrame(main_frame, text=" Configuration ", padding="15")
        fields_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Source Dir Row
        ttk.Label(fields_frame, text="Source Directory (Zdrojový adresář):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_entry = ttk.Entry(fields_frame, width=50)
        self.source_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        # Auto-fill source directory with workspace path for user convenience
        self.source_entry.insert(0, get_base_path())
        
        source_btn = ttk.Button(fields_frame, text="Browse...", command=self.browse_source)
        source_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # Target Dir Row
        ttk.Label(fields_frame, text="Target Directory (Cílový adresář - volitelně):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.target_entry = ttk.Entry(fields_frame, width=50)
        self.target_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        target_btn = ttk.Button(fields_frame, text="Browse...", command=self.browse_target)
        target_btn.grid(row=1, column=2, padx=5, pady=5)
        
        fields_frame.columnconfigure(1, weight=1)
        
        # Log Frame
        log_frame = ttk.LabelFrame(main_frame, text=" Progress Log (Průběh) ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        self.log_text = ScrolledText(log_frame, state='disabled', height=12, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Action Frame
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(action_frame, text="Ready", font=("Segoe UI", 10, "italic"))
        self.status_label.pack(side=tk.LEFT, pady=5)
        
        self.run_button = ttk.Button(action_frame, text="Run Merger (Spustit)", command=self.start_process)
        self.run_button.pack(side=tk.RIGHT, pady=5, ipadx=10)
        
        # Setup logging redirection
        self.setup_logging()
        
    def setup_logging(self):
        handler = TextHandler(self.log_text)
        handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
        logger.addHandler(handler)
        
        # Also log to stdout/stderr stream
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s"))
        logger.addHandler(console_handler)
        
    def browse_source(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, os.path.normpath(dir_path))
            
    def browse_target(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, os.path.normpath(dir_path))
            
    def update_status(self, text):
        self.status_label.config(text=text)
        
    def process_complete(self, status):
        self.run_button.config(state='normal')
        self.update_status(status)
        messagebox.showinfo("Finished", f"Process completed: {status}")
        
    def start_process(self):
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()
        
        if not source:
            messagebox.showerror("Error", "Please select a source directory.")
            return
            
        if not os.path.exists(source):
            messagebox.showerror("Error", f"Source directory does not exist:\n{source}")
            return
            
        if target and not os.path.exists(target):
            messagebox.showerror("Error", f"Target directory does not exist:\n{target}")
            return
            
        self.run_button.config(state='disabled')
        self.update_status("Processing...")
        
        # Clear log area
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')
        
        # Reset conversion cache for new run
        CONVERSION_CACHE.clear()
        
        # Run core merging process in a background thread to prevent GUI hang
        t = threading.Thread(target=main_process, args=(source, target if target else None, self.process_complete))
        t.daemon = True
        t.start()

def main():
    root = tk.Tk()
    app = MergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
