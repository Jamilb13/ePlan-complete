import os
import sys
import subprocess
import shutil
import time
import logging
import tempfile
import threading
import urllib.parse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pypdf
import win32com.client
import pythoncom
import fitz  # PyMuPDF

import ezdxf
from ezdxf.addons.drawing import Frontend, RenderContext, pymupdf, layout
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy

# Increase recursion limit for complex PDF outline structures
sys.setrecursionlimit(50000)

# Configuration
REPLACE_PLACEHOLDER = True  # If True, replace the placeholder page containing the link. If False, insert after it.

# Logger setup
logger = logging.getLogger("merge_doc")
logger.setLevel(logging.INFO)

def get_base_path():
    """Resolves the directory containing the running script or compiled executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

# Global caches for speed optimization
CONVERSION_CACHE = {}  # Maps target_path -> converted_pdf_path
SOURCE_FILE_INDEX = {} # Maps normalized filename -> absolute_path

def build_source_file_index(source_dir):
    """
    Crawls source_dir ONCE and builds a fast O(1) filename lookup index.
    Eliminates repetitive os.walk disk scans during link resolution.
    """
    global SOURCE_FILE_INDEX
    SOURCE_FILE_INDEX.clear()
    if not source_dir or not os.path.exists(source_dir):
        return

    logger.info(f"Indexing files in source directory for fast lookup: {source_dir}")
    file_count = 0
    start_time = time.time()
    
    for root, _, files in os.walk(source_dir):
        for f in files:
            norm_name = f.lower()
            if norm_name not in SOURCE_FILE_INDEX:
                SOURCE_FILE_INDEX[norm_name] = os.path.join(root, f)
            file_count += 1
            
    elapsed = time.time() - start_time
    logger.info(f"Indexed {file_count} file(s) in {elapsed:.2f}s.")

def find_target_file(filename, pdf_dir, source_dir):
    """Searches for the target file using O(1) index lookup and local directory check."""
    clean_name = os.path.basename(filename)
    clean_name_lower = clean_name.lower()
    raw_name_lower = filename.lower()

    # 1. Direct check in same folder as PDF
    p1 = os.path.join(pdf_dir, filename)
    if os.path.exists(p1):
        return p1
    p2 = os.path.join(pdf_dir, clean_name)
    if os.path.exists(p2):
        return p2

    # 2. Fast O(1) index lookup
    if clean_name_lower in SOURCE_FILE_INDEX:
        target = SOURCE_FILE_INDEX[clean_name_lower]
        logger.info(f"Found target file via fast index: {target}")
        return target
        
    if raw_name_lower in SOURCE_FILE_INDEX:
        target = SOURCE_FILE_INDEX[raw_name_lower]
        logger.info(f"Found target file via fast index: {target}")
        return target

    return None

def get_office_pdf_path(src_path):
    """Generates target PDF path in the same folder."""
    base, _ = os.path.splitext(src_path)
    return base + "_converted.pdf"

# Class managing persistent COM instances for Word & Excel across batch runs
class OfficeCOMManager:
    """Manages persistent MS Office COM instances to avoid opening/closing Word & Excel per file."""
    def __init__(self):
        self.word = None
        self.excel = None

    def get_word(self):
        if self.word is None:
            try:
                self.word = win32com.client.DispatchEx("Word.Application")
                self.word.Visible = False
                self.word.DisplayAlerts = False
            except Exception as e:
                logger.error(f"Failed to start MS Word COM instance: {e}")
                self.word = None
        return self.word

    def get_excel(self):
        if self.excel is None:
            try:
                self.excel = win32com.client.DispatchEx("Excel.Application")
                self.excel.Visible = False
                self.excel.DisplayAlerts = False
            except Exception as e:
                logger.error(f"Failed to start MS Excel COM instance: {e}")
                self.excel = None
        return self.excel

    def reset_word(self):
        if self.word:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None

    def reset_excel(self):
        if self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None

    def close_all(self):
        self.reset_word()
        self.reset_excel()

# Module-level COM manager instance
COM_MANAGER = OfficeCOMManager()

def convert_docx_to_pdf(docx_path):
    """Converts a Word document to PDF using persistent MS Word COM instance."""
    docx_path = os.path.abspath(docx_path)
    pdf_path = get_office_pdf_path(docx_path)
    
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for Word file: {pdf_path}")
        return pdf_path

    logger.info(f"Converting Word document: {docx_path} -> {pdf_path}")
    word = COM_MANAGER.get_word()
    if not word:
        logger.error(f"MS Word is unavailable. Cannot convert '{docx_path}'.")
        return None

    doc = None
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True, ConfirmConversions=False)
        doc.SaveAs(pdf_path, FileFormat=17) # wdFormatPDF = 17
        logger.info("Word conversion successful.")
        return pdf_path
    except Exception as e:
        logger.error(f"Failed to convert Word document {docx_path}: {e}")
        COM_MANAGER.reset_word()
        return None
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass

def convert_xlsx_to_pdf(xlsx_path):
    """Converts an Excel sheet to PDF using persistent MS Excel COM instance."""
    xlsx_path = os.path.abspath(xlsx_path)
    pdf_path = get_office_pdf_path(xlsx_path)
    
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for Excel file: {pdf_path}")
        return pdf_path

    logger.info(f"Converting Excel sheet: {xlsx_path} -> {pdf_path}")
    excel = COM_MANAGER.get_excel()
    if not excel:
        logger.error(f"MS Excel is unavailable. Cannot convert '{xlsx_path}'.")
        return None

    wb = None
    try:
        wb = excel.Workbooks.Open(xlsx_path, ReadOnly=True)
        wb.ExportAsFixedFormat(0, pdf_path) # xlTypePDF = 0
        logger.info("Excel conversion successful.")
        return pdf_path
    except Exception as e:
        logger.error(f"Failed to convert Excel sheet {xlsx_path}: {e}")
        COM_MANAGER.reset_excel()
        return None
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

def convert_dwg_to_pdf(dwg_path):
    """Converts a DWG drawing to DXF via ODAFileConverter and renders it to PDF using ezdxf + PyMuPDF."""
    dwg_path = os.path.abspath(dwg_path)
    base, _ = os.path.splitext(dwg_path)
    
    pre_converted_pdf = base + ".pdf"
    if os.path.exists(pre_converted_pdf):
        logger.info(f"Found pre-converted PDF for DWG file: {pre_converted_pdf}")
        return pre_converted_pdf

    pdf_path = base + "_converted.pdf"
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for DWG file: {pdf_path}")
        return pdf_path

    dxf_path = base + ".dxf"

    oda_converter = None
    if getattr(sys, 'frozen', False):
        temp_path = os.path.join(sys._MEIPASS, "bin", "ODAFileConverter.exe")
        if os.path.exists(temp_path):
            oda_converter = temp_path
            
    if not oda_converter:
        oda_converter = os.path.join(get_base_path(), "bin", "ODAFileConverter.exe")
        
    if not os.path.exists(oda_converter):
        logger.error(f"ODAFileConverter.exe not found at: {oda_converter}")
        return None

    logger.info(f"Attempting to convert DWG drawing to DXF: {dwg_path}")
    
    with tempfile.TemporaryDirectory(prefix="oda_conv_") as tmp_dir:
        temp_dwg = os.path.join(tmp_dir, "drawing.dwg")
        shutil.copy2(dwg_path, temp_dwg)
        
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

        temp_pdf = os.path.join(tmp_dir, "drawing.pdf")
        try:
            logger.info("Rendering DXF to PDF via Python (ezdxf + PyMuPDF)...")
            doc = ezdxf.readfile(temp_dxf)
            
            paperspace_layouts = []
            for name in doc.layout_names_in_taborder():
                lay = doc.layout(name)
                if lay.is_any_paperspace:
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
                    
                    page = layout.Page(0, 0)
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
                
                page = layout.Page(0, 0)
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

def convert_image_to_pdf(img_path):
    """Converts an image file (PNG, JPG, BMP, TIFF) to PDF using PyMuPDF."""
    img_path = os.path.abspath(img_path)
    pdf_path = get_office_pdf_path(img_path)
    if os.path.exists(pdf_path):
        logger.info(f"Using cached PDF for Image file: {pdf_path}")
        return pdf_path

    logger.info(f"Converting Image to PDF: {img_path} -> {pdf_path}")
    try:
        img_doc = fitz.open(img_path)
        pdf_bytes = img_doc.convert_to_pdf()
        img_doc.close()
        pdf_doc = fitz.open("pdf", pdf_bytes)
        pdf_doc.save(pdf_path)
        pdf_doc.close()
        logger.info("Image conversion successful.")
        return pdf_path
    except Exception as e:
        logger.error(f"Failed to convert image {img_path}: {e}")
        return None

def process_file_conversion(filename, pdf_dir, source_dir):
    """Resolves target file path, performs conversion to PDF if needed, and returns the converted PDF path."""
    target_path = find_target_file(filename, pdf_dir, source_dir)
    if not target_path:
        logger.warning(f"Target file not found anywhere in source directory: '{filename}'")
        return None

    if target_path in CONVERSION_CACHE:
        return CONVERSION_CACHE[target_path]

    _, ext = os.path.splitext(target_path.lower())
    converted_pdf = None
    
    try:
        if ext == ".pdf":
            logger.info(f"File is already PDF (direct merge): {target_path}")
            converted_pdf = target_path
        elif ext in (".docx", ".doc"):
            converted_pdf = convert_docx_to_pdf(target_path)
        elif ext in (".xlsx", ".xls"):
            converted_pdf = convert_xlsx_to_pdf(target_path)
        elif ext in (".dwg",):
            converted_pdf = convert_dwg_to_pdf(target_path)
        elif ext in (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"):
            converted_pdf = convert_image_to_pdf(target_path)
        else:
            logger.warning(f"Unsupported file format for conversion: {ext} ('{target_path}')")
    except Exception as e:
        logger.error(f"Error converting '{target_path}': {e}")
        return None

    if converted_pdf:
        CONVERSION_CACHE[target_path] = converted_pdf
        
    return converted_pdf

def extract_launch_link(page, page_idx, pypdf_reader=None):
    """
    Extracts external launch file link from a page using PyMuPDF first,
    falling back to pre-instantiated pypdf_reader if fitz finds none.
    """
    launch_filename = None

    # 1. Fast PyMuPDF link check
    try:
        for link in page.get_links():
            kind = link.get("kind")
            if kind in (fitz.LINK_LAUNCH, 5, 3):
                fname = link.get("file") or link.get("uri")
                if fname:
                    launch_filename = urllib.parse.unquote(fname).strip()
                    return launch_filename
    except Exception as e:
        logger.warning(f"Error reading PyMuPDF links on page {page_idx + 1}: {e}")

    # 2. PyPDF annotation fallback using shared reader instance
    if pypdf_reader and page_idx < len(pypdf_reader.pages):
        try:
            pypdf_page = pypdf_reader.pages[page_idx]
            annots = pypdf_page.get("/Annots")
            if annots:
                if isinstance(annots, pypdf.generic.IndirectObject):
                    annots = annots.get_object()
                for annot in annots:
                    if isinstance(annot, pypdf.generic.IndirectObject):
                        annot = annot.get_object()
                    action = annot.get("/A")
                    if action:
                        if isinstance(action, pypdf.generic.IndirectObject):
                            action = action.get_object()
                        if action.get("/S") == "/Launch":
                            f_spec = action.get("/F")
                            if f_spec:
                                if isinstance(f_spec, pypdf.generic.IndirectObject):
                                    f_spec = f_spec.get_object()
                                if isinstance(f_spec, dict):
                                    fname = f_spec.get("/F") or f_spec.get("/UF")
                                else:
                                    fname = f_spec
                                if fname:
                                    launch_filename = str(fname).strip()
                                    return launch_filename
        except Exception:
            pass

    return launch_filename

def get_pdf_structure_preview(pdf_path, source_dir):
    """
    Scans a PDF file for launch links and resolved target file locations.
    Returns a formatted string representing the document structure tree.
    """
    pdf_path = os.path.abspath(pdf_path)
    pdf_dir = os.path.dirname(pdf_path)
    filename_only = os.path.basename(pdf_path)

    items = []
    total_pages = 0

    pypdf_reader = None
    try:
        pypdf_reader = pypdf.PdfReader(pdf_path)
    except Exception:
        pypdf_reader = None

    try:
        src_doc = fitz.open(pdf_path)
        total_pages = len(src_doc)
        
        for page_idx in range(total_pages):
            page = src_doc[page_idx]
            launch_filename = extract_launch_link(page, page_idx, pypdf_reader)

            if launch_filename:
                target_path = find_target_file(launch_filename, pdf_dir, source_dir)
                status = "UNKNOWN"
                ext = os.path.splitext(launch_filename)[1].lower()
                
                if target_path and os.path.exists(target_path):
                    if ext == ".pdf":
                        status = "PDF (Přímé sloučení)"
                    elif ext in (".docx", ".doc"):
                        status = "Word Dokument (.doc/.docx)"
                    elif ext in (".xlsx", ".xls"):
                        status = "Excel Tabulka (.xls/.xlsx)"
                    elif ext in (".dwg", ".dxf"):
                        status = "DWG/DXF Výkres"
                    elif ext in (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"):
                        status = "Obrázek"
                    else:
                        status = f"Neznámý formát ({ext})"
                else:
                    status = "NENALEZENO (Soubor chybí ve zdrojové složce)"

                items.append({
                    "page": page_idx + 1,
                    "link": launch_filename,
                    "target_path": target_path,
                    "ext": ext,
                    "status": status
                })

        src_doc.close()
    except Exception as e:
        logger.error(f"Error scanning structure for '{pdf_path}': {e}")
        return f"Error reading PDF structure: {e}"

    lines = []
    lines.append("=================================================================================")
    lines.append("NÁHLED STRUKTURY DOKUMENTACE / DOCUMENTATION STRUCTURE PREVIEW")
    lines.append("=================================================================================")
    lines.append(f"📄 Hlavní PDF: {filename_only}")
    lines.append(f"📍 Umístění: {pdf_path}")
    lines.append(f"📑 Celkem stran hlavní dokumentace: {total_pages}")
    lines.append(f"🔗 Nalezeno odkazů na přílohy: {len(items)}")
    lines.append("─" * 81)

    if not items:
        lines.append("  (V dokumentu nebyly nalezeny žádné odkazované externí přílohy)")
    else:
        for idx, item in enumerate(items):
            is_last = (idx == len(items) - 1)
            prefix = "└── " if is_last else "├── "
            child_prefix = "    " if is_last else "│   "
            
            lines.append(f"{prefix}📄 Strana {item['page']}: Odkaz -> '{item['link']}'")
            if item['target_path']:
                lines.append(f"{child_prefix}├── 📍 Umístění: {item['target_path']}")
                lines.append(f"{child_prefix}└── ℹ️ Typ/Stav: {item['status']}")
            else:
                lines.append(f"{child_prefix}├── 📍 Umístění: NENALEZENO")
                lines.append(f"{child_prefix}└── ⚠️ Typ/Stav: {item['status']}")

    lines.append("=================================================================================")
    return "\n".join(lines)

def scan_and_merge_pdf(pdf_path, source_dir, target_dir=None):
    """Scans a single PDF file, converts external links, and merges them into a completed PDF with updated outlines."""
    pdf_path = os.path.abspath(pdf_path)
    pdf_dir = os.path.dirname(pdf_path)
    filename_only = os.path.basename(pdf_path)
    
    if filename_only.endswith("_complete.pdf") or "_converted" in filename_only:
        return

    logger.info(f"==================================================")
    logger.info(f"Processing PDF: {filename_only}")
    logger.info(f"==================================================")

    try:
        preview = get_pdf_structure_preview(pdf_path, source_dir)
        logger.info("\n" + preview)
    except Exception as e:
        logger.warning(f"Could not generate structure preview: {e}")

    try:
        src_doc = fitz.open(pdf_path)
    except Exception as e:
        logger.error(f"Failed to open PDF '{pdf_path}': {e}. Skipping this file.")
        return

    pypdf_reader = None
    try:
        pypdf_reader = pypdf.PdfReader(pdf_path)
    except Exception:
        pypdf_reader = None

    out_doc = fitz.open()
    
    has_merged_files = False
    page_mapping = {}
    merged_pages = {}
    added_merged = set()
    out_page_idx = 0
    
    for page_idx in range(len(src_doc)):
        try:
            page = src_doc[page_idx]
            launch_filename = extract_launch_link(page, page_idx, pypdf_reader)

            merged_successfully = False
            converted_pdf_path = None
            
            if launch_filename:
                logger.info(f"Page {page_idx + 1}: Found link to external file: '{launch_filename}'")
                try:
                    converted_pdf_path = process_file_conversion(launch_filename, pdf_dir, source_dir)
                except Exception as e:
                    logger.warning(f"Conversion error for '{launch_filename}': {e}. Skipping conversion.")

                if converted_pdf_path and os.path.exists(converted_pdf_path):
                    logger.info(f"Merging converted PDF into output...")
                    try:
                        ext_doc = fitz.open(converted_pdf_path)
                        page_count = len(ext_doc)
                        
                        merged_pages[page_idx] = launch_filename
                        
                        if not REPLACE_PLACEHOLDER:
                            out_doc.insert_pdf(src_doc, from_page=page_idx, to_page=page_idx)
                            page_mapping[page_idx] = out_page_idx
                            out_page_idx += 1
                        else:
                            page_mapping[page_idx] = out_page_idx
                            
                        out_doc.insert_pdf(ext_doc)
                        out_page_idx += page_count
                        logger.info(f"Merged {page_count} pages from '{launch_filename}' successfully.")
                        ext_doc.close()
                        has_merged_files = True
                        merged_successfully = True
                    except Exception as e:
                        logger.warning(f"Failed to insert PDF '{converted_pdf_path}': {e}. Keeping original placeholder page.")

            if not merged_successfully:
                if launch_filename and not converted_pdf_path:
                    logger.warning(f"Could not merge '{launch_filename}' (file missing, unsupported format, or conversion failed). Keeping original placeholder page.")
                out_doc.insert_pdf(src_doc, from_page=page_idx, to_page=page_idx)
                page_mapping[page_idx] = out_page_idx
                out_page_idx += 1

        except Exception as page_err:
            logger.error(f"Unhandled error processing page {page_idx + 1}: {page_err}. Keeping original page and continuing...")
            try:
                out_doc.insert_pdf(src_doc, from_page=page_idx, to_page=page_idx)
                page_mapping[page_idx] = out_page_idx
                out_page_idx += 1
            except Exception:
                pass

    if has_merged_files:
        try:
            orig_toc = src_doc.get_toc(simple=False)
            new_toc = []
            if orig_toc:
                for item in orig_toc:
                    lvl, title, p_num = item[0], item[1], item[2]
                    old_p_idx = p_num - 1
                    if old_p_idx in page_mapping:
                        new_p_idx = page_mapping[old_p_idx]
                        new_toc.append([lvl, title, new_p_idx + 1])
                        
                        if old_p_idx in merged_pages:
                            fname = merged_pages[old_p_idx]
                            if fname not in added_merged:
                                new_toc.append([lvl + 1, fname, new_p_idx + 1])
                                added_merged.add(fname)
                                
            for old_p_idx, fname in merged_pages.items():
                if fname not in added_merged:
                    new_p_idx = page_mapping[old_p_idx]
                    logger.info(f"Adding top-level bookmark for: {fname}")
                    new_toc.append([1, fname, new_p_idx + 1])
                    
            out_doc.set_toc(new_toc)
        except Exception as e:
            logger.error(f"Failed to rebuild outlines: {e}")

        out_folder = target_dir if target_dir else pdf_dir
        if not os.path.exists(out_folder):
            os.makedirs(out_folder, exist_ok=True)

        output_path = os.path.join(out_folder, filename_only.replace(".pdf", "_complete.pdf"))
        logger.info(f"Saving merged document to: {output_path}")
        try:
            out_doc.save(output_path, garbage=4, deflate=True)
            logger.info(f"Completed processing for: {filename_only}")
        except Exception as e:
            logger.error(f"Failed to save output PDF '{output_path}': {e}")
    else:
        logger.info(f"No changes made to {filename_only} (no launch links found or successfully merged).")

    try:
        out_doc.close()
        src_doc.close()
    except Exception:
        pass

def get_pdf_files_to_process(source_dir, source_file=None):
    """Finds list of PDF files to process from source_file or source_dir."""
    pdf_files = []
    if source_file and os.path.exists(source_file):
        source_file_abs = os.path.abspath(source_file)
        ext = os.path.splitext(source_file_abs)[1].lower()
        if ext == ".pdf":
            pdf_files.append(source_file_abs)
        elif ext in (".txt", ".csv"):
            logger.info(f"Reading PDF list from source file: {source_file_abs}")
            try:
                with open(source_file_abs, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#"):
                            parts = line.split(",") if ext == ".csv" else [line]
                            candidate = parts[0].strip().strip('"\'')
                            if not os.path.isabs(candidate):
                                candidate = os.path.join(source_dir, candidate)
                            if os.path.exists(candidate) and candidate.lower().endswith(".pdf"):
                                pdf_files.append(os.path.abspath(candidate))
                            else:
                                logger.warning(f"File listed in {source_file_abs} not found or not PDF: '{candidate}'")
            except Exception as e:
                logger.error(f"Failed to read source file '{source_file_abs}': {e}")
        else:
            logger.warning(f"Unsupported source file extension: {ext}. Searching source directory instead.")

    if not pdf_files and source_dir and os.path.exists(source_dir):
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.lower().endswith(".pdf") and not file.lower().endswith("_complete.pdf") and "_converted" not in file.lower():
                    pdf_files.append(os.path.join(root, file))

    return pdf_files

def preview_process(source_dir, source_file=None, progress_callback=None):
    """Scans and displays structure preview for all target PDFs."""
    pythoncom.CoInitialize()
    try:
        logger.info("Generuji náhled struktury dokumentace...")
        build_source_file_index(source_dir)
        pdf_files = get_pdf_files_to_process(source_dir, source_file)
        
        if not pdf_files:
            logger.warning("Nenalezeny žádné PDF soubory pro zobrazení struktury.")
            if progress_callback:
                progress_callback("Náhled dokončen (Žádná PDF nenalezena)")
            return

        for pdf_file in pdf_files:
            try:
                preview_text = get_pdf_structure_preview(pdf_file, source_dir)
                logger.info("\n" + preview_text)
            except Exception as e:
                logger.error(f"Chyba při generování náhledu pro '{pdf_file}': {e}")

        if progress_callback:
            progress_callback("Náhled struktury dokončen")
    finally:
        pythoncom.CoUninitialize()

def main_process(source_dir, target_dir=None, source_file=None, progress_callback=None):
    """Core process that scans files and merges them."""
    pythoncom.CoInitialize()
    try:
        logger.info("Starting documentation merging process...")
        logger.info(f"Source Directory: {source_dir}")
        if source_file:
            logger.info(f"Source File: {source_file}")
        if target_dir:
            logger.info(f"Target Directory: {target_dir}")
        else:
            logger.info("Target Directory: (Same folder as source PDF files)")

        # Fast O(1) file index building
        build_source_file_index(source_dir)

        pdf_files = get_pdf_files_to_process(source_dir, source_file)

        if not pdf_files:
            logger.warning("No target PDF files found.")
            if progress_callback:
                progress_callback("Finished (No PDFs found)")
            return

        logger.info(f"Found {len(pdf_files)} PDF file(s) to process.")
        for idx, pdf_file in enumerate(pdf_files):
            try:
                scan_and_merge_pdf(pdf_file, source_dir, target_dir)
            except Exception as e:
                logger.error(f"Error processing PDF '{pdf_file}': {e}. Skipping file and continuing...", exc_info=True)
                
        logger.info("Keeping temporary converted files for verification.")
        logger.info("Documentation merging process finished.")
        
        if progress_callback:
            progress_callback("Finished successfully!")
    finally:
        # Gracefully shut down MS Office COM processes
        COM_MANAGER.close_all()
        pythoncom.CoUninitialize()

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
        self.root.geometry("780x600")
        self.root.minsize(640, 480)
        
        self.style = ttk.Style()
        self.style.theme_use('vista')
        
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        header_label = ttk.Label(main_frame, text="ePlan Documentation Merger", font=("Segoe UI", 16, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 15))
        
        fields_frame = ttk.LabelFrame(main_frame, text=" Configuration ", padding="15")
        fields_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(fields_frame, text="Source Directory (Zdrojový adresář):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_entry = ttk.Entry(fields_frame, width=50)
        self.source_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        self.source_entry.insert(0, get_base_path())
        
        source_btn = ttk.Button(fields_frame, text="Browse Folder...", command=self.browse_source)
        source_btn.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(fields_frame, text="Source File (Zdrojový soubor - volitelně):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.source_file_entry = ttk.Entry(fields_frame, width=50)
        self.source_file_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        source_file_btn = ttk.Button(fields_frame, text="Browse File...", command=self.browse_source_file)
        source_file_btn.grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(fields_frame, text="Target Directory (Cílový adresář - volitelně):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.target_entry = ttk.Entry(fields_frame, width=50)
        self.target_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.EW)
        target_btn = ttk.Button(fields_frame, text="Browse Folder...", command=self.browse_target)
        target_btn.grid(row=2, column=2, padx=5, pady=5)
        
        fields_frame.columnconfigure(1, weight=1)
        
        log_frame = ttk.LabelFrame(main_frame, text=" Progress Log & Structure (Průběh a Náhled) ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        self.log_text = ScrolledText(log_frame, state='disabled', height=14, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(action_frame, text="Ready", font=("Segoe UI", 10, "italic"))
        self.status_label.pack(side=tk.LEFT, pady=5)
        
        self.preview_button = ttk.Button(action_frame, text="Náhled struktury (Preview)", command=self.start_preview)
        self.preview_button.pack(side=tk.RIGHT, pady=5, padx=(0, 5), ipadx=5)
        
        self.run_button = ttk.Button(action_frame, text="Run Merger (Spustit)", command=self.start_process)
        self.run_button.pack(side=tk.RIGHT, pady=5, ipadx=10)
        
        self.setup_logging()
        
    def setup_logging(self):
        # Prevent adding duplicate handlers
        if not logger.handlers:
            handler = TextHandler(self.log_text)
            handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
            logger.addHandler(handler)
            
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s"))
            logger.addHandler(console_handler)
        
    def browse_source(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, os.path.normpath(dir_path))
            
    def browse_source_file(self):
        initial_dir = self.source_entry.get() if self.source_entry.get() else get_base_path()
        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Select Source File",
            filetypes=[
                ("PDF or List Files", "*.pdf;*.txt;*.csv"),
                ("PDF Files (*.pdf)", "*.pdf"),
                ("Text/CSV List Files (*.txt;*.csv)", "*.txt;*.csv"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            file_path_norm = os.path.normpath(file_path)
            self.source_file_entry.delete(0, tk.END)
            self.source_file_entry.insert(0, file_path_norm)
            
            file_dir = os.path.dirname(file_path_norm)
            if file_dir:
                self.source_entry.delete(0, tk.END)
                self.source_entry.insert(0, file_dir)

    def browse_target(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, os.path.normpath(dir_path))
            
    def update_status(self, text):
        self.status_label.config(text=text)
        
    def validate_inputs(self):
        source_dir = self.source_entry.get().strip()
        source_file = self.source_file_entry.get().strip()
        target_dir = self.target_entry.get().strip()
        
        if source_file and not os.path.exists(source_file):
            messagebox.showerror("Error", f"Source file does not exist:\n{source_file}")
            return None, None, None

        if not source_dir and source_file:
            source_dir = os.path.dirname(os.path.abspath(source_file))
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, source_dir)

        if not source_dir:
            messagebox.showerror("Error", "Please select a source directory or source file.")
            return None, None, None
            
        if not os.path.exists(source_dir):
            messagebox.showerror("Error", f"Source directory does not exist:\n{source_dir}")
            return None, None, None
            
        if target_dir and not os.path.exists(target_dir):
            messagebox.showerror("Error", f"Target directory does not exist:\n{target_dir}")
            return None, None, None
            
        return source_dir, source_file, target_dir

    def process_complete(self, status):
        self.run_button.config(state='normal')
        self.preview_button.config(state='normal')
        self.update_status(status)
        messagebox.showinfo("Finished", f"Process completed: {status}")
        
    def preview_complete(self, status):
        self.run_button.config(state='normal')
        self.preview_button.config(state='normal')
        self.update_status(status)

    def start_preview(self):
        source_dir, source_file, _ = self.validate_inputs()
        if not source_dir:
            return

        self.run_button.config(state='disabled')
        self.preview_button.config(state='disabled')
        self.update_status("Generuji náhled...")
        
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')
        
        t = threading.Thread(
            target=preview_process,
            args=(source_dir, source_file if source_file else None, self.preview_complete)
        )
        t.daemon = True
        t.start()

    def start_process(self):
        source_dir, source_file, target_dir = self.validate_inputs()
        if not source_dir:
            return

        self.run_button.config(state='disabled')
        self.preview_button.config(state='disabled')
        self.update_status("Processing...")
        
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')
        
        CONVERSION_CACHE.clear()
        
        t = threading.Thread(
            target=main_process,
            args=(source_dir, target_dir if target_dir else None, source_file if source_file else None, self.process_complete)
        )
        t.daemon = True
        t.start()

def main():
    root = tk.Tk()
    app = MergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
