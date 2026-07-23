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

# Helper functions for PDF Splitting
def analyze_pdf_structure_for_split(src_path):
    import fitz
    import re
    doc = fitz.open(src_path)
    toc = doc.get_toc()
    total_pages = len(doc)

    page1_text = doc[0].get_text('text') if total_pages > 0 else ""
    m_base = re.search(r'\b(D\d{4,8})\b', page1_text)
    base_doc_num = m_base.group(1) if m_base else ""
    if not base_doc_num:
        m_file = re.search(r'\b(D\d{4,8})\b', os.path.basename(src_path))
        if m_file:
            base_doc_num = m_file.group(1)
    if not base_doc_num:
        base_doc_num = os.path.splitext(os.path.basename(src_path))[0]

    sections = []

    # Strategy 1: Bookmark scan for & tags (&TZ, &SM, &VV, &BS, &TZ1, etc.)
    toc_sections = []
    for item in toc:
        lvl, title, page = item[0], item[1], item[2]
        m = re.search(r'&([A-Za-z0-9_]+)\s*(.*)', title)
        if m:
            code = m.group(1).split('#')[0]
            name = m.group(2).strip()
            if not any(s['code'] == code for s in toc_sections):
                toc_sections.append({'code': code, 'name': name or f"Část {code}", 'start_page': page, 'source': 'Záložky'})

    if len(toc_sections) >= 2:
        sorted_toc = sorted(toc_sections, key=lambda x: x['start_page'])
        if sorted_toc[0]['start_page'] > 1:
            sections.append({
                'code': 'COVER',
                'name': 'Seznam příloh / Úvodní část',
                'start_page': 1,
                'end_page': sorted_toc[0]['start_page'] - 1,
                'source': 'Záložky'
            })
        for i, sec in enumerate(sorted_toc):
            start_p = sec['start_page']
            end_p = sorted_toc[i+1]['start_page'] - 1 if i + 1 < len(sorted_toc) else total_pages
            sections.append({
                'code': sec['code'],
                'name': sec['name'],
                'start_page': start_p,
                'end_page': end_p,
                'source': sec['source']
            })

    # Strategy 2: Text Scan Fallback (for merged files or lost TOC)
    if not sections:
        found_covers = []
        for pno in range(1, total_pages + 1):
            text = doc[pno-1].get_text('text')
            lines = [l.strip() for l in text.split('\n') if l.strip()]

            code_found = None
            name_found = None

            for idx, line in enumerate(lines[:15]):
                m1 = re.match(r'^\.([A-Za-z0-9_]{2,8})$', line)
                if m1:
                    code_found = m1.group(1).upper()
                    name_found = lines[1] if idx != 1 and len(lines) > 1 else (lines[0] if lines else "")
                    break
                m2 = re.match(r'^&([A-Za-z0-9_]{2,8})\b', line)
                if m2:
                    code_found = m2.group(1).upper()
                    name_found = line
                    break

            if code_found:
                found_covers.append({
                    'code': code_found,
                    'name': name_found or f"Část .{code_found}",
                    'start_page': pno
                })

        if found_covers:
            unique_covers = []
            for fc in found_covers:
                if not unique_covers or unique_covers[-1]['code'] != fc['code']:
                    unique_covers.append(fc)
            sorted_covers = sorted(unique_covers, key=lambda x: x['start_page'])
            if sorted_covers[0]['start_page'] > 1:
                sections.append({
                    'code': 'COVER',
                    'name': 'Seznam příloh / Úvodní část',
                    'start_page': 1,
                    'end_page': sorted_covers[0]['start_page'] - 1,
                    'source': 'Textový sken'
                })
            for i, sec in enumerate(sorted_covers):
                start_p = sec['start_page']
                end_p = sorted_covers[i+1]['start_page'] - 1 if i + 1 < len(sorted_covers) else total_pages
                sections.append({
                    'code': sec['code'],
                    'name': sec['name'],
                    'start_page': start_p,
                    'end_page': end_p,
                    'source': 'Textový sken'
                })

    # Discover full document number across sections (e.g. D09293501 or D231542633)
    full_doc_num = base_doc_num
    for sec in sections:
        code = sec['code']
        if code == 'COVER':
            continue
        candidate_pages = list(range(sec['start_page'], min(sec['start_page'] + 5, sec['end_page'] + 1)))
        for pno in candidate_pages:
            ptxt = doc[pno-1].get_text('text')
            m_full = re.search(r'\b(' + re.escape(base_doc_num) + r'\d{3,5})\b', ptxt)
            if m_full:
                full_doc_num = m_full.group(1)
                break
            m_sub = re.search(r'(\d{4})\s*\n\s*1\s*\n\s*\.' + re.escape(code), ptxt, re.IGNORECASE)
            if m_sub:
                full_doc_num = f"{base_doc_num}{m_sub.group(1)}"
                break
        if full_doc_num != base_doc_num:
            break

    # Assign output filenames according to user specification format: 0x_D231542633.TZ.pdf
    # where x is the bookmark index (00 for COVER, 01+ for sections)
    bookmark_idx = 0
    for sec in sections:
        code = sec['code']
        prefix = f"{bookmark_idx:02d}_"
        sec['bookmark_idx'] = bookmark_idx

        if code == 'COVER':
            sec['filename'] = f"{prefix}{base_doc_num}_Seznam_příloh.pdf"
        else:
            doc_num_found = None
            candidate_pages = list(range(sec['start_page'], min(sec['start_page'] + 5, sec['end_page'] + 1)))

            # Pass 1: Exact document code pattern match (e.g., D231542633.TZ1)
            for pno in candidate_pages:
                ptxt = doc[pno-1].get_text('text')
                # Match the core doc number without any existing prefix
                m_exact = re.search(r'\b(D\d{5,10}\.' + re.escape(code) + r')\b', ptxt, re.IGNORECASE)
                if m_exact:
                    doc_num_found = m_exact.group(1)
                    break

            # Pass 2: Header label match ("Dokument č.: ...")
            if not doc_num_found:
                for pno in candidate_pages:
                    ptxt = doc[pno-1].get_text('text')
                    m_doc_hdr = re.search(r'Dokument\s*č\.\s*:\s*([A-Za-z0-9_\.]+\.' + re.escape(code) + r')\b', ptxt, re.IGNORECASE)
                    if m_doc_hdr:
                        doc_num_found = m_doc_hdr.group(1)
                        break

            # Pass 3: Title block number fallback
            if not doc_num_found:
                for pno in candidate_pages:
                    ptxt = doc[pno-1].get_text('text')
                    m_sub = re.search(r'(\d{4})\s*\n\s*1\s*\n\s*\.' + re.escape(code), ptxt, re.IGNORECASE)
                    if m_sub:
                        doc_num_found = f"{base_doc_num}{m_sub.group(1)}.{code}"
                        break
                        
            if doc_num_found:
                sec['filename'] = f"{prefix}{doc_num_found}.pdf"
            else:
                sec['filename'] = f"{prefix}{full_doc_num}.{code}.pdf"

        bookmark_idx += 1

    doc.close()
    return sections

# GUI class with tabs for merging and splitting
class MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ePlan Documentation Merger & Splitter")
        self.root.geometry("860x720")
        self.root.minsize(720, 560)
        
        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except Exception:
            pass

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_label = ttk.Label(main_frame, text="ePlan Documentation Tool", font=("Segoe UI", 16, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 10))

        # Notebook tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_merge = ttk.Frame(self.notebook, padding="10")
        self.tab_split = ttk.Frame(self.notebook, padding="10")

        self.notebook.add(self.tab_merge, text=" 🔗 Slučování PDF ")
        self.notebook.add(self.tab_split, text=" ✂️ Rozdělování PDF ")

        self.setup_merge_tab()
        self.setup_split_tab()

        self.setup_logging()

    def setup_merge_tab(self):
        fields_frame = ttk.LabelFrame(self.tab_merge, text=" Configuration (Konfigurace) ", padding="15")
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
        
        log_frame = ttk.LabelFrame(self.tab_merge, text=" Progress Log & Structure (Průběh a Náhled) ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        self.log_text = ScrolledText(log_frame, state='disabled', height=14, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        action_frame = ttk.Frame(self.tab_merge)
        action_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(action_frame, text="Ready", font=("Segoe UI", 10, "italic"))
        self.status_label.pack(side=tk.LEFT, pady=5)
        
        self.preview_button = ttk.Button(action_frame, text="Náhled struktury (Preview)", command=self.start_preview)
        self.preview_button.pack(side=tk.RIGHT, pady=5, padx=(0, 5), ipadx=5)
        
        self.run_button = ttk.Button(action_frame, text="Run Merger (Spustit)", command=self.start_process)
        self.run_button.pack(side=tk.RIGHT, pady=5, ipadx=10)

    def setup_split_tab(self):
        split_cfg = ttk.LabelFrame(self.tab_split, text=" Konfigurace rozdělování ", padding="10")
        split_cfg.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(split_cfg, text="Zdrojový PDF soubor:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.split_src_entry = ttk.Entry(split_cfg, width=45)
        self.split_src_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        btn_src_browse = ttk.Button(split_cfg, text="Procházet...", command=self.browse_split_src)
        btn_src_browse.grid(row=0, column=2, padx=5, pady=5)

        self.btn_analyze = ttk.Button(split_cfg, text="🔍 Analýza struktury", command=self.start_split_analysis)
        self.btn_analyze.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(split_cfg, text="Cílová složka:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.split_dst_entry = ttk.Entry(split_cfg, width=45)
        self.split_dst_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        self.split_dst_entry.insert(0, get_base_path())

        btn_dst_browse = ttk.Button(split_cfg, text="Procházet...", command=self.browse_split_dst)
        btn_dst_browse.grid(row=1, column=2, padx=5, pady=5)

        split_cfg.columnconfigure(1, weight=1)

        # Options bar
        opt_frame = ttk.Frame(self.tab_split)
        opt_frame.pack(fill=tk.X, pady=(0, 5))

        btn_check_all = ttk.Button(opt_frame, text="Označit vše", command=self.split_check_all)
        btn_check_all.pack(side=tk.LEFT, padx=(0, 5))

        btn_uncheck_all = ttk.Button(opt_frame, text="Odznačit vše", command=self.split_uncheck_all)
        btn_uncheck_all.pack(side=tk.LEFT, padx=(0, 15))

        btn_edit_selected = ttk.Button(opt_frame, text="✏️ Přejmenovat vybraný soubor", command=self.edit_selected_split_filename)
        btn_edit_selected.pack(side=tk.LEFT, padx=(0, 15))

        self.split_export_cover_var = tk.BooleanVar(value=True)
        cb_cover = ttk.Checkbutton(opt_frame, text="Exportovat Seznam příloh / Úvodní část", variable=self.split_export_cover_var)
        cb_cover.pack(side=tk.LEFT)

        # Sections Table
        table_frame = ttk.LabelFrame(self.tab_split, text=" Nalezené části dokumentace ", padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        columns = ("export", "name", "code", "pages", "filename")
        self.split_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=8)
        
        self.split_tree.heading("export", text="Export")
        self.split_tree.heading("name", text="Část dokumentace")
        self.split_tree.heading("code", text="Kód")
        self.split_tree.heading("pages", text="Strany")
        self.split_tree.heading("filename", text="Výstupní název souboru (Dvojklik / Tlačítko pro úpravu)")

        self.split_tree.column("export", width=60, anchor="center")
        self.split_tree.column("name", width=220, anchor="w")
        self.split_tree.column("code", width=60, anchor="center")
        self.split_tree.column("pages", width=90, anchor="center")
        self.split_tree.column("filename", width=300, anchor="w")

        tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.split_tree.yview)
        self.split_tree.configure(yscrollcommand=tree_scroll.set)

        self.split_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.split_tree.bind("<Button-1>", self.on_split_tree_click)
        self.split_tree.bind("<Double-1>", self.on_split_tree_double_click)

        # Log & Action frame
        split_log_frame = ttk.LabelFrame(self.tab_split, text=" Průběh rozdělování ", padding="5")
        split_log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.split_log_text = ScrolledText(split_log_frame, state='disabled', height=6, font=("Consolas", 9))
        self.split_log_text.pack(fill=tk.BOTH, expand=True)

        split_action_frame = ttk.Frame(self.tab_split)
        split_action_frame.pack(fill=tk.X)

        self.split_status_lbl = ttk.Label(split_action_frame, text="Připraven k rozdělení", font=("Segoe UI", 10, "italic"))
        self.split_status_lbl.pack(side=tk.LEFT, pady=5)

        self.btn_open_dst = ttk.Button(split_action_frame, text="📂 Otevřít cílovou složku", command=self.open_split_dst)
        self.btn_open_dst.pack(side=tk.RIGHT, padx=(5, 0))

        self.btn_run_split = ttk.Button(split_action_frame, text="▶ Rozdělit PDF dokument", command=self.start_split_execution)
        self.btn_run_split.pack(side=tk.RIGHT, padx=5)

        self.split_detected_sections = []

    def setup_logging(self):
        if not logger.handlers:
            handler = TextHandler(self.log_text)
            handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
            logger.addHandler(handler)
            
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s"))
            logger.addHandler(console_handler)

    def _split_log(self, msg):
        def append():
            self.split_log_text.configure(state='normal')
            self.split_log_text.insert('end', msg + '\n')
            self.split_log_text.see('end')
            self.split_log_text.configure(state='disabled')
        self.split_log_text.after(0, append)

    def _split_log_clear(self):
        self.split_log_text.configure(state='normal')
        self.split_log_text.delete('1.0', tk.END)
        self.split_log_text.configure(state='disabled')

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

    def browse_split_src(self):
        initial_dir = os.path.dirname(self.split_src_entry.get()) if self.split_src_entry.get() else get_base_path()
        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Vyberte zdrojový PDF soubor k rozdělení",
            filetypes=[("PDF soubory (*.pdf)", "*.pdf"), ("Všechny soubory", "*.*")]
        )
        if file_path:
            norm_path = os.path.normpath(file_path)
            self.split_src_entry.delete(0, tk.END)
            self.split_src_entry.insert(0, norm_path)
            self.start_split_analysis()

    def browse_split_dst(self):
        initial_dir = self.split_dst_entry.get() if self.split_dst_entry.get() else get_base_path()
        dir_path = filedialog.askdirectory(initialdir=initial_dir, title="Vyberte složku pro uložení rozdělených PDF")
        if dir_path:
            self.split_dst_entry.delete(0, tk.END)
            self.split_dst_entry.insert(0, os.path.normpath(dir_path))

    def split_check_all(self):
        for item in self.split_detected_sections:
            item['export'] = True
        self.render_split_tree()

    def split_uncheck_all(self):
        for item in self.split_detected_sections:
            item['export'] = False
        self.render_split_tree()

    def edit_selected_split_filename(self):
        selected_item = self.split_tree.selection()
        if not selected_item:
            messagebox.showinfo("Výběr sekce", "Vyberte v tabulce řádek sekce, kterou chcete přejmenovat.")
            return
        idx = int(selected_item[0])
        if 0 <= idx < len(self.split_detected_sections):
            sec = self.split_detected_sections[idx]
            from tkinter.simpledialog import askstring
            new_name = askstring(
                "Přejmenování souboru",
                f"Zadejte nový výstupní název pro část '{sec['name']}' ({sec['code']}):",
                initialvalue=sec['filename']
            )
            if new_name and new_name.strip():
                clean_name = new_name.strip()
                if not clean_name.lower().endswith('.pdf'):
                    clean_name += '.pdf'
                sec['filename'] = clean_name
                self.render_split_tree()

    def on_split_tree_click(self, event):
        region = self.split_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.split_tree.identify_column(event.x)
            if column == "#1":
                item_id = self.split_tree.identify_row(event.y)
                if item_id:
                    idx = int(item_id)
                    if 0 <= idx < len(self.split_detected_sections):
                        self.split_detected_sections[idx]['export'] = not self.split_detected_sections[idx]['export']
                        self.render_split_tree()

    def on_split_tree_double_click(self, event):
        item_id = self.split_tree.identify_row(event.y)
        if item_id:
            idx = int(item_id)
            if 0 <= idx < len(self.split_detected_sections):
                sec = self.split_detected_sections[idx]
                from tkinter.simpledialog import askstring
                new_name = askstring(
                    "Přejmenování souboru",
                    f"Zadejte nový výstupní název pro část '{sec['name']}' ({sec['code']}):",
                    initialvalue=sec['filename']
                )
                if new_name and new_name.strip():
                    clean_name = new_name.strip()
                    if not clean_name.lower().endswith('.pdf'):
                        clean_name += '.pdf'
                    sec['filename'] = clean_name
                    self.render_split_tree()

    def start_split_analysis(self):
        src_path = self.split_src_entry.get().strip()
        if not src_path or not os.path.exists(src_path):
            messagebox.showerror("Chyba souboru", "Vyberte platný zdrojový PDF soubor k rozdělení.")
            return

        self._split_log_clear()
        self._split_log(f"Zahajuji analýzu struktury PDF: {os.path.basename(src_path)}")
        self.split_status_lbl.config(text="Skenuji strukturu PDF...")

        def _worker():
            try:
                sections = analyze_pdf_structure_for_split(src_path)
                for sec in sections:
                    sec['export'] = True
                self.split_detected_sections = sections
                self.root.after(0, self._render_analysis_results)
            except Exception as e:
                self.root.after(0, lambda: self._handle_analysis_error(str(e)))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _render_analysis_results(self):
        self.render_split_tree()
        count = len(self.split_detected_sections)
        self._split_log(f"Detekováno {count} částí dokumentace.")
        self.split_status_lbl.config(text=f"Analýza dokončena: Nalezeno {count} částí dokumentace.")

    def _handle_analysis_error(self, err_msg):
        self._split_log(f"❌ Chyba při analýze PDF: {err_msg}")
        self.split_status_lbl.config(text="Chyba při analýze PDF.")
        messagebox.showerror("Chyba analýzy", f"Nepodařilo se prozkoumat strukturu PDF:\n{err_msg}")

    def render_split_tree(self):
        for item in self.split_tree.get_children():
            self.split_tree.delete(item)

        for idx, sec in enumerate(self.split_detected_sections):
            check_str = "☑ Ano" if sec.get('export', True) else "☐ Ne"
            pages_str = f"{sec['start_page']}–{sec['end_page']}"
            self.split_tree.insert(
                "", "end", iid=str(idx),
                values=(check_str, sec['name'], sec['code'], pages_str, sec['filename'])
            )

    def start_split_execution(self):
        src_path = self.split_src_entry.get().strip()
        dst_dir = self.split_dst_entry.get().strip()

        if not src_path or not os.path.exists(src_path):
            messagebox.showerror("Chyba souboru", "Vyberte platný zdrojový PDF soubor.")
            return

        if not dst_dir:
            messagebox.showerror("Chyba složky", "Zadejte platnou cílovou složku pro uložení.")
            return

        selected_secs = [s for s in self.split_detected_sections if s.get('export', True)]

        if not self.split_export_cover_var.get():
            selected_secs = [s for s in selected_secs if s['code'] != 'COVER']

        if not selected_secs:
            messagebox.showwarning("Žádné sekce", "Nebyly vybrány žádné části dokumentace ke stažení/exportu.")
            return

        self.btn_run_split.config(state='disabled')
        self.btn_analyze.config(state='disabled')
        self.split_status_lbl.config(text="Rozdělování spuštěno...")
        self._split_log(f"Spouštím rozdělení PDF. Cílová složka: {dst_dir}")

        def _worker():
            try:
                import fitz
                os.makedirs(dst_dir, exist_ok=True)
                doc_src = fitz.open(src_path)
                total = len(selected_secs)

                for idx, sec in enumerate(selected_secs):
                    self._split_log(f"Exportuji ({idx+1}/{total}): {sec['filename']} (strany {sec['start_page']}–{sec['end_page']})...")
                    self.root.after(0, lambda i=idx, t=total: self.split_status_lbl.config(text=f"Exportuji ({i+1}/{t})..."))

                    doc_sub = fitz.open()
                    doc_sub.insert_pdf(doc_src, from_page=sec['start_page']-1, to_page=sec['end_page']-1)
                    out_path = os.path.join(dst_dir, sec['filename'])
                    doc_sub.save(out_path)
                    doc_sub.close()
                    self._split_log(f"  ✓ Uloženo: {out_path}")

                doc_src.close()
                self._split_log("🎉 Rozdělení dokumentace bylo úspěšně dokončeno.")
                self.root.after(0, lambda: self._finish_split_execution(True, "Rozdělení bylo dokončeno!"))
            except Exception as e:
                self.root.after(0, lambda err=str(e): self._finish_split_execution(False, f"Chyba: {err}"))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _finish_split_execution(self, success, msg):
        self.btn_run_split.config(state='normal')
        self.btn_analyze.config(state='normal')
        if success:
            self.split_status_lbl.config(text="Rozdělení PDF dokončeno.")
            messagebox.showinfo("Úspěch", f"Rozdělení PDF dokumentu bylo úspěšně dokončeno!\n\nVygenerované soubory naleznete ve složce:\n{self.split_dst_entry.get().strip()}")
        else:
            self.split_status_lbl.config(text="Chyba při rozdělování.")
            messagebox.showerror("Chyba", msg)

    def open_split_dst(self):
        dst_dir = self.split_dst_entry.get().strip()
        if dst_dir and os.path.exists(dst_dir):
            try:
                os.startfile(dst_dir)
            except Exception as e:
                messagebox.showerror("Chyba", f"Nelze otevřít složku:\n{e}")
        else:
            messagebox.showerror("Chyba", f"Složka neexistuje:\n{dst_dir}")

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
