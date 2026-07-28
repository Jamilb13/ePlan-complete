import os
import sys
import subprocess
import shutil
import time
import logging
import tempfile
import threading
import urllib.parse
import importlib
import webbrowser

def install_and_import(package_name, import_name=None):
    if import_name is None:
        import_name = package_name
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Instaluji chybějící knihovnu: {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            importlib.import_module(import_name)
        except Exception as e:
            print(f"Chyba při instalaci {package_name}: {e}")
            sys.exit(1)

install_and_import("customtkinter")
install_and_import("pillow", "PIL")

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Set GUI theme and style
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

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
    Eliminates repetitivše os.walk disk scans during link resolution.
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
    Scans a PDF file for launch links and resolvšed target file locations.
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
                    status = "NENALEZENO (Soubor chybí vše zdrojové složce)"

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
    lines.append(f"📋 Umístění: {pdf_path}")
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
                lines.append(f"{child_prefix}├── 📋 Umístění: {item['target_path']}")
            lines.append(f"{child_prefix}├── 📋 Umístění: NENALEZENO")
            lines.append(f"{child_prefix}└── ⚠️ Typ/Stav: {item['status']}")

    lines.append("=================================================================================")
    return "\n".join(lines)

def scan_and_merge_pdf(pdf_path, source_dir, target_dir=None, progress_callback=None):
    """Scans a single PDF file, converts external links, and merges them into a completed PDF with updated outlines."""
    pdf_path = os.path.abspath(pdf_path)
    pdf_dir = os.path.dirname(pdf_path)
    filename_only = os.path.basename(pdf_path)
    
    if filename_only.endswith("_complete.pdf") or "_converted" in filename_only:
        return None

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
        return None

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
    total_src_pages = len(src_doc)
    
    for page_idx in range(total_src_pages):
        if progress_callback:
            try:
                progress_callback(page_idx + 1, total_src_pages)
            except Exception:
                pass

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

    output_path = None
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
            output_path = None
    else:
        logger.info(f"No changes made to {filename_only} (no launch links found or successfully merged).")

    try:
        out_doc.close()
        src_doc.close()
    except Exception:
        pass

    return output_path

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
        if progress_callback:
            progress_callback("Indexuji soubory ve zdrojovém adresáři...", pct=0.02)
        build_source_file_index(source_dir)
        pdf_files = get_pdf_files_to_process(source_dir, source_file)
        
        if not pdf_files:
            logger.warning("Nenalezeny žádné PDF soubory pro zobrazení struktury.")
            if progress_callback:
                progress_callback("Náhled dokončen (Žádná PDF nenalezena)", pct=0.0)
            return

        total_files = len(pdf_files)
        for idx, pdf_file in enumerate(pdf_files):
            if progress_callback:
                progress_callback(f"Generuji náhled ({idx+1}/{total_files}): {os.path.basename(pdf_file)}", pct=(idx+1)/total_files)
            try:
                preview_text = get_pdf_structure_preview(pdf_file, source_dir)
                logger.info("\n" + preview_text)
            except Exception as e:
                logger.error(f"Chyba při generování náhledu pro '{pdf_file}': {e}")

        if progress_callback:
            progress_callback("Náhled struktury dokončen", pct=1.0)
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
        if progress_callback:
            progress_callback("Indexuji soubory ve zdrojovém adresáři...", pct=0.02)
        build_source_file_index(source_dir)

        pdf_files = get_pdf_files_to_process(source_dir, source_file)

        if not pdf_files:
            logger.warning("No target PDF files found.")
            if progress_callback:
                progress_callback("Dokončeno (Žádná PDF nenalezena)", pct=0.0)
            return

        total_files = len(pdf_files)
        logger.info(f"Found {total_files} PDF file(s) to process.")
        created_files = []

        for idx, pdf_file in enumerate(pdf_files):
            filename_only = os.path.basename(pdf_file)
            
            def file_page_callback(curr_p, total_p):
                pct = (idx + (curr_p / max(total_p, 1))) / total_files
                msg = f"Zpracovávám soubor {idx+1}/{total_files}: {filename_only} (strana {curr_p}/{total_p})"
                if progress_callback:
                    progress_callback(msg, pct=pct)

            try:
                out_path = scan_and_merge_pdf(pdf_file, source_dir, target_dir, progress_callback=file_page_callback)
                if out_path:
                    created_files.append(out_path)
            except Exception as e:
                logger.error(f"Error processing PDF '{pdf_file}': {e}. Skipping file and continuing...", exc_info=True)

            if progress_callback:
                progress_callback(f"Dokončen soubor {idx+1}/{total_files}: {filename_only}", pct=(idx+1)/total_files)

        logger.info("Keeping temporary converted files for verification.")
        logger.info("Documentation merging process finished.")
        
        if progress_callback:
            progress_callback("Slučování dokončeno úspěšně!", pct=1.0, created_files=created_files)
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

def get_resource_path(relative_path):
    """Resolvše resource paths for both devšelopment and PyInstaller frozen builds."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


# =========================================================================
# GUI class with tabs for merging, splitting, and help – customtkinter
# =========================================================================

class MergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Settings
        self.title("ePlan Documentation Merger & Splitter")
        self.geometry("950x750")
        self.minsize(800, 600)

        # Force window to front
        self.lift()
        self.attributes("-topmost", True)
        self.after(500, lambda: self.attributes("-topmost", False))

        # Set Window Icon
        try:
            ico_path = get_resource_path(os.path.join("grafika", "ikona.ico"))
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
            else:
                png_path = get_resource_path(os.path.join("grafika", "ikona.png"))
                if os.path.exists(png_path):
                    from PIL import Image, ImageTk
                    icon_photo = ImageTk.PhotoImage(Image.open(png_path))
                    self.wm_iconphoto(True, icon_photo)
                    self._icon_photo_ref = icon_photo
        except Exception as e:
            logger.error(f"Failed to set window icon: {e}")

        # State variables for split
        self.split_detected_sections = []
        self.split_section_rows = []

        # Build UI
        self.create_widgets()

        # Show splash screen overlay
        self.show_splash_overlay()

    def show_splash_overlay(self):
        """Displays splash screen as an overlay frame directly inside main window."""
        self.splash_overlay = ctk.CTkFrame(
            self,
            corner_radius=0,
            fg_color=["#1a1a1a", "#111111"]
        )
        self.splash_overlay.grid(row=0, column=0, sticky="nsew")
        self.splash_overlay.grid_columnconfigure(0, weight=1)
        self.splash_overlay.grid_rowconfigure(0, weight=1)

        # Centered splash card (650x420)
        self.splash_card = ctk.CTkFrame(
            self.splash_overlay,
            width=650,
            height=420,
            corner_radius=16,
            border_width=2,
            border_color=["#1f538d", "#60b0f4"]
        )
        self.splash_card.place(relx=0.5, rely=0.5, anchor="center")
        self.splash_card.grid_columnconfigure(0, weight=1)

        # Logo - enlarged 2x (logo_height = 200)
        try:
            logo_path = get_resource_path(os.path.join("grafika", "logo.png"))
            if os.path.exists(logo_path):
                from PIL import Image
                logo_img_raw = Image.open(logo_path)
                aspect_ratio = logo_img_raw.width / logo_img_raw.height
                logo_height = 200
                logo_width = int(logo_height * aspect_ratio)

                self.splash_logo_img = ctk.CTkImage(
                    light_image=logo_img_raw,
                    dark_image=logo_img_raw,
                    size=(logo_width, logo_height)
                )
                self.splash_logo_lbl = ctk.CTkLabel(
                    self.splash_card,
                    image=self.splash_logo_img,
                    text=""
                )
                self.splash_logo_lbl.pack(pady=(30, 10))
            else:
                raise FileNotFoundError
        except Exception:
            self.splash_logo_lbl = ctk.CTkLabel(
                self.splash_card,
                text="📋 ePlan Merger",
                font=ctk.CTkFont(family="Arial", size=48, weight="bold")
            )
            self.splash_logo_lbl.pack(pady=(40, 10))

        # Title
        ctk.CTkLabel(
            self.splash_card,
            text="ePlan Documentation Merger",
            font=ctk.CTkFont(family="Arial", size=24, weight="bold")
        ).pack(pady=(0, 5))

        # Subtitle
        ctk.CTkLabel(
            self.splash_card,
            text="Spouštění aplikace...",
            font=ctk.CTkFont(size=14, slant="italic"),
            text_color="gray"
        ).pack(pady=(0, 15))

        # Progress bar
        self.splash_progress = ctk.CTkProgressBar(self.splash_card, width=450, height=8)
        self.splash_progress.pack(pady=(0, 25))
        self.splash_progress.set(0)

        # Animation state
        self.splash_progress_val = 0
        self.animate_splash()

    def animate_splash(self):
        if not hasattr(self, 'splash_overlay') or not self.splash_overlay.winfo_exists():
            return
        if self.splash_progress_val < 1.0:
            self.splash_progress_val += 0.04
            if self.splash_progress_val > 1.0:
                self.splash_progress_val = 1.0
            self.splash_progress.set(self.splash_progress_val)
            self.after(35, self.animate_splash)
        else:
            self.after(200, self.hide_splash_overlay)

    def hide_splash_overlay(self):
        if hasattr(self, 'splash_overlay') and self.splash_overlay.winfo_exists():
            self.splash_overlay.destroy()

    def create_widgets(self):
        # Configure Grid Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Tabview ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=15, pady=(15, 15), sticky="nsew")

        self.tab_merge = self.tabview.add("🔗 Slučování")
        self.tab_split = self.tabview.add("✂️ Rozdělování")
        self.tab_help = self.tabview.add("❓ Nápověda")

        self.setup_merge_tab()
        self.setup_split_tab()
        self.setup_help_tab()
        self.setup_logging()

    # =========================================================================
    # MERGE TAB
    # =========================================================================

    def setup_merge_tab(self):
        self.tab_merge.grid_columnconfigure(0, weight=1)
        self.tab_merge.grid_rowconfigure(1, weight=1)

        # --- Configuration Frame ---
        config_frame = ctk.CTkFrame(self.tab_merge)
        config_frame.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="ew")
        config_frame.grid_columnconfigure(1, weight=1)

        # Row 0: Source Directory
        ctk.CTkLabel(config_frame, text="Zdrojový adresář:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(15, 5), pady=10, sticky="w"
        )

        self.source_entry = ctk.CTkEntry(config_frame, placeholder_text="Vyberte adresář se zdrojovými PDF soubory...")
        self.source_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        self.source_entry.insert(0, get_base_path())
        self.source_entry.bind("<FocusOut>", lambda e: self.sync_source_to_split(self.source_entry.get().strip()))

        ctk.CTkButton(
            config_frame, text="Procházet...", width=90, command=self.browse_source
        ).grid(row=0, column=2, padx=5, pady=10, sticky="e")

        ctk.CTkButton(
            config_frame, text="📂 Otevřít", width=90, command=self.open_merge_src
        ).grid(row=0, column=3, padx=(5, 15), pady=10, sticky="e")

        # Row 1: Source File (optional)
        ctk.CTkLabel(config_frame, text="Zdrojový soubor (volitelně):", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=(15, 5), pady=10, sticky="w"
        )

        self.source_file_entry = ctk.CTkEntry(config_frame, placeholder_text="Vyberte konkrétní PDF soubor (nebo ponechte prázdné pro zpracování všech)...")
        self.source_file_entry.grid(row=1, column=1, padx=5, pady=10, sticky="ew")

        ctk.CTkButton(
            config_frame, text="Procházet...", width=90, command=self.browse_source_file
        ).grid(row=1, column=2, padx=(5, 15), pady=10, sticky="e", columnspan=2)

        # Row 2: Target Directory (optional)
        ctk.CTkLabel(config_frame, text="Cílový adresář (volitelně):", font=ctk.CTkFont(weight="bold")).grid(
            row=2, column=0, padx=(15, 5), pady=10, sticky="w"
        )

        self.target_entry = ctk.CTkEntry(config_frame, placeholder_text="Ponechte prázdné pro uložení vedle zdroje s příponou _complete.pdf...")
        self.target_entry.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

        ctk.CTkButton(
            config_frame, text="Procházet...", width=90, command=self.browse_target
        ).grid(row=2, column=2, padx=5, pady=10, sticky="e")

        ctk.CTkButton(
            config_frame, text="📂 Otevřít", width=90, command=self.open_merge_dst
        ).grid(row=2, column=3, padx=(5, 15), pady=10, sticky="e")

        # --- Log Frame ---
        log_frame = ctk.CTkFrame(self.tab_merge)
        log_frame.grid(row=1, column=0, padx=20, pady=(5, 10), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_frame, text="Průběh a náhled struktury:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, padx=15, pady=(10, 2), sticky="w")

        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(family="Courier", size=11),
            state="disabled"
        )
        self.log_text.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")

        # --- Bottom Action Frame ---
        action_frame = ctk.CTkFrame(self.tab_merge, fg_color="transparent")
        action_frame.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            action_frame, text="Připraven",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray"
        )
        self.status_label.grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.merge_progress_bar = ctk.CTkProgressBar(action_frame, orientation="horizontal", height=8)
        self.merge_progress_bar.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.merge_progress_bar.set(0)

        self.preview_button = ctk.CTkButton(
            action_frame,
            text="📋 Náhled struktury (Preview)",
            width=200,
            font=ctk.CTkFont(weight="bold"),
            command=self.start_preview
        )
        self.preview_button.grid(row=2, column=0, sticky="w")

        self.run_button = ctk.CTkButton(
            action_frame,
            text="▶  Spustit slučování (Run Merger)",
            width=220,
            font=ctk.CTkFont(weight="bold"),
            fg_color="#2B7A78",
            hover_color="#175856",
            command=self.start_process
        )
        self.run_button.grid(row=2, column=2, sticky="e")

    # =========================================================================
    # SPLIT TAB
    # =========================================================================

    def setup_split_tab(self):
        self.tab_split.grid_columnconfigure(0, weight=1)
        self.tab_split.grid_rowconfigure(2, weight=1)

        # --- Top Section: Source PDF and Output Folder ---
        top_frame = ctk.CTkFrame(self.tab_split)
        top_frame.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="ew")
        top_frame.grid_columnconfigure(1, weight=1)

        # Row 0: Source PDF
        ctk.CTkLabel(top_frame, text="Zdrojový PDF soubor:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(15, 5), pady=10, sticky="w"
        )

        self.split_src_entry = ctk.CTkEntry(
            top_frame,
            placeholder_text="Vyberte PDF soubor pro rozdělení (např. D23154_V7.1_20250306110226_complete.pdf)..."
        )
        self.split_src_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        self.split_src_entry.bind("<FocusOut>", lambda e: self._on_split_src_focus_out())

        ctk.CTkButton(
            top_frame, text="Procházet...", width=100, command=self.browse_split_src
        ).grid(row=0, column=2, padx=(5, 5), pady=10, sticky="e")

        self.btn_analyze = ctk.CTkButton(
            top_frame, text="🔍 Analýza struktury", width=140,
            font=ctk.CTkFont(weight="bold"),
            fg_color="#1f538d", hover_color="#174170",
            command=self.start_split_analysis
        )
        self.btn_analyze.grid(row=0, column=3, padx=(5, 15), pady=10, sticky="e")

        # Row 1: Destination Folder
        ctk.CTkLabel(top_frame, text="Cílový adresář:", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=(15, 5), pady=(0, 10), sticky="w"
        )

        self.split_dst_entry = ctk.CTkEntry(
            top_frame, placeholder_text="Vyberte složku pro uložení rozdělených PDF..."
        )
        self.split_dst_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=(0, 10), sticky="ew")
        self.split_dst_entry.insert(0, get_base_path())
        self.split_dst_entry.bind("<FocusOut>", lambda e: self.sync_split_to_source(self.split_dst_entry.get().strip()))

        ctk.CTkButton(
            top_frame, text="Procházet...", width=100, command=self.browse_split_dst
        ).grid(row=1, column=3, padx=(5, 15), pady=(0, 10), sticky="e")

        # --- Middle Section: Options & Table Header ---
        mid_frame = ctk.CTkFrame(self.tab_split)
        mid_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        mid_frame.grid_columnconfigure(0, weight=1)

        self.split_title_lbl = ctk.CTkLabel(
            mid_frame, text="Nalezené části dokumentace (vyberte PDF soubor pro analýzu):",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.split_title_lbl.pack(anchor="w", padx=15, pady=(8, 4))

        # Check/Uncheck all & options
        opts_subframe = ctk.CTkFrame(mid_frame, fg_color="transparent")
        opts_subframe.pack(fill="x", padx=15, pady=(0, 5))

        ctk.CTkButton(
            opts_subframe, text="Označit vše", width=100, height=26,
            font=ctk.CTkFont(size=11), command=self.split_check_all
        ).pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            opts_subframe, text="Odznačit vše", width=100, height=26,
            font=ctk.CTkFont(size=11), command=self.split_uncheck_all
        ).pack(side="left", padx=(0, 15))

        self.split_export_cover_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            opts_subframe, text="Exportovat Seznam příloh / Úvodní část",
            variable=self.split_export_cover_var
        ).pack(side="left", padx=(0, 15))

        # Table Header Row
        header_row = ctk.CTkFrame(mid_frame, height=28)
        header_row.pack(fill="x", padx=15, pady=(5, 0))

        ctk.CTkLabel(header_row, text="Export", font=ctk.CTkFont(size=11, weight="bold"), width=55).pack(side="left", padx=(10, 5))
        ctk.CTkLabel(header_row, text="Část dokumentace", font=ctk.CTkFont(size=11, weight="bold"), width=200, anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(header_row, text="Kód", font=ctk.CTkFont(size=11, weight="bold"), width=60, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_row, text="Strany", font=ctk.CTkFont(size=11, weight="bold"), width=80, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_row, text="Výstupní název souboru (klikněte pro úpravu)", font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(side="left", padx=5, fill="x", expand=True)

        # --- Scrollable Sections Table ---
        self.split_sections_scroll_frame = ctk.CTkScrollableFrame(self.tab_split, height=200)
        self.split_sections_scroll_frame.grid(row=2, column=0, padx=20, pady=5, sticky="nsew")

        # --- Bottom: Log & Actions ---
        bottom_frame = ctk.CTkFrame(self.tab_split)
        bottom_frame.grid(row=3, column=0, padx=20, pady=(5, 15), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(2, weight=1)

        self.split_status_lbl = ctk.CTkLabel(
            bottom_frame, text="Připraven k rozdělení",
            font=ctk.CTkFont(size=12, slant="italic"), text_color="gray"
        )
        self.split_status_lbl.grid(row=0, column=0, padx=15, pady=(8, 2), sticky="w")

        self.split_progress_bar = ctk.CTkProgressBar(bottom_frame, orientation="horizontal", height=8)
        self.split_progress_bar.grid(row=1, column=0, padx=10, pady=(2, 5), sticky="ew")
        self.split_progress_bar.set(0)

        self.split_log_text = ctk.CTkTextbox(
            bottom_frame,
            font=ctk.CTkFont(family="Courier", size=11),
            state="disabled",
            height=100
        )
        self.split_log_text.grid(row=2, column=0, padx=10, pady=(5, 5), sticky="nsew")

        # Action buttons
        btn_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        btn_frame.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")

        self.btn_run_split = ctk.CTkButton(
            btn_frame,
            text="▶  Rozdělit PDF dokument",
            font=ctk.CTkFont(weight="bold"),
            fg_color="#2B7A78", hover_color="#175856",
            command=self.start_split_execution
        )
        self.btn_run_split.pack(side="left", padx=(0, 10))

        self.btn_open_dst = ctk.CTkButton(
            btn_frame,
            text="📂 Otevřít cílovou složku",
            command=self.open_split_dst
        )
        self.btn_open_dst.pack(side="left", padx=(0, 10))

    # =========================================================================
    # HELP TAB
    # =========================================================================

    def setup_help_tab(self):
        self.tab_help.grid_columnconfigure(0, weight=1)
        self.tab_help.grid_rowconfigure(0, weight=1)

        help_scroll = ctk.CTkScrollableFrame(self.tab_help)
        help_scroll.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        help_scroll.grid_columnconfigure(0, weight=1)

        # Header
        header_frame = ctk.CTkFrame(help_scroll, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(15, 10))

        ctk.CTkLabel(
            header_frame,
            text="Uživatelská příručka k programu ePlan Documentation Merger",
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Nástroj pro automatické sloučení a kompletaci projektové dokumentace z ePlan.",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray"
        ).pack(anchor="w", pady=(2, 0))

        # --- Section 1: Spuštění ---
        self._help_card(help_scroll, "🚀 Jak program spustit",
            "• Stáhněte a rozbalte archiv ePlan_Documentation_Merger_portable.zip.\n"
            "• Spusťte soubor ePlan_Documentation_Merger_portable.exe.\n"
            "• Pokud se zobrazí upozornění filtru SmartScreen, klikněte na 'Více informací' → 'Spustit i přesto'.\n\n"
            "ℹ️ Pro převod Word (.docx, .doc) a Excel (.xlsx, .xls) je nutný MS Office.\n"
            "   Převod výkresů DWG funguje samostatně bez AutoCADu."
        )

        # --- Section 2: Slučování ---
        self._help_card(help_scroll, "🔗 Slučování PDF – záložka 'Slučování'",
            "• Zdrojový adresář: Složka s hlavními PDF soubory a všemi přílohami.\n"
            "• Zdrojový soubor (volitelně): Konkrétní PDF soubor, nebo ponechte prázdné pro zpracování všech.\n"
            "• Cílový adresář (volitelně): Kam se uloží sloučené PDF. Pokud prázdné, uloží se vedle zdroje s příponou _complete.pdf.\n\n"
            "• Náhled struktury (Preview): Zobrazí stromový náhled dokumentu s přílohami bez spuštění slučování.\n"
            "• Run Merger (Spustit): Spustí kompletní proces slučování na pozadí."
        )

        # --- Section 3: Formáty ---
        self._help_card(help_scroll, "📄 Jak program zpracovává jednotlivé formáty",
            "• PDF dokumenty: Sloučí přímo bez konverze.\n"
            "• Word / Excel: Otevře přes MS Office COM rozhraní, exportuje do PDF.\n"
            "• DWG / DXF výkresy:\n"
            "  - Převede DWG na DXF pomocí ODAFileConverter.\n"
            "  - Pokud výkres obsahuje rozvržení (Paperspace) s výřezy, vyrenderuje je.\n"
            "  - Pokud jsou rozvržení prázdná, vyrenderuje celý Model.\n"
            "• Obrázky (PNG, JPG, BMP, TIFF): Automaticky převede na PDF stránku."
        )

        # --- Section 4: Rozdělování ---
        self._help_card(help_scroll, "✂️ Rozdělování PDF podle částí dokumentace",
            "• Automatická detekce: Identifikuje sekce (&TZ, &SM, &VV, &BS, &TZ1 atd.) podle záložek i textového skenu.\n"
            "• Přesné názvy: Výstupní PDF se pojmenují dle archivačních čísel (např. D231542633.TZ.pdf).\n"
            "• Úvodní část: Seznam příloh se exportuje zvlášť (např. D23154_00_Seznam_příloh.pdf).\n"
            "• Náhled před exportem: V tabulce můžete změnit názvy nebo odškrtnout sekce."
        )

        # --- Section 5: Troubleshooting ---
        self._help_card(help_scroll, "🛠️ Řešení problémů",
            "• Chyba / Přeskočení souboru: Program chybu přeskočí a pokračuje dál.\n"
            "  U chybějícího souboru ponechá původní zástupnou stránku s odkazem.\n"
            "• 'Target file not found': Ujistěte se, že se soubor jmenuje přesně tak, jak je v odkazu,\n"
            "  a nachází se ve zdrojovém adresáři nebo jeho podsložkách."
        )

        # --- Section 6: GitHub Repozitář ---
        card_git = ctk.CTkFrame(help_scroll)
        card_git.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            card_git, text="🌐 GitHub Repozitář a zdrojový kód",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=["#1f538d", "#60b0f4"]
        ).pack(anchor="w", padx=15, pady=(10, 5))

        ctk.CTkLabel(
            card_git,
            text="Projekt, zdrojové kódy a vývojový repozitář naleznete na GitHubu:\nhttps://github.com/Jamilb13/ePlan-complete",
            justify="left", anchor="w",
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=15, pady=(0, 10))

        ctk.CTkButton(
            card_git,
            text="🔗 Otevřít GitHub repozitář",
            font=ctk.CTkFont(weight="bold"),
            width=220,
            command=lambda: webbrowser.open("https://github.com/Jamilb13/ePlan-complete")
        ).pack(anchor="w", padx=15, pady=(0, 15))

    def _help_card(self, parent, title, text):
        """Helper to create a styled help card."""
        card = ctk.CTkFrame(parent)
        card.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            card, text=title,
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=["#1f538d", "#60b0f4"]
        ).pack(anchor="w", padx=15, pady=(10, 5))

        ctk.CTkLabel(
            card, text=text,
            justify="left", anchor="w",
            font=ctk.CTkFont(size=12),
            wraplength=720
        ).pack(anchor="w", padx=15, pady=(0, 15))

    # =========================================================================
    # LOGGING SETUP
    # =========================================================================

    def setup_logging(self):
        if not logger.handlers:
            handler = TextHandler(self.log_text)
            handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
            logger.addHandler(handler)

            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s"))
            logger.addHandler(console_handler)

    # =========================================================================
    # DIRECTORY SYNC & OPEN HELPERS
    # =========================================================================

    def sync_source_to_split(self, src_dir):
        """Synchronizes source directory from Merge tab to Split tab."""
        if src_dir and os.path.exists(src_dir):
            if os.path.isfile(src_dir):
                src_dir = os.path.dirname(src_dir)
            current_split_dst = self.split_dst_entry.get().strip()
            if not current_split_dst or current_split_dst == get_base_path():
                self.split_dst_entry.delete(0, tk.END)
                self.split_dst_entry.insert(0, src_dir)

    def sync_split_to_source(self, dir_path):
        """Synchronizes directory from Split tab to Merge tab."""
        if dir_path and os.path.exists(dir_path):
            if os.path.isfile(dir_path):
                dir_path = os.path.dirname(dir_path)
            current_merge_src = self.source_entry.get().strip()
            if not current_merge_src or current_merge_src == get_base_path():
                self.source_entry.delete(0, tk.END)
                self.source_entry.insert(0, dir_path)

    def _on_split_src_focus_out(self):
        val = self.split_src_entry.get().strip()
        if val and os.path.exists(val):
            src_dir = os.path.dirname(val) if os.path.isfile(val) else val
            self.sync_split_to_source(src_dir)

    def open_merge_src(self):
        src_dir = self.source_entry.get().strip()
        if not src_dir and self.source_file_entry.get().strip():
            src_dir = os.path.dirname(self.source_file_entry.get().strip())
        if src_dir and os.path.exists(src_dir):
            if os.path.isfile(src_dir):
                src_dir = os.path.dirname(src_dir)
            try:
                os.startfile(src_dir)
            except Exception as e:
                messagebox.showerror("Chyba", f"Nelze otevřít složku:\n{e}")
        else:
            messagebox.showerror("Chyba", f"Zdrojová složka neexistuje:\n{src_dir}")

    def open_merge_dst(self):
        dst_dir = self.target_entry.get().strip()
        if not dst_dir:
            dst_dir = self.source_entry.get().strip()
        if dst_dir and os.path.exists(dst_dir):
            if os.path.isfile(dst_dir):
                dst_dir = os.path.dirname(dst_dir)
            try:
                os.startfile(dst_dir)
            except Exception as e:
                messagebox.showerror("Chyba", f"Nelze otevřít složku:\n{e}")
        else:
            messagebox.showerror("Chyba", f"Cílová složka neexistuje nebo nebyla zadána:\n{dst_dir}")

    # =========================================================================
    # SPLIT TAB – logging helpers
    # =========================================================================

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

    # =========================================================================
    # MERGE TAB – browse and action callbacks
    # =========================================================================

    def browse_source(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            norm_dir = os.path.normpath(dir_path)
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, norm_dir)
            self.sync_source_to_split(norm_dir)

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
                self.sync_source_to_split(file_dir)

    def browse_target(self):
        dir_path = filedialog.askdirectory(initialdir=self.source_entry.get())
        if dir_path:
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, os.path.normpath(dir_path))

    def update_merge_progress(self, status, pct=None):
        def _gui_update():
            if status is not None:
                if pct is not None and pct > 0:
                    pct_int = min(100, int(pct * 100))
                    self.status_label.configure(text=f"{pct_int}% - {status}")
                else:
                    self.status_label.configure(text=status)
            if pct is not None:
                self.merge_progress_bar.set(pct)
        self.after(0, _gui_update)

    def update_status(self, text):
        self.update_merge_progress(text)

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

    def process_complete(self, status, pct=1.0, created_files=None):
        def _finish():
            self.run_button.configure(state='normal')
            self.preview_button.configure(state='normal')
            self.update_merge_progress(status, pct=1.0)

            if created_files and len(created_files) > 0:
                first_out = created_files[0]
                if os.path.exists(first_out):
                    self.split_src_entry.delete(0, tk.END)
                    self.split_src_entry.insert(0, os.path.normpath(first_out))
                    dst_folder = os.path.dirname(first_out)
                    self.split_dst_entry.delete(0, tk.END)
                    self.split_dst_entry.insert(0, os.path.normpath(dst_folder))
                    self.start_split_analysis()

            messagebox.showinfo("Hotovo", f"Proces slučování dokončen:\n{status}")
        self.after(0, _finish)

    def preview_complete(self, status, pct=1.0):
        def _finish():
            self.run_button.configure(state='normal')
            self.preview_button.configure(state='normal')
            self.update_merge_progress(status, pct=pct)
        self.after(0, _finish)

    def start_preview(self):
        source_dir, source_file, _ = self.validate_inputs()
        if not source_dir:
            return

        self.run_button.configure(state='disabled')
        self.preview_button.configure(state='disabled')
        self.update_merge_progress("Generuji náhled...", pct=0.0)

        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

        t = threading.Thread(
            target=preview_process,
            args=(source_dir, source_file if source_file else None, lambda status, pct=1.0: self.preview_complete(status, pct=pct))
        )
        t.daemon = True
        t.start()

    def start_process(self):
        source_dir, source_file, target_dir = self.validate_inputs()
        if not source_dir:
            return

        self.run_button.configure(state='disabled')
        self.preview_button.configure(state='disabled')
        self.update_merge_progress("Zpracovávám...", pct=0.0)

        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

        CONVERSION_CACHE.clear()

        def _callback(status, pct=None, created_files=None):
            if pct is not None and pct < 1.0 and not created_files:
                self.update_merge_progress(status, pct=pct)
            else:
                self.process_complete(status, pct=pct if pct is not None else 1.0, created_files=created_files)

        t = threading.Thread(
            target=main_process,
            args=(source_dir, target_dir if target_dir else None, source_file if source_file else None, _callback)
        )
        t.daemon = True
        t.start()

    # =========================================================================
    # SPLIT TAB – browse callbacks
    # =========================================================================

    def browse_split_src(self):
        initial_dir = os.path.dirname(self.split_src_entry.get()) if self.split_src_entry.get() else (self.source_entry.get() if self.source_entry.get() else get_base_path())
        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Vyberte zdrojový PDF soubor k rozdělení",
            filetypes=[("PDF soubory (*.pdf)", "*.pdf"), ("Všechny soubory", "*.*")]
        )
        if file_path:
            norm_path = os.path.normpath(file_path)
            self.split_src_entry.delete(0, tk.END)
            self.split_src_entry.insert(0, norm_path)
            src_dir = os.path.dirname(norm_path)
            self.sync_split_to_source(src_dir)
            self.start_split_analysis()

    def browse_split_dst(self):
        initial_dir = self.split_dst_entry.get() if self.split_dst_entry.get() else (self.source_entry.get() if self.source_entry.get() else get_base_path())
        dir_path = filedialog.askdirectory(initialdir=initial_dir, title="Vyberte složku pro uložení rozdělených PDF")
        if dir_path:
            norm_dir = os.path.normpath(dir_path)
            self.split_dst_entry.delete(0, tk.END)
            self.split_dst_entry.insert(0, norm_dir)
            self.sync_split_to_source(norm_dir)

    # =========================================================================
    # SPLIT TAB – section table management
    # =========================================================================

    def split_check_all(self):
        for item in self.split_detected_sections:
            item['export'] = True
        self.render_split_table()

    def split_uncheck_all(self):
        for item in self.split_detected_sections:
            item['export'] = False
        self.render_split_table()

    def render_split_table(self):
        """Render the sections table using custom CTk widgets inside a scrollable frame."""
        # Clear existing rows
        for widget in self.split_sections_scroll_frame.winfo_children():
            widget.destroy()
        self.split_section_rows = []

        for idx, sec in enumerate(self.split_detected_sections):
            row_frame = ctk.CTkFrame(self.split_sections_scroll_frame, height=32)
            row_frame.pack(fill="x", padx=5, pady=2)

            # Export checkbox
            var = tk.BooleanVar(value=sec.get('export', True))
            cb = ctk.CTkCheckBox(
                row_frame, text="", variable=var, width=30,
                command=lambda i=idx, v=var: self._toggle_section_export(i, v)
            )
            cb.pack(side="left", padx=(10, 5))

            # Name
            ctk.CTkLabel(
                row_frame, text=sec['name'],
                font=ctk.CTkFont(size=12), width=200, anchor="w"
            ).pack(side="left", padx=5)

            # Code
            ctk.CTkLabel(
                row_frame, text=sec['code'],
                font=ctk.CTkFont(size=12, weight="bold"), width=60, anchor="center"
            ).pack(side="left", padx=5)

            # Pages
            pages_str = f"{sec['start_page']}–{sec['end_page']}"
            ctk.CTkLabel(
                row_frame, text=pages_str,
                font=ctk.CTkFont(size=12), width=80, anchor="center"
            ).pack(side="left", padx=5)

            # Filename (editable entry)
            fn_entry = ctk.CTkEntry(row_frame, font=ctk.CTkFont(size=12))
            fn_entry.pack(side="left", padx=5, fill="x", expand=True)
            fn_entry.insert(0, sec['filename'])
            fn_entry.bind("<FocusOut>", lambda e, i=idx, ent=fn_entry: self._update_section_filename(i, ent))
            fn_entry.bind("<Return>", lambda e, i=idx, ent=fn_entry: self._update_section_filename(i, ent))

            self.split_section_rows.append({
                'frame': row_frame,
                'checkbox_var': var,
                'filename_entry': fn_entry
            })

    def _toggle_section_export(self, idx, var):
        if 0 <= idx < len(self.split_detected_sections):
            self.split_detected_sections[idx]['export'] = var.get()

    def _update_section_filename(self, idx, entry_widget):
        if 0 <= idx < len(self.split_detected_sections):
            new_name = entry_widget.get().strip()
            if new_name:
                if not new_name.lower().endswith('.pdf'):
                    new_name += '.pdf'
                self.split_detected_sections[idx]['filename'] = new_name

    # =========================================================================
    # SPLIT TAB – analysis and execution
    # =========================================================================

    def start_split_analysis(self):
        src_path = self.split_src_entry.get().strip()
        if not src_path or not os.path.exists(src_path):
            messagebox.showerror("Chyba souboru", "Vyberte platný zdrojový PDF soubor k rozdělení.")
            return

        self._split_log_clear()
        self._split_log(f"Zahajuji analýzu struktury PDF: {os.path.basename(src_path)}")
        self.split_status_lbl.configure(text="Skenuji strukturu PDF...")

        def _worker():
            try:
                sections = analyze_pdf_structure_for_split(src_path)
                for sec in sections:
                    sec['export'] = True
                self.split_detected_sections = sections
                self.after(0, self._render_analysis_results)
            except Exception as e:
                self.after(0, lambda: self._handle_analysis_error(str(e)))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _render_analysis_results(self):
        self.render_split_table()
        count = len(self.split_detected_sections)
        self._split_log(f"Detekováno {count} částí dokumentace.")
        self.split_status_lbl.configure(text=f"Analýza dokončena: Nalezeno {count} částí dokumentace.")
        self.split_title_lbl.configure(text=f"Nalezené části dokumentace ({count}):")

    def _handle_analysis_error(self, err_msg):
        self._split_log(f"✖ Chyba při analýze PDF: {err_msg}")
        self.split_status_lbl.configure(text="Chyba při analýze PDF.")
        messagebox.showerror("Chyba analýzy", f"Nepodařilo se prozkoumat strukturu PDF:\n{err_msg}")

    def start_split_execution(self):
        src_path = self.split_src_entry.get().strip()
        dst_dir = self.split_dst_entry.get().strip()

        if not src_path or not os.path.exists(src_path):
            messagebox.showerror("Chyba souboru", "Vyberte platný zdrojový PDF soubor.")
            return

        if not dst_dir:
            messagebox.showerror("Chyba složky", "Zadejte platnou cílovou složku pro uložení.")
            return

        # Sync filename entries back to data
        for idx, row in enumerate(self.split_section_rows):
            if idx < len(self.split_detected_sections):
                fn = row['filename_entry'].get().strip()
                if fn:
                    if not fn.lower().endswith('.pdf'):
                        fn += '.pdf'
                    self.split_detected_sections[idx]['filename'] = fn

        selected_secs = [s for s in self.split_detected_sections if s.get('export', True)]

        if not self.split_export_cover_var.get():
            selected_secs = [s for s in selected_secs if s['code'] != 'COVER']

        if not selected_secs:
            messagebox.showwarning("Žádné sekce", "Nebyly vybrány žádné části dokumentace ke stažení/exportu.")
            return

        self.btn_run_split.configure(state='disabled')
        self.btn_analyze.configure(state='disabled')
        self.split_status_lbl.configure(text="Rozdělování spuštěno...")
        self.split_progress_bar.set(0)
        self._split_log(f"Spouštím rozdělení PDF. Cílová složka: {dst_dir}")

        def _worker():
            try:
                import fitz as fitz_split
                os.makedirs(dst_dir, exist_ok=True)
                doc_src = fitz_split.open(src_path)
                total = len(selected_secs)

                for idx, sec in enumerate(selected_secs):
                    pct = (idx + 1) / total
                    msg_log = f"Exportuji ({idx+1}/{total}): {sec['filename']} (strany {sec['start_page']}–{sec['end_page']})..."
                    msg_lbl = f"{int(pct * 100)}% - Exportuji ({idx+1}/{total}): {sec['filename']}..."
                    self._split_log(msg_log)
                    self.after(0, lambda m=msg_lbl, p=pct: self._update_split_progress(m, p))

                    doc_sub = fitz_split.open()
                    doc_sub.insert_pdf(doc_src, from_page=sec['start_page']-1, to_page=sec['end_page']-1)
                    out_path = os.path.join(dst_dir, sec['filename'])
                    doc_sub.save(out_path)
                    doc_sub.close()
                    self._split_log(f"  ✓ Uloženo: {out_path}")

                doc_src.close()
                self._split_log("🎉 Rozdělení dokumentace bylo úspěšně dokončeno.")
                self.after(0, lambda: self._finish_split_execution(True, "Rozdělení bylo dokončeno!"))
            except Exception as e:
                self.after(0, lambda err=str(e): self._finish_split_execution(False, f"Chyba: {err}"))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _update_split_progress(self, msg, pct):
        self.split_status_lbl.configure(text=msg)
        self.split_progress_bar.set(pct)

    def _finish_split_execution(self, success, msg):
        self.btn_run_split.configure(state='normal')
        self.btn_analyze.configure(state='normal')
        if success:
            self.split_progress_bar.set(1.0)
            self.split_status_lbl.configure(text="Rozdělení PDF dokončeno.")
            messagebox.showinfo("Úspěch", f"Rozdělení PDF dokumentu bylo úspěšně dokončeno!\n\nVygenerované soubory naleznete ve složce:\n{self.split_dst_entry.get().strip()}")
        else:
            self.split_progress_bar.set(0.0)
            self.split_status_lbl.configure(text="Chyba při rozdělování.")
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


# =========================================================================
# ENTRY POINT
# =========================================================================

def main():
    app = MergerApp()
    app.mainloop()

if __name__ == "__main__":
    main()

