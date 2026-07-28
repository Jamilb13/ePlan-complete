import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, HRFlowable, KeepTogether
)
from reportlab.pdfgen import canvas

class NumberedCanvas(canvas.Canvas):
    """Two-pass canvas to dynamically compute and render 'Page X of Y' page numbers and running headers."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        # Do not draw headers/footers on page 1 (Title/Cover Page)
        if self._pageNumber == 1:
            return

        self.saveState()
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(colors.HexColor("#1F538D"))

        # Running Header Line
        self.setLineWidth(0.75)
        self.setStrokeColor(colors.HexColor("#1F538D"))
        self.line(40, 802, 555, 802)
        self.drawString(40, 807, "ePlan Documentation Merger & Splitter — Uživatelský manuál")

        # Running Footer Line
        self.setLineWidth(0.5)
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.line(40, 45, 555, 45)

        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#555555"))
        self.drawString(40, 32, "AUTEL, a.s.  |  https://github.com/Jamilb13/ePlan-complete")

        page_str = f"Strana {self._pageNumber} z {page_count}"
        self.drawRightString(555, 32, page_str)
        self.restoreState()

def create_callout_box(title, text, bg_color="#EBF3FB", border_color="#1F538D", title_color="#1F538D", style_title=None, style_text=None):
    """Helper to generate styled callout note/tip/warning boxes."""
    content = [
        Paragraph(title, style_title),
        Spacer(1, 4),
        Paragraph(text, style_text)
    ]
    t = Table([[content]], colWidths=[505])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(bg_color)),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor(border_color)),
        ('LINELEFT', (0,0), (0,0), 4, colors.HexColor(border_color)),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    return t

def build_pdf():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_filename = os.path.join(base_dir, "MANUAL.pdf")

    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=50,
        bottomMargin=50
    )

    styles = getSampleStyleSheet()

    # Custom typography styles
    style_cover_title = ParagraphStyle(
        'CoverTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=30,
        alignment=1, # Center
        textColor=colors.HexColor('#1F538D')
    )

    style_cover_subtitle = ParagraphStyle(
        'CoverSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=15,
        leading=20,
        alignment=1, # Center
        textColor=colors.HexColor('#2B7A78')
    )

    style_cover_meta = ParagraphStyle(
        'CoverMeta',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=15,
        alignment=1, # Center
        textColor=colors.HexColor('#555555')
    )

    style_h1 = ParagraphStyle(
        'CustomH1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#1F538D'),
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    )

    style_h2 = ParagraphStyle(
        'CustomH2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#2B7A78'),
        spaceBefore=10,
        spaceAfter=6,
        keepWithNext=True
    )

    style_body = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        textColor=colors.HexColor('#222222'),
        spaceAfter=6
    )

    style_body_bold = ParagraphStyle(
        'CustomBodyBold',
        parent=style_body,
        fontName='Helvetica-Bold'
    )

    style_bullet = ParagraphStyle(
        'CustomBullet',
        parent=style_body,
        leftIndent=15,
        spaceAfter=4
    )

    style_callout_title = ParagraphStyle(
        'CalloutTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#1F538D')
    )

    style_callout_text = ParagraphStyle(
        'CalloutText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#333333')
    )

    style_table_header = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.white,
        alignment=0
    )

    style_table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#222222')
    )

    story = []

    # =========================================================================
    # COVER / TITLE PAGE
    # =========================================================================
    story.append(Spacer(1, 30))
    logo_path = os.path.join(base_dir, "grafika", "logo.png")
    if os.path.exists(logo_path):
        story.append(Image(logo_path, width=380, height=140))

    story.append(Spacer(1, 35))
    story.append(Paragraph("ePlan Documentation Merger & Splitter", style_cover_title))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Grafický uživatelský manuál a příručka", style_cover_subtitle))
    story.append(Spacer(1, 20))

    story.append(HRFlowable(width="80%", thickness=2, color=colors.HexColor('#1F538D'), spaceBefore=10, spaceAfter=25))

    story.append(Spacer(1, 40))

    meta_text = (
        "<b>Aplikace pro kompletaci a rozdělování ePlan dokumentace</b><br/>"
        "Verze 1.0  |  AUTEL, a.s.<br/>"
        "GitHub repozitář: https://github.com/Jamilb13/ePlan-complete"
    )
    story.append(Paragraph(meta_text, style_cover_meta))
    story.append(PageBreak())

    # =========================================================================
    # SECTION 1: OVERVIEW & ARCHITECTURE
    # =========================================================================
    story.append(Paragraph("1. O programu a architektuře zpracování", style_h1))
    story.append(Paragraph(
        "Aplikace <b>ePlan Documentation Merger & Splitter</b> slouží k automatizovanému vyhledávání, "
        "převodu a slučování externích příloh (výkresy DWG, dokumenty MS Word, tabulky MS Excel, obrázky) "
        "odkazovaných z hlavních PDF dokumentů vygenerovaných systémem ePlan. Výsledný kompletní dokument "
        "lze následně automaticky rozdělit na samostatné pojmenované archivní sekce.",
        style_body
    ))
    story.append(Spacer(1, 6))

    infographic_path = os.path.join(base_dir, "grafika", "workflow_infographic.png")
    if os.path.exists(infographic_path):
        story.append(Image(infographic_path, width=505, height=220))
        story.append(Spacer(1, 10))

    story.append(Paragraph("<b>Hlavní vlastnosti nástroje:</b>", style_body))
    story.append(Paragraph("• <b>Automatická detekce odkazů:</b> Skenuje akce typu Launch v PDF a vyhledává odkazované soubory.", style_bullet))
    story.append(Paragraph("• <b>Převod MS Office:</b> Dávková konverze Word a Excel do PDF via nativní COM rozhraní.", style_bullet))
    story.append(Paragraph("• <b>Konverze DWG/DXF:</b> Převod DWG do DXF (via ODA Converter) a inteligentní renderování výkresových rozvržení (Paperspace) nebo Modelspace.", style_bullet))
    story.append(Paragraph("• <b>Ukazatel postupu v reálném čase:</b> Plynulý Progress Bar a stavový text (procenta, aktuální soubor/strana).", style_bullet))
    story.append(Paragraph("• <b>Rychlé otvírání složek:</b> Tlačítka 📂 Otevřít pro přímý přechod do zdrojového i cílového adresáře.", style_bullet))
    story.append(Paragraph("• <b>Oboustranná synchronizace:</b> Automatické sdílení cest a výstupního _complete.pdf mezi záložkami.", style_bullet))

    story.append(Spacer(1, 10))

    # =========================================================================
    # SECTION 2: GUI OVERVIEW
    # =========================================================================
    story.append(Paragraph("2. Uživatelské rozhraní (GUI Overview)", style_h1))
    story.append(Paragraph(
        "Grafické rozhraní je navrženo pro maximální přehlednost a plynulost práce. "
        "Veškeré operace běží na pozadí a stav je zobrazován v logu a na ukazateli postupu.",
        style_body
    ))
    story.append(Spacer(1, 6))

    gui_overview_path = os.path.join(base_dir, "grafika", "gui_overview.png")
    if os.path.exists(gui_overview_path):
        story.append(Image(gui_overview_path, width=505, height=225))
        story.append(Spacer(1, 10))

    # =========================================================================
    # SECTION 3: MERGING TAB
    # =========================================================================
    story.append(Paragraph("3. Záložka Slučování PDF (PDF Merger)", style_h1))
    story.append(Paragraph(
        "Tato záložka slouží k sestavení jednoho kompletního PDF dokumentu ze všech odkazovaných příloh.",
        style_body
    ))
    story.append(Paragraph("<b>Postup práce:</b>", style_body_bold))
    story.append(Paragraph("1. <b>Zdrojový adresář:</b> Vyberte složku s hlavními PDF a přílohami. Složku můžete ihned otevřít tlačítkem <b>📂 Otevřít</b>.", style_bullet))
    story.append(Paragraph("2. <b>Zdrojový soubor (volitelně):</b> Můžete vybrat konkrétní PDF nebo seznam (.txt, .csv). Pokud je prázdné, zpracují se všechna PDF v adresáři.", style_bullet))
    story.append(Paragraph("3. <b>Cílový adresář (volitelně):</b> Kam se má výstup uložit. Pokud je prázdný, uloží se vedle zdroje s příponou <i>_complete.pdf</i>.", style_bullet))
    story.append(Paragraph("4. <b>Náhled struktury (Preview):</b> Vykreslí stromový náhled dokumentu a stav odkazů bez spuštění konverze.", style_bullet))
    story.append(Paragraph("5. <b>Spustit slučování (Run Merger):</b> Spustí konverzi a sloučení. Progress bar zobrazuje procenta a zpracovávaný soubor.", style_bullet))

    story.append(Spacer(1, 6))
    note_box = create_callout_box(
        "💡 Užitečný tip – Rychlé otvírání složek",
        "Pomocí tlačítek <b>📂 Otevřít</b> u zdrojového a cílového adresáře se okamžitě přepnete do vybrané složky v Průzkumníku Windows.",
        bg_color="#EBF3FB", border_color="#1F538D", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(note_box)
    story.append(Spacer(1, 10))

    # =========================================================================
    # SECTION 4: SPLITTING TAB
    # =========================================================================
    story.append(Paragraph("4. Záložka Rozdělování PDF (PDF Splitter)", style_h1))
    story.append(Paragraph(
        "Umožňuje rozdělit sloučený PDF dokument na jednotlivé výstupní sekce (např. Technická zpráva, "
        "Souhrnná metodika, Výkaz výměr, Bezpečnostní specifikace) podle archivačních kódů.",
        style_body
    ))
    story.append(Paragraph("<b>Postup práce:</b>", style_body_bold))
    story.append(Paragraph("1. <b>Načtení vstupu:</b> Po dokončení slučování se vytvořený <i>_complete.pdf</i> automaticky předvyplní do pole <b>Zdrojový PDF soubor</b>.", style_bullet))
    story.append(Paragraph("2. <b>Analýza struktury:</b> Kliknutím na <b>🔍 Analýza struktury</b> se automaticky identifikují sekce (&TZ, &SM, &VV, &BS, &TZ1 atd.).", style_bullet))
    story.append(Paragraph("3. <b>Úprava názvů a výběr:</b> V tabulce můžete přímo upravit názvy výstupních souborů a odškrtnout neužívané částí.", style_bullet))
    story.append(Paragraph("4. <b>Spuštění exportu:</b> Klikněte na <b>▶ Rozdělit PDF dokument</b>. Průběh exportu sekcí je zobrazen na ukazateli postupu.", style_bullet))

    story.append(Spacer(1, 10))

    # =========================================================================
    # SECTION 5: FORMAT CONVERSIONS & TROUBLESHOOTING
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("5. Přehled konverzí formátů a Řešení problémů", style_h1))

    table_data = [
        [Paragraph("Formát přílohy", style_table_header), Paragraph("Způsob zpracování", style_table_header), Paragraph("Požadavky a poznámka", style_table_header)],
        [Paragraph("<b>PDF</b>", style_table_cell), Paragraph("Přímé sloučení bez konverze", style_table_cell), Paragraph("Původní záložky se zachovají a aktualizují", style_table_cell)],
        [Paragraph("<b>MS Word</b> (.docx, .doc)", style_table_cell), Paragraph("Nativní COM export přes Word.Application", style_table_cell), Paragraph("Vyžaduje MS Office (Word) na počítači", style_table_cell)],
        [Paragraph("<b>MS Excel</b> (.xlsx, .xls)", style_table_cell), Paragraph("Nativní COM export přes Excel.Application", style_table_cell), Paragraph("Vyžaduje MS Office (Excel) na počítači", style_table_cell)],
        [Paragraph("<b>DWG / DXF výkresy</b>", style_table_cell), Paragraph("Převod DWG->DXF (ODA) + render Layouts nebo Modelspace", style_table_cell), Paragraph("Funguje samostatně bez nutnosti mat AutoCAD", style_table_cell)],
        [Paragraph("<b>Obrázky</b> (PNG, JPG, BMP)", style_table_cell), Paragraph("Převod obrázku na PDF stránku a vložení", style_table_cell), Paragraph("Automatické přizpůsobení velikosti stránky", style_table_cell)],
    ]
    t_conv = Table(table_data, colWidths=[120, 205, 180])
    t_conv.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F538D')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8F9FA')]),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_conv)

    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>Řešení častých dotazů (Troubleshooting):</b>", style_h2))

    warn_box1 = create_callout_box(
        "⚠️ Varování Windows SmartScreen při spuštění",
        "Pokud systém Windows zobrazí upozornění SmartScreen při spuštění přenosné verze (.exe), klikněte na tlačítko <b>Více informací</b> a následně <b>Spustit i přesto</b>.",
        bg_color="#FFF8E7", border_color="#E6A100", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(warn_box1)
    story.append(Spacer(1, 8))

    warn_box2 = create_callout_box(
        "🛠 Chyba 'Target file not found anywhere in source directory'",
        "Odkazovaný soubor se nepodařilo najít. Ujistěte se, že se soubor jmenuje <b>přesně tak</b>, jak je uvedeno v odkazu v PDF, a že se nachází ve zdrojovém adresáři nebo v jeho podsložkách.",
        bg_color="#FDF2F2", border_color="#DE350B", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(warn_box2)

    story.append(Spacer(1, 20))
    story.append(Paragraph("<b>🌐 GitHub Repozitář:</b>", style_h2))
    story.append(Paragraph(
        "Zdrojový kód, aktualizace a vývojový repozitář naleznete na:<br/>"
        "<b>https://github.com/Jamilb13/ePlan-complete</b>",
        style_body
    ))

    # Build document
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"PDF Manual successfully created at: {pdf_filename}")

if __name__ == "__main__":
    build_pdf()
