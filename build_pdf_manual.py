import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, HRFlowable, KeepTogether
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# -----------------------------------------------------------------------------
# FONT REGISTRATION (Segoe UI with Unicode Czech support)
# -----------------------------------------------------------------------------
FONT_REG = "SegoeUI"
FONT_BOLD = "SegoeUI-Bold"
FONT_ITALIC = "SegoeUI-Italic"

fonts_dir = "C:/Windows/Fonts"
if os.path.exists(os.path.join(fonts_dir, "segoeui.ttf")):
    pdfmetrics.registerFont(TTFont("SegoeUI", os.path.join(fonts_dir, "segoeui.ttf")))
    pdfmetrics.registerFont(TTFont("SegoeUI-Bold", os.path.join(fonts_dir, "segoeuib.ttf")))
    pdfmetrics.registerFont(TTFont("SegoeUI-Italic", os.path.join(fonts_dir, "segoeuii.ttf")))
else:
    pdfmetrics.registerFont(TTFont("SegoeUI", os.path.join(fonts_dir, "arial.ttf")))
    pdfmetrics.registerFont(TTFont("SegoeUI-Bold", os.path.join(fonts_dir, "arialbd.ttf")))
    pdfmetrics.registerFont(TTFont("SegoeUI-Italic", os.path.join(fonts_dir, "ariali.ttf")))

# -----------------------------------------------------------------------------
# TWO-PASS NUMBERED CANVAS (Page X of Y, Header & Footer)
# -----------------------------------------------------------------------------
class ModernNumberedCanvas(canvas.Canvas):
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
            self.draw_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_decorations(self, page_count):
        # Skip header and footer on Cover / Title page
        if self._pageNumber == 1:
            return

        self.saveState()
        
        # Running Header
        self.setFont(FONT_BOLD, 8)
        self.setFillColor(colors.HexColor("#0F172A"))
        self.drawString(40, 808, "ePlan Documentation Merger & Splitter")
        
        self.setFont(FONT_REG, 8)
        self.setFillColor(colors.HexColor("#0284C7"))
        self.drawRightString(555, 808, "Uživatelská příručka")
        
        self.setLineWidth(0.75)
        self.setStrokeColor(colors.HexColor("#0284C7"))
        self.line(40, 800, 555, 800)

        # Running Footer
        self.setLineWidth(0.5)
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.line(40, 42, 555, 42)

        self.setFont(FONT_REG, 8)
        self.setFillColor(colors.HexColor("#64748B"))
        self.drawString(40, 28, "AUTEL, a.s.  |  https://github.com/Jamilb13/ePlan-complete")

        page_str = f"Strana {self._pageNumber} z {page_count}"
        self.drawRightString(555, 28, page_str)
        self.restoreState()

# -----------------------------------------------------------------------------
# HELPER COMPONENTS
# -----------------------------------------------------------------------------
def create_card_box(title, text, bg_color="#F8FAFC", border_color="#0284C7", style_title=None, style_text=None):
    """Generates a styled visual card container without unsupported emoji characters."""
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

# -----------------------------------------------------------------------------
# MAIN PDF BUILDER
# -----------------------------------------------------------------------------
def build_pdf():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_filename = os.path.join(base_dir, "MANUAL.pdf")

    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=48,
        bottomMargin=48
    )

    styles = getSampleStyleSheet()

    # Custom Typography with Segoe UI (Clean Czech Accents & Proper Spacing)
    style_cover_title = ParagraphStyle(
        'CoverTitle',
        fontName=FONT_BOLD,
        fontSize=24,
        leading=30,
        alignment=1, # Centered
        textColor=colors.HexColor('#0F172A')
    )

    style_cover_subtitle = ParagraphStyle(
        'CoverSubtitle',
        fontName=FONT_BOLD,
        fontSize=12.5,
        leading=17,
        alignment=1, # Centered
        textColor=colors.HexColor('#0284C7')
    )

    style_cover_desc = ParagraphStyle(
        'CoverDesc',
        fontName=FONT_REG,
        fontSize=10.5,
        leading=16,
        alignment=1, # Centered
        textColor=colors.HexColor('#334155')
    )

    style_h1 = ParagraphStyle(
        'CustomH1',
        fontName=FONT_BOLD,
        fontSize=14.5,
        leading=18.5,
        textColor=colors.HexColor('#0F172A'),
        spaceBefore=12,
        spaceAfter=6,
        keepWithNext=True
    )

    style_h2 = ParagraphStyle(
        'CustomH2',
        fontName=FONT_BOLD,
        fontSize=11,
        leading=15,
        textColor=colors.HexColor('#0284C7'),
        spaceBefore=8,
        spaceAfter=5,
        keepWithNext=True
    )

    style_body = ParagraphStyle(
        'CustomBody',
        fontName=FONT_REG,
        fontSize=9.2,
        leading=13.5,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=5
    )

    style_body_bold = ParagraphStyle(
        'CustomBodyBold',
        parent=style_body,
        fontName=FONT_BOLD
    )

    style_bullet = ParagraphStyle(
        'CustomBullet',
        parent=style_body,
        leftIndent=12,
        spaceAfter=3
    )

    style_callout_title = ParagraphStyle(
        'CalloutTitle',
        fontName=FONT_BOLD,
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#0F172A')
    )

    style_callout_text = ParagraphStyle(
        'CalloutText',
        fontName=FONT_REG,
        fontSize=8.8,
        leading=13,
        textColor=colors.HexColor('#334155')
    )

    style_table_header = ParagraphStyle(
        'TableHeader',
        fontName=FONT_BOLD,
        fontSize=9,
        leading=12,
        textColor=colors.white,
        alignment=0
    )

    style_table_cell = ParagraphStyle(
        'TableCell',
        fontName=FONT_REG,
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#1E293B')
    )

    story = []

    # =========================================================================
    # 1. TITULNÍ / ÚVODNÍ STRANA (COVER PAGE)
    # =========================================================================
    story.append(Spacer(1, 10))
    
    # Top Accent Bar
    top_bar = Table([[""]], colWidths=[505], rowHeights=[6])
    top_bar.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#0284C7')),
    ]))
    story.append(top_bar)
    story.append(Spacer(1, 30))

    logo_path = os.path.join(base_dir, "grafika", "logo.png")
    if os.path.exists(logo_path):
        story.append(Image(logo_path, width=360, height=130))

    story.append(Spacer(1, 35))
    story.append(Paragraph("ePlan Documentation Merger &amp; Splitter", style_cover_title))
    story.append(Spacer(1, 10))
    
    # Subtitle Badge Box
    sub_badge = Table([[Paragraph("UŽIVATELSKÁ PŘÍRUČKA A TECHNICKÝ MANUÁL", style_cover_subtitle)]], colWidths=[420])
    sub_badge.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F0F9FF')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#BAE6FD')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(sub_badge)

    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "Nástroj pro automatickou kompletaci, převod odkazovaných příloh (DWG, Word, Excel, obrázky) "
        "a rozdělování projektové dokumentace vygenerované ze systému ePlan.",
        style_cover_desc
    ))

    story.append(Spacer(1, 30))

    # Feature Highlights Grid Box (Clean typography without missing emoji glyphs)
    grid_data = [
        [
            Paragraph("<b>Sloučení příloh</b><br/><font size=8 color='#64748B'>Automatický převod výkresů DWG, dokumentů Word, tabulek Excel i obrázků přímo do PDF.</font>", style_callout_text),
            Paragraph("<b>Ukazatel postupu (Real-time)</b><br/><font size=8 color='#64748B'>Plynulý Progress Bar a detailní informace o stavu zpracování v reálném čase.</font>", style_callout_text)
        ],
        [
            Paragraph("<b>Otevření složek &amp; Synchronizace</b><br/><font size=8 color='#64748B'>Tlačítka pro okamžitý přechod do adresářů v Průzkumníku Windows a sdílení cest.</font>", style_callout_text),
            Paragraph("<b>Detekce sekcí &amp; Export</b><br/><font size=8 color='#64748B'>Automatická identifikace částí dokumentace a dávkové rozdělení na samostatné soubory.</font>", style_callout_text)
        ]
    ]
    grid_table = Table(grid_data, colWidths=[245, 245])
    grid_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E2E8F0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(grid_table)

    story.append(Spacer(1, 40))

    # Footer Metadata Card
    meta_box_data = [[
        Paragraph("<b>Vývoj &amp; Správa:</b> AUTEL, a.s. &nbsp;|&nbsp; <b>Verze:</b> 1.0 (2026)<br/>"
                  "<b>GitHub repozitář:</b> https://github.com/Jamilb13/ePlan-complete", style_callout_text)
    ]]
    meta_table = Table(meta_box_data, colWidths=[505])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(meta_table)

    story.append(PageBreak())

    # =========================================================================
    # 2. O PROGRAMU A ARCHITEKTURA ZPRACOVÁNÍ
    # =========================================================================
    story.append(Paragraph("1. O programu a architektuře zpracování", style_h1))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=8))
    
    story.append(Paragraph(
        "Aplikace <b>ePlan Documentation Merger &amp; Splitter</b> slouží k automatickému vyhledávání, "
        "převodu a sloučení externích příloh (výkresy DWG/DXF, dokumenty MS Word, tabulky MS Excel, obrázky) "
        "odkazovaných z hlavních PDF dokumentů vygenerovaných systémem ePlan. Výsledný kompletní dokument "
        "lze následně automaticky rozdělit na samostatné pojmenované archivní sekce.",
        style_body
    ))
    story.append(Spacer(1, 4))

    infographic_path = os.path.join(base_dir, "grafika", "workflow_infographic.png")
    if os.path.exists(infographic_path):
        img_w = Table([[Image(infographic_path, width=495, height=210)]], colWidths=[505])
        img_w.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E2E8F0')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(img_w)
        story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Hlavní přednosti aplikace:</b>", style_body_bold))
    story.append(Paragraph("• <b>Automatická detekce odkazů:</b> Skenuje akce typu Launch v PDF a vyhledává odkazované soubory.", style_bullet))
    story.append(Paragraph("• <b>Převod MS Office:</b> Dávková konverze Word a Excel do PDF přes nativní COM rozhraní MS Office.", style_bullet))
    story.append(Paragraph("• <b>Konverze DWG/DXF:</b> Převod DWG do DXF (via ODA Converter) a inteligentní renderování výkresových rozvržení (Paperspace) nebo Modelspace.", style_bullet))
    story.append(Paragraph("• <b>Ukazatel postupu v reálném čase:</b> Plynulý Progress Bar a stavový text (procenta, aktuální soubor a strana).", style_bullet))
    story.append(Paragraph("• <b>Rychlé otvírání složek:</b> Tlačítka <b>Otevřít</b> pro přímý přechod do zdrojového i cílového adresáře v Průzkumníku.", style_bullet))
    story.append(Paragraph("• <b>Oboustranná synchronizace:</b> Automatické sdílení cest a vygenerovaného _complete.pdf mezi záložkami.", style_bullet))

    story.append(Spacer(1, 10))

    # =========================================================================
    # 3. UŽIVATELSKÉ ROZHRANÍ (GUI OVERVIEW)
    # =========================================================================
    story.append(Paragraph("2. Popis uživatelského rozhraní", style_h1))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=8))
    
    story.append(Paragraph(
        "Grafické rozhraní je navrženo pro maximální přehlednost a plynulost práce. "
        "Veškeré náročné operace běží v samostatném vlákně na pozadí, takže okno programu nezamrzá a stav "
        "je v reálném čase zobrazován na ukazateli postupu i v logovacím okně.",
        style_body
    ))
    story.append(Spacer(1, 4))

    gui_overview_path = os.path.join(base_dir, "grafika", "gui_overview.png")
    if os.path.exists(gui_overview_path):
        img_gui = Table([[Image(gui_overview_path, width=495, height=210)]], colWidths=[505])
        img_gui.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E2E8F0')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(img_gui)
        story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Aplikace obsahuje tři hlavní záložky:</b>", style_body_bold))
    story.append(Paragraph("1. <b>Slučování:</b> Sloučení hlavního PDF s externími přílohami a zobrazíení stromového náhledu.", style_bullet))
    story.append(Paragraph("2. <b>Rozdělování:</b> Detekce sekcí a export samostatných pojmenovaných PDF dokumentů.", style_bullet))
    story.append(Paragraph("3. <b>Nápověda:</b> Uživatelská příručka, řešení problémů a odkaz na repozitář GitHub.", style_bullet))

    story.append(Spacer(1, 10))

    # =========================================================================
    # 4. ZÁLOŽKA SLUČOVÁNÍ PDF (PDF MERGER)
    # =========================================================================
    story.append(Paragraph("3. Záložka Slučování PDF (PDF Merger)", style_h1))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=8))
    
    story.append(Paragraph(
        "Tato záložka slouží k sestavení kompletního PDF dokumentu se všemi odkazovanými přílohami.",
        style_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("<b>Postup práce:</b>", style_body_bold))
    story.append(Paragraph("1. <b>Zdrojový adresář:</b> Vyberte složku s hlavními PDF a přílohami. Tlačítkem <b>Otevřít</b> můžete složku ihned zobraziť v Průzkumníku Windows.", style_bullet))
    story.append(Paragraph("2. <b>Zdrojový soubor (volitelně):</b> Můžete vybrat konkrétní PDF nebo seznam (.txt, .csv). Pokud je prázdné, zpracují se všechna PDF v adresáři.", style_bullet))
    story.append(Paragraph("3. <b>Cílový adresář (volitelně):</b> Složka pro uložení výstupu. Pokud je prázdná, uloží se vedle zdroje s příponou <i>_complete.pdf</i>.", style_bullet))
    story.append(Paragraph("4. <b>Náhled struktury (Preview):</b> Vykreslí stromový náhled dokumentu s výpisem odkazů a stavu příloh bez spuštění plné konverze.", style_bullet))
    story.append(Paragraph("5. <b>Spustit slučování (Run Merger):</b> Spustí kompletaci na pozadí. Progress bar zobrazuje procenta i právě zpracovávanou stránku.", style_bullet))

    story.append(Spacer(1, 6))
    tip_box = create_card_box(
        "Tip – Rychlé otvírání složek a automatická synchronizace",
        "Pomocí tlačítek <b>Otevřít</b> u zdrojového i cílového adresáře se okamžitě přepnete do vybrané složky. "
        "Po dokončení slučování se vytvořené PDF (např. <i>D23154_complete.pdf</i>) automaticky předvyplní do záložky Rozdělování.",
        bg_color="#F0F9FF", border_color="#0284C7", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(tip_box)
    story.append(Spacer(1, 10))

    # =========================================================================
    # 5. ZÁLOŽKA ROZDĚLOVÁNÍ PDF (PDF SPLITTER)
    # =========================================================================
    story.append(Paragraph("4. Záložka Rozdělování PDF (PDF Splitter)", style_h1))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=8))
    
    story.append(Paragraph(
        "Umožňuje rozdělit kompletní sloučený PDF dokument na samostatné pojmenované archivní sekce "
        "(např. Technická zpráva, Souhrnná metodika, Výkaz výměr, Bezpečnostní specifikace) podle kódů sekcí.",
        style_body
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("<b>Postup práce:</b>", style_body_bold))
    story.append(Paragraph("1. <b>Načtení vstupu:</b> Po dokončení slučování se vygenerovaný soubor <i>_complete.pdf</i> automaticky předvyplní do pole <b>Zdrojový PDF soubor</b>.", style_bullet))
    story.append(Paragraph("2. <b>Analýza struktury:</b> Klikněte na <b>Analýza struktury</b>. Program identifikuje sekce podle záložek i textových skenů (&amp;TZ, &amp;SM, &amp;VV, &amp;BS, &amp;TZ1 atd.).", style_bullet))
    story.append(Paragraph("3. <b>Úprava názvů a výběr:</b> V tabulce můžete přímo upravit výstupní názvy souborů (např. <i>01_D231542633.TZ.pdf</i>) a odškrtnout neužívané sekce.", style_bullet))
    story.append(Paragraph("4. <b>Spuštění exportu:</b> Klikněte na <b>Rozdělit PDF dokument</b>. Progress bar zobrazuje stav exportu jednotlivých souborů.", style_bullet))

    story.append(Spacer(1, 6))
    split_box = create_card_box(
        "Úvodní část a Seznam příloh",
        "Úvodní část dokumentace se Seznamem příloh se automaticky detekuje a exportuje jako samostatný soubor (např. <i>00_D23154_Seznam_příloh.pdf</i>).",
        bg_color="#F8FAFC", border_color="#64748B", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(split_box)

    story.append(Spacer(1, 10))

    # =========================================================================
    # 6. PŘEHLED KONVERZÍ A ŘEŠENÍ PROBLÉMŮ
    # =========================================================================
    story.append(Paragraph("5. Přehled konverzí formátů a Řešení problémů", style_h1))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=8))
    
    table_data = [
        [Paragraph("Formát přílohy", style_table_header), Paragraph("Způsob zpracování", style_table_header), Paragraph("Požadavky a poznámka", style_table_header)],
        [Paragraph("<b>PDF dokumenty</b>", style_table_cell), Paragraph("Přímé sloučení bez konverze", style_table_cell), Paragraph("Původní záložky se zachovají a aktualizují", style_table_cell)],
        [Paragraph("<b>MS Word</b> (.docx, .doc)", style_table_cell), Paragraph("Nativní COM export přes Word.Application", style_table_cell), Paragraph("Vyžaduje MS Office (Word) na počítači", style_table_cell)],
        [Paragraph("<b>MS Excel</b> (.xlsx, .xls)", style_table_cell), Paragraph("Nativní COM export přes Excel.Application", style_table_cell), Paragraph("Vyžaduje MS Office (Excel) na počítači", style_table_cell)],
        [Paragraph("<b>DWG / DXF výkresy</b>", style_table_cell), Paragraph("Převod DWG->DXF (ODA) + render Layouts nebo Modelspace", style_table_cell), Paragraph("Funguje samostatně bez nutnosti mít AutoCAD", style_table_cell)],
        [Paragraph("<b>Obrázky</b> (PNG, JPG, BMP)", style_table_cell), Paragraph("Převod obrázku na PDF stránku a vložení", style_table_cell), Paragraph("Automatické přizpůsobení rozměru stránky", style_table_cell)],
    ]
    t_conv = Table(table_data, colWidths=[125, 200, 180])
    t_conv.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_conv)

    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Řešení častých dotazů (Troubleshooting):</b>", style_h2))

    warn_box1 = create_card_box(
        "Varování Windows SmartScreen při spuštění",
        "Pokud systém Windows zobrazí upozornění SmartScreen při spuštění přenosné verze (.exe), "
        "klikněte na tlačítko <b>Více informací</b> a následně zvolte <b>Spustit i přesto</b>.",
        bg_color="#FFFBEB", border_color="#F59E0B", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(warn_box1)
    story.append(Spacer(1, 6))

    warn_box2 = create_card_box(
        "Chyba 'Target file not found anywhere in source directory'",
        "Odkazovaný soubor se nepodařilo najít. Ujistěte se, že se soubor jmenuje <b>přesně tak</b>, "
        "jak je uvedeno v odkazu v PDF, a že se nachází ve zdrojovém adresáři nebo v jeho podsložkách.",
        bg_color="#FEF2F2", border_color="#EF4444", style_title=style_callout_title, style_text=style_callout_text
    )
    story.append(warn_box2)

    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>GitHub Repozitář a podpora:</b>", style_h2))
    story.append(Paragraph(
        "Zdrojové kódy, dokumentaci a případné aktualizace aplikace naleznete na vývojovém repozitáři:<br/>"
        "<b>https://github.com/Jamilb13/ePlan-complete</b>",
        style_body
    ))

    # Build PDF with Modern Numbered Canvas
    doc.build(story, canvasmaker=ModernNumberedCanvas)
    print(f"Modern PDF Manual successfully compiled at: {pdf_filename}")

if __name__ == "__main__":
    build_pdf()
