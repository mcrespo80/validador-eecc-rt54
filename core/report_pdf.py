import os
from datetime import datetime
from typing import Dict, Any
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Fonts (Century Gothic)
font_dir = r"C:\Windows\Fonts"
if os.path.exists(os.path.join(font_dir, "GOTHIC.TTF")):
    pdfmetrics.registerFont(TTFont('CenturyGothic', os.path.join(font_dir, "GOTHIC.TTF")))
    pdfmetrics.registerFont(TTFont('CenturyGothic-Bold', os.path.join(font_dir, "GOTHICB.TTF")))
    pdfmetrics.registerFont(TTFont('CenturyGothic-Italic', os.path.join(font_dir, "GOTHICI.TTF")))
    pdfmetrics.registerFont(TTFont('CenturyGothic-BoldItalic', os.path.join(font_dir, "GOTHICBI.TTF")))
    FONT_REG = 'CenturyGothic'
    FONT_BOLD = 'CenturyGothic-Bold'
    FONT_ITA = 'CenturyGothic-Italic'
else:
    FONT_REG = 'Helvetica'
    FONT_BOLD = 'Helvetica-Bold'
    FONT_ITA = 'Helvetica-Oblique'

# Brand Palette (MIC DIGITAL CONSULTING)
NAVY_PRIMARY = colors.HexColor("#0D1F3C")
NAVY_SECONDARY = colors.HexColor("#1A365D")
NAVY_LIGHT = colors.HexColor("#F0F4F8")
GOLD_PRIMARY = colors.HexColor("#C59B27")
GOLD_MUTED = colors.HexColor("#B58D20")
GOLD_TINT = colors.HexColor("#FDFBF7")
TEXT_DARK = colors.HexColor("#1E293B")
TEXT_MUTED = colors.HexColor("#64748B")
BG_LIGHT_ROW = colors.HexColor("#F8FAFC")
LINE_COLOR = colors.HexColor("#E2E8F0")
WHITE = colors.HexColor("#FFFFFF")

STATUS_GREEN_BG = colors.HexColor("#ECFDF5")
STATUS_GREEN_TXT = colors.HexColor("#065F46")
STATUS_YELLOW_BG = colors.HexColor("#FFFBEB")
STATUS_YELLOW_TXT = colors.HexColor("#92400E")
STATUS_RED_BG = colors.HexColor("#FEF2F2")
STATUS_RED_TXT = colors.HexColor("#991B1B")


class NumberedCanvas(canvas.Canvas):
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
        self.saveState()
        w, h = A4

        # 1. Top Decorative Bar
        self.setFillColor(NAVY_PRIMARY)
        self.rect(0, h - 4*mm, w, 4*mm, fill=True, stroke=False)
        self.setFillColor(GOLD_PRIMARY)
        self.rect(0, h - 5*mm, w, 1*mm, fill=True, stroke=False)

        # 2. Bottom Decorative Footer
        footer_y = 13 * mm
        self.setStrokeColor(LINE_COLOR)
        self.setLineWidth(0.7)
        self.line(15*mm, footer_y + 4.5*mm, w - 15*mm, footer_y + 4.5*mm)

        self.setFillColor(GOLD_PRIMARY)
        self.circle(15*mm, footer_y + 4.5*mm, 1.2, fill=True, stroke=False)
        self.circle(w - 15*mm, footer_y + 4.5*mm, 1.2, fill=True, stroke=False)

        # Footer Text
        self.setFont(FONT_BOLD, 8)
        self.setFillColor(NAVY_PRIMARY)
        self.drawString(15*mm, footer_y, "MIC DIGITAL CONSULTING")

        self.setFont(FONT_REG, 7.2)
        self.setFillColor(TEXT_MUTED)
        self.drawString(56*mm, footer_y, "•   micdigital.com.ar   •   contacto@micdigital.com.ar   •   Tel: 156-111-7714")

        page_str = f"Página {self._pageNumber} de {page_count}"
        self.setFont(FONT_BOLD, 8)
        self.setFillColor(NAVY_PRIMARY)
        self.drawRightString(w - 15*mm, footer_y, page_str)

        self.restoreState()


def generate_audit_report_pdf(validation_result: Dict[str, Any], output_path: str, logo_path: str = "data/brand/emblem_clean.png") -> str:
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=15*mm,
        rightMargin=15*mm,
        topMargin=11*mm,
        bottomMargin=19*mm
    )

    content_width = A4[0] - 30*mm
    story = []

    # Styles
    brand_title_style = ParagraphStyle('BrandTitle', fontName=FONT_REG, fontSize=13, leading=16, textColor=NAVY_PRIMARY)
    brand_sub_style = ParagraphStyle('BrandSub', fontName=FONT_BOLD, fontSize=6.5, leading=8.5, textColor=GOLD_MUTED)
    doc_title_style = ParagraphStyle('DocTitle', fontName=FONT_BOLD, fontSize=10, leading=13, textColor=NAVY_PRIMARY, alignment=TA_RIGHT)
    doc_sub_style = ParagraphStyle('DocSub', fontName=FONT_ITA, fontSize=7.2, leading=9.5, textColor=TEXT_MUTED, alignment=TA_RIGHT)

    body_style = ParagraphStyle('Body', fontName=FONT_REG, fontSize=7.2, leading=9.6, textColor=TEXT_DARK)
    body_bold = ParagraphStyle('BodyBold', fontName=FONT_BOLD, fontSize=7.2, leading=9.6, textColor=NAVY_PRIMARY)
    section_head_style = ParagraphStyle('SecHead', fontName=FONT_BOLD, fontSize=8.0, leading=10.5, textColor=WHITE)

    # Header
    logo_img = RLImage(logo_path, width=25*mm, height=20*mm) if os.path.exists(logo_path) else Paragraph("<b>MIC</b>", brand_title_style)
    left_flow = [
        Paragraph("MIC DIGITAL CONSULTING", brand_title_style),
        Spacer(1, 1*mm),
        Paragraph("SERVICIOS CONTABLES Y DE AUTOMATIZACIÓN CON IA", brand_sub_style)
    ]
    right_flow = [
        Paragraph("AUDITORÍA PRE-LEGALIZACIÓN DE EECC", doc_title_style),
        Spacer(1, 0.8*mm),
        Paragraph("Control Normativo RT 54 (NUA) • RT 37 (mod. RT 53) • CPCECABA", doc_sub_style)
    ]
    header_table = Table([[logo_img, left_flow, right_flow]], colWidths=[27*mm, 76.5*mm, 76.5*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5*mm),
        ('LINEBELOW', (0,0), (-1,-1), 0.6, LINE_COLOR),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 2.5*mm))

    # Summary Banner
    summary = validation_result["summary"]
    meta = validation_result["entity_metadata"]
    finfo = validation_result["file_info"]

    semaforo = summary["semaforo"]
    if semaforo == "VERDE":
        bg_col, border_col, txt_col = STATUS_GREEN_BG, STATUS_GREEN_TXT, STATUS_GREEN_TXT
    elif semaforo == "AMARILLO":
        bg_col, border_col, txt_col = STATUS_YELLOW_BG, STATUS_YELLOW_TXT, STATUS_YELLOW_TXT
    else:
        bg_col, border_col, txt_col = STATUS_RED_BG, STATUS_RED_TXT, STATUS_RED_TXT

    dictamen_text = (
        f"<b>DICTAMEN DE PRE-LEGALIZACIÓN: {summary['dictamen']}</b><br/>"
        f"Puntaje de Cumplimiento Normativo: <b>{summary['score']}/100</b> &nbsp;|&nbsp; "
        f"Checks Aprobados: <b>{summary['passed']}</b> &nbsp;|&nbsp; "
        f"Advertencias: <b>{summary['warnings']}</b> &nbsp;|&nbsp; "
        f"Errores Críticos: <b>{summary['errors']}</b>"
    )
    t_dictamen = Table([[Paragraph(dictamen_text, ParagraphStyle('Dict', fontName=FONT_REG, fontSize=8.0, leading=11.5, textColor=txt_col))]], colWidths=[content_width])
    t_dictamen.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg_col),
        ('BOX', (0,0), (-1,-1), 1.0, border_col),
        ('TOPPADDING', (0,0), (-1,-1), 2.2*mm),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.2*mm),
        ('LEFTPADDING', (0,0), (-1,-1), 3.5*mm),
        ('RIGHTPADDING', (0,0), (-1,-1), 3.5*mm),
    ]))
    story.append(t_dictamen)
    story.append(Spacer(1, 2.5*mm))

    def make_banner(title):
        t = Table([[Paragraph(title.upper(), section_head_style)]], colWidths=[content_width])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), NAVY_PRIMARY),
            ('TOPPADDING', (0,0), (-1,-1), 1.5*mm),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1.5*mm),
            ('LEFTPADDING', (0,0), (-1,-1), 2.5*mm),
        ]))
        return t

    # Entity & File Info Table
    story.append(make_banner("Datos Identificatorios Extraídos del Documento"))
    info_rows = [
        [Paragraph("Archivo Analizado:", body_bold), Paragraph(f"{finfo['file_name']} ({finfo['format']} - {finfo['file_size_mb']} MB)", body_style),
         Paragraph("Fecha de Análisis:", body_bold), Paragraph(datetime.now().strftime("%d/%m/%Y %H:%M"), body_style)],
        [Paragraph("Denominación Entidad:", body_bold), Paragraph(str(meta.get("denominacion", "-")), body_style),
         Paragraph("CUIT Detectada:", body_bold), Paragraph(str(meta.get("cuit", "-")), body_style)],
        [Paragraph("Fecha Cierre Ejercicio:", body_bold), Paragraph(str(meta.get("fecha_cierre", "-")), body_style),
         Paragraph("Jurisdicción Objetivo:", body_bold), Paragraph(str(meta.get("jurisdiccion", "-")), body_style)],
        [Paragraph("Tipo de Entidad:", body_bold), Paragraph(str(meta.get("tipo_ente", "-")), body_style),
         Paragraph("Categoría RT 54 (NUA):", body_bold), Paragraph(str(meta.get("categoria_rt54", "-")), body_style)],
    ]
    t_info = Table(info_rows, colWidths=[36*mm, 54*mm, 36*mm, 54*mm])
    t_info.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, LINE_COLOR),
        ('BACKGROUND', (0,0), (0,-1), BG_LIGHT_ROW),
        ('BACKGROUND', (2,0), (2,-1), BG_LIGHT_ROW),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.2*mm),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.2*mm),
        ('LEFTPADDING', (0,0), (-1,-1), 2*mm),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 2.5*mm))

    # Detailed Checklist Table
    story.append(make_banner("Detalle de Controles Normativos y Aritméticos (RT 54 / RT 37 / CPCECABA)"))

    check_rows = [[
        Paragraph("Cód.", body_bold),
        Paragraph("Estado", body_bold),
        Paragraph("Control y Diagnóstico Técnico", body_bold),
        Paragraph("Marco Normativo", body_bold)
    ]]

    for c in validation_result["checks"]:
        st = c["status"]
        if st == "PASS":
            badge = "<font color='#065F46'><b>APROBADO</b></font>"
        elif st == "WARNING":
            badge = "<font color='#92400E'><b>ADVERTENCIA</b></font>"
        elif st == "ERROR":
            badge = "<font color='#991B1B'><b>RECHAZO / ERROR</b></font>"
        else:
            badge = "<font color='#475569'><b>INFO</b></font>"

        detail_html = f"<b>{c['title']}</b><br/>{c['detail']}"
        if c.get("recommendation"):
            detail_html += f"<br/><font color='#1A365D'><i>Acción sugerida: {c['recommendation']}</i></font>"

        check_rows.append([
            Paragraph(c["code"], body_bold),
            Paragraph(badge, body_style),
            Paragraph(detail_html, body_style),
            Paragraph(c.get("norm_ref", ""), ParagraphStyle('Ref', fontName=FONT_ITA, fontSize=6.5, leading=8.5, textColor=TEXT_MUTED))
        ])

    t_checks = Table(check_rows, colWidths=[16*mm, 25*mm, 105*mm, 34*mm], repeatRows=1)
    t_checks.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.4, LINE_COLOR),
        ('BACKGROUND', (0,0), (-1,0), NAVY_LIGHT),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 1.4*mm),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.4*mm),
        ('LEFTPADDING', (0,0), (-1,-1), 1.8*mm),
        ('RIGHTPADDING', (0,0), (-1,-1), 1.8*mm),
    ]))
    story.append(t_checks)

    doc.build(story, canvasmaker=NumberedCanvas)
    return output_path
