# -*- coding: utf-8 -*-
# templates/pdf/pdf_Hinderniskarte.py
# Hinderniskarte: R-Nr. (backNumber), Reiter/Pferd, Wertungsspalten, Richterbox
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageTemplate, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
from reportlab.pdfgen import canvas
from datetime import datetime
import os
from PIL import Image as PILImage

WEEKDAY_MAP = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag",
    "Sunday": "Sonntag"
}

MONTH_MAP = {
    "January": "Januar", "February": "Februar", "March": "März", "April": "April",
    "May": "Mai", "June": "Juni", "July": "Juli", "August": "August",
    "September": "September", "October": "Oktober", "November": "November", "December": "Dezember"
}

def _fmt_header_datetime(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        weekday = WEEKDAY_MAP.get(dt.strftime("%A"), dt.strftime("%A"))
        month   = MONTH_MAP.get(dt.strftime("%B"),   dt.strftime("%B"))
        return f"{weekday}, {dt.day}. {month} {dt.year} um {dt.strftime('%H:%M')} Uhr"
    except Exception:
        return str(iso)

def _process_information_text(text):
    if not text:
        return ""
    text = text.lstrip('\n')
    text = text.replace('\n', '<br/>')
    return text

def get_sponsor_bar_height():
    sponsor_path = "logos/sponsorenleiste.png"
    if os.path.exists(sponsor_path):
        try:
            from reportlab.lib.utils import ImageReader
            img = ImageReader(sponsor_path)
            w_px, h_px = img.getSize()
            return (190 * mm) * (h_px / w_px)
        except:
            return 25 * mm
    return 0


class FooterCanvas(canvas.Canvas):
    _print_options = {}

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.sponsor_height = get_sponsor_bar_height() if self._print_options.get("show_sponsor_bar", True) else 0
        self.page_num   = 0
        self.banner_path   = None
        self.banner_height = 0
        if self._print_options.get("show_banner", True):
            bp = os.path.join("logos", "banner.png")
            if os.path.exists(bp):
                try:
                    pil_img = PILImage.open(bp)
                    img_w, img_h = pil_img.size
                    dpi   = pil_img.info.get('dpi', (72, 72))
                    dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
                    if dpi_x > 150:
                        disp_w = img_w * 72.0 / dpi_x
                        disp_h = img_h * 72.0 / dpi_x
                    else:
                        disp_w, disp_h = img_w, img_h
                    self.banner_width  = A4[0]
                    self.banner_height = A4[0] * (disp_h / disp_w)
                    self.banner_path   = bp
                except:
                    pass

    def showPage(self):
        self.draw_banner()
        self.draw_footer()
        self.page_num += 1
        canvas.Canvas.showPage(self)

    def draw_banner(self):
        if self.page_num != 0:
            return
        if self.banner_path and os.path.exists(self.banner_path):
            try:
                self.drawImage(self.banner_path, 0, A4[1] - self.banner_height,
                               width=self.banner_width, height=self.banner_height,
                               preserveAspectRatio=True, mask='auto')
            except:
                pass

    def draw_footer(self):
        if not self._print_options.get("show_sponsor_bar", True):
            return
        sponsor_path = "logos/sponsorenleiste.png"
        if os.path.exists(sponsor_path):
            try:
                img_width  = 190 * mm
                img_height = self.sponsor_height
                x = (A4[0] - img_width) / 2
                self.drawImage(sponsor_path, x, 4 * mm, width=img_width, height=img_height,
                               preserveAspectRatio=True, mask='auto')
            except:
                pass


def render(starterlist: dict, filename: str, logo_max_width_cm: float = 5.0):
    print(f"DEBUG: Starting PDF render (Hinderniskarte) for {filename}")

    # --- Druckoptionen ---
    print_options    = starterlist.get("printOptions", {})
    sponsor_top      = print_options.get("sponsor_top",      False)
    sponsor_bottom   = print_options.get("sponsor_bottom",   False)
    single_sided     = print_options.get("single_sided",     False)
    show_banner      = print_options.get("show_banner",      True)
    show_sponsor_bar = print_options.get("show_sponsor_bar", True)
    show_title       = print_options.get("show_title",       True)
    show_header      = print_options.get("show_header",      True)

    FooterCanvas._print_options = print_options

    has_sponsor_paper    = sponsor_top or sponsor_bottom
    needs_custom_margins = has_sponsor_paper or not show_header

    # --- Doc / Seitenränder ---
    if needs_custom_margins:
        spacing_top_cm    = starterlist.get("spacingTopCm",    3.0)
        spacing_bottom_cm = starterlist.get("spacingBottomCm", 2.0)

        if show_sponsor_bar:
            sponsor_height = get_sponsor_bar_height()
            min_bottom_mm  = (4 * mm + sponsor_height + 1 * mm) / mm
        else:
            min_bottom_mm = 10

        top_margin_front = (spacing_top_cm * 10) if sponsor_top or not show_header else 10.0
        top_margin_back  = 10.0

        bottom_margin_front = max(spacing_bottom_cm * 10, min_bottom_mm) if (sponsor_bottom or not show_header) else min_bottom_mm
        bottom_margin_back  = min_bottom_mm

        if single_sided:
            from reportlab.platypus import BaseDocTemplate

            class CustomDocTemplate(BaseDocTemplate):
                def __init__(self, fname, **kw):
                    self.allowSplitting = 1
                    BaseDocTemplate.__init__(self, fname, **kw)
                    frame_all = Frame(
                        8 * mm, bottom_margin_front * mm,
                        A4[0] - 16 * mm, A4[1] - top_margin_front * mm - bottom_margin_front * mm,
                        id='all', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
                    )
                    self.addPageTemplates([PageTemplate(id='AllPages', frames=[frame_all])])

            doc = CustomDocTemplate(filename, pagesize=A4, rightMargin=0, leftMargin=0, topMargin=0, bottomMargin=0)
        else:
            from reportlab.platypus import BaseDocTemplate

            class CustomDocTemplate(BaseDocTemplate):
                def __init__(self, fname, **kw):
                    self.allowSplitting = 1
                    BaseDocTemplate.__init__(self, fname, **kw)
                    frame_front = Frame(
                        8 * mm, bottom_margin_front * mm,
                        A4[0] - 16 * mm, A4[1] - top_margin_front * mm - bottom_margin_front * mm,
                        id='front', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
                    )
                    frame_back = Frame(
                        8 * mm, bottom_margin_back * mm,
                        A4[0] - 16 * mm, A4[1] - top_margin_back * mm - bottom_margin_back * mm,
                        id='back', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
                    )
                    self.addPageTemplates([
                        PageTemplate(id='Front', frames=[frame_front]),
                        PageTemplate(id='Back',  frames=[frame_back])
                    ])

                def handle_pageBegin(self):
                    self.pageTemplate = self.pageTemplates[0 if self.page % 2 == 0 else 1]
                    BaseDocTemplate.handle_pageBegin(self)

            doc = CustomDocTemplate(filename, pagesize=A4, rightMargin=0, leftMargin=0, topMargin=0, bottomMargin=0)

        page_width = A4[0] - 16 * mm
    else:
        sponsor_height = get_sponsor_bar_height() if show_sponsor_bar else 0
        bottom_margin  = 4 * mm + sponsor_height + 1 * mm if show_sponsor_bar else 8 * mm
        doc = SimpleDocTemplate(
            filename, pagesize=A4,
            leftMargin=8 * mm, rightMargin=8 * mm,
            topMargin=8 * mm, bottomMargin=bottom_margin
        )
        page_width = A4[0] - doc.leftMargin - doc.rightMargin

    # --- Styles ---
    styles = getSampleStyleSheet()
    style_show  = ParagraphStyle("show",  parent=styles["Heading1"], fontSize=14, leading=16, spaceAfter=2)
    style_comp  = ParagraphStyle("comp",  parent=styles["Normal"],   fontSize=11, leading=13, spaceAfter=4)
    style_sub   = ParagraphStyle("sub",   parent=styles["Normal"],   fontSize=10, leading=12, spaceAfter=4)
    style_info  = ParagraphStyle("info",  parent=styles["Normal"],   fontSize=10, leading=12, spaceAfter=4)
    style_hdr   = ParagraphStyle("hdr",   parent=styles["Normal"],   fontSize=7,  alignment=TA_CENTER, leading=8)
    style_hdr_l = ParagraphStyle("hdr_l", parent=styles["Normal"],   fontSize=9,  alignment=TA_LEFT)
    style_pos   = ParagraphStyle("pos",   parent=styles["Normal"],   fontSize=9,  alignment=TA_CENTER)
    style_rider = ParagraphStyle("rider", parent=styles["Normal"],   fontSize=8,  leading=10)
    style_pause = ParagraphStyle("pause", parent=styles["Normal"],   fontSize=9,  alignment=TA_CENTER)
    style_date_r= ParagraphStyle("date_r",parent=styles["Normal"],   fontSize=10, leading=12,
                                 fontName='Helvetica-Bold', alignment=TA_RIGHT)

    elements = []

    # --- Banner-Spacer ---
    if show_banner:
        banner_path = os.path.join("logos", "banner.png")
        if os.path.exists(banner_path):
            try:
                pil_img = PILImage.open(banner_path)
                img_w, img_h = pil_img.size
                dpi   = pil_img.info.get('dpi', (72, 72))
                dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
                if dpi_x > 150:
                    disp_w = img_w * 72.0 / dpi_x
                    disp_h = img_h * 72.0 / dpi_x
                else:
                    disp_w, disp_h = img_w, img_h
                elements.append(Spacer(1, A4[0] * (disp_h / disp_w) - 6 * mm))
            except:
                pass

    # --- Prüfungskopf ---
    logo_path    = starterlist.get("logoPath") if show_header else None
    header_parts = []
    show = starterlist.get("show")        or {}
    comp = starterlist.get("competition") or {}

    if show_header:
        if show_title:
            show_title_text = show.get("title") or starterlist.get("showTitle") or "Unbenannte Veranstaltung"
            header_parts.append(Paragraph(f"<b>{show_title_text}</b>", style_show))

        comp_no         = comp.get("number")         or starterlist.get("competitionNumber") or ""
        comp_title_txt  = comp.get("title")          or starterlist.get("competitionTitle")  or ""
        division        = comp.get("divisionNumber") or starterlist.get("divisionNumber")
        div_text        = ""
        subtitle        = comp.get("subtitle")        or starterlist.get("subtitle")
        informationText = comp.get("informationText") or starterlist.get("informationText")
        location        = comp.get("location")        or starterlist.get("location")

        try:
            if division is not None and str(division) != "" and int(division) > 0:
                div_text = f"{int(division)}. Abt. "
        except:
            div_text = f"{division} " if division else ""

        if comp_no or comp_title_txt:
            comp_line = f"Prüfung {comp_no}"
            if div_text:
                comp_line += f" - {div_text}{comp_title_txt}"
            elif comp_title_txt:
                comp_line += f" - {comp_title_txt}"
            header_parts.append(Paragraph(f"<b>{comp_line}</b>", style_comp))

        if informationText:
            header_parts.append(Paragraph(_process_information_text(informationText), style_info))
        if subtitle:
            header_parts.append(Paragraph(subtitle, style_sub))
    else:
        location = comp.get("location") or starterlist.get("location")

    # Datum/Zeit/Ort
    start_raw    = None
    division_num = starterlist.get('divisionNumber')
    divisions    = starterlist.get('divisions', [])
    if divisions and division_num is not None:
        try:
            div_num = int(division_num)
            for div in divisions:
                if div.get("number") == div_num:
                    ds = div.get("start")
                    if ds:
                        start_raw = ds
                        break
        except (ValueError, TypeError):
            pass
    elif divisions:
        ds = divisions[0].get("start")
        if ds:
            start_raw = ds
    if not start_raw:
        start_raw = starterlist.get('start')

    date_line_text = None
    if start_raw:
        date_line = _fmt_header_datetime(start_raw)
        if location:
            date_line = f"{date_line}  -  {location}"
        date_line_text = date_line

    # Logo + Header
    if logo_path and os.path.exists(logo_path):
        try:
            logo     = Image(logo_path)
            max_size = logo_max_width_cm * 10 * mm
            if logo.drawWidth > max_size or logo.drawHeight > max_size:
                scale = min(max_size / logo.drawWidth, max_size / logo.drawHeight)
                logo.drawWidth  *= scale
                logo.drawHeight *= scale
            logo_col_w = max(logo.drawWidth + 5 * mm, 35 * mm)
            ht = Table([[header_parts, logo]], colWidths=[page_width - logo_col_w, logo_col_w])
            ht.setStyle(TableStyle([
                ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
                ('ALIGN',        (0, 0), (0,  0),  'LEFT'),
                ('ALIGN',        (1, 0), (1,  0),  'RIGHT'),
                ('LEFTPADDING',  (0, 0), (0,  0),  0),
                ('RIGHTPADDING', (1, 0), (1,  0),  0),
            ]))
            elements.append(ht)
        except Exception as e:
            print(f"PDF LOGO DEBUG: {e}")
            for p in header_parts:
                elements.append(p)
    else:
        for p in header_parts:
            elements.append(p)

    elements.append(Spacer(1, 6))

    # Starterliste + Datum
    elements.append(Spacer(1, 1 * mm))
    sl_left = Paragraph("<b>Starterliste</b>", style_comp)
    if date_line_text:
        date_r = Paragraph(f"<b>{date_line_text}</b>", style_date_r)
        so_t = Table([[sl_left, date_r]], colWidths=[page_width * 0.5, page_width * 0.5])
        so_t.setStyle(TableStyle([
            ('VALIGN',        (0, 0), (-1, -1), 'BOTTOM'),
            ('ALIGN',         (0, 0), (0,  0),  'LEFT'),
            ('ALIGN',         (1, 0), (1,  0),  'RIGHT'),
            ('TOPPADDING',    (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ]))
        elements.append(so_t)
    else:
        elements.append(sl_left)
    elements.append(Spacer(1, 2 * mm))

    # --- HINDERNIS + Linie --- gleiche Tabellenstruktur wie Starterliste für exakt gleichen Einzug
    hind_style = ParagraphStyle("hind_combo", parent=styles["Normal"],
                                fontSize=13, leading=15, fontName='Helvetica', spaceAfter=0)
    hind_para = Paragraph(
        "<b>HINDERNIS</b>  <u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u>",
        hind_style
    )
    hind_t = Table([[hind_para, Paragraph("", hind_style)]],
                   colWidths=[page_width * 0.5, page_width * 0.5])
    hind_t.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'BOTTOM'),
        ('ALIGN',         (0, 0), (0,  0),  'LEFT'),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    elements.append(hind_t)
    elements.append(Spacer(1, 2 * mm))

    # --- Starter aufbauen ---
    starters        = starterlist.get("starters") or []
    breaks          = starterlist.get("breaks")   or []
    breaks_by_after = {}
    for b in breaks:
        after_num = b.get("afterNumberInCompetition")
        k = -1 if after_num is None else int(after_num)
        breaks_by_after.setdefault(k, []).append(b)

    # Spaltenbreiten: schmale Wertungsspalten, etwas mehr für außerhalb+Hindernis
    col_rnr    = 13 * mm
    col_reiter = 58 * mm
    col_ff     = 13 * mm   # fehlerfrei
    col_verw   = 14 * mm   # 1.+2. Verweigerung
    col_aussen = 19 * mm   # außerhalb der Flagge
    col_hind   = 18 * mm   # Hindernis ausgel.
    col_sturz  = 13 * mm   # Sturz Reiter/Pferd
    col_aufg   = 13 * mm   # aufgegeben
    col_widths = [
        col_rnr, col_reiter,
        col_ff, col_verw, col_verw,
        col_aussen, col_hind,
        col_sturz, col_sturz, col_aufg
    ]

    # Spaltenüberschrift dynamisch: R-Nr. wenn backNumber vorhanden, K-Nr. wenn cno, sonst Nr.
    has_back_nr = any(s.get("backNumber") is not None for s in starters)
    has_cno     = any((s.get("horses") or [{}])[0].get("cno") for s in starters) if not has_back_nr else False
    if has_back_nr:
        col1_label = "<b>R-Nr.</b>"
    elif has_cno:
        col1_label = "<b>K-Nr.</b>"
    else:
        col1_label = "<b>Nr.</b>"

    # Kopfzeile
    def hdr_p(txt):
        return Paragraph(txt, style_hdr)

    header_row = [
        hdr_p(col1_label),
        Paragraph("<b>Reiter / Rider</b><br/><font size=6>Pferd / Horse</font>", style_hdr_l),
        hdr_p("<b>fehler-\nfrei</b>"),
        hdr_p("<b>1.\nVerwei-\ngerung</b>"),
        hdr_p("<b>2.\nVerwei-\ngerung</b>"),
        hdr_p("<b>außerhalb<br/>der<br/>Flagge</b>"),
        hdr_p("<b>Hindernis\nausge-\nlassen</b>"),
        hdr_p("<b>Sturz\nReiter</b>"),
        hdr_p("<b>Sturz\nPferd</b>"),
        hdr_p("<b>aufge-\ngeben</b>"),
    ]

    data_rows = [header_row]
    meta      = [{"type": "header"}]

    # Pause vor erstem Starter
    if 0 in breaks_by_after:
        for br in breaks_by_after[0]:
            pause_text = br.get("informationText", "Pause")
            data_rows.append([Paragraph(pause_text, style_pause)] + [Paragraph("", style_pos)] * 9)
            meta.append({"type": "pause"})

    current_group = None

    for s in starters:
        # Abteilungs-Header
        starter_group = s.get("groupNumber")
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            data_rows.append(
                [Paragraph(f"<b>Abteilung {starter_group}</b>", style_hdr_l)]
                + [Paragraph("", style_pos)] * 9
            )
            meta.append({"type": "group"})
            current_group = starter_group

        withdrawn     = bool(s.get("withdrawn",    False))
        hors_concours = bool(s.get("horsConcours", False))

        # Reiter + Pferd (nur Name, kein Verein/Zucht)
        athlete    = s.get("athlete") or {}
        rider_name = athlete.get("name", "")

        horses     = s.get("horses") or []
        horse      = horses[0] if horses else {}
        horse_name = horse.get("name", "")

        # Nummer: backNumber → cno (Kopfnummer) → startNumber
        back_nr  = s.get("backNumber")
        start_nr = s.get("startNumber") or ""
        cno_val  = horse.get("cno", "") if horse else ""
        if back_nr is not None:
            rnr_display = str(back_nr)
        elif cno_val:
            rnr_display = str(cno_val)
        else:
            rnr_display = str(start_nr)

        reiter_line = f"<b>{rider_name}</b>"
        if horse_name:
            reiter_line += f"<br/><font size=7>{horse_name}</font>"

        if withdrawn:
            reiter_line = f"<strike>{reiter_line}</strike>"
            rnr_display = f"<strike>{rnr_display}</strike>"

        # AK in fehlerfrei-Spalte
        ff_cell = Paragraph("AK", style_pos) if hors_concours else Paragraph("", style_pos)

        row = (
            [Paragraph(rnr_display, style_pos),
             Paragraph(reiter_line, style_rider),
             ff_cell]
            + [Paragraph("", style_pos)] * 7
        )
        data_rows.append(row)
        meta.append({"type": "starter", "withdrawn": withdrawn})

        # Pausen nach diesem Starter
        try:
            cur = int(start_nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = br.get("informationText", "Pause")
                data_rows.append([Paragraph(pause_text, style_pause)] + [Paragraph("", style_pos)] * 9)
                meta.append({"type": "pause"})

    # --- Tabelle ---
    t = Table(data_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("GRID",          (0, 0), (-1, -1), 0.25, colors.black),
        ("BACKGROUND",    (0, 0), (-1,  0), colors.lightgrey),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ALIGN",         (0, 0), (0,  -1), "CENTER"),
        ("ALIGN",         (2, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])

    starter_count = 0
    for ri in range(1, len(data_rows)):
        m = meta[ri]
        if m["type"] == "group":
            ts.add("SPAN",       (0, ri), (-1, ri))
            ts.add("BACKGROUND", (0, ri), (-1, ri), colors.Color(0.88, 0.88, 0.88))
            ts.add("ALIGN",      (0, ri), (-1, ri), "LEFT")
            starter_count = 0
        elif m["type"] == "pause":
            ts.add("SPAN",  (0, ri), (-1, ri))
            ts.add("ALIGN", (0, ri), (-1, ri), "CENTER")
        elif m["type"] == "starter":
            if starter_count % 2 == 1:
                ts.add("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#f0f0f0"))
            if m.get("withdrawn"):
                ts.add("TEXTCOLOR", (0, ri), (-1, ri), colors.darkgrey)
            starter_count += 1

    t.setStyle(ts)
    elements.append(t)

    # --- Richterbox ---
    elements.append(Spacer(1, 6 * mm))
    richter_style = ParagraphStyle("richter", parent=styles["Normal"], fontSize=9, leading=12)
    richter_para = Paragraph(
        "Name und Telefonnummer Hindernisrichter:",
        richter_style
    )
    # Linie zum Schreiben mit viel Abstand darunter
    linie_style = ParagraphStyle("linie", parent=styles["Normal"], fontSize=9, leading=28)
    linie_para = Paragraph(
        "_" * 70,
        linie_style
    )
    richter_box = Table([[richter_para]], colWidths=[page_width])
    richter_box.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, colors.black),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 28),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ]))
    elements.append(richter_box)

    # --- PDF bauen ---
    print("DEBUG: Building PDF (Hinderniskarte)...")
    try:
        if show_banner or show_sponsor_bar:
            doc.build(elements, canvasmaker=FooterCanvas)
        else:
            doc.build(elements)
        print("DEBUG: PDF Hinderniskarte completed successfully!")
    except Exception as e:
        print(f"DEBUG: PDF build error: {e}")
        raise
