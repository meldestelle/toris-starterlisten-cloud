# -*- coding: utf-8 -*-
# templates/pdf/pdf_dre_3_logo.py
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, Flowable, PageTemplate, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from datetime import datetime
import os
from io import BytesIO

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

# Mapping für Geschlecht - VERBESSERT
SEX_MAP = {
    "MARE": "Stute",
    "GELDING": "Wallach", 
    "STALLION": "Hengst",
    "M": "Wallach",
    "F": "Stute",
    "": ""
}

class LogoFlowable(Flowable):
    def __init__(self, logo_path):
        self.logo_path = logo_path
        self.width = 35*mm
        self.height = 18*mm
    
    def wrap(self, availWidth, availHeight):
        # Reserviert Platz für das Logo
        return (self.width, self.height)
    
    def draw(self):
        # Zeichnet das Logo an der aktuellen Position
        self.canv.drawImage(self.logo_path, 0, 0, self.width, self.height, preserveAspectRatio=True, mask='auto')

def _safe_get(d, key, default=""):
    if not d:
        return default
    return d.get(key, default) if isinstance(d, dict) else default

def _fmt_time(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%H:%M:%S")  # Erweitert auf HH:MM:SS Format
    except Exception:
        try:
            return str(iso).split("T")[-1][:8]  # Nimm die ersten 8 Zeichen für HH:MM:SS
        except:
            return str(iso)

def _fmt_header_datetime(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        weekday = WEEKDAY_MAP.get(dt.strftime("%A"), dt.strftime("%A"))
        month = MONTH_MAP.get(dt.strftime("%B"), dt.strftime("%B"))
        return f"{weekday}, {dt.day}. {month} {dt.year} um {dt.strftime('%H:%M')} Uhr"
    except Exception:
        return str(iso)

def _format_pause_text(total_seconds, info):
    secs = int(total_seconds or 0)
    if secs == 0:
        return info or "Pause"
    if secs >= 3600:
        h = secs // 3600
        m = (secs % 3600) // 60
        base = f"{h} Std. {m:02d} Min. Pause"
    elif secs >= 60:
        base = f"{secs//60} Minuten Pause"
    else:
        base = f"{secs} Sekunden Pause"
    return f"{base} - {info}" if info else base

def _process_information_text(text):
    """Verarbeitet informationText und konvertiert \n zu <br/> für ReportLab"""
    if not text:
        return ""
    text = text.lstrip('\n')
    text = text.replace('\n', '<br/>')
    return text

def _get_ordered_judge_positions_main_table(judges):
    """Bestimmt die Richter für die Haupttabelle: E H C M B in fester Reihenfolge, maximal 3 Spalten für das 7-Spalten Layout"""
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    available_positions = set()
    judges_by_position = {}
    
    for judge in judges:
        position = judge.get("position", "")
        if isinstance(position, int):
            pos_label = pos_map.get(position, str(position))
        else:
            pos_label = pos_map.get(str(position), str(position))
        
        if pos_label in ["E", "H", "C", "M", "B"]:
            available_positions.add(pos_label)
            judges_by_position[pos_label] = judge
    
    fixed_order = ["E", "H", "C", "M", "B"]
    ordered_positions = [pos for pos in fixed_order if pos in available_positions]
    
    # Fülle auf 3 Spalten auf (für 7-Spalten Layout)
    while len(ordered_positions) < 3:
        ordered_positions.append("")
    
    return ordered_positions[:3]

def _get_ordered_judges_all(judges):
    """Sortiert alle Richter: E H C M B zuerst in fester Reihenfolge, dann alle anderen"""
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    dressage_positions = ["E", "H", "C", "M", "B"]
    dressage_judges = []
    other_judges = []
    
    for judge in judges:
        position = judge.get("position", "")
        if isinstance(position, int):
            pos_label = pos_map.get(position, str(position))
        else:
            pos_label = pos_map.get(str(position), str(position))
        
        judge_with_pos = judge.copy()
        judge_with_pos["pos_label"] = pos_label
        
        if pos_label in dressage_positions:
            dressage_judges.append(judge_with_pos)
        else:
            other_judges.append(judge_with_pos)
    
    ordered_dressage = []
    for pos in dressage_positions:
        for judge in dressage_judges:
            if judge["pos_label"] == pos:
                ordered_dressage.append(judge)
                break
    
    other_judges.sort(key=lambda j: j["pos_label"])
    return ordered_dressage + other_judges


def get_sponsor_bar_height():
    """Ermittelt die tatsächliche Höhe der Sponsorenleiste"""
    sponsor_path = "logos/sponsorenleiste.png"
    if os.path.exists(sponsor_path):
        try:
            from reportlab.lib.utils import ImageReader
            img = ImageReader(sponsor_path)
            img_width_px, img_height_px = img.getSize()
            
            # Zielbreite: 190mm (7.5 inches)
            target_width = 190*mm
            
            # Berechne proportionale Höhe basierend auf Originalverhältnis
            aspect_ratio = img_height_px / img_width_px
            calculated_height = target_width * aspect_ratio
            
            return calculated_height
        except:
            return 25*mm  # Fallback
    return 0  # Keine Sponsorenleiste

class FooterCanvas(canvas.Canvas):
    """Canvas mit Sponsorenleiste im Footer"""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.sponsor_height = get_sponsor_bar_height()
        
    def showPage(self):
        self.draw_footer()
        canvas.Canvas.showPage(self)
        
    def draw_footer(self):
        sponsor_path = "logos/sponsorenleiste.png"
        if os.path.exists(sponsor_path):
            try:
                # Sponsorenleiste unten zentriert - DYNAMISCHE HÖHE
                img_width = 190*mm  # 7.5 inches = ca. 190mm
                img_height = self.sponsor_height  # Dynamisch berechnet
                x = (A4[0] - img_width) / 2  # Zentriert
                y = 7*mm  # 7mm vom unteren Rand
                self.drawImage(sponsor_path, x, y, width=img_width, height=img_height, 
                              preserveAspectRatio=True, mask='auto')
            except:
                pass

def render(starterlist: dict, filename: str, logo_max_width_cm: float = 5.0):
    # Berechne dynamischen bottomMargin basierend auf Sponsorenleiste
    sponsor_height = get_sponsor_bar_height()
    footer_space = 7*mm  # Abstand der Sponsorenleiste vom unteren Rand
    margin_above_sponsor = 3*mm  # 3mm Abstand zwischen Tabelle und Sponsorenleiste
    bottom_margin = footer_space + sponsor_height + margin_above_sponsor
    
    doc = SimpleDocTemplate(
        filename, pagesize=A4,
        leftMargin=8*mm, rightMargin=8*mm,
        topMargin=8*mm, bottomMargin=bottom_margin  # Dynamisch berechnet
    )

    styles = getSampleStyleSheet()
    style_show = ParagraphStyle("show", parent=styles["Heading1"], fontSize=14, leading=16, spaceAfter=2)
    style_comp = ParagraphStyle("comp", parent=styles["Normal"], fontSize=11, leading=13, spaceAfter=4)
    style_sub = ParagraphStyle("sub", parent=styles["Normal"], fontSize=10, leading=12, spaceAfter=4)
    style_info = ParagraphStyle("info", parent=styles["Normal"], fontSize=10, leading=12, spaceAfter=4)
    style_hdr = ParagraphStyle("hdr", parent=styles["Normal"], fontSize=9, alignment=1)
    style_hdr_left = ParagraphStyle("hdr_left", parent=styles["Normal"], fontSize=9, alignment=0)
    style_pos = ParagraphStyle("pos", parent=styles["Normal"], fontSize=9, alignment=1)
    style_rider = ParagraphStyle("rider", parent=styles["Normal"], fontSize=8, leading=9)
    style_horse = ParagraphStyle("horse", parent=styles["Normal"], fontSize=8, leading=9)
    style_pause = ParagraphStyle("pause", parent=styles["Normal"], fontSize=9, alignment=1)

    elements = []

    # LOGO DIREKT AM ANFANG - Prüfungsspezifisches Logo-System mit DPI-Korrektur
    logo_path = starterlist.get("logoPath")
    if logo_path and os.path.exists(logo_path):
        try:
            # Logo in Originalgröße laden (ohne width/height)
            # Logo in Originalgröße laden
            # Logo in Originalgröße laden
            logo = Image(logo_path)
            
            # Maximale Größe: von Parameter (Standard 5cm = 50mm)
            max_size = logo_max_width_cm * 10 * mm
            
            # Wenn das Logo größer als max_size ist, verkleinern (aber nicht vergrößern!)
            if logo.drawWidth > max_size or logo.drawHeight > max_size:
                # Proportional skalieren
                scale = min(max_size / logo.drawWidth, max_size / logo.drawHeight)
                logo.drawWidth = logo.drawWidth * scale
                logo.drawHeight = logo.drawHeight * scale
            
            logo.hAlign = 'RIGHT'
            elements.append(logo)
            
            # Negativer Spacer proportional zur Logo-Höhe
            # Mindestens 5mm Abstand oben lassen, dann Logo-Höhe abziehen
            negative_spacer = max(-logo.drawHeight - 5*mm, -25*mm)
            elements.append(Spacer(1, negative_spacer))
            
            print(f"PDF STANDARD LOGO DEBUG: Logo positioned at top right: {logo_path} (Größe: {logo.drawWidth:.1f}x{logo.drawHeight:.1f}pt, Spacer: {negative_spacer/mm:.1f}mm)")
        except Exception as e:
            print(f"PDF STANDARD LOGO DEBUG: Logo error: {e}")

    # --- KOPF ---
    show = starterlist.get("show") or {}
    comp = starterlist.get("competition") or {}

    show_title = show.get("title") or starterlist.get("showTitle") or "Unbenannte Veranstaltung"
    elements.append(Paragraph(f"<b>{show_title}</b>", style_show))

    comp_no = comp.get("number") or starterlist.get("competitionNumber") or ""
    comp_title = comp.get("title") or starterlist.get("competitionTitle") or ""
    division = comp.get("divisionNumber") or starterlist.get("divisionNumber")
    div_text = ""

    # --- Zusatzfelder ---
    subtitle = comp.get("subtitle") or starterlist.get("subtitle")
    informationText = comp.get("informationText") or starterlist.get("informationText")
    location = comp.get("location") or starterlist.get("location")

    try:
        if division is not None and str(division) != "" and int(division) > 0:
            div_text = f"{int(division)}. Abt. "
    except:
        div_text = f"{division} " if division else ""

    # 2. Zeile: Prüfungszeile FETT
    if comp_no or comp_title:
        comp_line = f"Prüfung {comp_no}"
        if div_text:
            comp_line += f" - {div_text}{comp_title}"
        elif comp_title:
            comp_line += f" - {comp_title}"
        elements.append(Paragraph(f"<b>{comp_line}</b>", style_comp))

    # 3. Zeile: informationText (nicht mehr fett)
    if informationText:
        processed_info_text = _process_information_text(informationText)
        elements.append(Paragraph(processed_info_text, style_info))

    if subtitle:
        elements.append(Paragraph(subtitle, style_sub))

    # ERWEITERTE DIVISION-START-ZEIT BEHANDLUNG VOM STANDARD TEMPLATE
    start_raw = None
    division_num = starterlist.get('divisionNumber')
    divisions = starterlist.get('divisions', [])
    
    if divisions and division_num is not None:
        try:
            div_num = int(division_num)
            for div in divisions:
                if div.get("number") == div_num:
                    division_start = div.get("start")
                    if division_start:
                        start_raw = division_start
                        break
        except (ValueError, TypeError):
            pass
    elif divisions and len(divisions) > 0:
        first_division = divisions[0]
        division_start = first_division.get("start")
        if division_start:
            start_raw = division_start
    
    if not start_raw:
        start_raw = comp.get("start") or starterlist.get("start") or show.get("start")
    
    if start_raw:
        date_line = _fmt_header_datetime(start_raw)
        if location:
            date_line = f"{date_line}  -  {location}"
        elements.append(Paragraph(f"<b>{date_line}</b>", style_sub))

    elements.append(Spacer(1, 6))

    # --- TABELLE MIT 3 RICHTER-SPALTEN (7 Spalten total) MIT GRUPPIERUNGSLOGIK ---
    starters = starterlist.get("starters") or []
    breaks = starterlist.get("breaks") or []
    breaks_by_after = {}
    for b in breaks:
        try:
            # WICHTIG: afterNumberInCompetition kann 0 sein!
            after_num = b.get("afterNumberInCompetition")
            if after_num is None:
                k = -1
            else:
                k = int(after_num)
            breaks_by_after.setdefault(k, []).append(b)
        except:
            continue

    judges = comp.get("judges") or starterlist.get("judges") or []
    judge_positions = _get_ordered_judge_positions_main_table(judges)

    data_texts = []
    meta = []
    data_texts.append(["<b>Start</b>", "<b>KoNr.</b>", "<b>Reiter / Pferd</b><br/><font size=7>Geschlecht / Farbe / Geboren / Vater x Muttervater / Besitzer / Züchter</font>", f"<b>{judge_positions[0]}</b>", f"<b>{judge_positions[1]}</b>", f"<b>{judge_positions[2]}</b>", "<b>Total</b>"])
    meta.append({"type":"header"})
    
    # WICHTIG: Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_by_after:
        for br in breaks_by_after[0]:
            pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append([pause_text, "", "", "", "", "", ""])
            meta.append({"type":"pause"})

    # GRUPPIERUNGSLOGIK VOM STANDARD TEMPLATE ÜBERNOMMEN - ABER VEREINFACHT
    current_group = None
    group_start_time_shown = False

    for s in starters:
        # Prüfe auf Gruppenwechsel BEVOR der Starter hinzugefügt wird
        starter_group = s.get("groupNumber")
        
        # Nur wenn groupNumber existiert und > 0
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            # Neue Gruppe erkannt - Gruppen-Header hinzufügen
            group_text = f"Abteilung {starter_group}"
            data_texts.append([group_text, "", "", "", "", "", ""])
            meta.append({"type":"group"})
            
            current_group = starter_group
            group_start_time_shown = False  # Reset für neue Gruppe

        nr = s.get("startNumber") or ""
        hors_concours = bool(s.get("horsConcours", False))  # Außer Konkurrenz
        tstr = _fmt_time(s.get("startTime"))
        
        # Zeit nur beim ersten Starter pro Gruppe zeigen, oder wenn keine Gruppierung
        if starter_group is None or starter_group == 0 or not group_start_time_shown:
            start_time_display = tstr
            if starter_group is not None and starter_group > 0:
                group_start_time_shown = True  # Markiere Zeit als gezeigt für diese Gruppe
        else:
            start_time_display = ""  # Verstecke Zeit für weitere Starter in derselben Gruppe
        
        pos_html = f"<b>{nr}</b><br/><font size=9>{start_time_display}</font>" if start_time_display else f"<b>{nr}</b>"

        horses = s.get("horses") or []
        cno = ""
        if horses:
            horse = horses[0]
            cno = str(horse.get("cno", ""))

        # VERBESSERTE ABSTAMMUNGSDARSTELLUNG
        athlete = s.get("athlete") or {}
        rider_name = _safe_get(athlete, "name", "")
        
        content_parts = []
        if rider_name:
            content_parts.append(f"<b>{rider_name.upper()}</b>")
        
        if horses:
            horse = horses[0]
            horse_name = _safe_get(horse, "name", "")
            studbook = _safe_get(horse, "studbook", "")
            
            horse_line = ""
            if horse_name:
                horse_line = f"<b>{horse_name.upper()}</b>"
            if studbook:
                if horse_line:
                    horse_line += f" - {studbook}"
                else:
                    horse_line = studbook
            if horse_line:
                content_parts.append(horse_line)
            
            details_parts = []
            sex = _safe_get(horse, "sex", "")
            sex_german = SEX_MAP.get(str(sex).upper(), sex)
            if sex_german:
                details_parts.append(sex_german)
            
            color = _safe_get(horse, "color", "")
            if color:
                details_parts.append(color)
                
            year = _safe_get(horse, "breedingSeason", "")
            if year:
                details_parts.append(str(year))
                
            sire = _safe_get(horse, "sire", "")
            damsire = _safe_get(horse, "damSire", "")
            if sire:
                pedigree = sire
                if damsire:
                    pedigree += f" x {damsire}"
                details_parts.append(pedigree)
            
            if details_parts:
                details_text = " / ".join(details_parts)
                content_parts.append(f"<font size=7>{details_text}</font>")
            
            # KORREKTE BESITZER/ZÜCHTER REIHENFOLGE
            owner = _safe_get(horse, "owner", "")
            breeder = _safe_get(horse, "breeder", "")
            
            owner_breeder_parts = []
            if owner:
                owner_breeder_parts.append(f"B: {owner}")
            if breeder:
                owner_breeder_parts.append(f"Z: {breeder}")
            
            if owner_breeder_parts:
                owner_breeder_text = " / ".join(owner_breeder_parts)
                content_parts.append(f"<font size=7>{owner_breeder_text}</font>")

        combined_content = "<br/>".join(content_parts)

        # AK in Total-Spalte (letzte Spalte)
        total_text = "AK" if hors_concours else ""
        data_texts.append([pos_html, cno, combined_content, "", "", "", total_text])
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type":"starter","withdrawn":withdrawn_flag, "horsConcours": hors_concours})

        # Breaks verarbeiten
        try:
            cur = int(nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                data_texts.append([pause_text, "", "", "", "", "", ""])
                meta.append({"type":"pause"})

    # --- Tabelle erstellen (7 Spalten) ---
    page_width = A4[0] - doc.leftMargin - doc.rightMargin
    col1 = 18*mm
    col2 = 16*mm      # KoNr. breiter gemacht (von 12mm auf 16mm)
    col_judge = 12*mm
    col_total = 15*mm
    col3 = page_width - col1 - col2 - (3*col_judge) - col_total
    col_widths = [col1, col2, col3, col_judge, col_judge, col_judge, col_total]

    table_rows = []
    for i, row in enumerate(data_texts):
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr),
                Paragraph(row[2], style_hdr_left), Paragraph(row[3], style_hdr),
                Paragraph(row[4], style_hdr), Paragraph(row[5], style_hdr),
                Paragraph(row[6], style_hdr)
            ])
        elif m["type"] == "group":
            # GRUPPIERUNGSLOGIK VOM STANDARD TEMPLATE - LINKSBÜNDIG
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_hdr_left), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub)
            ])
        elif m["type"] == "pause":
            table_rows.append([
                Paragraph(row[0], style_pause), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub)
            ])
        else:
            withdrawn = m.get("withdrawn", False)
            def maybe_strike(text, s):
                if not text: return Paragraph("", s)
                return Paragraph(f"<strike>{text}</strike>" if withdrawn else text, s)
            table_rows.append([
                maybe_strike(row[0], style_pos),
                maybe_strike(row[1], style_pos),
                maybe_strike(row[2], style_horse),
                maybe_strike(row[3], style_pos),
                maybe_strike(row[4], style_pos),
                maybe_strike(row[5], style_pos),
                maybe_strike(row[6], style_pos),
            ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("ALIGN", (1,0), (1,-1), "CENTER"),
        ("ALIGN", (2,0), (2,-1), "LEFT"),
        ("ALIGN", (3,0), (-1,-1), "CENTER"),
    ])

    # Styling für spezielle Zeilen mit Zebra-Streifen
    starter_count = 0
    
    for ri in range(1, len(table_rows)):
        if ri < len(meta):
            m = meta[ri]
            if m and m.get("type") == "group":
                ts.add("SPAN", (0,ri), (-1,ri))
                ts.add("BACKGROUND", (0,ri), (-1,ri), colors.Color(0.9, 0.9, 0.9))
                ts.add("ALIGN", (0,ri), (-1,ri), "LEFT")
                starter_count = 0
            elif m and m.get("type") == "pause":
                ts.add("SPAN", (0,ri), (-1,ri))
                ts.add("ALIGN", (0,ri), (-1,ri), "CENTER")
                prev_was_gray = (starter_count - 1) % 2 == 1
                if not prev_was_gray:
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor("#f5f5f5"))
                starter_count -= 1
            elif m and m.get("type") == "starter":
                is_gray_zebra = starter_count % 2 == 1
                if is_gray_zebra:
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor("#f5f5f5"))
                
                if m.get("withdrawn", False):
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor("#f2f2f2"))
                    ts.add("TEXTCOLOR", (0,ri), (-1,ri), colors.darkgrey)
                
                starter_count += 1

    t.setStyle(ts)
    elements.append(t)

    # --- Richter ---
    judges = comp.get("judges") or starterlist.get("judges") or []
    judging_rule = comp.get("judgingRule") or starterlist.get("judgingRule")
    if judges:
        title = "Richter" + (f" ({judging_rule})" if judging_rule else "")
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>{title}</b>", style_sub))
        
        jrows = [[Paragraph("<b>Pos.</b>", style_hdr), Paragraph("<b>Name</b>", style_sub)]]
        ordered_judges = _get_ordered_judges_all(judges)
        
        for judge in ordered_judges:
            pos_label = judge["pos_label"]
            jrows.append([Paragraph(pos_label, style_pos), Paragraph(judge.get("name",""), style_rider)])
        
        jt = Table(jrows, colWidths=[25*mm, page_width-25*mm])
        jt.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ALIGN",(0,0),(0,-1),"CENTER"),
        ]))
        elements.append(jt)

    doc.build(elements, canvasmaker=FooterCanvas)