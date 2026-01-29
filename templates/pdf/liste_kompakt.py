# templates/pdf/liste_standard_logo.py
# Listen-Template: Nur Tabelle + Richter, ohne Kopfbereich
# Basierend auf pdf_standard_logo.py aber ohne Kopf
import os
import json
from datetime import datetime

def _fmt_header_datetime(iso):
    """Format datetime for header: Samstag, 6. September 2025 um 08:00 Uhr"""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        weekday_map = {
            "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
            "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
        }
        month_map = {
            "January": "Januar", "February": "Februar", "March": "März", "April": "April",
            "May": "Mai", "June": "Juni", "July": "Juli", "August": "August",
            "September": "September", "October": "Oktober", "November": "November", "December": "Dezember"
        }
        weekday = weekday_map.get(dt.strftime("%A"), dt.strftime("%A"))
        month = month_map.get(dt.strftime("%B"), dt.strftime("%B"))
        return f"{weekday}, {dt.day}. {month} {dt.year} um {dt.strftime('%H:%M')} Uhr"
    except Exception:
        return str(iso)

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageTemplate, Frame, NextPageTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

OUTPUT_DIR = "Ausgabe"

def _ensure_output_dir():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def _safe_get(obj, key, default=""):
    """Safely get value from dict or return default"""
    if obj is None:
        return default
    return obj.get(key, default)

def _format_time(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%H:%M:%S")
    except Exception:
        try:
            return str(iso).split("T")[-1][:8]
        except:
            return str(iso)

def _maybe_strike(text, style, withdrawn=False):
    if not text:
        return Paragraph("", style)
    if withdrawn:
        return Paragraph(f"<strike>{text}</strike>", style)
    else:
        return Paragraph(text, style)

def _get_ordered_judges_all(judges):
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    dressage_positions = ["E", "H", "C", "M", "B"]
    judges_by_position = {}
    other_judges = []
    
    for judge in judges:
        position = judge.get("position", "")
        if isinstance(position, int):
            pos_label = pos_map.get(position, str(position))
        else:
            pos_label = pos_map.get(str(position), str(position))
        
        judge_copy = judge.copy()
        judge_copy["pos_label"] = pos_label
        
        if pos_label in dressage_positions:
            if pos_label not in judges_by_position:
                judges_by_position[pos_label] = []
            judges_by_position[pos_label].append(judge_copy)
        else:
            other_judges.append(judge_copy)
    
    result = []
    for pos in dressage_positions:
        if pos in judges_by_position:
            result.extend(judges_by_position[pos])
    
    other_judges.sort(key=lambda j: j["pos_label"])
    result.extend(other_judges)
    
    return result

def render(starterlist: dict, filename: str):
    """Creates a PDF with only starter list table and judges, using flexible spacing"""
    print("PDF LISTE STANDARD DEBUG: Starting render function")
    
    if starterlist is None:
        raise ValueError("starterlist is None")

    if isinstance(starterlist, str):
        try:
            starterlist = json.loads(starterlist)
        except Exception as e:
            raise ValueError(f"starterlist is a string but not valid JSON: {e}")

    _ensure_output_dir()
    
    # Extrahiere Prüfungsnummer und Abteilung für XXY Format
    comp = starterlist.get("competition") or {}
    comp_number_raw = comp.get("number") or starterlist.get("competitionNumber") or "99"
    
    # Extrahiere nur die Zahl aus comp_number (z.B. "14start" → "14", "5a" → "5")
    import re
    match = re.search(r'(\d+)', str(comp_number_raw))
    comp_number = match.group(1) if match else "99"
    
    # Hole Abteilungsnummer (falls vorhanden)
    # Prüfe erst in starterlist, dann in competition
    division = starterlist.get("division") or starterlist.get("divisionNumber") or comp.get("division") or 0
    
    # Formatiere als XXY: XX = Prüfung (2-stellig), Y = Abteilung (1-stellig)
    # Beispiele: 5 → 050, 10 → 100, 5/Abt.2 → 052, 17 → 170, "14start" → 140
    try:
        comp_num = int(comp_number)
        div_num = int(division) if division else 0
        output_filename = f"{comp_num:02d}{div_num:01d}.pdf"
    except (ValueError, TypeError):
        # Fallback wenn Konvertierung fehlschlägt
        output_filename = "990.pdf"
    
    # Verwende übergebenen filename statt filename
    
    
    # Hole Abstände aus starterlist (in cm)
    spacing_top_cm = starterlist.get("spacingTopCm", 2.0)  # Nur für Seite 1 - gleich wie unten
    spacing_bottom_cm = starterlist.get("spacingBottomCm", 2.0)  # Für alle Seiten
    
    print(f"PDF LISTE DEBUG: Abstände - Seite 1 Oben: {spacing_top_cm}cm, Alle Seiten Unten: {spacing_bottom_cm}cm, Ab Seite 2 Oben: 1cm")
    
    # Konvertiere cm zu mm für ReportLab
    top_margin_first = spacing_top_cm * 10  # cm -> mm (nur Seite 1)
    top_margin_later = 1.0 * 10  # 1cm für alle weiteren Seiten
    bottom_margin = spacing_bottom_cm * 10  # cm -> mm (alle Seiten)
    
    # Erstelle Custom DocTemplate mit unterschiedlichen Seitenvorlagen
    class ListeDocTemplate(SimpleDocTemplate):
        def __init__(self, filename, **kw):
            self.allowSplitting = 1
            SimpleDocTemplate.__init__(self, filename, **kw)
            
        def build(self, flowables):
            # Erste Seite mit großem oberen Rand
            frame_first = Frame(
                8*mm, bottom_margin*mm,
                A4[0] - 18*mm, A4[1] - top_margin_first*mm - bottom_margin*mm,
                id='first',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
            )
            
            # Folgeseiten mit kleinem oberen Rand (1cm)
            frame_later = Frame(
                8*mm, bottom_margin*mm,
                A4[0] - 18*mm, A4[1] - top_margin_later*mm - bottom_margin*mm,
                id='later',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
            )
            
            # Definiere die Seitenvorlagen
            self.addPageTemplates([
                PageTemplate(id='First', frames=[frame_first]),
                PageTemplate(id='Later', frames=[frame_later])
            ])
            
            SimpleDocTemplate.build(self, flowables)
    
    doc = ListeDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=0, leftMargin=0,
        topMargin=0, bottomMargin=0
    )
    
    styles = getSampleStyleSheet()
    
    style_hdr = ParagraphStyle('Header', parent=styles['Normal'], fontSize=9, 
                              alignment=TA_CENTER, textColor=colors.black, leading=10)
    style_normal = ParagraphStyle('TableCell', parent=styles['Normal'], fontSize=9, 
                                 alignment=TA_CENTER, leading=10)
    style_start = ParagraphStyle('StartNumber', parent=styles['Normal'], fontSize=9, 
                                alignment=TA_CENTER, leading=10)
    style_time = ParagraphStyle('Time', parent=styles['Normal'], fontSize=9,
                               alignment=TA_CENTER, leading=10)
    style_rider = ParagraphStyle('Rider', parent=styles['Normal'], fontSize=9, 
                                alignment=TA_LEFT, leading=10)
    style_sub = ParagraphStyle('CustomSubtitle', parent=styles['Heading2'], fontSize=14, leading=16,
                              spaceAfter=6, alignment=TA_LEFT)
    
    elements = []
    
    # Nach der ersten Seite wechseln wir zum "Later" Template
    elements.append(NextPageTemplate('Later'))

    # Competition-Daten
    comp = starterlist.get("competition") or {}
    show = starterlist.get("show") or {}
    
    # --- KOPFZEILE: STARTERLISTE (links) und Datum/Ort (rechts) ---
    starters = starterlist.get("starters") or []
    
    # Zeit-Logik wie im PDF-Template
    divisions = comp.get("divisions") or starterlist.get("divisions") or []
    division_num = comp.get("divisionNumber") or starterlist.get("divisionNumber")
    
    start_raw = None
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
    
    location = comp.get("location") or starterlist.get("location") or ""
    
    header_left = "<b>STARTERLISTE</b>"
    header_right = ""
    if start_raw:
        formatted_time = _fmt_header_datetime(start_raw)
        if location:
            header_right = f"{formatted_time} - {location}"
        else:
            header_right = formatted_time
    elif location:
        header_right = location
    
    if header_right:
        style_header_left = ParagraphStyle("header_left", parent=styles["Normal"], fontSize=10, alignment=0)
        style_header_right = ParagraphStyle("header_right", parent=styles["Normal"], fontSize=10, alignment=2)
        
        header_table = Table(
            [[Paragraph(header_left, style_header_left), Paragraph(header_right, style_header_right)]],
            colWidths=[A4[0]/2 - 9*mm, A4[0]/2 - 9*mm]
        )
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 3*mm))

    # Table setup - 6 columns wie im Original
    page_width = A4[0] - 18*mm
    col_widths = [12*mm, 22*mm, 16*mm, 62*mm, 62*mm, 22*mm]

    headers = ["Start", "Zeit", "KoNr.", "Reiter", "Pferd", "Ergeb."]
    header_row = [Paragraph(f"<b>{h}</b>", style_hdr) for h in headers]

    data = [header_row]

    # Process starters
    starters = starterlist.get("starters", [])
    breaks = starterlist.get("breaks", [])
    
    print(f"PDF DEBUG: Found {len(starters)} starters in starterlist")
    
    breaks_map = {}
    for br in breaks:
        try:
            # WICHTIG: afterNumberInCompetition kann 0 sein!
            after_num = br.get("afterNumberInCompetition")
            if after_num is None:
                k = -1
            else:
                k = int(after_num)
            breaks_map[k] = br
        except:
            continue
    
    # WICHTIG: Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_map:
        break_info = breaks_map[0]
        total_seconds = int(break_info.get("totalSeconds", 0) or 0)
        info_text = break_info.get("informationText", "")
        
        if total_seconds > 0:
            if total_seconds >= 3600:
                hours = total_seconds // 3600
                mins = (total_seconds % 3600) // 60
                time_text = f"{hours} Std. {mins:02d} Min. Pause"
            else:
                mins = total_seconds // 60
                time_text = f"{mins} Minuten Pause"
            break_text = f"{time_text} – {info_text}" if info_text else time_text
        else:
            break_text = info_text or "Pause"
        
        # Pause-Zeile hinzufügen (6 Spalten, über alle Spalten)
        pause_row = [Paragraph(f"<i>{break_text}</i>", style_normal)] + [""] * 5
        data.append(pause_row)

    current_group = None

    for starter_index, starter in enumerate(starters):
        starter_group = starter.get("groupNumber")
        
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            group_data = ["", "", "", "", "", ""]
            group_data[0] = f"Abteilung {starter_group}"
            
            group_row_data = []
            for i, cell_content in enumerate(group_data):
                if cell_content:
                    group_style = ParagraphStyle('GroupHeader', parent=styles['Normal'], fontSize=10, 
                                               alignment=TA_LEFT, textColor=colors.black)
                    group_row_data.append(Paragraph(f"<b>{cell_content}</b>", group_style))
                else:
                    group_row_data.append(Paragraph("", style_normal))
            
            data.append(group_row_data)
            current_group = starter_group

        # Starter data - KORREKTE FELDNAMEN wie im Original!
        withdrawn = starter.get("withdrawn", False)
        start_number = str(starter.get("startNumber", ""))
        hors_concours = bool(starter.get("horsConcours", False))  # Außer Konkurrenz
        time_value = _format_time(starter.get("startTime", ""))  # startTime nicht time!
        
        horses = starter.get("horses", [])
        ko_nr = horses[0].get("cno", "") if horses else ""  # horses[0].cno nicht competitorNumber!
        
        athlete = starter.get("athlete", {})  # athlete nicht rider!
        rider_name = athlete.get("name", "") if athlete else ""
        
        horse_name = horses[0].get("name", "") if horses else ""  # horses[0].name
        
        # Startnummer OHNE AK (AK kommt in letzte Spalte)
        start_number_display = start_number
        
        # AK in letzte Spalte mittig
        result_text = "AK" if hors_concours else ""
        
        data.append([
            _maybe_strike(start_number_display, style_start, withdrawn),
            _maybe_strike(time_value, style_time, withdrawn),
            _maybe_strike(str(ko_nr), style_normal, withdrawn),
            _maybe_strike(rider_name, style_rider, withdrawn),
            _maybe_strike(horse_name, style_rider, withdrawn),
            Paragraph(result_text, style_normal)
        ])
        
        # Break row - KORREKTE LOGIK aus Original
        try:
            starter_num = int(start_number)
            if starter_num in breaks_map:
                break_info = breaks_map[starter_num]
                
                total_seconds = int(break_info.get("totalSeconds", 0) or 0)
                info_text = break_info.get("informationText", "")
                
                if total_seconds > 0:
                    if total_seconds >= 3600:
                        hours = total_seconds // 3600
                        mins = (total_seconds % 3600) // 60
                        time_text = f"{hours} Std. {mins:02d} Min. Pause"
                    else:
                        mins = total_seconds // 60
                        time_text = f"{mins} Minuten Pause"
                    
                    break_text = f"{time_text} – {info_text}" if info_text else time_text
                else:
                    break_text = info_text or "Pause"
                
                # Add break row
                break_row_data = [Paragraph(f"<i>{break_text}</i>", style_normal)] + [""] * 5
                data.append(break_row_data)
        except:
            pass

    # Create table
    table = Table(data, colWidths=col_widths)
    
    # Style table
    table_style = [
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (2, -1), "CENTER"),
        ("ALIGN", (3, 1), (4, -1), "LEFT"),
        ("ALIGN", (5, 1), (5, -1), "CENTER"),
        # Kompakte Paddings
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]
    
    # Apply styling for groups and breaks + ZEBRA
    row_index = 1
    starter_count = 0
    
    # Prüfe ob Pause vor erstem Starter existiert (row_index = 1)
    if 0 in breaks_map:
        table_style.append(('SPAN', (0, row_index), (-1, row_index)))
        table_style.append(('ALIGN', (0, row_index), (-1, row_index), 'CENTER'))
        # Pause bekommt entgegengesetzte Farbe vom vorherigen Starter
        prev_was_gray = (starter_count - 1) % 2 == 1
        if not prev_was_gray:  # Vorheriger war weiß → Pause grau
            table_style.append(('BACKGROUND', (0, row_index), (-1, row_index), colors.HexColor("#f5f5f5")))
        # WICHTIG: Counter um 1 zurücksetzen, damit nächste Zeile gleiche Farbe hat!
        starter_count -= 1
        row_index += 1
    
    current_group = None
    
    for starter_index, starter in enumerate(starters):
        starter_group = starter.get("groupNumber")
        
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            light_gray = colors.Color(0.9, 0.9, 0.9)
            table_style.append(('BACKGROUND', (0, row_index), (-1, row_index), light_gray))
            table_style.append(('SPAN', (0, row_index), (-1, row_index)))
            table_style.append(('ALIGN', (0, row_index), (0, row_index), 'LEFT'))
            row_index += 1
            starter_count = 0
            current_group = starter_group
        
        # Zebra für Starter
        is_gray_zebra = starter_count % 2 == 1
        if is_gray_zebra:
            table_style.append(('BACKGROUND', (0, row_index), (-1, row_index), colors.HexColor("#f5f5f5")))
        
        # Withdrawn nur TEXTCOLOR
        if starter.get("withdrawn", False):
            table_style.append(('TEXTCOLOR', (0, row_index), (-1, row_index), colors.darkgrey))
        
        starter_count += 1
        row_index += 1
        
        try:
            start_num = str(starter.get("startNumber", ""))
            starter_num = int(start_num)
            if starter_num in breaks_map:
                table_style.append(('SPAN', (0, row_index), (-1, row_index)))
                table_style.append(('ALIGN', (0, row_index), (-1, row_index), 'CENTER'))
                # Pause zebra
                prev_was_gray = (starter_count - 1) % 2 == 1
                if not prev_was_gray:
                    table_style.append(('BACKGROUND', (0, row_index), (-1, row_index), colors.HexColor("#f5f5f5")))
                starter_count -= 1
                row_index += 1
        except:
            pass

    table.setStyle(TableStyle(table_style))
    table.hAlign = 'LEFT'
    elements.append(table)

    # Add judges section
    judges = starterlist.get("judges", [])
    judging_rule = starterlist.get("judgingRule")
    if judges:
        title = "Richter" + (f" ({judging_rule})" if judging_rule else "")
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>{title}</b>", style_sub))
        
        jrows = [[Paragraph("<b>Pos.</b>", style_hdr), Paragraph("<b>Name</b>", style_sub)]]
        
        ordered_judges = _get_ordered_judges_all(judges)
        
        for judge in ordered_judges:
            pos_label = judge["pos_label"]
            jrows.append([Paragraph(pos_label, style_normal), Paragraph(judge.get("name",""), style_rider)])
        
        jt = Table(jrows, colWidths=[25*mm, page_width-25*mm])
        jt.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ALIGN",(0,0),(0,-1),"CENTER"),
        ]))
        jt.hAlign = 'LEFT'
        elements.append(jt)

    doc.build(elements)
    print(f"PDF LISTE DEBUG: Created {filename}")
    return filename
