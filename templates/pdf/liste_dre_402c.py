# -*- coding: utf-8 -*-
# templates/pdf/liste_dre_402c_logo.py - Speziell für Richtverfahren 402.C
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageTemplate, Frame, NextPageTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from datetime import datetime
import os

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

SEX_MAP = {
    "MARE": "Stute",
    "GELDING": "Wallach", 
    "STALLION": "Hengst",
    "M": "Wallach",
    "F": "Stute",
    "": ""
}

def _safe_get(d, key, default=""):
    if not d:
        return default
    return d.get(key, default) if isinstance(d, dict) else default

def _fmt_time(iso):
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
    if not text:
        return ""
    text = text.lstrip('\n')
    text = text.replace('\n', '<br/>')
    text = f"<b>{text}</b>"
    return text

def _get_judge_data_for_display_402c(judges, starterlist):
    """Prozessiert Richter für 402.C mit Aufgaben aus dressageTests"""
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    standard_positions = ["E", "H", "C", "M", "B", "K", "V", "S", "R", "P", "F", "A"]
    
    dressage_tests = starterlist.get("dressageTests", [])
    
    position_to_task = {}
    for test in dressage_tests:
        task_name = test.get("name", "")
        judge_positions = test.get("judgePositions", [])
        
        for pos in judge_positions:
            if task_name:
                position_to_task[pos] = task_name
    
    result = []
    for judge in judges:
        original_position = judge.get("position", "")
        
        if isinstance(original_position, int):
            pos_label = pos_map.get(original_position, str(original_position))
        else:
            pos_label = pos_map.get(original_position, original_position)
        
        judge_copy = judge.copy()
        judge_copy["pos_label"] = pos_label
        
        if pos_label in standard_positions:
            judge_copy["task"] = position_to_task.get(original_position, "")
        else:
            judge_copy["task"] = ""
        
        result.append(judge_copy)
    
    dressage_positions = ["C", "E", "H", "M", "B"]
    
    def sort_key(j):
        pos = j["pos_label"]
        if pos in dressage_positions:
            return (0, dressage_positions.index(pos))
        elif pos in standard_positions:
            return (1, pos)
        return (2, pos)
    
    result.sort(key=sort_key)
    return result


def get_country_name(ioc_code):
    """Returns full country name in German - Complete list for all 250 countries"""
    if not ioc_code:
        return ""
    
    names = {
        # A
        "AFG": "Afghanistan", "AIA": "Anguilla", "ALB": "Albanien", "ALG": "Algerien",
        "AND": "Andorra", "ANG": "Angola", "ANT": "Antigua und Barbuda", "ARG": "Argentinien",
        "ARM": "Armenien", "ARU": "Aruba", "ASA": "Amerikanisch-Samoa", "AUS": "Australien",
        "AUT": "Österreich", "AZE": "Aserbaidschan",
        
        # B
        "BAH": "Bahamas", "BAN": "Bangladesch", "BAR": "Barbados", "BDI": "Burundi",
        "BEL": "Belgien", "BEN": "Benin", "BER": "Bermuda", "BHU": "Bhutan",
        "BIH": "Bosnien und Herzegowina", "BIZ": "Belize", "BLR": "Belarus", "BOL": "Bolivien",
        "BOT": "Botswana", "BRA": "Brasilien", "BRN": "Bahrain", "BRU": "Brunei",
        "BUL": "Bulgarien", "BUR": "Burkina Faso",
        
        # C
        "CAF": "Zentralafrikanische Republik", "CAM": "Kambodscha", "CAN": "Kanada",
        "CAY": "Kaimaninseln", "CGO": "Kongo", "CHA": "Tschad", "CHI": "Chile",
        "CHN": "China", "CIV": "Elfenbeinküste", "CMR": "Kamerun", "COD": "DR Kongo",
        "COK": "Cookinseln", "COL": "Kolumbien", "COM": "Komoren", "CPV": "Kap Verde",
        "CRC": "Costa Rica", "CRO": "Kroatien", "CUB": "Kuba", "CYP": "Zypern",
        "CZE": "Tschechien",
        
        # D-E
        "DEN": "Dänemark", "DJI": "Dschibuti", "DMA": "Dominica", "DOM": "Dominikanische Republik",
        "ECU": "Ecuador", "EGY": "Ägypten", "ERI": "Eritrea", "ESA": "El Salvador",
        "ESP": "Spanien", "EST": "Estland", "ETH": "Äthiopien",
        
        # F
        "FAR": "Färöer", "FIJ": "Fidschi", "FIN": "Finnland", "FRA": "Frankreich",
        "FSM": "Mikronesien",
        
        # G
        "GAB": "Gabun", "GAM": "Gambia", "GBR": "Großbritannien", "GBS": "Guinea-Bissau",
        "GEO": "Georgien", "GEQ": "Äquatorialguinea", "GER": "Deutschland", "GHA": "Ghana",
        "GRE": "Griechenland", "GRN": "Grenada", "GUA": "Guatemala", "GUI": "Guinea",
        "GUM": "Guam", "GUY": "Guyana",
        
        # H-I
        "HAI": "Haiti", "HKG": "Hongkong", "HON": "Honduras", "HUN": "Ungarn",
        "INA": "Indonesien", "IND": "Indien", "IRI": "Iran", "IRL": "Irland",
        "IRQ": "Irak", "ISL": "Island", "ISR": "Israel", "ISV": "Amerikanische Jungferninseln",
        "ITA": "Italien", "IVB": "Britische Jungferninseln",
        
        # J-K
        "JAM": "Jamaika", "JOR": "Jordanien", "JPN": "Japan", "KAZ": "Kasachstan",
        "KEN": "Kenia", "KGZ": "Kirgisistan", "KIR": "Kiribati", "KOR": "Südkorea",
        "KOS": "Kosovo", "KSA": "Saudi-Arabien", "KUW": "Kuwait",
        
        # L
        "LAO": "Laos", "LAT": "Lettland", "LBA": "Libyen", "LBN": "Libanon",
        "LBR": "Liberia", "LCA": "St. Lucia", "LES": "Lesotho", "LIE": "Liechtenstein",
        "LTU": "Litauen", "LUX": "Luxemburg",
        
        # M
        "MAC": "Macau", "MAD": "Madagaskar", "MAR": "Marokko", "MAS": "Malaysia",
        "MAW": "Malawi", "MDA": "Moldau", "MDV": "Malediven", "MEX": "Mexiko",
        "MGL": "Mongolei", "MHL": "Marshallinseln", "MKD": "Nordmazedonien", "MLI": "Mali",
        "MLT": "Malta", "MNE": "Montenegro", "MON": "Monaco", "MOZ": "Mosambik",
        "MRI": "Mauritius", "MTN": "Mauretanien", "MYA": "Myanmar",
        
        # N
        "NAM": "Namibia", "NCA": "Nicaragua", "NED": "Niederlande", "NEP": "Nepal",
        "NGR": "Nigeria", "NIG": "Niger", "NOR": "Norwegen", "NRU": "Nauru",
        "NZL": "Neuseeland",
        
        # O-P
        "OMA": "Oman", "PAK": "Pakistan", "PAN": "Panama", "PAR": "Paraguay",
        "PER": "Peru", "PHI": "Philippinen", "PLE": "Palästina", "PLW": "Palau",
        "PNG": "Papua-Neuguinea", "POL": "Polen", "POR": "Portugal", "PRK": "Nordkorea",
        "PUR": "Puerto Rico",
        
        # Q-R
        "QAT": "Katar", "ROU": "Rumänien", "RSA": "Südafrika", "RUS": "Russland",
        "RWA": "Ruanda",
        
        # S
        "SAM": "Samoa", "SEN": "Senegal", "SEY": "Seychellen", "SGP": "Singapur",
        "SKN": "St. Kitts und Nevis", "SLE": "Sierra Leone", "SLO": "Slowenien",
        "SMR": "San Marino", "SOL": "Salomonen", "SOM": "Somalia", "SRB": "Serbien",
        "SRI": "Sri Lanka", "SSD": "Südsudan", "STP": "São Tomé und Príncipe",
        "SUD": "Sudan", "SUI": "Schweiz", "SUR": "Suriname", "SVK": "Slowakei",
        "SWE": "Schweden", "SWZ": "Eswatini", "SYR": "Syrien",
        
        # T
        "TAN": "Tansania", "TCA": "Turks- und Caicosinseln", "TGA": "Tonga",
        "THA": "Thailand", "TJK": "Tadschikistan", "TKM": "Turkmenistan",
        "TLS": "Timor-Leste", "TOG": "Togo", "TPE": "Taiwan", "TTO": "Trinidad und Tobago",
        "TUN": "Tunesien", "TUR": "Türkei", "TUV": "Tuvalu",
        
        # U-V
        "UAE": "Vereinigte Arabische Emirate", "UGA": "Uganda", "UKR": "Ukraine",
        "URU": "Uruguay", "USA": "USA", "UZB": "Usbekistan", "VAN": "Vanuatu",
        "VEN": "Venezuela", "VIE": "Vietnam", "VIN": "St. Vincent und die Grenadinen",
        
        # Y-Z
        "YEM": "Jemen", "ZAM": "Sambia", "ZIM": "Simbabwe",
        
        # Backwards compatibility
        "GB": "Großbritannien", "DE": "Deutschland", "NL": "Niederlande", "CH": "Schweiz",
        "DK": "Dänemark", "AT": "Österreich", "BE": "Belgien", "FR": "Frankreich",
        "IT": "Italien", "ES": "Spanien", "SE": "Schweden", "NO": "Norwegen",
        "PL": "Polen", "CZ": "Tschechien", "HU": "Ungarn", "RO": "Rumänien",
        "IE": "Irland", "PT": "Portugal",
    }
    
    return names.get(ioc_code.upper(), ioc_code)

def render(starterlist: dict, filename: str):
    # Hole Abstände aus starterlist (in cm)
    spacing_top_cm = starterlist.get("spacingTopCm", 3.0)
    spacing_bottom_cm = starterlist.get("spacingBottomCm", 2.0)
    
    print(f"PDF LISTE DEBUG: Abstände - Seite 1 Oben: {spacing_top_cm}cm, Alle Seiten Unten: {spacing_bottom_cm}cm, Ab Seite 2 Oben: 1cm")
    print(f"PDF LISTE DEBUG: Erstelle PDF: {filename}")
    
    top_margin_first = spacing_top_cm * 10
    top_margin_later = 1.0 * 10
    bottom_margin = spacing_bottom_cm * 10
    
    class ListeDocTemplate(SimpleDocTemplate):
        def __init__(self, filename, **kw):
            self.allowSplitting = 1
            SimpleDocTemplate.__init__(self, filename, **kw)
            
        def build(self, flowables):
            frame_first = Frame(
                8*mm, bottom_margin*mm,
                A4[0] - 16*mm, A4[1] - top_margin_first*mm - bottom_margin*mm,
                id='first',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
            )
            
            frame_later = Frame(
                8*mm, bottom_margin*mm,
                A4[0] - 16*mm, A4[1] - top_margin_later*mm - bottom_margin*mm,
                id='later',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
            )
            
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
    style_show = ParagraphStyle("show", parent=styles["Heading1"], fontSize=18, leading=15, spaceAfter=10)
    style_comp = ParagraphStyle("comp", parent=styles["Normal"], fontSize=10, leading=15)
    style_sub = ParagraphStyle("sub", parent=styles["Normal"], fontSize=9, leading=15)
    style_info = ParagraphStyle("info", parent=styles["Normal"], fontSize=11, leading=15)
    style_hdr = ParagraphStyle("hdr", parent=styles["Normal"], fontSize=9, alignment=1)
    style_hdr_left = ParagraphStyle("hdr_left", parent=styles["Normal"], fontSize=9, alignment=0)
    style_pos = ParagraphStyle("pos", parent=styles["Normal"], fontSize=9, alignment=1)
    style_horse = ParagraphStyle("horse", parent=styles["Normal"], fontSize=8, leading=9)
    style_pause = ParagraphStyle("pause", parent=styles["Normal"], fontSize=9, alignment=1)

    elements = []
    
    # Nach erster Seite zu Later-Template wechseln
    elements.append(NextPageTemplate('Later'))
    
    # KEINE Kopfzeile - direkt zur Tabelle
    
    # Competition-Daten
    comp = starterlist.get("competition") or {}
    
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
            colWidths=[A4[0]/2 - 8*mm, A4[0]/2 - 8*mm]
        )
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 3*mm))


    starters = starterlist.get("starters") or []
    breaks = starterlist.get("breaks") or []
    breaks_by_after = {}
    for b in breaks:
        try:
            # WICHTIG: afterNumberInCompetition kann 0 sein! Nicht "or -1" verwenden!
            after_num = b.get("afterNumberInCompetition")
            if after_num is None:
                k = -1
            else:
                k = int(after_num)
            breaks_by_after.setdefault(k, []).append(b)
        except:
            continue

    data_texts = []
    meta = []
    data_texts.append(["<b>Start</b>", "<b>KoNr.</b>", "<b>Reiter / Pferd</b><br/><font size=7>Geschlecht / Farbe / Geboren / Vater x Muttervater / Besitzer / Züchter</font>", "<b>Aufgabe</b>", "<b>Qualität</b>", "<b>Total</b>"])
    meta.append({"type":"header"})
    
    # WICHTIG: Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_by_after:
        for br in breaks_by_after[0]:
            pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append([pause_text, "", "", "", "", ""])  # 6 Spalten
            meta.append({"type":"pause"})

    # Gruppierung
    current_group = None
    group_start_time_shown = False

    for s in starters:
        starter_group = s.get("groupNumber")
        
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            group_text = f"Abteilung {starter_group}"
            data_texts.append([group_text, "", "", "", "", ""])
            meta.append({"type":"group"})
            
            current_group = starter_group
            group_start_time_shown = False

        nr = s.get("startNumber") or ""
        hors_concours = bool(s.get("horsConcours", False))  # Außer Konkurrenz
        tstr = _fmt_time(s.get("startTime"))
        
        if starter_group is None or starter_group == 0 or not group_start_time_shown:
            start_time_display = tstr
            if starter_group is not None and starter_group > 0:
                group_start_time_shown = True
        else:
            start_time_display = ""
        
        # Startnummer mit optionalem AK-Zusatz
        if hors_concours:
            if start_time_display:
                pos_html = f"<b>{nr}</b><br/><font size=7>AK</font><br/><font size=9>{start_time_display}</font>"
            else:
                pos_html = f"<b>{nr}</b><br/><font size=7>AK</font>"
        else:
            pos_html = f"<b>{nr}</b><br/><font size=9>{start_time_display}</font>" if start_time_display else f"<b>{nr}</b>"

        horses = s.get("horses") or []
        cno = ""
        if horses:
            horse = horses[0]
            cno = str(horse.get("cno", ""))

        athlete = s.get("athlete") or {}
        rider_name = _safe_get(athlete, "name", "")
        club = _safe_get(athlete, "club", "")
        nationality = _safe_get(athlete, "nation", "")
        
        content_parts = []
        if rider_name:
            # Reitername fett
            rider_line = f"<b>{rider_name.upper()}</b>"
            
            # Logik: Ausländer zeigen Land nur wenn Verein leer ODER "Gastlizenz GER"
            if nationality and nationality.upper() != "GER":
                # Ausländer
                if not club or club.strip() == "" or club.strip().upper() == "GASTLIZENZ GER":
                    # Kein Verein oder Gastlizenz → Land ausgeschrieben anzeigen
                    country_full = get_country_name(nationality)
                    if country_full:
                        rider_line += f" - <font size=7>{country_full}</font>"
                else:
                    # Hat einen Verein (nicht Gastlizenz) → Verein anzeigen
                    rider_line += f" - <font size=7>{club}</font>"
            elif club:
                # Deutsche: Verein
                rider_line += f" - <font size=7>{club}</font>"
            
            content_parts.append(rider_line)
        
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
            
            owner = _safe_get(horse, "owner", "")
            breeder = _safe_get(horse, "breeder", "")
            
            if owner and breeder:
                # Beide vorhanden
                if owner.strip() == breeder.strip():
                    # Identisch: "B u. Z: Name"
                    content_parts.append(f"<font size=6.5><i>B u. Z: {owner}</i></font>")
                else:
                    # Unterschiedlich: "B: Name / Z: Name"
                    content_parts.append(f"<font size=6.5><i>B: {owner} / Z: {breeder}</i></font>")
            elif owner:
                # Nur Besitzer
                content_parts.append(f"<font size=6.5><i>B: {owner}</i></font>")
            elif breeder:
                # Nur Züchter
                content_parts.append(f"<font size=6.5><i>Z: {breeder}</i></font>")

        combined_content = "<br/>".join(content_parts)

        data_texts.append([pos_html, cno, combined_content, "", "", ""])
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type":"starter","withdrawn":withdrawn_flag, "horsConcours": hors_concours})

        # Breaks
        try:
            cur = int(nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                data_texts.append([pause_text, "", "", "", "", ""])
                meta.append({"type":"pause"})

    # Tabelle erstellen - 6 SPALTEN
    page_width = A4[0] - 16*mm  # 8mm links + 8mm rechts
    col1 = 18*mm
    col2 = 16*mm
    col_aufgabe = 20*mm
    col_qualitaet = 20*mm
    col_total = 20*mm
    col3 = page_width - col1 - col2 - col_aufgabe - col_qualitaet - col_total
    col_widths = [col1, col2, col3, col_aufgabe, col_qualitaet, col_total]

    table_rows = []
    for i, row in enumerate(data_texts):
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr),
                Paragraph(row[2], style_hdr_left), Paragraph(row[3], style_hdr),
                Paragraph(row[4], style_hdr), Paragraph(row[5], style_hdr)
            ])
        elif m["type"] == "group":
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_hdr_left), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub)
            ])
        elif m["type"] == "pause":
            table_rows.append([
                Paragraph(row[0], style_pause), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub)
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
                maybe_strike(row[5], style_pos)
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

    starter_row_count = 0  # Zähler für Starter-Zeilen (für Zebra-Streifen)
    
    for ri in range(1, len(table_rows)):
        if ri < len(meta):
            m = meta[ri]
            if m and m.get("type") == "group":
                ts.add("SPAN", (0,ri), (-1,ri))
                ts.add("BACKGROUND", (0,ri), (-1,ri), colors.Color(0.9, 0.9, 0.9))
                ts.add("ALIGN", (0,ri), (-1,ri), "LEFT")
                # Counter zurücksetzen für neue Abteilung
                starter_row_count = 0
            elif m and m.get("type") == "pause":
                ts.add("SPAN", (0,ri), (-1,ri))
                # Prüfe ob vorheriger Starter grau war
                prev_was_gray = (starter_row_count - 1) % 2 == 1
                if not prev_was_gray:  # Vorherige war weiß → Pause grau
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor('#E8E8E8'))
                ts.add("ALIGN", (0,ri), (-1,ri), "CENTER")
                # WICHTIG: Counter um 1 zurücksetzen, damit nächste Zeile gleiche Farbe hat!
                starter_row_count -= 1
            elif m and m.get("type") == "starter":
                # Zebra: ungerade Starter sind grau (1, 3, 5, ...)
                if starter_row_count % 2 == 1:
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor('#E8E8E8'))
                
                # Withdrawn: Kein extra Hintergrund - nur durchgestrichen im Text
                
                starter_row_count += 1

    t.setStyle(ts)
    elements.append(t)

    # RICHTER MIT AUFGABEN
    judges = comp.get("judges") or starterlist.get("judges") or []
    judging_rule = comp.get("judgingRule") or starterlist.get("judgingRule")
    
    if judges:
        title = f"Richter (Richtverfahren {judging_rule})" if judging_rule else "Richter"
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>{title}</b>", style_sub))
        
        jrows = [[Paragraph("<b>Pos.</b>", style_hdr), Paragraph("<b>Name</b>", style_sub), Paragraph("<b>Aufgabe</b>", style_sub)]]
        ordered_judges = _get_judge_data_for_display_402c(judges, starterlist)
        
        for judge in ordered_judges:
            pos_label = judge["pos_label"]
            name = judge.get("name", "")
            task = judge.get("task", "")
            jrows.append([
                Paragraph(pos_label, style_pos), 
                Paragraph(name, style_horse),
                Paragraph(f"<font size=7>{task}</font>" if task else "", style_horse)
            ])
        
        jt = Table(jrows, colWidths=[18*mm, 57*mm, page_width-75*mm])
        jt.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ALIGN",(0,0),(0,-1),"CENTER"),
            ("ALIGN",(2,0),(2,-1),"LEFT"),
        ]))
        elements.append(jt)

    doc.build(elements)