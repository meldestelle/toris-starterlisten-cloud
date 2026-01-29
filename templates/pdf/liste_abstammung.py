# -*- coding: utf-8 -*-
# templates/pdf/liste_abstammung_logo.py
# Listen-Version: Nur Tabelle, kein Kopfbereich
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

# Mapping für Geschlecht - VERBESSERT
SEX_MAP = {
    "MARE": "Stute",
    "GELDING": "Wallach", 
    "STALLION": "Hengst",
    "M": "Wallach",
    "F": "Stute",
    "": ""
}

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

def _map_sex(s):
    if not s:
        return ""
    s_up = str(s).upper()
    if "GELD" in s_up or s_up == "M":
        return "Wallach"
    if "MARE" in s_up or "STUTE" in s_up or s_up == "F":
        return "Stute"
    if "STALL" in s_up or "STALLION" in s_up:
        return "Hengst"
    return s

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
    """Verarbeitet informationText und konvertiert \n zu <br/> für ReportLab, macht Text fett"""
    if not text:
        return ""
    text = text.lstrip('\n')
    text = text.replace('\n', '<br/>')
    text = f"<b>{text}</b>"
    return text

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

def _get_ordered_judges_all(judges):
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    dressage_positions = ["E", "H", "C", "M", "B"]
    judges_by_position = {}  # Dictionary of Lists!
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
            # Füge zu Liste hinzu statt zu überschreiben
            if pos_label not in judges_by_position:
                judges_by_position[pos_label] = []
            judges_by_position[pos_label].append(judge_copy)
        else:
            other_judges.append(judge_copy)
    
    # Return judges in E H C M B order first, then others
    result = []
    for pos in dressage_positions:
        if pos in judges_by_position:
            # Füge ALLE Richter dieser Position hinzu
            result.extend(judges_by_position[pos])
    
    # Add other judges sorted by position label
    other_judges.sort(key=lambda j: j["pos_label"])
    result.extend(other_judges)
    
    return result
def render(starterlist: dict, filename: str):
    print(f"DEBUG: Starting PDF render for {filename}")
    
    # Ausgabe-Pfad
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
    spacing_top_cm = starterlist.get("spacingTopCm", 3.0)
    spacing_bottom_cm = starterlist.get("spacingBottomCm", 2.0)
    
    print(f"PDF LISTE DEBUG: Abstände - Seite 1 Oben: {spacing_top_cm}cm, Alle Seiten Unten: {spacing_bottom_cm}cm, Ab Seite 2 Oben: 1cm")
    
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
    style_show = ParagraphStyle("show", parent=styles["Heading1"], fontSize=18, leading=15,spaceAfter=10)
    style_comp = ParagraphStyle("comp", parent=styles["Normal"], fontSize=10, leading=15)
    style_sub = ParagraphStyle("sub", parent=styles["Normal"], fontSize=9, leading=15)
    style_info = ParagraphStyle("info", parent=styles["Normal"], fontSize=11, leading=15)
    style_hdr = ParagraphStyle("hdr", parent=styles["Normal"], fontSize=9, alignment=1)
    style_hdr_left = ParagraphStyle("hdr_left", parent=styles["Normal"], fontSize=9, alignment=0)  # linksbündig für Gruppen
    style_pos = ParagraphStyle("pos", parent=styles["Normal"], fontSize=9, alignment=1)
    style_rider = ParagraphStyle("rider", parent=styles["Normal"], fontSize=8, leading=9)
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
    data_texts.append(["Start", "KoNr.", "Reiter", "Pferd", "Ergeb."])
    meta.append({"type":"header"})
    
    # WICHTIG: Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_by_after:
        for br in breaks_by_after[0]:
            pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append([pause_text, "", "", "", ""])
            meta.append({"type":"pause"})

    print(f"DEBUG: Processing {len(starters)} starters...")

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
            data_texts.append([group_text, "", "", "", ""])
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
        
        # Startnummer OHNE AK (AK kommt in letzte Spalte)
        pos_html = f"<b>{nr}</b><br/><font size=9>{start_time_display}</font>" if start_time_display else f"<b>{nr}</b>"

        # Get KoNr. (cno) from horse data
        horses = s.get("horses") or []
        cno = ""
        if horses:
            horse = horses[0]
            cno = str(horse.get("cno", ""))

        athlete = s.get("athlete") or {}
        rider_html = "<br/>".join([
            f"<b>{athlete.get('name','')}</b>" if athlete.get("name") else "",
            f"<font size=7>{athlete.get('club')}</font>" if athlete.get("club") else "",
            f"<font size=7>{athlete.get('nation')}</font>" if athlete.get("nation") else "",
        ])

        horse_html = ""
        for h in horses:
            # Pferdename fett, Zuchtgebiet normal mit " - " getrennt
            name_line = f"<b>{h.get('name','')}</b>"
            if h.get("studbook"):
                name_line += f" - {h.get('studbook')}"
            
            # Kombiniere alle Details in einer Zeile mit " / " getrennt
            details = []
            if h.get("breedingSeason"): 
                details.append(str(h.get("breedingSeason")))
            if h.get("color"): 
                details.append(h.get("color"))
            if h.get("sex"): 
                sex_german = SEX_MAP.get(str(h.get("sex")).upper(), h.get("sex"))
                details.append(sex_german)
            
            # Pedigree hinzufügen falls vorhanden
            if h.get("sire") or h.get("damSire"):
                ped = " x ".join([p for p in (h.get("sire"), h.get("damSire")) if p])
                details.append(ped)
            
            # Alle Details in einer Zeile
            if details:
                name_line += f"<br/><font size=7>{' / '.join(details)}</font>"
            
            # Besitzer und Züchter kompakt in einer Zeile (wie im Beispiel-Template)
            owner = h.get("owner")
            breeder = h.get("breeder")
            
            if owner and breeder:
                # Beide vorhanden
                if owner.strip() == breeder.strip():
                    # Identisch: "B u. Z: Name"
                    name_line += f"<br/><font size=6.5><i>B u. Z: {owner}</i></font>"
                else:
                    # Unterschiedlich: "B: Name / Z: Name"
                    name_line += f"<br/><font size=6.5><i>B: {owner} / Z: {breeder}</i></font>"
            elif owner:
                # Nur Besitzer
                name_line += f"<br/><font size=6.5><i>B: {owner}</i></font>"
            elif breeder:
                # Nur Züchter
                name_line += f"<br/><font size=6.5><i>Z: {breeder}</i></font>"
            
            horse_html += name_line + ("<br/><br/>" if len(horses) > 1 else "")

        # AK in letzte Spalte (Ergebnis-Spalte) mittig
        result_html = "AK" if hors_concours else ""
        
        data_texts.append([pos_html, cno, rider_html or "", horse_html or "", result_html])
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type":"starter","withdrawn":withdrawn_flag, "horsConcours": hors_concours})

        try:
            cur = int(nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                data_texts.append([pause_text, "", "", "", ""])
                meta.append({"type":"pause"})

    print(f"DEBUG: Created data_texts with {len(data_texts)} rows")

    # --- ZURÜCK ZUR EINFACHEN TABELLENLÖSUNG (ohne komplexe Gruppierung) ---
    
    # Spaltenbreiten definieren
    # WICHTIG: Da wir PageTemplate mit Frames verwenden, müssen wir die Ränder manuell abziehen
    page_width = A4[0] - 16*mm  # 8mm links + 8mm rechts
    col1 = 18*mm
    col2 = 16*mm      # KoNr. breiter gemacht (von 12mm auf 16mm)
    col3 = 45*mm
    col5 = 20*mm
    col4 = page_width - (col1+col2+col3+col5)
    if col4 < 50*mm: col4 = 50*mm
    col_widths = [col1, col2, col3, col4, col5]

    table_rows = []
    for i, row in enumerate(data_texts):
        if i >= len(meta):  # Sicherheitsprüfung
            continue
            
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr),
                Paragraph(row[2], style_sub), Paragraph(row[3], style_sub),
                Paragraph(row[4], style_hdr)
            ])
        elif m["type"] == "group":
            # GRUPPIERUNGSLOGIK - LINKSBÜNDIG (nur wenn Gruppen vorhanden)
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_hdr_left), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub)
            ])
        elif m["type"] == "pause":
            table_rows.append([
                Paragraph(row[0], style_pause), Paragraph("", style_sub),
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
                maybe_strike(row[2], style_rider),
                maybe_strike(row[3], style_horse),
                maybe_strike(row[4], style_pos),
            ])

    print(f"DEBUG: About to create table with {len(table_rows)} rows")

    try:
        t = Table(table_rows, colWidths=col_widths, repeatRows=1)
        ts = TableStyle([
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ALIGN", (0,0), (0,-1), "CENTER"),
            ("ALIGN", (1,0), (1,-1), "CENTER"),
            ("ALIGN", (-1,0), (-1,-1), "CENTER"),
        ])

        # Styling für spezielle Zeilen (OHNE PageBreak-Versuche)
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
                    
                    # Withdrawn: Kein extra Hintergrund - nur durchgestrichen im Text (via maybe_strike)
                    
                    starter_row_count += 1

        t.setStyle(ts)
        elements.append(t)
        print("DEBUG: Table created and added successfully")
    except Exception as e:
        print(f"DEBUG: Table creation error: {e}")
        raise

    # --- Richter ---
    judges = comp.get("judges") or starterlist.get("judges") or []
    judging_rule = comp.get("judgingRule") or starterlist.get("judgingRule")
    if judges:
        print(f"DEBUG: Adding {len(judges)} judges...")
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
        print("DEBUG: Judges table added successfully")

    print("DEBUG: About to build PDF document...")
    try:
        doc.build(elements)
        print("DEBUG: PDF build completed successfully!")
    except Exception as e:
        print(f"DEBUG: PDF build error: {e}")
        raise