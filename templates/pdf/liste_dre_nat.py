# -*- coding: utf-8 -*-
# templates/pdf/liste_nat.py - Deutsches Design mit Flaggen
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageTemplate, Frame, NextPageTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
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

import os

WEEKDAY_MAP = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
}
MONTH_MAP = {
    "January": "Januar", "February": "Februar", "March": "März", "April": "April",
    "May": "Mai", "June": "Juni", "July": "Juli", "August": "August",
    "September": "September", "October": "Oktober", "November": "November", "December": "Dezember"
}

def get_nationality_code(nationality_str):
    """IOC → ISO 3166-1 alpha-3"""
    if not nationality_str:
        return ""
    code = str(nationality_str).strip().upper()
    ioc_to_iso = {
        "GER": "DEU", "NED": "NLD", "SUI": "CHE", "DEN": "DNK", "CRO": "HRV", "GRE": "GRC",
        "BUL": "BGR", "RSA": "ZAF", "POR": "PRT", "LAT": "LVA", "UAE": "ARE", "CHI": "CHL",
        "URU": "URY", "SLO": "SVN", "MAS": "MYS",
        # 2-letter codes (ISO 3166-1 alpha-2) that might be used
        "GB": "GBR", "DE": "DEU", "NL": "NLD", "CH": "CHE", "DK": "DNK", "AT": "AUT",
        "BE": "BEL", "FR": "FRA", "IT": "ITA", "ES": "ESP", "SE": "SWE", "NO": "NOR",
        "PL": "POL", "CZ": "CZE", "HU": "HUN", "RO": "ROU", "IE": "IRL", "PT": "PRT",
        # 3-letter codes
        "ALB": "ALB", "ALG": "DZA", "ARG": "ARG", "ARM": "ARM", "AUS": "AUS", "AUT": "AUT",
        "AZE": "AZE", "BEL": "BEL", "BIH": "BIH", "BLR": "BLR", "BRA": "BRA", "CAN": "CAN",
        "CHN": "CHN", "COL": "COL", "CZE": "CZE", "EGY": "EGY", "ESP": "ESP", "EST": "EST",
        "FIN": "FIN", "FRA": "FRA", "GBR": "GBR", "GEO": "GEO", "HUN": "HUN", "IND": "IND",
        "IRL": "IRL", "ISL": "ISL", "ISR": "ISR", "ITA": "ITA", "JPN": "JPN", "KAZ": "KAZ",
        "KEN": "KEN", "KOR": "KOR", "LTU": "LTU", "LUX": "LUX", "MAR": "MAR", "MDA": "MDA",
        "MEX": "MEX", "MKD": "MKD", "MNE": "MNE", "NOR": "NOR", "NZL": "NZL", "POL": "POL",
        "QAT": "QAT", "ROU": "ROU", "RUS": "RUS", "SGP": "SGP", "SRB": "SRB", "SVK": "SVK",
        "SWE": "SWE", "THA": "THA", "TUR": "TUR", "UKR": "UKR", "USA": "USA", "VEN": "VEN",
    }
    return ioc_to_iso.get(code, code[:3] if len(code) >= 3 else code)

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


def find_flag_image(nat_code):
    """Sucht Flagge - WICHTIG: Flaggen haben IOC-Namen (ger.png), nicht ISO (deu.png)!"""
    if not nat_code:
        return None
    
    # Rück-Konvertierung ISO → IOC/Flaggen-Dateinamen (alle 69 Länder)
    iso_to_ioc = {
        # Spezielle Mappings (IOC ≠ ISO)
        "DEU": "ger", "NLD": "ned", "CHE": "sui", "DNK": "den", "HRV": "cro", "GRC": "gre",
        "BGR": "bul", "ZAF": "rsa", "PRT": "por", "LVA": "lat", "ARE": "uae", "CHL": "chi",
        "URY": "uru", "SVN": "slo", "MYS": "mas",
        # Alle anderen (ISO = IOC, aber vollständig für Klarheit)
        "ALB": "alb", "DZA": "alg", "ARG": "arg", "ARM": "arm", "AUS": "aus", "AUT": "aut",
        "AZE": "aze", "BEL": "bel", "BIH": "bih", "BLR": "blr", "BRA": "bra", "CAN": "can",
        "CHN": "chn", "COL": "col", "CZE": "cze", "EGY": "egy", "ESP": "esp", "EST": "est",
        "FIN": "fin", "FRA": "fra", "GBR": "gbr", "GEO": "geo", "HUN": "hun", "IND": "ind",
        "IRL": "irl", "ISL": "isl", "ISR": "isr", "ITA": "ita", "JPN": "jpn", "KAZ": "kaz",
        "KEN": "ken", "KOR": "kor", "LTU": "ltu", "LUX": "lux", "MAR": "mar", "MDA": "mda",
        "MEX": "mex", "MKD": "mkd", "MNE": "mne", "NOR": "nor", "NZL": "nzl", "POL": "pol",
        "QAT": "qat", "ROU": "rou", "RUS": "rus", "SGP": "sgp", "SRB": "srb", "SVK": "svk",
        "SWE": "swe", "THA": "tha", "TUR": "tur", "UKR": "ukr", "USA": "usa", "VEN": "ven",
    }
    
    # Wenn ISO-Code, zurück zu IOC konvertieren
    flag_code = iso_to_ioc.get(nat_code.upper(), nat_code.lower())
    
    for path in [f"flags/{flag_code}.png", f"C:/Python/flags/{flag_code}.png"]:
        if os.path.exists(path):
            return path
    return None

def translate_sex(sex):
    sex_map = {"STALLION": "Hengst", "GELDING": "Wallach", "MARE": "Stute"}
    return sex_map.get(str(sex).upper() if sex else "", sex or "")

def format_time(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%H:%M:%S")
    except:
        return str(iso).split("T")[-1][:8] if "T" in str(iso) else str(iso)

def format_datetime_german(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        wd = WEEKDAY_MAP.get(dt.strftime("%A"), dt.strftime("%A"))
        mo = MONTH_MAP.get(dt.strftime("%B"), dt.strftime("%B"))
        return f"{wd}, {dt.day}. {mo} {dt.year} {dt.strftime('%H:%M')} Uhr"
    except:
        return str(iso)

def format_pause_text(total_seconds, info):
    secs = int(total_seconds or 0)
    if secs == 0:
        return info or "Pause"
    if secs >= 3600:
        h = secs // 3600
        m = (secs % 3600) // 60
        return f"Pause ({h} Std. {m:02d} Min.){' - ' + info if info else ''}"
    elif secs >= 60:
        return f"Pause ({secs//60} Min.){' - ' + info if info else ''}"
    else:
        return f"Pause ({secs} Sek.){' - ' + info if info else ''}"

class FooterCanvas(canvas.Canvas):
    """Canvas mit Sponsorenleiste im Footer"""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        
    def showPage(self):
        self.draw_footer()
        canvas.Canvas.showPage(self)
        
    def draw_footer(self):
        sponsor_path = "logos/sponsorenleiste.png"
        if os.path.exists(sponsor_path):
            try:
                # Sponsorenleiste unten zentriert - GLEICHE GRÖSSE WIE WORD (7.5 inches)
                img_width = 190*mm  # 7.5 inches = ca. 190mm
                img_height = 25*mm  # Proportional etwas höher
                x = (A4[0] - img_width) / 2  # Zentriert
                y = 7*mm  # 7mm vom unteren Rand (näher am Rand)
                self.drawImage(sponsor_path, x, y, width=img_width, height=img_height, 
                              preserveAspectRatio=True, mask='auto')
            except:
                pass

def render(starterlist, filename):
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
                10*mm, bottom_margin*mm,
                A4[0] - 20*mm, A4[1] - top_margin_first*mm - bottom_margin*mm,
                id='first',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
            )
            
            frame_later = Frame(
                10*mm, bottom_margin*mm,
                A4[0] - 20*mm, A4[1] - top_margin_later*mm - bottom_margin*mm,
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

    # Styles definieren
    styles = getSampleStyleSheet()
    style_comp = ParagraphStyle('Comp', fontSize=11, leading=13, fontName='Helvetica', spaceAfter=2)
    style_info = ParagraphStyle('Info', fontSize=10, leading=12, fontName='Helvetica-Bold', spaceAfter=2)
    style_sub = ParagraphStyle('Sub', fontSize=10, leading=12, fontName='Helvetica', spaceAfter=2)
    style_hdr = ParagraphStyle('Hdr', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER, textColor=colors.white)
    style_hdr_left = ParagraphStyle('HdrLeft', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_LEFT, textColor=colors.white)
    style_pos = ParagraphStyle('Pos', fontSize=8, leading=10, fontName='Helvetica', alignment=TA_CENTER)
    style_rider = ParagraphStyle('Rider', fontSize=9, leading=11, fontName='Helvetica', alignment=TA_LEFT)
    style_horse = ParagraphStyle('Horse', fontSize=9, leading=11, fontName='Helvetica', alignment=TA_LEFT)
    style_pause = ParagraphStyle('Pause', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER)
    style_group = ParagraphStyle('Group', fontSize=10, leading=12, fontName='Helvetica-Bold', alignment=TA_LEFT, textColor=colors.white)
    
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


    starters = starterlist.get("starters", [])
    breaks = starterlist.get("breaks", [])
    breaks_map = {}
    for b in breaks:
        try:
            # WICHTIG: afterNumberInCompetition kann 0 sein!
            after_num = b.get("afterNumberInCompetition")
            if after_num is None:
                k = -1
            else:
                k = int(after_num)
            breaks_map.setdefault(k, []).append(b)
        except:
            pass
    
    data_texts = []
    meta = []
    
    # Header
    data_texts.append(["#", "Zeit", "KNR", "Pferd", "Reiter", "Nat."])
    meta.append({"type": "header"})
    
    # WICHTIG: Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_map:
        for br in breaks_map[0]:
            pause_text = format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append([pause_text, "", "", "", "", "", ""])
            meta.append({"type": "pause"})
    
    # Gruppierungslogik (wie in pdf_abstammung_logo)
    current_group = None
    group_start_time_shown = False
    
    for s in starters:
        # Prüfe auf Gruppenwechsel (Abteilung) BEVOR der Starter hinzugefügt wird
        starter_group = s.get("groupNumber")
        
        # Nur wenn groupNumber existiert und > 0
        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            # Neue Gruppe erkannt - Gruppen-Header hinzufügen
            group_text = f"Abteilung {starter_group}"
            data_texts.append([group_text, "", "", "", "", ""])
            meta.append({"type": "group"})
            
            current_group = starter_group
            group_start_time_shown = False  # Reset für neue Gruppe
        
        nr = str(s.get("startNumber", ""))
        hors_concours = bool(s.get("horsConcours", False))  # Außer Konkurrenz  # Convert to string immediately!
        
        # Zeit nur beim ersten Starter pro Gruppe zeigen, oder wenn keine Gruppierung
        if starter_group is None or starter_group == 0 or not group_start_time_shown:
            time_str = format_time(s.get("startTime", ""))
            if starter_group is not None and starter_group > 0:
                group_start_time_shown = True  # Markiere Zeit als gezeigt
        else:
            time_str = ""  # Verstecke Zeit für weitere Starter in derselben Gruppe
        
        horses = s.get("horses", [])
        horse = horses[0] if horses else {}
        cno = str(horse.get("cno", ""))
        
        # Pferd
        horse_html = ""
        if horse:
            horse_name = horse.get("name", "")
            horse_html = f"<b>{horse_name}</b>"
            
            details = []
            breeding_season = horse.get("breedingSeason")
            if breeding_season:
                try:
                    current_year = datetime.now().year
                    age = current_year - int(breeding_season)
                    details.append(f"{age}jähr.")
                except:
                    pass
            
            color = horse.get("color", "")
            if color:
                details.append(color)
            
            sex = horse.get("sex", "")
            if sex:
                details.append(translate_sex(sex))
            
            studbook = horse.get("studbook", "")
            if studbook:
                details.append(studbook)
            
            sire = horse.get("sire", "")
            dam_sire = horse.get("damSire", "")
            if sire and dam_sire:
                details.append(f"{sire} x {dam_sire}")
            elif sire:
                details.append(sire)
            
            if details:
                horse_html += f"<br/><font size=7>{' / '.join(details)}</font>"
            
            owner = horse.get("owner", "")
            breeder = horse.get("breeder", "")
            if owner:
                horse_html += f"<br/><font size=7><i>Besitzer: {owner}</i></font>"
            if breeder:
                horse_html += f"<br/><font size=7><i>Züchter: {breeder}</i></font>"
        
        # Reiter mit Club/Land
        athlete = s.get("athlete", {})
        athlete_name = str(athlete.get("name", ""))  # Ensure string
        club = athlete.get("club", "")
        nationality = athlete.get("nation", "")
        
        # Reiter-HTML: Name + Club/Land darunter
        athlete_html = f"<b>{athlete_name}</b>" if athlete_name else ""
        if nationality and nationality.upper() != "GER":
            # Ausländer: Land ausgeschrieben
            country_full = get_country_name(nationality)
            if country_full:
                athlete_html += f"<br/><font size=7>{country_full}</font>"
        elif club:
            # Deutsche: Verein
            athlete_html += f"<br/><font size=7>{club}</font>"
        
        # Nationalität mit Flagge UND Kürzel
        nat_code_iso = get_nationality_code(nationality) if nationality else ""  # ISO für Mapping
        
        # Display-Code: Rück-Konvertierung ISO → IOC für Anzeige
        iso_to_display = {
            "DEU": "GER", "NLD": "NED", "CHE": "SUI", "DNK": "DEN", "HRV": "CRO", "GRC": "GRE",
            "BGR": "BUL", "ZAF": "RSA", "PRT": "POR", "LVA": "LAT", "ARE": "UAE", "CHL": "CHI",
            "URY": "URU", "SVN": "SLO", "MYS": "MAS", "GBR": "GBR",
        }
        nat_code_display = iso_to_display.get(nat_code_iso, nat_code_iso)  # Anzeige-Code (GER, GBR, etc.)
        
        flag_path = find_flag_image(nat_code_iso)
        
        # Flagge + Kürzel in Mini-Tabelle
        if flag_path and os.path.exists(flag_path):
            try:
                flag_img = Image(flag_path, width=5*mm, height=3.5*mm)
                # Mini-Tabelle: Flagge oben, Code unten
                nat_mini_data = [[flag_img], [Paragraph(f'<font size="6">{nat_code_display}</font>', style_pos)]]
                nat_cell = Table(nat_mini_data, colWidths=[5*mm])
                nat_cell.setStyle(TableStyle([
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                    ('RIGHTPADDING', (0,0), (-1,-1), 0),
                    ('TOPPADDING', (0,0), (-1,-1), 0),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 1),
                ]))
                print(f"DEBUG: Flagge+Code für {nat_code_display}: {flag_path}")
            except Exception as e:
                print(f"DEBUG: Flagge FEHLER für {nat_code_display}: {e}")
                nat_cell = nat_code_display
        else:
            print(f"DEBUG: Keine Flagge für {nat_code_display}")
            nat_cell = nat_code_display
        
        # Startnummer mit optionalem AK-Zusatz
        nr_display = nr
        if hors_concours:
            nr_display = f"{nr}<br/><font size=7>AK</font>"
        
        data_texts.append([nr_display, time_str, cno, horse_html, athlete_html, nat_cell])
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type": "starter", "withdrawn": withdrawn_flag, "horsConcours": hors_concours})
        
        # Break
        try:
            cur = int(nr)
            if cur in breaks_map:
                for br in breaks_map[cur]:
                    pause_text = format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                    data_texts.append([pause_text, "", "", "", "", ""])
                    meta.append({"type": "pause"})
        except:
            pass
    
    # Spaltenbreiten
    page_width = A4[0] - 20*mm  # 10mm links + 10mm rechts
    col_widths = [10*mm, 16*mm, 12*mm, 85*mm, 50*mm, 13*mm]  # +5mm Horse, +5mm Athlete
    
    table_rows = []
    for i, row in enumerate(data_texts):
        if i >= len(meta):
            continue
        
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr),
                Paragraph(row[2], style_hdr), Paragraph(row[3], style_hdr_left),
                Paragraph(row[4], style_hdr_left), Paragraph(row[5], style_hdr)
            ])
        elif m["type"] == "group":
            # Abteilungs-Header (linksbündig, dunkler Hintergrund, weiße Schrift)
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_group), Paragraph("", style_sub),
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
                if not text:
                    return Paragraph("", s)
                return Paragraph(f"<strike>{text}</strike>" if withdrawn else text, s)
            
            # Nat-Spalte: Wenn Table-Objekt (Flagge+Code) direkt nutzen, sonst Paragraph
            nat_value = row[5]
            if isinstance(nat_value, Table):
                nat_cell_final = nat_value
            elif isinstance(nat_value, Image):
                nat_cell_final = nat_value
            else:
                nat_cell_final = Paragraph(str(nat_value), style_pos)
            
            table_rows.append([
                maybe_strike(row[0], style_pos),
                maybe_strike(row[1], style_pos),
                maybe_strike(row[2], style_pos),
                maybe_strike(row[3], style_horse),
                maybe_strike(row[4], style_rider),
                nat_cell_final
            ])
    
    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("LINEBELOW", (0,0), (-1,-1), 0.5, colors.black),  # Nur horizontale Linien!
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#404040')),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (2,-1), "CENTER"),
        ("ALIGN", (3,0), (4,-1), "LEFT"),
        ("ALIGN", (5,0), (5,-1), "CENTER"),
    ])
    
    # SPAN + Zebra - mit Gruppen-Logik!
    starter_row_count = 0
    for ri in range(1, len(table_rows)):
        if ri < len(meta):
            m = meta[ri]
            if m.get("type") == "group":
                # Abteilungs-Header: SPAN über alle Spalten, dunkler Hintergrund wie Header
                ts.add("SPAN", (0,ri), (5,ri))
                ts.add("BACKGROUND", (0,ri), (5,ri), colors.HexColor('#404040'))
                # Counter zurücksetzen für neue Abteilung
                starter_row_count = 0
            elif m.get("type") == "starter":
                # Zebra: ungerade Starter sind grau (1, 3, 5, ...)
                if starter_row_count % 2 == 1:
                    ts.add("BACKGROUND", (0,ri), (5,ri), colors.HexColor('#E8E8E8'))
                starter_row_count += 1
            elif m.get("type") == "pause":
                ts.add("SPAN", (0,ri), (5,ri))
                prev_was_gray = (starter_row_count - 1) % 2 == 1
                if not prev_was_gray:  # Vorherige war weiß → Pause grau
                    ts.add("BACKGROUND", (0,ri), (5,ri), colors.HexColor('#E8E8E8'))
                # WICHTIG: Counter um 1 zurücksetzen, damit nächste Zeile gleiche Farbe hat!
                starter_row_count -= 1
    
    t.setStyle(ts)
    elements.append(t)
    
    # Build mit Footer-Canvas
    doc.build(elements)
    print(f"PDF NAT: {filename}")
