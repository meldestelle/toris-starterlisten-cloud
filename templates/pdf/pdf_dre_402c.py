# -*- coding: utf-8 -*-
# templates/pdf/pdf_dre_402c_logo.py - Speziell für Richtverfahren 402.C
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
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
    pass  # text already processed
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


def get_nationality_code(nationality_str):
    """Konvertiert IOC-Codes zu ISO 3166-1 alpha-3 Codes"""
    if not nationality_str:
        return ""
    
    ioc_to_iso = {
        "GER": "DEU", "NED": "NLD", "SUI": "CHE", "DEN": "DNK", "CRO": "HRV", "GRE": "GRC",
        "BUL": "BGR", "RSA": "ZAF", "POR": "PRT", "LAT": "LVA", "UAE": "ARE", "CHI": "CHL",
        "URU": "URY", "SLO": "SVN", "MAS": "MYS", "GBR": "GBR", "AUT": "AUT", "BEL": "BEL",
        "FRA": "FRA", "ITA": "ITA", "ESP": "ESP", "SWE": "SWE", "NOR": "NOR", "POL": "POL",
        "CZE": "CZE", "HUN": "HUN", "ROU": "ROU", "IRL": "IRL", "USA": "USA", "CAN": "CAN",
        "AUS": "AUS", "NZL": "NZL", "JPN": "JPN", "KOR": "KOR", "CHN": "CHN", "BRA": "BRA",
        "ARG": "ARG", "MEX": "MEX", "RUS": "RUS", "UKR": "UKR", "TUR": "TUR", "ISR": "ISR",
    }
    
    code = nationality_str.upper()
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
    """Sucht Flagge - Vollständiges ISO→IOC Mapping"""
    if not nat_code:
        return None
    
    iso_to_ioc = {
        "AFG": "AFG", "ALB": "ALB", "DZA": "ALG", "ASM": "ASA", "AND": "AND", "AGO": "ANG",
        "AUT": "AUT", "AZE": "AZE", "BHS": "BAH", "BHR": "BRN", "BGD": "BAN", "BRB": "BAR",
        "BLR": "BLR", "BEL": "BEL", "BRA": "BRA", "BGR": "BUL", "CAN": "CAN", "CHL": "CHI",
        "CHN": "CHN", "HRV": "CRO", "CUB": "CUB", "CYP": "CYP", "CZE": "CZE", "DNK": "DEN",
        "EGY": "EGY", "EST": "EST", "FIN": "FIN", "FRA": "FRA", "DEU": "GER", "GRC": "GRE",
        "HUN": "HUN", "ISL": "ISL", "IND": "IND", "IDN": "INA", "IRN": "IRI", "IRL": "IRL",
        "ISR": "ISR", "ITA": "ITA", "JAM": "JAM", "JPN": "JPN", "KAZ": "KAZ", "KOR": "KOR",
        "LVA": "LAT", "LTU": "LTU", "LUX": "LUX", "MYS": "MAS", "MEX": "MEX", "NLD": "NED",
        "NZL": "NZL", "NOR": "NOR", "POL": "POL", "PRT": "POR", "ROU": "ROU", "RUS": "RUS",
        "ZAF": "RSA", "SVN": "SLO", "ESP": "ESP", "SWE": "SWE", "CHE": "SUI", "TUR": "TUR",
        "UKR": "UKR", "ARE": "UAE", "GBR": "GBR", "USA": "USA", "URY": "URU",
        # Direct IOC codes fallback
        "GER": "GER", "NED": "NED", "SUI": "SUI", "DEN": "DEN", "CRO": "CRO",
    }
    
    flag_code = iso_to_ioc.get(nat_code.upper(), nat_code.upper())
    
    for path in [f"flags/{flag_code}.png", f"C:/Python/flags/{flag_code}.png"]:
        if os.path.exists(path):
            return path
    return None


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
        self.page_num = 0
        self.banner_path = os.path.join("logos", "banner.png")
        self.banner_height = 0
        if os.path.exists(self.banner_path):
            try:
                pil_img = PILImage.open(self.banner_path)
                img_width, img_height = pil_img.size
                dpi = pil_img.info.get('dpi', (72, 72))
                dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
                if dpi_x > 150:
                    display_width = img_width * 72.0 / dpi_x
                    display_height = img_height * 72.0 / dpi_x
                else:
                    display_width = img_width
                    display_height = img_height
                self.banner_width = A4[0]
                self.banner_height = A4[0] * (display_height / display_width)
            except:
                self.banner_path = None
        else:
            self.banner_path = None
        
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
                self.drawImage(self.banner_path, 0, A4[1] - self.banner_height, width=self.banner_width, height=self.banner_height, preserveAspectRatio=True, mask='auto')
            except:
                pass
        
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
    style_horse = ParagraphStyle("horse", parent=styles["Normal"], fontSize=8, leading=9)
    style_pause = ParagraphStyle("pause", parent=styles["Normal"], fontSize=9, alignment=1)

    elements = []
    banner_path = os.path.join("logos", "banner.png")
    has_banner = False
    if os.path.exists(banner_path):
        try:
            pil_img = PILImage.open(banner_path)
            img_width, img_height = pil_img.size
            dpi = pil_img.info.get('dpi', (72, 72))
            dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
            if dpi_x > 150:
                display_width = img_width * 72.0 / dpi_x
                display_height = img_height * 72.0 / dpi_x
            else:
                display_width = img_width
                display_height = img_height
            banner_height = A4[0] * (display_height / display_width)
            elements.append(Spacer(1, banner_height - 5*mm))
            has_banner = True
        except:
            pass

    # LOGO
    logo_path = starterlist.get("logoPath")
    
    # Header-Parts sammeln (werden später mit Logo in Tabelle gepackt)
    header_parts = []
    
    show = starterlist.get("show") or {}
    comp = starterlist.get("competition") or {}

    if not has_banner:
        show_title = show.get("title") or starterlist.get("showTitle") or "Unbenannte Veranstaltung"
        header_parts.append(Paragraph(f"<b>{show_title}</b>", style_show))

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
        header_parts.append(Paragraph(f"<b>{comp_line}</b>", style_comp))

    # 3. Zeile: informationText (nicht mehr fett)
    if informationText:
        processed_info_text = _process_information_text(informationText)
        header_parts.append(Paragraph(processed_info_text, style_info))

    if subtitle:
        header_parts.append(Paragraph(subtitle, style_sub))

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
        header_parts.append(Paragraph(f"<b>{date_line}</b>", style_sub))

    # Logo + Header in Tabelle (wie pdf_std_spr_zucht_komp)
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path)
            max_size = logo_max_width_cm * 10 * mm
            
            if logo.drawWidth > max_size or logo.drawHeight > max_size:
                scale = min(max_size / logo.drawWidth, max_size / logo.drawHeight)
                logo.drawWidth = logo.drawWidth * scale
                logo.drawHeight = logo.drawHeight * scale
            
            # Tabelle: Text links, Logo rechts
            logo_col_width = max(logo.drawWidth + 5*mm, 35*mm)
            text_col_width = 185*mm - logo_col_width
            header_table_data = [[header_parts, logo]]
            header_table = Table(header_table_data, colWidths=[text_col_width, logo_col_width])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (0,0), (0,0), 'LEFT'),
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
                ('LEFTPADDING', (0,0), (0,0), 0),
                ('RIGHTPADDING', (1,0), (1,0), 0),
            ]))
            elements.append(header_table)
        except:
            # Fallback ohne Logo
            for p in header_parts:
                elements.append(p)
    else:
        # Kein Logo
        for p in header_parts:
            elements.append(p)

    elements.append(Spacer(1, 6))


    # TABELLE - 6 SPALTEN für 402.C (Start|KoNr|Reiter/Pferd|Aufgabe|Qualität|Total)
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
    data_texts = []
    meta = []
    # 402.C hat Aufgabe + Qualität + Total
    data_texts.append(["<b># / KNr</b>", "<b>Reiter / Pferd</b><br/><font size=7>Alter / Farbe / Geschlecht / Vater x Muttervater / Besitzer / Züchter</font>", "<b><font size=6>Nat</font></b>", "<b>Aufgabe</b>", "<b>Qualität</b>", "<b>Total</b>"])
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
            data_texts.append([group_text, "", "", "", "", ""])  # 6 Spalten
            meta.append({"type":"group"})
            
            current_group = starter_group
            group_start_time_shown = False

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
        
        horses = s.get("horses") or []
        cno = ""
        if horses:
            horse = horses[0]
            cno = str(horse.get("cno", ""))
        
        # Prüfe withdrawn-Status VOR der Mini-Tabelle Erstellung
        withdrawn_flag = bool(s.get("withdrawn", False))
        
        # Kombinierte Start+KNr Zelle als Mini-Tabelle
        # Aufbau: [Zeit (wenn vorhanden)]
        #         [Nr | KNr]
        start_knr_data = []
        
        # Zeile 1: Startzeit (wenn vorhanden)
        if start_time_display:
            if withdrawn_flag:
                time_text = f'<font size=8><strike>{start_time_display}</strike></font>'
            else:
                time_text = f'<font size=8>{start_time_display}</font>'
            start_knr_data.append([Paragraph(time_text, style_pos)])
        
        # Zeile 2: Nr und KNr nebeneinander
        # Links: Startnummer (mit optionalem AK)
        if withdrawn_flag:
            if hors_concours:
                nr_text = f'<strike><b>{nr}</b></strike><br/><strike><font size=7>AK</font></strike>'
            else:
                nr_text = f'<strike><b>{nr}</b></strike>'
            cno_text = f'<strike><font size=8>{cno}</font></strike>'
        else:
            if hors_concours:
                nr_text = f'<b>{nr}</b><br/><font size=7>AK</font>'
            else:
                nr_text = f'<b>{nr}</b>'
            cno_text = f'<font size=8>{cno}</font>'
        
        nr_cell = Paragraph(nr_text, style_pos)
        knr_cell = Paragraph(cno_text, style_pos)
        start_knr_data.append([nr_cell, knr_cell])
        
        # Erstelle die Mini-Tabelle
        if start_time_display:
            # Mit Zeit: 2 Zeilen
            start_knr_table = Table(start_knr_data, colWidths=[11*mm, 11*mm])
            start_knr_table.setStyle(TableStyle([
                ('SPAN', (0,0), (1,0)),  # Zeit über beide Spalten
                ('GRID', (0,1), (1,1), 0.25, colors.grey),  # Nur untere Zeile hat Trennlinie
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('LEFTPADDING', (0,0), (-1,-1), 1),
                ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('TOPPADDING', (0,0), (-1,-1), 1),
                ('BOTTOMPADDING', (0,0), (-1,-1), 1),
            ]))
        else:
            # Ohne Zeit: nur 1 Zeile
            start_knr_table = Table(start_knr_data, colWidths=[11*mm, 11*mm])
            start_knr_table.setStyle(TableStyle([
                ('GRID', (0,0), (1,0), 0.25, colors.grey),  # Trennlinie zwischen Nr und KNr
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('LEFTPADDING', (0,0), (-1,-1), 1),
                ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('TOPPADDING', (0,0), (-1,-1), 1),
                ('BOTTOMPADDING', (0,0), (-1,-1), 1),
            ]))

        # VERBESSERTE ABSTAMMUNGSDARSTELLUNG
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
            
            # DYNAMISCHES ALTER statt Jahr - ZUERST
            breeding_season = _safe_get(horse, "breedingSeason", "")
            if breeding_season:
                try:
                    from datetime import datetime
                    current_year = datetime.now().year
                    age = current_year - int(breeding_season)
                    details_parts.append(f"{age}jähr.")
                except:
                    pass
            
            # FARBE - ZWEITE
            color = _safe_get(horse, "color", "")
            if color:
                details_parts.append(color)
            
            # GESCHLECHT - DRITTE
            sex = _safe_get(horse, "sex", "")
            sex_german = SEX_MAP.get(str(sex).upper(), sex)
            if sex_german:
                details_parts.append(sex_german)
                
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

        # Nationalität mit Flagge + Kürzel (nationality wurde bereits oben definiert)
        nat_code_iso = get_nationality_code(nationality) if nationality else ""
        
        # Display-Code: Rück-Konvertierung ISO → IOC für Anzeige
        iso_to_display = {
            "DEU": "GER", "NLD": "NED", "CHE": "SUI", "DNK": "DEN", "HRV": "CRO", "GRC": "GRE",
            "BGR": "BUL", "ZAF": "RSA", "PRT": "POR", "LVA": "LAT", "ARE": "UAE", "CHL": "CHI",
            "URY": "URU", "SVN": "SLO", "MYS": "MAS", "GBR": "GBR", "AUT": "AUT", "BEL": "BEL",
            "FRA": "FRA", "ITA": "ITA", "ESP": "ESP", "SWE": "SWE", "NOR": "NOR", "POL": "POL",
            "CZE": "CZE", "HUN": "HUN", "ROU": "ROU", "IRL": "IRL", "USA": "USA", "CAN": "CAN",
        }
        nat_code_display = iso_to_display.get(nat_code_iso, nat_code_iso)
        
        flag_path = find_flag_image(nat_code_iso)
        
        # Flagge + Kürzel in Mini-Tabelle
        nat_cell = ""  # Default fallback
        
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
            except Exception as e:
                # Fallback: Nur Text-Kürzel
                nat_cell = Paragraph(f'<font size="6">{nat_code_display}</font>', style_pos) if nat_code_display else ""
        else:
            # Fallback: Nur Text-Kürzel wenn vorhanden
            if nat_code_display:
                nat_cell = Paragraph(f'<font size="6">{nat_code_display}</font>', style_pos)

        data_texts.append([start_knr_table, combined_content, nat_cell, "", "", ""])  # 6 Spalten
        meta.append({"type":"starter","withdrawn":withdrawn_flag, "horsConcours": hors_concours})


        # Breaks
        try:
            cur = int(nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                data_texts.append([pause_text, "", "", "", "", ""])  # 6 Spalten
                meta.append({"type":"pause"})

    # Tabelle erstellen - 6 SPALTEN
    # Spaltenbreiten - MIT NAT-SPALTE, Start+KNr kombiniert
    page_width = A4[0] - doc.leftMargin - doc.rightMargin
    col1 = 22*mm  # Start+KNr kombiniert
    col_nat = 8*mm  # Nat (Flagge+Kürzel)
    col_aufgabe = 20*mm
    col_qualitaet = 20*mm
    col_total = 20*mm
    col2 = page_width - col1 - col_nat - col_aufgabe - col_qualitaet - col_total  # Reiter+Pferd
    col_widths = [col1, col2, col_nat, col_aufgabe, col_qualitaet, col_total]  # 6 Spalten

    table_rows = []
    for i, row in enumerate(data_texts):
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr_left),
                Paragraph(row[2], style_hdr), Paragraph(row[3], style_hdr),
                Paragraph(row[4], style_hdr), Paragraph(row[5], style_hdr)  # 6 Spalten
            ])
        elif m["type"] == "group":
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_hdr_left), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub)  # 6 Spalten
            ])
        elif m["type"] == "pause":
            table_rows.append([
                Paragraph(row[0], style_pause), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub),
                Paragraph("", style_sub), Paragraph("", style_sub)  # 6 Spalten
            ])
        else:
            withdrawn = m.get("withdrawn", False)
            def maybe_strike(text, s):
                if not text: return Paragraph("", s)
                return Paragraph(f"<strike>{text}</strike>" if withdrawn else text, s)
            
            # Spalte 1 (Start+KNr) ist Table
            start_knr_value = row[0]
            if isinstance(start_knr_value, Table):
                start_knr_display = start_knr_value
            else:
                start_knr_display = Paragraph(str(start_knr_value) if start_knr_value else "", style_pos)
            
            # Spalte 3 (Nat) ist Table/Image/Paragraph
            nat_value = row[2]
            if isinstance(nat_value, (Image, Table)):
                nat_cell_display = nat_value
            elif isinstance(nat_value, Paragraph):
                nat_cell_display = nat_value
            else:
                nat_cell_display = Paragraph(str(nat_value) if nat_value else "", style_pos)
            
            table_rows.append([
                start_knr_display,  # Start+KNr Mini-Tabelle
                maybe_strike(row[1], style_horse),  # Reiter+Pferd
                nat_cell_display,  # Nat
                maybe_strike(row[3], style_pos), maybe_strike(row[4], style_pos),
                maybe_strike(row[5], style_pos)  # 6 Spalten
            ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (0,-1), "CENTER"),  # Start+KNr
        ("ALIGN", (1,0), (1,-1), "LEFT"),    # Reiter+Pferd
        ("ALIGN", (2,0), (-1,-1), "CENTER"), # Rest
    ])

    # Zebra-Streifen mit korrekter Pausen-Logik
    # Zebra wie in pdf_dre_5: Pause hat entgegengesetzte Farbe, counter wird um 1 reduziert
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
                # Pause bekommt entgegengesetzte Farbe vom vorherigen Starter
                prev_was_gray = (starter_count - 1) % 2 == 1
                if not prev_was_gray:  # Vorheriger war weiß → Pause grau
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor("#f5f5f5"))
                # WICHTIG: Counter um 1 zurücksetzen, damit nächste Zeile gleiche Farbe hat!
                starter_count -= 1
            elif m and m.get("type") == "starter":
                # Zebra: ungerade Starter sind grau (1, 3, 5...)
                is_gray_zebra = starter_count % 2 == 1
                
                if is_gray_zebra:
                    ts.add("BACKGROUND", (0,ri), (-1,ri), colors.HexColor("#f5f5f5"))
                
                if m.get("withdrawn", False):
                    ts.add("TEXTCOLOR", (0,ri), (-1,ri), colors.darkgrey)
                
                starter_count += 1

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

    doc.build(elements, canvasmaker=FooterCanvas)