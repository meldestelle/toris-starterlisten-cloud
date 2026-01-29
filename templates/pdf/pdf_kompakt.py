# -*- coding: utf-8 -*-
# templates/pdf/pdf_nat.py - Deutsches Design mit Flaggen
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageTemplate, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
from datetime import datetime
import os
from PIL import Image as PILImage

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
        
        # Backwards compatibility - 2-letter codes
        "GB": "Großbritannien", "DE": "Deutschland", "NL": "Niederlande", "CH": "Schweiz",
        "DK": "Dänemark", "AT": "Österreich", "BE": "Belgien", "FR": "Frankreich",
        "IT": "Italien", "ES": "Spanien", "SE": "Schweden", "NO": "Norwegen",
        "PL": "Polen", "CZ": "Tschechien", "HU": "Ungarn", "RO": "Rumänien",
        "IE": "Irland", "PT": "Portugal",
        
        # Alternative spellings
        "DZA": "Algerien",  # ISO code for Algeria
    }
    
    return names.get(ioc_code.upper(), ioc_code)


def find_flag_image(nat_code):
    """Sucht Flagge - Vollständiges ISO→IOC Mapping für alle 250 Länder"""
    if not nat_code:
        return None
    
    # Vollständiges ISO 3166-1 alpha-3 → IOC Code Mapping
    iso_to_ioc = {
        "AFG": "AFG", "ALB": "ALB", "DZA": "ALG", "ASM": "ASA", "AND": "AND", "AGO": "ANG",
        "AIA": "AIA", "ATG": "ANT", "ARG": "ARG", "ARM": "ARM", "ABW": "ARU", "AUS": "AUS",
        "AUT": "AUT", "AZE": "AZE", "BHS": "BAH", "BHR": "BRN", "BGD": "BAN", "BRB": "BAR",
        "BLR": "BLR", "BEL": "BEL", "BLZ": "BIZ", "BEN": "BEN", "BMU": "BER", "BTN": "BHU",
        "BOL": "BOL", "BIH": "BIH", "BWA": "BOT", "BRA": "BRA", "VGB": "IVB", "BRN": "BRU",
        "BGR": "BUL", "BFA": "BUR", "BDI": "BDI", "KHM": "CAM", "CMR": "CMR", "CAN": "CAN",
        "CPV": "CPV", "CYM": "CAY", "CAF": "CAF", "TCD": "CHA", "CHL": "CHI", "CHN": "CHN",
        "COL": "COL", "COM": "COM", "COG": "CGO", "COD": "COD", "COK": "COK", "CRI": "CRC",
        "CIV": "CIV", "HRV": "CRO", "CUB": "CUB", "CYP": "CYP", "CZE": "CZE", "DNK": "DEN",
        "DJI": "DJI", "DMA": "DMA", "DOM": "DOM", "ECU": "ECU", "EGY": "EGY", "SLV": "ESA",
        "GNQ": "GEQ", "ERI": "ERI", "EST": "EST", "ETH": "ETH", "FRO": "FAR", "FJI": "FIJ",
        "FIN": "FIN", "FRA": "FRA", "GAB": "GAB", "GMB": "GAM", "GEO": "GEO", "DEU": "GER",
        "GHA": "GHA", "GBR": "GBR", "GRC": "GRE", "GRD": "GRN", "GUM": "GUM", "GTM": "GUA",
        "GIN": "GUI", "GNB": "GBS", "GUY": "GUY", "HTI": "HAI", "HND": "HON", "HKG": "HKG",
        "HUN": "HUN", "ISL": "ISL", "IND": "IND", "IDN": "INA", "IRN": "IRI", "IRQ": "IRQ",
        "IRL": "IRL", "ISR": "ISR", "ITA": "ITA", "JAM": "JAM", "JPN": "JPN", "JOR": "JOR",
        "KAZ": "KAZ", "KEN": "KEN", "KIR": "KIR", "PRK": "PRK", "KOR": "KOR", "KWT": "KUW",
        "KGZ": "KGZ", "LAO": "LAO", "LVA": "LAT", "LBN": "LBN", "LSO": "LES", "LBR": "LBR",
        "LBY": "LBA", "LIE": "LIE", "LTU": "LTU", "LUX": "LUX", "MAC": "MAC", "MDG": "MAD",
        "MWI": "MAW", "MYS": "MAS", "MDV": "MDV", "MLI": "MLI", "MLT": "MLT", "MHL": "MHL",
        "MRT": "MTN", "MUS": "MRI", "MEX": "MEX", "FSM": "FSM", "MDA": "MDA", "MCO": "MON",
        "MNG": "MGL", "MNE": "MNE", "MAR": "MAR", "MOZ": "MOZ", "MMR": "MYA", "NAM": "NAM",
        "NRU": "NRU", "NPL": "NEP", "NLD": "NED", "NZL": "NZL", "NIC": "NCA", "NER": "NIG",
        "NGA": "NGR", "NOR": "NOR", "OMN": "OMA", "PAK": "PAK", "PLW": "PLW", "PSE": "PLE",
        "PAN": "PAN", "PNG": "PNG", "PRY": "PAR", "PER": "PER", "PHL": "PHI", "POL": "POL",
        "PRT": "POR", "PRI": "PUR", "QAT": "QAT", "MKD": "MKD", "ROU": "ROU", "RUS": "RUS",
        "RWA": "RWA", "KNA": "SKN", "LCA": "LCA", "VCT": "VIN", "WSM": "SAM", "SMR": "SMR",
        "STP": "STP", "SAU": "KSA", "SEN": "SEN", "SRB": "SRB", "SYC": "SEY", "SLE": "SLE",
        "SGP": "SGP", "SVK": "SVK", "SVN": "SLO", "SLB": "SOL", "SOM": "SOM", "ZAF": "RSA",
        "SSD": "SSD", "ESP": "ESP", "LKA": "SRI", "SDN": "SUD", "SUR": "SUR", "SWZ": "SWZ",
        "SWE": "SWE", "CHE": "SUI", "SYR": "SYR", "TWN": "TPE", "TJK": "TJK", "TZA": "TAN",
        "THA": "THA", "TLS": "TLS", "TGO": "TOG", "TON": "TGA", "TTO": "TTO", "TUN": "TUN",
        "TUR": "TUR", "TKM": "TKM", "TCA": "TCA", "TUV": "TUV", "UGA": "UGA", "UKR": "UKR",
        "ARE": "UAE", "USA": "USA", "URY": "URU", "UZB": "UZB", "VUT": "VAN", "VEN": "VEN",
        "VNM": "VIE", "VIR": "ISV", "YEM": "YEM", "ZMB": "ZAM", "ZWE": "ZIM", "XKX": "KOS"
    }
    
    # Wenn ISO-Code übergeben wird, zu IOC konvertieren
    flag_code = iso_to_ioc.get(nat_code.upper(), nat_code.upper())
    
    # Suche in verschiedenen Pfaden
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

# Alias for compatibility with liste_kompakt
_format_time = format_time

def _maybe_strike(text, style, withdrawn=False):
    """Helper function to strikethrough text for withdrawn entries"""
    if not text:
        return Paragraph("", style)
    if withdrawn:
        return Paragraph(f"<strike>{text}</strike>", style)
    else:
        return Paragraph(text, style)

def _get_ordered_judges_all(judges):
    """Order judges: E H C M B first, then others alphabetically"""
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

def get_sponsor_bar_height():
    sponsor_path = "logos/sponsorenleiste.png"
    if not os.path.exists(sponsor_path): return 25*mm
    try:
        pil_img = PILImage.open(sponsor_path)
        img_width, img_height = pil_img.size
        target_width_mm = 190
        aspect_ratio = img_height / img_width
        return target_width_mm * aspect_ratio * mm
    except: return 25*mm

class BannerCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
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
                    scale_factor = 72.0 / dpi_x
                    display_width = img_width * scale_factor
                    display_height = img_height * scale_factor
                else:
                    display_width = img_width
                    display_height = img_height
                page_width = A4[0]
                aspect_ratio = display_height / display_width
                self.banner_width = page_width
                self.banner_height = page_width * aspect_ratio
            except: self.banner_path = None
        else: self.banner_path = None
        self.sponsor_height = get_sponsor_bar_height()
    def showPage(self):
        self.draw_banner(); self.draw_footer(); self.page_num += 1; canvas.Canvas.showPage(self)
    def draw_banner(self):
        if self.page_num != 0: return
        if self.banner_path and os.path.exists(self.banner_path):
            try: self.drawImage(self.banner_path, 0, A4[1] - self.banner_height, width=self.banner_width, height=self.banner_height, preserveAspectRatio=True, mask='auto')
            except: pass
    def draw_footer(self):
        sponsor_path = "logos/sponsorenleiste.png"
        if os.path.exists(sponsor_path):
            try: self.drawImage(sponsor_path, (A4[0] - 190*mm) / 2, 4*mm, width=190*mm, height=self.sponsor_height, preserveAspectRatio=True, mask='auto')
            except: pass
def render(starterlist, filename, logo_max_width_cm=5.0):
    sponsor_height = get_sponsor_bar_height(); bottom_margin = 4*mm + sponsor_height + 1*mm
    doc = SimpleDocTemplate(
        filename, pagesize=A4, 
        rightMargin=10*mm, leftMargin=10*mm,  # Kleiner für mehr Platz!
        topMargin=5*mm, bottomMargin=bottom_margin
    )
    
    elements = []
    banner_path = os.path.join("logos", "banner.png"); has_banner = False
    if os.path.exists(banner_path):
        try:
            pil_img = PILImage.open(banner_path); img_width, img_height = pil_img.size
            dpi = pil_img.info.get('dpi', (72, 72)); dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
            if dpi_x > 150: display_width = img_width * 72.0 / dpi_x; display_height = img_height * 72.0 / dpi_x
            else: display_width = img_width; display_height = img_height
            banner_height = A4[0] * (display_height / display_width)
            elements.append(Spacer(1, banner_height - 3*mm)); has_banner = True
        except: pass
    styles = getSampleStyleSheet()
    
    # Styles
    style_title = ParagraphStyle('Title', fontSize=14, leading=16, fontName='Helvetica-Bold', spaceAfter=2, alignment=TA_LEFT)
    style_comp = ParagraphStyle('Comp', fontSize=11, leading=13, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)  # Nicht mehr Bold, mit Abstand!
    style_info = ParagraphStyle('Info', fontSize=10, leading=12, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)  # Nicht mehr BOLD, mit Abstand!
    style_sub = ParagraphStyle('Sub', fontSize=10, leading=12, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)  # Gleicher Abstand wie comp/info
    style_hdr = ParagraphStyle('Hdr', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER, textColor=colors.white)
    style_hdr_left = ParagraphStyle('HdrLeft', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_LEFT, textColor=colors.white)
    style_pos = ParagraphStyle('Pos', fontSize=8, leading=10, fontName='Helvetica', alignment=TA_CENTER)
    style_rider = ParagraphStyle('Rider', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_LEFT)
    style_normal = ParagraphStyle('Normal', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_CENTER)
    style_start = ParagraphStyle('StartNumber', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_CENTER)
    style_time = ParagraphStyle('Time', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_CENTER)
    style_horse = ParagraphStyle('Horse', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_LEFT)
    style_pause = ParagraphStyle('Pause', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER)
    style_group = ParagraphStyle('Group', fontSize=10, leading=12, fontName='Helvetica-Bold', alignment=TA_LEFT, textColor=colors.white)
    
    # Logo und Header nebeneinander
    logo_path = starterlist.get("logoPath")
    
    # Header-Text sammeln
    header_parts = []
    if not has_banner:
        show_title = starterlist.get("showTitle", "")
        if show_title:
            header_parts.append(Paragraph(f"<b>{show_title}</b>", style_title))
    
    # Competition-Objekt holen (wie in pdf_abstammung_logo)
    comp = starterlist.get("competition") or {}
    
    # Prüfungsnummer und Name
    comp_no = comp.get("number") or starterlist.get("competitionNumber") or ""
    comp_title = comp.get("title") or starterlist.get("competitionName") or starterlist.get("competitionTitle") or ""
    division = comp.get("divisionNumber") or starterlist.get("divisionNumber")
    
    # Division-Text erstellen
    div_text = ""
    try:
        if division is not None and str(division) != "" and int(division) > 0:
            div_text = f"{int(division)}. Abt. "
    except:
        div_text = f"{division} " if division else ""
    
    # 2. Zeile: Prüfungszeile FETT
    if comp_no or comp_title:
        # Mit Division: "Prüfung 17 - 1. Abt. Dressurpferdeprüfung Kl. L"
        # Ohne Division: "Prüfung 17 — Dressurpferdeprüfung Kl. L"
        comp_line = f"Prüfung {comp_no}"
        if div_text:
            comp_line += f" - {div_text}{comp_title}"
        elif comp_title:
            comp_line += f" — {comp_title}"
        header_parts.append(Paragraph(f"<b>{comp_line}</b>", style_comp))  # JETZT FETT!
    
    # 3. Zeile: informationText (nicht mehr fett)
    comp_info = starterlist.get("informationText", "")
    if comp_info:
        # Zeilenumbrüche in <br/> umwandeln für HTML-Darstellung
        comp_info_html = comp_info.replace("\n", "<br/>")
        header_parts.append(Paragraph(comp_info_html, style_info))  # style_info - nicht mehr BOLD!
    
    subtitle = starterlist.get("subtitle", "")
    if subtitle:
        header_parts.append(Paragraph(subtitle, style_sub))
    
    # Division-Start-Zeit (wie in pdf_abstammung_logo)
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
        start_raw = starterlist.get("start")
    
    # Location kann auch competitionLocation heißen
    location = starterlist.get("competitionLocation", "") or starterlist.get("location", "")
    if start_raw:
        date_line = format_datetime_german(start_raw)
        if location:
            date_line = f"{date_line} - {location}"
        header_parts.append(Paragraph(f"<b>{date_line}</b>", style_sub))  # FETT!
    
    # Richter (VOR Logo-Tabelle, damit sie in header_parts landen!)
    judges = starterlist.get("judges", [])
    if judges:
        pos_map = {0: "E", 1: "H", 2: "C", 3: "M", 4: "B"}
        judges_by_pos = {}  
        for j in judges:
            pos = j.get("position", "")
            name = j.get("name", "")
            if isinstance(pos, int):
                pos_label = pos_map.get(pos, "")
            else:
                pos_label = str(pos)
            if pos_label and name:
                if pos_label not in judges_by_pos:
                    judges_by_pos[pos_label] = []
                judges_by_pos[pos_label].append(name)
        
        jury_parts = []
        displayed_positions = [p for p in ["E", "H", "C", "M", "B"] if p in judges_by_pos]
        only_c = len(displayed_positions) == 1 and displayed_positions[0] == "C"
        print(f"DEBUG: displayed_positions={displayed_positions}, only_c={only_c}, judges_by_pos={judges_by_pos}")
        
        for p in ["E", "H", "C", "M", "B"]:
            if p in judges_by_pos:
                names = ' & '.join(judges_by_pos[p])
                if only_c:
                    print(f"DEBUG: Adding judge WITHOUT position box: {names}")
                    jury_parts.append(names)
                else:
                    print(f"DEBUG: Adding judge WITH position box: {p} {names}")
                    # Use Unicode non-breaking space (\u00A0) to keep position letter with name
                    jury_parts.append(f'<font name="Helvetica-Bold" backColor="black" color="white">\u00A0{p}\u00A0</font>\u00A0{names}')
        
        if jury_parts:
            # Join with regular space between judges
            judges_text = ' '.join(jury_parts)
            header_parts.append(Paragraph(f"<b>Richter:</b> {judges_text}", style_sub))
    
    # Logo + Header in einer Zeile
    if logo_path and os.path.exists(logo_path):
        try:
            # Logo in Originalgröße laden
            logo = Image(logo_path)
            
            # Maximale Größe: von Parameter (Standard 5cm = 50mm)
            max_size = logo_max_width_cm * 10 * mm  # cm in mm umrechnen
            
            # Wenn das Logo größer als max_size ist, verkleinern (aber nicht vergrößern!)
            if logo.drawWidth > max_size or logo.drawHeight > max_size:
                # Proportional skalieren
                scale = min(max_size / logo.drawWidth, max_size / logo.drawHeight)
                logo.drawWidth = logo.drawWidth * scale
                logo.drawHeight = logo.drawHeight * scale
            
            # Tabelle: Text links, Logo rechts
            # Breite dynamisch anpassen an Logobreite + etwas Puffer
            logo_col_width = max(logo.drawWidth + 5*mm, 35*mm)  # Mindestens 35mm
            text_col_width = 185*mm - logo_col_width
            header_table_data = [[header_parts, logo]]
            header_table = Table(header_table_data, colWidths=[text_col_width, logo_col_width])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (0,0), (0,0), 'LEFT'),
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
                ('LEFTPADDING', (0,0), (0,0), 0),  # Kein Padding links für Ausrichtung mit Tabelle
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
    
    # Spacer und "Starterliste" Überschrift
    elements.append(Spacer(1, 1*mm))
    elements.append(Paragraph("<b>Starterliste</b>", style_comp))
    elements.append(Spacer(1, 2*mm))
    
    # Tabelle - wie pdf_abstammung_logo
    starters = starterlist.get("starters", [])
    breaks = starterlist.get("breaks", [])
    breaks_map = {}
    for b in breaks:
        try:
            k = int(b.get("afterNumberInCompetition", -1))
            breaks_map.setdefault(k, []).append(b)
        except:
            pass
    
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
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4d4d4d")),  # Dunkelgrau statt lightgrey
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

    # Judges are already displayed in the header, no need for table at bottom


    doc.build(elements, canvasmaker=BannerCanvas)
    print(f"PDF NAT: {filename}")
