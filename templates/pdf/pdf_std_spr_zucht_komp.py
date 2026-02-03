# -*- coding: utf-8 -*-
# templates/pdf/pdf_std_spr_zucht_komp.py - Deutsches Design mit Flaggen
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
    """Berechnet die Höhe der Sponsorenleiste basierend auf dem Seitenverhältnis"""
    sponsor_path = "logos/sponsorenleiste.png"
    if not os.path.exists(sponsor_path):
        return 25*mm  # Fallback
    
    try:
        pil_img = PILImage.open(sponsor_path)
        img_width, img_height = pil_img.size
        
        # Zielbreite: 190mm
        target_width_mm = 190
        aspect_ratio = img_height / img_width
        calculated_height_mm = target_width_mm * aspect_ratio
        
        return calculated_height_mm * mm
    except:
        return 25*mm  # Fallback bei Fehler

class BannerCanvas(canvas.Canvas):
    """Canvas mit Banner oben (nur Seite 1, wenn vorhanden) und Sponsorenleiste im Footer"""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        
        # Page counter
        self.page_num = 0
        
        # Banner-Pfad und Höhe berechnen
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
            except Exception as e:
                print(f"DEBUG: Banner-Fehler: {e}")
                self.banner_path = None
        else:
            self.banner_path = None
        
        # Sponsorenleiste Höhe dynamisch berechnen
        self.sponsor_height = get_sponsor_bar_height()
        
    def showPage(self):
        # Banner VOR Increment zeichnen, dann zählen
        self.draw_banner()
        self.draw_footer()
        self.page_num += 1
        canvas.Canvas.showPage(self)
    
    def draw_banner(self):
        """Zeichnet Banner oben - NUR AUF SEITE 1"""
        # page_num ist 0 auf erster Seite (vor Increment)
        if self.page_num != 0:
            return
            
        if self.banner_path and os.path.exists(self.banner_path):
            try:
                x = 0
                y = A4[1] - self.banner_height
                self.drawImage(
                    self.banner_path,
                    x, y,
                    width=self.banner_width,
                    height=self.banner_height,
                    preserveAspectRatio=True,
                    mask='auto'
                )
            except Exception as e:
                print(f"DEBUG: Banner draw error: {e}")
        
    def draw_footer(self):
        """Zeichnet Sponsorenleiste unten mit dynamischer Höhe"""
        sponsor_path = "logos/sponsorenleiste.png"
        if os.path.exists(sponsor_path):
            try:
                img_width = 190*mm
                img_height = self.sponsor_height  # Dynamisch berechnet
                x = (A4[0] - img_width) / 2
                y = 4*mm  # 4mm vom Seitenende
                self.drawImage(sponsor_path, x, y, width=img_width, height=img_height, 
                              preserveAspectRatio=True, mask='auto')
            except:
                pass

def render(starterlist, filename, logo_max_width_cm=5.0):
    # Berechne dynamischen bottomMargin basierend auf Sponsorenleiste
    sponsor_height = get_sponsor_bar_height()
    footer_space = 4*mm  # Abstand der Sponsorenleiste vom unteren Rand
    margin_above_sponsor = 1*mm  # 1mm Abstand zwischen Tabelle und Sponsorenleiste
    bottom_margin = footer_space + sponsor_height + margin_above_sponsor
    
    doc = SimpleDocTemplate(
        filename, pagesize=A4, 
        rightMargin=10*mm, leftMargin=10*mm,  # Kleiner für mehr Platz!
        topMargin=5*mm, bottomMargin=bottom_margin  # Dynamisch berechnet
    )
    
    elements = []
    
    # Banner-Höhe berechnen und Spacer einfügen - ODER Show Title wenn kein Banner
    banner_path = os.path.join("logos", "banner.png")
    banner_height = 0
    has_banner = False
    
    if os.path.exists(banner_path):
        try:
            pil_img = PILImage.open(banner_path)
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
            banner_height = page_width * aspect_ratio
            # Spacer für Banner-Höhe auf erster Seite
            elements.append(Spacer(1, banner_height - 3*mm))  # -3mm da schon 5mm topMargin
            has_banner = True
        except:
            pass
    
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

    style_horse = ParagraphStyle('Horse', fontSize=9, leading=10, fontName='Helvetica', alignment=TA_LEFT)
    style_pause = ParagraphStyle('Pause', fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER)
    style_group = ParagraphStyle('Group', fontSize=10, leading=12, fontName='Helvetica-Bold', alignment=TA_LEFT, textColor=colors.white)
    
    # Logo und Header nebeneinander
    logo_path = starterlist.get("logoPath")
    
    # Header-Text sammeln
    header_parts = []
    
    # Show title NUR wenn KEIN Banner vorhanden
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
        judges_by_pos = {}  # Dict mit Listen für mehrere Richter pro Position
        for j in judges:
            pos = j.get("position", "")
            name = j.get("name", "")
            if isinstance(pos, int):
                pos_label = pos_map.get(pos, "")
            else:
                pos_label = str(pos)
            if pos_label and name:
                # Liste erstellen wenn noch nicht vorhanden
                if pos_label not in judges_by_pos:
                    judges_by_pos[pos_label] = []
                judges_by_pos[pos_label].append(name)
        
        jury_parts = []
        
        # Prüfen, ob nur "C" bei den angezeigten Positionen (E,H,C,M,B) verwendet wird
        displayed_positions = [p for p in ["E", "H", "C", "M", "B"] if p in judges_by_pos]
        only_c = len(displayed_positions) == 1 and displayed_positions[0] == "C"
        print(f"DEBUG: displayed_positions={displayed_positions}, only_c={only_c}, judges_by_pos={judges_by_pos}")
        
        for p in ["E", "H", "C", "M", "B"]:
            if p in judges_by_pos:
                # Mehrere Richter mit & verbinden
                names = ' & '.join(judges_by_pos[p])
                
                # Wenn nur "C" verwendet wird, zeige nur die Namen ohne Position
                if only_c:
                    print(f"DEBUG: Adding judge WITHOUT position box: {names}")
                    jury_parts.append(names)
                else:
                    print(f"DEBUG: Adding judge WITH position box: {p} {names}")
                    jury_parts.append(f'<font name="Helvetica-Bold" backColor="black" color="white"> {p} </font> {names}')
        
        if jury_parts:
            header_parts.append(Paragraph(f"<b>Richter:</b> {' '.join(jury_parts)}", style_sub))
    
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
    
    data_texts = []
    meta = []
    
    # Header (ohne Zeit-Spalte)
    data_texts.append(["#", "KNR", "Pferd", "Reiter", "Nat."])
    meta.append({"type": "header"})
    
    # Prüfe ob es eine Pause VOR dem ersten Starter gibt (afterNumberInCompetition=0)
    if 0 in breaks_map:
        for br in breaks_map[0]:
            pause_text = format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append([pause_text, "", "", "", ""])  # 5 Spalten
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
            data_texts.append([group_text, "", "", "", ""])  # 5 Spalten
            meta.append({"type": "group"})
            
            current_group = starter_group
            group_start_time_shown = False  # Reset für neue Gruppe
        
        nr = str(s.get("startNumber", ""))  # Convert to string immediately!
        hors_concours = bool(s.get("horsConcours", False))  # Außer Konkurrenz prüfen
        
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
            breed = horse.get("breed", "") or horse.get("studbook", "")  # breed, oder studbook als Fallback
            
            # Pferdename FETT, dann breed in kleinerer, normaler Schrift
            if breed:
                horse_html = f"<b>{horse_name}</b> - <font size=7>{breed}</font>"
            else:
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
            # Nur hinzufügen wenn nicht schon als breed angezeigt
            if studbook and studbook != breed:
                details.append(studbook)
            
            sire = horse.get("sire", "")
            dam_sire = horse.get("damSire", "")
            if sire and dam_sire:
                details.append(f"{sire} x {dam_sire}")
            elif sire:
                details.append(sire)
            
            if details:
                horse_html += f"<br/><font size=7>{' / '.join(details)}</font>"
            
            # Besitzer und Züchter kompakt in einer Zeile
            owner = horse.get("owner", "")
            breeder = horse.get("breeder", "")
            
            if owner and breeder:
                # Beide vorhanden
                if owner.strip() == breeder.strip():
                    # Identisch: "B u. Z: Name"
                    horse_html += f"<br/><font size=6.5><i>B u. Z: {owner}</i></font>"
                else:
                    # Unterschiedlich: "B: Name / Z: Name"
                    horse_html += f"<br/><font size=6.5><i>B: {owner} / Z: {breeder}</i></font>"
            elif owner:
                # Nur Besitzer
                horse_html += f"<br/><font size=6.5><i>B: {owner}</i></font>"
            elif breeder:
                # Nur Züchter
                horse_html += f"<br/><font size=6.5><i>Z: {breeder}</i></font>"
        
        # Reiter mit Club/Land
        athlete = s.get("athlete", {})
        athlete_name = str(athlete.get("name", ""))  # Ensure string
        club = athlete.get("club", "")
        nationality = athlete.get("nation", "")
        
        # Reiter-HTML: Name + Club/Land darunter
        athlete_html = f"<b>{athlete_name}</b>" if athlete_name else ""
        
        # Neue Logik: Ausländer zeigen Land nur wenn Verein leer ODER "Gastlizenz GER"
        if nationality and nationality.upper() != "GER":
            # Ausländer
            if not club or club.strip() == "" or club.strip().upper() == "GASTLIZENZ GER":
                # Kein Verein oder Gastlizenz → Land ausgeschrieben anzeigen
                country_full = get_country_name(nationality)
                if country_full:
                    athlete_html += f"<br/><font size=7>{country_full}</font>"
            else:
                # Hat einen Verein (nicht Gastlizenz) → Verein anzeigen
                athlete_html += f"<br/><font size=7>{club}</font>"
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
        
        data_texts.append([nr_display, cno, horse_html, athlete_html, nat_cell])  # Ohne Zeit
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type": "starter", "withdrawn": withdrawn_flag, "horsConcours": hors_concours})
        
        # Break
        try:
            cur = int(nr)
            if cur in breaks_map:
                for br in breaks_map[cur]:
                    pause_text = format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                    data_texts.append([pause_text, "", "", "", ""])  # 5 Spalten
                    meta.append({"type": "pause"})
        except:
            pass
    
    # Spaltenbreiten
    page_width = A4[0] - doc.leftMargin - doc.rightMargin
    col_widths = [10*mm, 12*mm, 100*mm, 60*mm, 13*mm]  # Ohne Zeit: Pferd breiter (100mm), Reiter breiter (60mm)
    
    table_rows = []
    for i, row in enumerate(data_texts):
        if i >= len(meta):
            continue
        
        m = meta[i]
        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr), Paragraph(row[1], style_hdr),
                Paragraph(row[2], style_hdr_left), Paragraph(row[3], style_hdr_left),
                Paragraph(row[4], style_hdr)
            ])
        elif m["type"] == "group":
            # Abteilungs-Header (linksbündig, dunkler Hintergrund, weiße Schrift)
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_group), Paragraph("", style_sub),
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
                if not text:
                    return Paragraph("", s)
                return Paragraph(f"<strike>{text}</strike>" if withdrawn else text, s)
            
            # Nat-Spalte: Wenn Table-Objekt (Flagge+Code) direkt nutzen, sonst Paragraph
            nat_value = row[4]  # Jetzt Index 4 statt 5
            if isinstance(nat_value, Table):
                nat_cell_final = nat_value
            elif isinstance(nat_value, Image):
                nat_cell_final = nat_value
            else:
                nat_cell_final = Paragraph(str(nat_value), style_pos)
            
            table_rows.append([
                maybe_strike(row[0], style_pos),  # #
                maybe_strike(row[1], style_pos),  # KNR
                maybe_strike(row[2], style_horse),  # Pferd
                maybe_strike(row[3], style_rider),  # Reiter
                nat_cell_final  # Nat.
            ])
    
    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("LINEBELOW", (0,0), (-1,-1), 0.5, colors.black),  # Nur horizontale Linien!
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#404040')),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (1,-1), "CENTER"),
        ("ALIGN", (2,0), (3,-1), "LEFT"),
        ("ALIGN", (4,0), (4,-1), "CENTER"),
    ])
    
    # SPAN + Zebra - mit Gruppen-Logik!
    starter_row_count = 0
    for ri in range(1, len(table_rows)):
        if ri < len(meta):
            m = meta[ri]
            if m.get("type") == "group":
                # Abteilungs-Header: SPAN über alle Spalten, dunkler Hintergrund wie Header
                ts.add("SPAN", (0,ri), (4,ri))
                ts.add("BACKGROUND", (0,ri), (4,ri), colors.HexColor('#404040'))
                # Counter zurücksetzen für neue Abteilung
                starter_row_count = 0
            elif m.get("type") == "starter":
                # Zebra: ungerade Starter sind grau (1, 3, 5, ...)
                if starter_row_count % 2 == 1:
                    ts.add("BACKGROUND", (0,ri), (4,ri), colors.HexColor('#E8E8E8'))
                starter_row_count += 1
            elif m.get("type") == "pause":
                ts.add("SPAN", (0,ri), (4,ri))
                prev_was_gray = (starter_row_count - 1) % 2 == 1
                if not prev_was_gray:  # Vorherige war weiß → Pause grau
                    ts.add("BACKGROUND", (0,ri), (4,ri), colors.HexColor('#E8E8E8'))
                # WICHTIG: Counter um 1 zurücksetzen, damit nächste Zeile gleiche Farbe hat!
                starter_row_count -= 1
    
    t.setStyle(ts)
    elements.append(t)
    
    # Build mit Footer-Canvas
    doc.build(elements, canvasmaker=BannerCanvas)
    print(f"PDF NAT: {filename}")
