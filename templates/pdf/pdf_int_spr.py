# -*- coding: utf-8 -*-
# templates/pdf/pdf_int.py - International Design with flags (English)
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

# No translation needed - English is default

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

def get_country_name_english(ioc_code):
    """Returns full country name in English - Complete list for all 250 countries"""
    if not ioc_code:
        return ""
    
    names = {
        # A
        "AFG": "Afghanistan", "AIA": "Anguilla", "ALB": "Albania", "ALG": "Algeria",
        "AND": "Andorra", "ANG": "Angola", "ANT": "Antigua and Barbuda", "ARG": "Argentina",
        "ARM": "Armenia", "ARU": "Aruba", "ASA": "American Samoa", "AUS": "Australia",
        "AUT": "Austria", "AZE": "Azerbaijan",
        
        # B
        "BAH": "Bahamas", "BAN": "Bangladesh", "BAR": "Barbados", "BDI": "Burundi",
        "BEL": "Belgium", "BEN": "Benin", "BER": "Bermuda", "BHU": "Bhutan",
        "BIH": "Bosnia and Herzegovina", "BIZ": "Belize", "BLR": "Belarus", "BOL": "Bolivia",
        "BOT": "Botswana", "BRA": "Brazil", "BRN": "Bahrain", "BRU": "Brunei",
        "BUL": "Bulgaria", "BUR": "Burkina Faso",
        
        # C
        "CAF": "Central African Republic", "CAM": "Cambodia", "CAN": "Canada",
        "CAY": "Cayman Islands", "CGO": "Congo", "CHA": "Chad", "CHI": "Chile",
        "CHN": "China", "CIV": "Ivory Coast", "CMR": "Cameroon", "COD": "DR Congo",
        "COK": "Cook Islands", "COL": "Colombia", "COM": "Comoros", "CPV": "Cape Verde",
        "CRC": "Costa Rica", "CRO": "Croatia", "CUB": "Cuba", "CYP": "Cyprus",
        "CZE": "Czech Republic",
        
        # D-E
        "DEN": "Denmark", "DJI": "Djibouti", "DMA": "Dominica", "DOM": "Dominican Republic",
        "ECU": "Ecuador", "EGY": "Egypt", "ERI": "Eritrea", "ESA": "El Salvador",
        "ESP": "Spain", "EST": "Estonia", "ETH": "Ethiopia",
        
        # F
        "FAR": "Faroe Islands", "FIJ": "Fiji", "FIN": "Finland", "FRA": "France",
        "FSM": "Micronesia",
        
        # G
        "GAB": "Gabon", "GAM": "Gambia", "GBR": "Great Britain", "GBS": "Guinea-Bissau",
        "GEO": "Georgia", "GEQ": "Equatorial Guinea", "GER": "Germany", "GHA": "Ghana",
        "GRE": "Greece", "GRN": "Grenada", "GUA": "Guatemala", "GUI": "Guinea",
        "GUM": "Guam", "GUY": "Guyana",
        
        # H-I
        "HAI": "Haiti", "HKG": "Hong Kong", "HON": "Honduras", "HUN": "Hungary",
        "INA": "Indonesia", "IND": "India", "IRI": "Iran", "IRL": "Ireland",
        "IRQ": "Iraq", "ISL": "Iceland", "ISR": "Israel", "ISV": "US Virgin Islands",
        "ITA": "Italy", "IVB": "British Virgin Islands",
        
        # J-K
        "JAM": "Jamaica", "JOR": "Jordan", "JPN": "Japan", "KAZ": "Kazakhstan",
        "KEN": "Kenya", "KGZ": "Kyrgyzstan", "KIR": "Kiribati", "KOR": "South Korea",
        "KOS": "Kosovo", "KSA": "Saudi Arabia", "KUW": "Kuwait",
        
        # L
        "LAO": "Laos", "LAT": "Latvia", "LBA": "Libya", "LBN": "Lebanon",
        "LBR": "Liberia", "LCA": "Saint Lucia", "LES": "Lesotho", "LIE": "Liechtenstein",
        "LTU": "Lithuania", "LUX": "Luxembourg",
        
        # M
        "MAC": "Macau", "MAD": "Madagascar", "MAR": "Morocco", "MAS": "Malaysia",
        "MAW": "Malawi", "MDA": "Moldova", "MDV": "Maldives", "MEX": "Mexico",
        "MGL": "Mongolia", "MHL": "Marshall Islands", "MKD": "North Macedonia", "MLI": "Mali",
        "MLT": "Malta", "MNE": "Montenegro", "MON": "Monaco", "MOZ": "Mozambique",
        "MRI": "Mauritius", "MTN": "Mauritania", "MYA": "Myanmar",
        
        # N
        "NAM": "Namibia", "NCA": "Nicaragua", "NED": "Netherlands", "NEP": "Nepal",
        "NGR": "Nigeria", "NIG": "Niger", "NOR": "Norway", "NRU": "Nauru",
        "NZL": "New Zealand",
        
        # O-P
        "OMA": "Oman", "PAK": "Pakistan", "PAN": "Panama", "PAR": "Paraguay",
        "PER": "Peru", "PHI": "Philippines", "PLE": "Palestine", "PLW": "Palau",
        "PNG": "Papua New Guinea", "POL": "Poland", "POR": "Portugal", "PRK": "North Korea",
        "PUR": "Puerto Rico",
        
        # Q-R
        "QAT": "Qatar", "ROU": "Romania", "RSA": "South Africa", "RUS": "Russia",
        "RWA": "Rwanda",
        
        # S
        "SAM": "Samoa", "SEN": "Senegal", "SEY": "Seychelles", "SGP": "Singapore",
        "SKN": "Saint Kitts and Nevis", "SLE": "Sierra Leone", "SLO": "Slovenia",
        "SMR": "San Marino", "SOL": "Solomon Islands", "SOM": "Somalia", "SRB": "Serbia",
        "SRI": "Sri Lanka", "SSD": "South Sudan", "STP": "São Tomé and Príncipe",
        "SUD": "Sudan", "SUI": "Switzerland", "SUR": "Suriname", "SVK": "Slovakia",
        "SWE": "Sweden", "SWZ": "Eswatini", "SYR": "Syria",
        
        # T
        "TAN": "Tanzania", "TCA": "Turks and Caicos Islands", "TGA": "Tonga",
        "THA": "Thailand", "TJK": "Tajikistan", "TKM": "Turkmenistan",
        "TLS": "Timor-Leste", "TOG": "Togo", "TPE": "Taiwan", "TTO": "Trinidad and Tobago",
        "TUN": "Tunisia", "TUR": "Turkey", "TUV": "Tuvalu",
        
        # U-V
        "UAE": "United Arab Emirates", "UGA": "Uganda", "UKR": "Ukraine",
        "URU": "Uruguay", "USA": "USA", "UZB": "Uzbekistan", "VAN": "Vanuatu",
        "VEN": "Venezuela", "VIE": "Vietnam", "VIN": "Saint Vincent and the Grenadines",
        
        # Y-Z
        "YEM": "Yemen", "ZAM": "Zambia", "ZIM": "Zimbabwe",
        
        # Backwards compatibility - 2-letter codes
        "GB": "Great Britain", "DE": "Germany", "NL": "Netherlands", "CH": "Switzerland",
        "DK": "Denmark", "AT": "Austria", "BE": "Belgium", "FR": "France",
        "IT": "Italy", "ES": "Spain", "SE": "Sweden", "NO": "Norway",
        "PL": "Poland", "CZ": "Czech Republic", "HU": "Hungary", "RO": "Romania",
        "IE": "Ireland", "PT": "Portugal",
        
        # Alternative spellings
        "DZA": "Algeria",  # ISO code for Algeria
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


def format_time(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%H:%M:%S")
    except:
        return str(iso).split("T")[-1][:8] if "T" in str(iso) else str(iso)

def format_datetime_english(iso):
    """English: Wednesday, March 19, 2025 14:00"""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%A, %B %d, %Y %H:%M")
    except:
        return str(iso)

def format_pause_text(total_seconds, info):
    secs = int(total_seconds or 0)
    if secs == 0:
        return info or "Pause"
    if secs >= 3600:
        h = secs // 3600
        m = (secs % 3600) // 60
        return f"Break ({h} h {m:02d} min){' - ' + info if info else ''}"
    elif secs >= 60:
        return f"Break ({secs//60} min){' - ' + info if info else ''}"
    else:
        return f"Break ({secs} Sek.){' - ' + info if info else ''}"

def get_sponsor_bar_height():
    sponsor_path = "logos/sponsorenleiste.png"
    if not os.path.exists(sponsor_path):
        return 25*mm
    try:
        pil_img = PILImage.open(sponsor_path)
        img_width, img_height = pil_img.size
        target_width_mm = 190
        aspect_ratio = img_height / img_width
        return target_width_mm * aspect_ratio * mm
    except:
        return 25*mm

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
        self.sponsor_height = get_sponsor_bar_height()
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
                self.drawImage(sponsor_path, (A4[0] - 190*mm) / 2, 4*mm, width=190*mm, height=self.sponsor_height, preserveAspectRatio=True, mask='auto')
            except:
                pass


def render(starterlist, filename, logo_max_width_cm=5.0):
    sponsor_height = get_sponsor_bar_height()
    footer_space = 4*mm
    margin_above_sponsor = 1*mm
    bottom_margin = footer_space + sponsor_height + margin_above_sponsor
    doc = SimpleDocTemplate(
        filename, pagesize=A4, 
        rightMargin=10*mm, leftMargin=10*mm, 
        topMargin=5*mm, bottomMargin=bottom_margin  # Kompakt oben und unten
    )
    
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
            elements.append(Spacer(1, banner_height - 3*mm))
            has_banner = True
        except:
            pass
    styles = getSampleStyleSheet()
    
    # Styles
    style_title = ParagraphStyle('Title', fontSize=14, leading=16, fontName='Helvetica-Bold', spaceAfter=2, alignment=TA_LEFT)
    style_comp = ParagraphStyle('Comp', fontSize=11, leading=13, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)  # Nicht mehr Bold!
    style_info = ParagraphStyle('Info', fontSize=10, leading=12, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)  # Kleiner für informationText
    style_sub = ParagraphStyle('Sub', fontSize=10, leading=12, fontName='Helvetica', spaceAfter=4, alignment=TA_LEFT)
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
    if not has_banner:
        show_title = starterlist.get("showTitle", "")
        if show_title:
            header_parts.append(Paragraph(f"<b>{show_title}</b>", style_title))
    
    # Competition-Objekt holen (wie in pdf_abstammung_logo)
    comp = starterlist.get("competition") or {}
    
    # Competition number and name
    comp_no = comp.get("number") or starterlist.get("competitionNumber") or ""
    comp_title = comp.get("title") or starterlist.get("competitionName") or starterlist.get("competitionTitle") or ""
    division = comp.get("divisionNumber") or starterlist.get("divisionNumber")
    
    # Division text
    div_text = ""
    try:
        if division is not None and str(division) != "" and int(division) > 0:
            div_text = f"{int(division)}. Div. "
    except:
        div_text = f"{division} " if division else ""
    
    # 2. Zeile: Competition line FETT
    if comp_no or comp_title:
        comp_line = f"Competition {comp_no}"
        if div_text:
            comp_line += f" - {div_text}{comp_title}"
        elif comp_title:
            comp_line += f" — {comp_title}"
        header_parts.append(Paragraph(f"<b>{comp_line}</b>", style_comp))
    
    # 3. Zeile: informationText (nicht mehr fett)
    comp_info = starterlist.get("informationText", "")
    if comp_info:
        # Zeilenumbrüche in <br/> umwandeln für HTML-Darstellung
        comp_info_html = comp_info.replace("\n", "<br/>")
        header_parts.append(Paragraph(comp_info_html, style_info))  # style_comp statt style_sub!
    

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
    
    location = starterlist.get("competitionLocation", "") or starterlist.get("location", "")
    if start_raw:
        date_line = format_datetime_english(start_raw)
        if location:
            date_line = f"{date_line} - {location}"
        header_parts.append(Paragraph(f"<b>{date_line}</b>", style_sub))
    
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
    
    # Kein extra Spacer nötig - spaceAfter=4 in style_sub reicht
    
    # Jury
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
        for p in ["E", "H", "C", "M", "B"]:
            if p in judges_by_pos:
                # Mehrere Richter mit & verbinden
                names = ' & '.join(judges_by_pos[p])
                jury_parts.append(f'<font name="Helvetica-Bold" backColor="black" color="white"> {p} </font> {names}')
        
        if jury_parts:
            elements.append(Paragraph(f"<b>Jury:</b> {' '.join(jury_parts)}", style_sub))
    
    elements.append(Spacer(1, 1*mm))
    elements.append(Paragraph("<b>Starting Order</b>", style_comp))
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
    
    # Header
    data_texts.append(["#", "CNO", "Horse", "Athlete", "Nat."])
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
            group_text = f"Division {starter_group}"
            data_texts.append([group_text, "", "", "", ""])
            meta.append({"type": "group"})
            
            current_group = starter_group
            group_start_time_shown = False  # Reset für neue Gruppe
        
        nr = str(s.get("startNumber", ""))
        hors_concours = bool(s.get("horsConcours", False))  # Außer Konkurrenz  # Convert to string immediately!
        
        # Time column removed - no longer needed
        time_str = ""  # Time column removed
        
        horses = s.get("horses", [])
        horse = horses[0] if horses else {}
        cno = str(horse.get("cno", ""))
        
        # Pferd
        horse_html = ""
        if horse:
            horse_name = horse.get("name", "")
            studbook = horse.get("studbook", "")
            horse_html = f"<b>{horse_name}</b>"
            if studbook:
                horse_html += f" - <font size=7>{studbook}</font>"
            
            details = []
            breeding_season = horse.get("breedingSeason")
            if breeding_season:
                try:
                    current_year = datetime.now().year
                    age = current_year - int(breeding_season)
                    details.append(f"{age}Y")
                except:
                    pass
            
            color = horse.get("color", "")
            if color:
                details.append(color)
            
            sex = horse.get("sex", "")
            if sex:
                details.append(sex.upper() if sex else "")
            
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
                horse_html += f"<br/><font size=6.5><i>Owner: {owner}</i></font>"
            if breeder:
                horse_html += f"<br/><font size=6.5><i>Breeder: {breeder}</i></font>"
        
        # Reiter mit Club/Land
        athlete = s.get("athlete", {})
        athlete_name = str(athlete.get("name", ""))  # Ensure string
        club = athlete.get("club", "")
        nationality = athlete.get("nation", "")
        
        # Reiter-HTML: Name + Club/Land darunter
        athlete_html = f"<b>{athlete_name}</b>" if athlete_name else ""
        
        # New logic: Show country for foreigners only if club is empty OR "Gastlizenz GER"
        if nationality and nationality.upper() != "GER":
            # Foreigner
            if not club or club.strip() == "" or club.strip().upper() == "GASTLIZENZ GER":
                # No club or guest license → Show country name
                country_full = get_country_name_english(nationality)
                if country_full:
                    athlete_html += f"<br/><font size=7>{country_full}</font>"
            else:
                # Has a club (not guest license) → Show club
                athlete_html += f"<br/><font size=7>{club}</font>"
        elif club:
            # German: Show club
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
        
        data_texts.append([nr_display, cno, horse_html, athlete_html, nat_cell])
        withdrawn_flag = bool(s.get("withdrawn", False))
        meta.append({"type": "starter", "withdrawn": withdrawn_flag, "horsConcours": hors_concours})
        
        # Break
        try:
            cur = int(nr)
            if cur in breaks_map:
                for br in breaks_map[cur]:
                    pause_text = format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                    data_texts.append([pause_text, "", "", "", ""])
                    meta.append({"type": "pause"})
        except:
            pass
    
    # Spaltenbreiten
    page_width = A4[0] - doc.leftMargin - doc.rightMargin
    col_widths = [10*mm, 12*mm, 100*mm, 60*mm, 13*mm]  # Breiter wie deutsche Version  # Time removed, +8mm to Horse and Athlete
    
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
            # Abteilungs-Header (fett, grau)
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
            nat_value = row[4]
            if isinstance(nat_value, Table):
                nat_cell_final = nat_value
            elif isinstance(nat_value, Image):
                nat_cell_final = nat_value
            else:
                nat_cell_final = Paragraph(str(nat_value), style_pos)
            
            table_rows.append([
                maybe_strike(row[0], style_pos),
                maybe_strike(row[1], style_pos),
                maybe_strike(row[2], style_horse),
                maybe_strike(row[3], style_rider),
                nat_cell_final
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
        # Kompaktere Zeilen: Reduzierte Paddings
        ("TOPPADDING", (0,0), (-1,-1), 2),      # Standard war 6
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),   # Standard war 6,
    ])
    
    # SPAN + Zebra - mit Gruppen-Logik!
    starter_row_count = 0
    for ri in range(1, len(table_rows)):
        if ri < len(meta):
            m = meta[ri]
            if m.get("type") == "group":
                # Abteilungs-Header: SPAN über alle Spalten, grauer Hintergrund
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
    print(f"PDF INT: {filename}")
