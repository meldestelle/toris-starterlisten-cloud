# -*- coding: utf-8 -*-
# templates/pdf/liste_int.py - International Design with flags (English)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageTemplate, Frame, NextPageTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
from datetime import datetime
import os

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

def _fmt_header_datetime(iso):
    """Format datetime for header: Saturday, January 3, 2026 at 22:00"""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        weekday_map = {
            "Monday": "Monday", "Tuesday": "Tuesday", "Wednesday": "Wednesday",
            "Thursday": "Thursday", "Friday": "Friday", "Saturday": "Saturday", "Sunday": "Sunday"
        }
        month_map = {
            "January": "January", "February": "February", "March": "March", "April": "April",
            "May": "May", "June": "June", "July": "July", "August": "August",
            "September": "September", "October": "October", "November": "November", "December": "December"
        }
        weekday = weekday_map.get(dt.strftime("%A"), dt.strftime("%A"))
        month = month_map.get(dt.strftime("%B"), dt.strftime("%B"))
        return f"{weekday}, {month} {dt.day}, {dt.year} at {dt.strftime('%H:%M')}"
    except Exception:
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
    
