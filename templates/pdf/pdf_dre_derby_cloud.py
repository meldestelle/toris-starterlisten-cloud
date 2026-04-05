# -*- coding: utf-8 -*-
# templates/pdf/pdf_dre_derby.py - Derby-Format: 2 Halbfinale + Finale
#
# Structure:
#   Starter 1 vs 2  → 1. Halbfinale
#   Starter 3 vs 4  → 2. Halbfinale
#   Starter 5 vs 6  → Finale (Reiternamen werden nicht angezeigt, da noch nicht feststehend)
#
# Pop-up-Eingabe:  Prüfungsbeginn (vor Starter 1) + Finale-Uhrzeit (vor Starter 5)
# Richterzeile:    Inline-Format aus pdf_nat
# Starterliste:    Spaltenformat aus pdf_dre_5

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    Image, PageTemplate, Frame, BaseDocTemplate
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
from reportlab.pdfgen import canvas
from datetime import datetime
import os
from PIL import Image as PILImage

# ---------------------------------------------------------------------------
# Übersetzungs-Maps
# ---------------------------------------------------------------------------
WEEKDAY_MAP = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag",
    "Sunday": "Sonntag"
}
MONTH_MAP = {
    "January": "Januar", "February": "Februar", "March": "März",
    "April": "April", "May": "Mai", "June": "Juni", "July": "Juli",
    "August": "August", "September": "September", "October": "Oktober",
    "November": "November", "December": "Dezember"
}
SEX_MAP = {
    "MARE": "Stute", "GELDING": "Wallach", "STALLION": "Hengst",
    "M": "Wallach", "F": "Stute", "": ""
}

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------
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
        month   = MONTH_MAP.get(dt.strftime("%B"), dt.strftime("%B"))
        return f"{weekday}, {dt.day}. {month} {dt.year}"
    except Exception:
        return str(iso)


def get_nationality_code(nationality_str):
    if not nationality_str:
        return ""
    ioc_to_iso = {
        "GER": "DEU", "NED": "NLD", "SUI": "CHE", "DEN": "DNK", "CRO": "HRV",
        "GRE": "GRC", "BUL": "BGR", "RSA": "ZAF", "POR": "PRT", "LAT": "LVA",
        "UAE": "ARE", "CHI": "CHL", "URU": "URY", "SLO": "SVN", "MAS": "MYS",
        "GBR": "GBR", "AUT": "AUT", "BEL": "BEL", "FRA": "FRA", "ITA": "ITA",
        "ESP": "ESP", "SWE": "SWE", "NOR": "NOR", "POL": "POL", "CZE": "CZE",
        "HUN": "HUN", "ROU": "ROU", "IRL": "IRL", "USA": "USA", "CAN": "CAN",
        "AUS": "AUS", "NZL": "NZL", "JPN": "JPN", "KOR": "KOR", "CHN": "CHN",
        "BRA": "BRA", "ARG": "ARG", "MEX": "MEX", "RUS": "RUS", "UKR": "UKR",
        "TUR": "TUR", "ISR": "ISR",
    }
    code = nationality_str.upper()
    return ioc_to_iso.get(code, code[:3] if len(code) >= 3 else code)


def get_country_name(ioc_code):
    """Returns full country name in English - Complete list for all 250 countries"""
    if not ioc_code:
        return ""
    names = {
        # A
        "AFG": "Afghanistan", "AIA": "Anguilla", "ALB": "Albania", "ALG": "Algeria",
        "AND": "Andorra", "ANG": "Angola", "ANT": "Antigua and Barbuda", "ARG": "Argentina",
        "ARM": "Armenia", "ARU": "Aruba", "ASA": "American Samoa", "AUS": "Australia",
        "AUT": "Österreich", "AZE": "Azerbaijan",
        # B
        "BAH": "Bahamas", "BAN": "Bangladesh", "BAR": "Barbados", "BDI": "Burundi",
        "BEL": "Belgien", "BEN": "Benin", "BER": "Bermuda", "BHU": "Bhutan",
        "BIH": "Bosnia and Herzegovina", "BIZ": "Belize", "BLR": "Belarus", "BOL": "Bolivia",
        "BOT": "Botswana", "BRA": "Brasilien", "BRN": "Bahrain", "BRU": "Brunei",
        "BUL": "Bulgarien", "BUR": "Burkina Faso",
        # C
        "CAF": "Central African Republic", "CAM": "Cambodia", "CAN": "Kanada",
        "CAY": "Cayman Islands", "CGO": "Congo", "CHA": "Chad", "CHI": "Chile",
        "CHN": "China", "CIV": "Ivory Coast", "CMR": "Cameroon", "COD": "DR Congo",
        "COK": "Cook Islands", "COL": "Colombia", "COM": "Comoros", "CPV": "Cape Verde",
        "CRC": "Costa Rica", "CRO": "Kroatien", "CUB": "Cuba", "CYP": "Cyprus",
        "CZE": "Tschechien",
        # D-E
        "DEN": "Dänemark", "DJI": "Djibouti", "DMA": "Dominica", "DOM": "Dominican Republic",
        "ECU": "Ecuador", "EGY": "Egypt", "ERI": "Eritrea", "ESA": "El Salvador",
        "ESP": "Spanien", "EST": "Estland", "ETH": "Ethiopia",
        # F
        "FAR": "Faroe Islands", "FIJ": "Fiji", "FIN": "Finnland", "FRA": "Frankreich",
        "FSM": "Micronesia",
        # G
        "GAB": "Gabon", "GAM": "Gambia", "GBR": "Großbritannien", "GBS": "Guinea-Bissau",
        "GEO": "Georgia", "GEQ": "Equatorial Guinea", "GER": "Deutschland", "GHA": "Ghana",
        "GRE": "Griechenland", "GRN": "Grenada", "GUA": "Guatemala", "GUI": "Guinea",
        "GUM": "Guam", "GUY": "Guyana",
        # H-I
        "HAI": "Haiti", "HKG": "Hong Kong", "HON": "Honduras", "HUN": "Ungarn",
        "INA": "Indonesia", "IND": "India", "IRI": "Iran", "IRL": "Irland",
        "IRQ": "Iraq", "ISL": "Island", "ISR": "Israel", "ISV": "US Virgin Islands",
        "ITA": "Italien", "IVB": "British Virgin Islands",
        # J-K
        "JAM": "Jamaica", "JOR": "Jordan", "JPN": "Japan", "KAZ": "Kasachstan",
        "KEN": "Kenya", "KGZ": "Kyrgyzstan", "KIR": "Kiribati", "KOR": "Südkorea",
        "KOS": "Kosovo", "KSA": "Saudi Arabia", "KUW": "Kuwait",
        # L
        "LAO": "Laos", "LAT": "Lettland", "LBA": "Libya", "LBN": "Lebanon",
        "LBR": "Liberia", "LCA": "St. Lucia", "LES": "Lesotho", "LIE": "Liechtenstein",
        "LTU": "Litauen", "LUX": "Luxemburg",
        # M
        "MAC": "Macau", "MAD": "Madagascar", "MAR": "Morocco", "MAS": "Malaysia",
        "MAW": "Malawi", "MDA": "Moldova", "MDV": "Maldives", "MEX": "Mexiko",
        "MGL": "Mongolia", "MHL": "Marshall Islands", "MKD": "North Macedonia", "MLI": "Mali",
        "MLT": "Malta", "MNE": "Montenegro", "MON": "Monaco", "MOZ": "Mozambique",
        "MRI": "Mauritius", "MTN": "Mauritania", "MYA": "Myanmar",
        # N
        "NAM": "Namibia", "NCA": "Nicaragua", "NED": "Niederlande", "NEP": "Nepal",
        "NGR": "Nigeria", "NIG": "Niger", "NOR": "Norwegen", "NRU": "Nauru",
        "NZL": "Neuseeland",
        # O-P
        "OMA": "Oman", "PAK": "Pakistan", "PAN": "Panama", "PAR": "Paraguay",
        "PER": "Peru", "PHI": "Philippines", "PLE": "Palestine", "PLW": "Palau",
        "PNG": "Papua New Guinea", "POL": "Polen", "POR": "Portugal", "PRK": "North Korea",
        "PUR": "Puerto Rico",
        # Q-R
        "QAT": "Qatar", "ROU": "Rumänien", "RSA": "Südafrika", "RUS": "Russland",
        "RWA": "Rwanda",
        # S
        "SAM": "Samoa", "SEN": "Senegal", "SEY": "Seychelles", "SGP": "Singapore",
        "SKN": "St. Kitts and Nevis", "SLE": "Sierra Leone", "SLO": "Slowenien",
        "SMR": "San Marino", "SOL": "Solomon Islands", "SOM": "Somalia", "SRB": "Serbia",
        "SRI": "Sri Lanka", "SSD": "South Sudan", "STP": "São Tomé and Príncipe",
        "SUD": "Sudan", "SUI": "Schweiz", "SUR": "Suriname", "SVK": "Slovakia",
        "SWE": "Schweden", "SWZ": "Eswatini", "SYR": "Syria",
        # T
        "TAN": "Tanzania", "TCA": "Turks and Caicos Islands", "TGA": "Tonga",
        "THA": "Thailand", "TJK": "Tajikistan", "TKM": "Turkmenistan",
        "TLS": "Timor-Leste", "TOG": "Togo", "TPE": "Taiwan", "TTO": "Trinidad and Tobago",
        "TUN": "Tunisia", "TUR": "Türkei", "TUV": "Tuvalu",
        # U-V
        "UAE": "Vereinigte Arab. Emirate", "UGA": "Uganda", "UKR": "Ukraine",
        "URU": "Uruguay", "USA": "USA", "UZB": "Uzbekistan", "VAN": "Vanuatu",
        "VEN": "Venezuela", "VIE": "Vietnam", "VIN": "St. Vincent and the Grenadines",
        # Y-Z
        "YEM": "Yemen", "ZAM": "Zambia", "ZIM": "Zimbabwe",
        # Backwards compatibility
        "GB": "Großbritannien", "DE": "Deutschland", "NL": "Niederlande", "CH": "Schweiz",
        "DK": "Dänemark", "AT": "Österreich", "BE": "Belgien", "FR": "Frankreich",
        "IT": "Italien", "ES": "Spanien", "SE": "Schweden", "NO": "Norwegen",
        "PL": "Polen", "CZ": "Tschechien", "HU": "Ungarn", "RO": "Rumänien",
        "IE": "Irland", "PT": "Portugal",
    }
    return names.get(ioc_code.upper(), ioc_code)


def find_flag_image(nat_code):
    """Sucht Flagge - identisch zu pdf_dre_5.py"""
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


def _get_ordered_judge_positions(judges):
    pos_map = {0: "E", 1: "H", 2: "C", 3: "M", 4: "B"}
    available = set()
    for j in judges:
        pos = j.get("position", "")
        if isinstance(pos, int):
            lbl = pos_map.get(pos, "")
        else:
            lbl = str(pos)
        if lbl in ("E", "H", "C", "M", "B"):
            available.add(lbl)
    result = [p for p in ["E", "H", "C", "M", "B"] if p in available]
    while len(result) < 5:
        result.append("")
    return result[:5]


def _get_ordered_judges_all(judges):
    pos_map = {0: "E", 1: "H", 2: "C", 3: "M", 4: "B",
               5: "K", 6: "V", 7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
               "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"}
    dressage = []
    others   = []
    for j in judges:
        pos = j.get("position", "")
        lbl = pos_map.get(pos, str(pos)) if isinstance(pos, int) else pos_map.get(str(pos), str(pos))
        jj = j.copy()
        jj["pos_label"] = lbl
        if lbl in ("E", "H", "C", "M", "B"):
            dressage.append(jj)
        else:
            others.append(jj)
    ordered = []
    for p in ["E", "H", "C", "M", "B"]:
        for jj in dressage:
            if jj["pos_label"] == p:
                ordered.append(jj)
                break
    others.sort(key=lambda x: x["pos_label"])
    return ordered + others


# ---------------------------------------------------------------------------
# Pop-up: Uhrzeiten abfragen
# ---------------------------------------------------------------------------
def _ask_times():
    """Cloud-Version: liest Zeiten aus dem starterlist-Dict (gesetzt von app5_cloud.py)."""
    # Werte werden von app5_cloud.py über derby_config in starterlist eingetragen
    # Fallback: leere Strings
    return "", ""


# ---------------------------------------------------------------------------
# Sponsoren-Höhe
# ---------------------------------------------------------------------------
def get_sponsor_bar_height():
    path = "logos/sponsorenleiste.png"
    if not os.path.exists(path):
        return 25 * mm
    try:
        pil = PILImage.open(path)
        w, h = pil.size
        return (190 * h / w) * mm
    except:
        return 25 * mm


# ---------------------------------------------------------------------------
# Canvas mit Banner + Sponsorenleiste
# ---------------------------------------------------------------------------
class BannerCanvas(canvas.Canvas):
    _print_options = {}

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.sponsor_height = get_sponsor_bar_height() if self._print_options.get("show_sponsor_bar", True) else 0
        self.page_num  = 0
        self.banner_path   = None
        self.banner_height = 0
        if self._print_options.get("show_banner", True):
            bp = os.path.join("logos", "banner.png")
            if os.path.exists(bp):
                try:
                    pil = PILImage.open(bp)
                    iw, ih = pil.size
                    dpi = pil.info.get("dpi", (72, 72))
                    dx = dpi[0] if isinstance(dpi, tuple) else dpi
                    dw = iw * 72.0 / dx if dx > 150 else iw
                    dh = ih * 72.0 / dx if dx > 150 else ih
                    self.banner_path   = bp
                    self.banner_height = A4[0] * (dh / dw)
                except:
                    pass

    def showPage(self):
        self._draw_banner()
        self._draw_footer()
        self.page_num += 1
        canvas.Canvas.showPage(self)

    def _draw_banner(self):
        if self.page_num != 0 or not self.banner_path:
            return
        try:
            self.drawImage(self.banner_path, 0, A4[1] - self.banner_height,
                           width=A4[0], height=self.banner_height,
                           preserveAspectRatio=True, mask="auto")
        except:
            pass

    def _draw_footer(self):
        if not self._print_options.get("show_sponsor_bar", True):
            return
        sp = "logos/sponsorenleiste.png"
        if os.path.exists(sp):
            try:
                self.drawImage(sp, (A4[0] - 190 * mm) / 2, 4 * mm,
                               width=190 * mm, height=self.sponsor_height,
                               preserveAspectRatio=True, mask="auto")
            except:
                pass


# ---------------------------------------------------------------------------
# Flaggen-Zelle bauen (wie pdf_dre_5)
# ---------------------------------------------------------------------------
def _build_nat_cell(nationality, style_pos):
    nat_code_iso = get_nationality_code(nationality) if nationality else ""
    iso_to_display = {
        "DEU": "GER", "NLD": "NED", "CHE": "SUI", "DNK": "DEN", "HRV": "CRO",
        "GRC": "GRE", "BGR": "BUL", "ZAF": "RSA", "PRT": "POR", "LVA": "LAT",
        "ARE": "UAE", "CHL": "CHI", "URY": "URU", "SVN": "SLO", "MYS": "MAS",
        "GBR": "GBR", "AUT": "AUT", "BEL": "BEL", "FRA": "FRA", "ITA": "ITA",
        "ESP": "ESP", "SWE": "SWE", "NOR": "NOR", "POL": "POL", "CZE": "CZE",
        "HUN": "HUN", "ROU": "ROU", "IRL": "IRL", "USA": "USA", "CAN": "CAN",
    }
    nat_code_display = iso_to_display.get(nat_code_iso, nat_code_iso)
    flag_path = find_flag_image(nat_code_iso)  # gibt IOC-Pfad zurück wie in pdf_dre_5
    if flag_path and os.path.exists(flag_path):
        try:
            flag_img = Image(flag_path, width=5 * mm, height=3.5 * mm)
            mini = Table(
                [[flag_img], [Paragraph(f'<font size="6">{nat_code_display}</font>', style_pos)]],
                colWidths=[5 * mm]
            )
            mini.setStyle(TableStyle([
                ("ALIGN",  (0,0), (-1,-1), "CENTER"),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("LEFTPADDING",   (0,0), (-1,-1), 0),
                ("RIGHTPADDING",  (0,0), (-1,-1), 0),
                ("TOPPADDING",    (0,0), (-1,-1), 0),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ]))
            return mini
        except:
            pass
    if nat_code_display:
        return Paragraph(f'<font size="6">{nat_code_display}</font>', style_pos)
    return ""


# ---------------------------------------------------------------------------
# Starter-Inhalt (Reiter + Pferd) aufbauen (wie pdf_dre_5)
# ---------------------------------------------------------------------------
def _build_starter_content(s):
    athlete    = s.get("athlete") or {}
    rider_name = _safe_get(athlete, "name", "")
    club       = _safe_get(athlete, "club", "")
    nationality = _safe_get(athlete, "nation", "")
    horses     = s.get("horses") or []

    content_parts = []
    if rider_name:
        rider_line = f"<b>{rider_name.upper()}</b>"
        if nationality and nationality.upper() != "GER":
            if not club or club.strip().upper() in ("", "GASTLIZENZ GER"):
                country_full = get_country_name(get_nationality_code(nationality))
                if country_full:
                    rider_line += f" - <font size=7>{country_full}</font>"
            else:
                rider_line += f" - <font size=7>{club}</font>"
        elif club:
            rider_line += f" - <font size=7>{club}</font>"
        content_parts.append(rider_line)

    if horses:
        horse      = horses[0]
        horse_name = _safe_get(horse, "name", "")
        studbook   = _safe_get(horse, "studbook", "")
        horse_line = f"<b>{horse_name.upper()}</b>" if horse_name else ""
        if studbook and horse_line:
            horse_line += f" - {studbook}"
        if horse_line:
            content_parts.append(horse_line)

        details = []
        bs = _safe_get(horse, "breedingSeason", "")
        if bs:
            try:
                details.append(f"{datetime.now().year - int(bs)}jähr.")
            except:
                pass
        color = _safe_get(horse, "color", "")
        if color:
            details.append(color)
        sex = _safe_get(horse, "sex", "")
        sex_en = SEX_MAP.get(str(sex).upper(), sex)
        if sex_en:
            details.append(sex_en)
        sire    = _safe_get(horse, "sire", "")
        damsire = _safe_get(horse, "damSire", "")
        if sire:
            details.append(f"{sire} x {damsire}" if damsire else sire)
        if details:
            content_parts.append(f"<font size=7>{' / '.join(details)}</font>")

        owner   = _safe_get(horse, "owner", "")
        breeder = _safe_get(horse, "breeder", "")
        ob = []
        if owner:
            ob.append(f"B: {owner}")
        if breeder:
            ob.append(f"Z: {breeder}")
        if ob:
            content_parts.append(f"<font size=7>{' / '.join(ob)}</font>")

    return "<br/>".join(content_parts), nationality


# ---------------------------------------------------------------------------
# Hauptfunktion render()
# ---------------------------------------------------------------------------
def render(starterlist, filename, logo_max_width_cm=5.0):
    # Druckoptionen
    print_options  = starterlist.get("printOptions", {})
    show_banner    = print_options.get("show_banner",    True)
    show_sponsor   = print_options.get("show_sponsor_bar", True)
    show_title     = print_options.get("show_title",    True)
    show_header    = print_options.get("show_header",   True)

    BannerCanvas._print_options = print_options

    # Cloud-Version: Zeiten aus derby_config im starterlist-Dict lesen
    derby_cfg = starterlist.get("derby_config", {})
    begin_time = derby_cfg.get("begin_time", "")
    final_time = derby_cfg.get("final_time", "")

    # --- Dokument-Setup (wie pdf_nat) ---
    sponsor_height  = get_sponsor_bar_height() if show_sponsor else 0
    bottom_margin   = 4 * mm + sponsor_height + 1 * mm if show_sponsor else 8 * mm

    doc = SimpleDocTemplate(
        filename, pagesize=A4,
        rightMargin=10 * mm, leftMargin=10 * mm,
        topMargin=5 * mm, bottomMargin=bottom_margin
    )
    page_width = A4[0] - doc.leftMargin - doc.rightMargin
    elements   = []

    # --- Banner-Platzhalter ---
    if show_banner:
        bp = os.path.join("logos", "banner.png")
        if os.path.exists(bp):
            try:
                pil = PILImage.open(bp)
                iw, ih = pil.size
                dpi = pil.info.get("dpi", (72, 72))
                dx  = dpi[0] if isinstance(dpi, tuple) else dpi
                dw  = iw * 72.0 / dx if dx > 150 else iw
                dh  = ih * 72.0 / dx if dx > 150 else ih
                bh  = A4[0] * (dh / dw)
                elements.append(Spacer(1, bh - 6 * mm))
            except:
                pass

    # --- Styles (aus pdf_nat / pdf_dre_5) ---
    styles = getSampleStyleSheet()
    style_title    = ParagraphStyle("Title",  fontSize=14, leading=16, fontName="Helvetica-Bold", spaceAfter=2, alignment=TA_LEFT)
    style_comp     = ParagraphStyle("Comp",   fontSize=11, leading=13, fontName="Helvetica",      spaceAfter=4, alignment=TA_LEFT)
    style_info     = ParagraphStyle("Info",   fontSize=10, leading=12, fontName="Helvetica",      spaceAfter=4, alignment=TA_LEFT)
    style_sub      = ParagraphStyle("Sub",    fontSize=10, leading=12, fontName="Helvetica",      spaceAfter=4, alignment=TA_LEFT)
    style_hdr      = ParagraphStyle("Hdr",    fontSize=9,  leading=11, fontName="Helvetica-Bold", alignment=TA_CENTER,  textColor=colors.white)
    style_hdr_left = ParagraphStyle("HdrL",   fontSize=9,  leading=11, fontName="Helvetica-Bold", alignment=TA_LEFT,    textColor=colors.white)
    style_pos      = ParagraphStyle("Pos",    fontSize=8,  leading=10, fontName="Helvetica",      alignment=TA_CENTER)
    style_rider    = ParagraphStyle("Rider",  fontSize=9,  leading=10, fontName="Helvetica",      alignment=TA_LEFT)
    style_horse    = ParagraphStyle("Horse",  fontSize=9,  leading=10, fontName="Helvetica",      alignment=TA_LEFT)
    style_pause    = ParagraphStyle("Pause",  fontSize=9,  leading=11, fontName="Helvetica-Bold", alignment=TA_CENTER)
    style_section  = ParagraphStyle("Sect",   fontSize=11, leading=14, fontName="Helvetica-Bold", alignment=TA_LEFT, textColor=colors.black)
    style_date_right = ParagraphStyle("DateR", fontSize=10, leading=12, fontName="Helvetica-Bold", alignment=TA_RIGHT)

    # --- Header-Block ---
    comp = starterlist.get("competition") or {}
    logo_path   = starterlist.get("logoPath") if show_header else None
    header_parts = []

    if show_header:
        if show_title:
            show_title_text = starterlist.get("showTitle", "")
            if show_title_text:
                header_parts.append(Paragraph(f"<b>{show_title_text}</b>", style_title))

        comp_no    = comp.get("number") or starterlist.get("competitionNumber") or ""
        comp_title = comp.get("title")  or starterlist.get("competitionName")   or ""
        division   = comp.get("divisionNumber") or starterlist.get("divisionNumber")
        div_text   = ""
        try:
            if division is not None and int(division) > 0:
                div_text = f"{int(division)}. Abt. "
        except:
            div_text = f"{division} " if division else ""

        if comp_no or comp_title:
            comp_line = f"Prüfung {comp_no}"
            if div_text:
                comp_line += f" - {div_text}{comp_title}"
            elif comp_title:
                comp_line += f" — {comp_title}"
            header_parts.append(Paragraph(f"<b>{comp_line}</b>", style_comp))

        comp_info = starterlist.get("informationText", "")
        if comp_info:
            header_parts.append(Paragraph(comp_info.replace("\n", "<br/>"), style_info))

        subtitle = starterlist.get("subtitle", "")
        if subtitle:
            header_parts.append(Paragraph(subtitle, style_sub))

    # Datum/Ort
    start_raw    = comp.get("start") or starterlist.get("start")
    location     = starterlist.get("competitionLocation", "") or starterlist.get("location", "")
    date_line_text = None
    if start_raw:
        dl = _fmt_header_datetime(start_raw)
        date_line_text = f"{dl}  -  {location}" if location else dl

    # Logo + Header-Tabelle
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path)
            max_s = logo_max_width_cm * 10 * mm
            if logo.drawWidth > max_s or logo.drawHeight > max_s:
                scale = min(max_s / logo.drawWidth, max_s / logo.drawHeight)
                logo.drawWidth  *= scale
                logo.drawHeight *= scale
            lw  = max(logo.drawWidth + 5 * mm, 35 * mm)
            ht  = Table([[header_parts, logo]], colWidths=[page_width - lw, lw])
            ht.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("ALIGN",  (1,0), (1,0),   "RIGHT"),
                ("RIGHTPADDING", (1,0), (1,0), 0),
            ]))
            elements.append(ht)
        except:
            for p in header_parts:
                elements.append(p)
    else:
        for p in header_parts:
            elements.append(p)

    # --- Richterzeile (Inline-Format aus pdf_nat) ---
    judges = comp.get("judges") or starterlist.get("judges") or []
    if judges:
        pos_map = {0: "E", 1: "H", 2: "C", 3: "M", 4: "B"}
        judges_by_pos = {}
        for j in judges:
            pos = j.get("position", "")
            lbl = pos_map.get(pos, str(pos)) if isinstance(pos, int) else str(pos)
            name = j.get("name", "")
            if lbl and name:
                judges_by_pos.setdefault(lbl, []).append(name)

        displayed = [p for p in ["E", "H", "C", "M", "B"] if p in judges_by_pos]
        only_c    = len(displayed) == 1 and displayed[0] == "C"
        jury_parts = []
        for p in ["E", "H", "C", "M", "B"]:
            if p in judges_by_pos:
                names = " & ".join(judges_by_pos[p])
                if only_c:
                    jury_parts.append(names)
                else:
                    jury_parts.append(
                        f'<font name="Helvetica-Bold" backColor="black" color="white"> {p} </font> {names}'
                    )
        if jury_parts:
            judges_para  = Paragraph(f"<b>Richter:</b> {' '.join(jury_parts)}", style_sub)
            judges_table = Table([[judges_para]], colWidths=[page_width])
            judges_table.setStyle(TableStyle([
                ("LEFTPADDING",   (0,0), (-1,-1), 6),
                ("RIGHTPADDING",  (0,0), (-1,-1), 0),
                ("TOPPADDING",    (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            elements.append(Spacer(1, 3 * mm))
            elements.append(judges_table)
            elements.append(Spacer(1, 3 * mm))

    # --- "Starterliste" + Datum ---
    elements.append(Spacer(1, 1 * mm))
    starterliste_left = Paragraph("<b>Starterliste</b>", style_comp)
    if date_line_text:
        date_right = Paragraph(f"<b>{date_line_text}</b>", style_date_right)
        so_table   = Table([[starterliste_left, date_right]],
                           colWidths=[page_width * 0.5, page_width * 0.5])
        so_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "BOTTOM"),
            ("ALIGN",  (0,0), (0,0),  "LEFT"),
            ("ALIGN",  (1,0), (1,0),  "RIGHT"),
            ("TOPPADDING",    (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ]))
        elements.append(so_table)
    else:
        elements.append(starterliste_left)
    elements.append(Spacer(1, 2 * mm))

    # -----------------------------------------------------------------------
    # Starter-Tabelle aufbauen
    # -----------------------------------------------------------------------
    starters       = starterlist.get("starters") or []
    judge_positions = _get_ordered_judge_positions(judges)

    # Spaltenbreiten (identisch zu pdf_dre_5)
    col1      = 22 * mm            # Start + KNr
    col_nat   = 8 * mm             # Flagge
    col_judges = [12 * mm] * 6    # 5 Richter + Total
    col2      = page_width - col1 - col_nat - sum(col_judges)  # Reiter + Pferd

    col_widths = [col1, col2, col_nat] + col_judges  # 9 Spalten

    data_texts = []
    meta       = []

    # Tabellenkopf
    data_texts.append([
        "<b># / KNr</b>",
        "<b>Reiter / Pferd</b><br/><font size=7>Alter / Farbe / Geschlecht / Vater x Muttervater / Besitzer / Züchter</font>",
        "<b><font size=6>Nat</font></b>",
        f"<b>{judge_positions[0]}</b>", f"<b>{judge_positions[1]}</b>",
        f"<b>{judge_positions[2]}</b>", f"<b>{judge_positions[3]}</b>",
        f"<b>{judge_positions[4]}</b>", "<b>Total</b>"
    ])
    meta.append({"type": "header"})

    # "Prüfungsbeginn"-Zeile (vor Starter 1)
    begin_label = f"Prüfungsbeginn {begin_time}".strip() if begin_time else "Prüfungsbeginn"
    data_texts.append([begin_label, "", "", "", "", "", "", "", ""])
    meta.append({"type": "section_begin"})

    # Starters 1-6 verarbeiten
    for idx, s in enumerate(starters[:6]):
        star_num = idx + 1  # 1-basiert

        # Leerzeile nach Starter 2 (Pause zwischen Halbfinale 1 und 2)
        if star_num == 3:
            data_texts.append(["", "", "", "", "", "", "", "", ""])
            meta.append({"type": "spacer"})

        # Trennzeile vor Finale (vor Starter 5)
        if star_num == 5:
            finale_label = f"Finale {final_time}".strip() if final_time else "Finale"
            data_texts.append([finale_label, "", "", "", "", "", "", "", ""])
            meta.append({"type": "section_finale"})

        nr            = str(s.get("startNumber", star_num))
        hors_concours = bool(s.get("horsConcours", False))
        withdrawn     = bool(s.get("withdrawn", False))
        horses        = s.get("horses") or []
        cno           = str(horses[0].get("cno", "")) if horses else ""

        # Startzeit aus den Daten (wie pdf_dre_5)
        tstr = _fmt_time(s.get("startTime"))

        if withdrawn:
            nr_text  = f"<strike><b>{nr}</b></strike>"
            cno_text = f"<strike><font size=8>{cno}</font></strike>"
        else:
            nr_text  = f"<b>{nr}</b><br/><font size=7>AK</font>" if hors_concours else f"<b>{nr}</b>"
            cno_text = f"<font size=8>{cno}</font>"

        # Start+KNr Mini-Tabelle: Nr|KNr oben, Uhrzeit unten (fett, zentriert)
        start_knr_data = []

        # Zeile 1: Nr und KNr nebeneinander
        start_knr_data.append([Paragraph(nr_text, style_pos), Paragraph(cno_text, style_pos)])

        # Zeile 2 (unten): Uhrzeit fett zentriert
        if tstr:
            time_text = f'<b><strike>{tstr}</strike></b>' if withdrawn else f'<b>{tstr}</b>'
            start_knr_data.append([Paragraph(time_text, style_pos)])

        if tstr:
            start_knr_table = Table(start_knr_data, colWidths=[11 * mm, 11 * mm])
            start_knr_table.setStyle(TableStyle([
                ("GRID",          (0,0), (1,0),   0.25, colors.grey),
                ("SPAN",          (0,1), (1,1)),
                ("VALIGN",        (0,0), (1,0),   "MIDDLE"),
                ("VALIGN",        (0,1), (1,1),   "BOTTOM"),
                ("ALIGN",         (0,0), (-1,-1),  "CENTER"),
                ("LEFTPADDING",   (0,0), (-1,-1),  1),
                ("RIGHTPADDING",  (0,0), (-1,-1),  1),
                ("TOPPADDING",    (0,0), (-1,-1),  1),
                ("BOTTOMPADDING", (0,0), (-1,-1),  1),
                ("TOPPADDING",    (0,1), (1,1),    6),
            ]))
        else:
            start_knr_table = Table(start_knr_data, colWidths=[11 * mm, 11 * mm])
            start_knr_table.setStyle(TableStyle([
                ("GRID",   (0,0), (1,0), 0.25, colors.grey),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ALIGN",  (0,0), (-1,-1), "CENTER"),
                ("LEFTPADDING",   (0,0), (-1,-1), 1),
                ("RIGHTPADDING",  (0,0), (-1,-1), 1),
                ("TOPPADDING",    (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ]))

        # Finale-Starter (5 + 6): kein Reitername, stattdessen Platzhalter + Pferd normal
        if star_num >= 5:
            # Pferd-Daten aufbauen ohne Reiternamen
            _, nationality = _build_starter_content(s)
            # Platzhalter-Zeile für Reiter
            rider_placeholder = "Reiter: ___________________________"
            # Pferddaten separat zusammenbauen
            horse_parts = []
            if horses:
                horse = horses[0]
                horse_name = _safe_get(horse, "name", "")
                studbook   = _safe_get(horse, "studbook", "")
                horse_line = f"<b>{horse_name.upper()}</b>" if horse_name else ""
                if studbook and horse_line:
                    horse_line += f" - {studbook}"
                if horse_line:
                    horse_parts.append(horse_line)

                details = []
                bs = _safe_get(horse, "breedingSeason", "")
                if bs:
                    try:
                        details.append(f"{datetime.now().year - int(bs)}jähr.")
                    except:
                        pass
                color = _safe_get(horse, "color", "")
                if color:
                    details.append(color)
                sex = _safe_get(horse, "sex", "")
                sex_en = SEX_MAP.get(str(sex).upper(), sex)
                if sex_en:
                    details.append(sex_en)
                sire    = _safe_get(horse, "sire", "")
                damsire = _safe_get(horse, "damSire", "")
                if sire:
                    details.append(f"{sire} x {damsire}" if damsire else sire)
                if details:
                    horse_parts.append(f"<font size=7>{' / '.join(details)}</font>")
                owner   = _safe_get(horse, "owner", "")
                breeder = _safe_get(horse, "breeder", "")
                ob = []
                if owner:
                    ob.append(f"B: {owner}")
                if breeder:
                    ob.append(f"Z: {breeder}")
                if ob:
                    horse_parts.append(f"<font size=7>{' / '.join(ob)}</font>")

            combined_content = "Reiter: _____________________________________________<br/>" + ("<br/>".join(horse_parts) if horse_parts else "")
            nat_cell = ""  # Keine Flagge für Finale
        else:
            combined_content, nationality = _build_starter_content(s)
            nat_cell = _build_nat_cell(nationality, style_pos)

        data_texts.append([start_knr_table, combined_content, nat_cell, "", "", "", "", "", ""])
        meta.append({"type": "starter", "withdrawn": withdrawn, "horsConcours": hors_concours})

    # -----------------------------------------------------------------------
    # Tabelle rendern
    # -----------------------------------------------------------------------
    table_rows = []
    for i, row in enumerate(data_texts):
        if i >= len(meta):
            continue
        m = meta[i]

        if m["type"] == "header":
            table_rows.append([
                Paragraph(row[0], style_hdr),
                Paragraph(row[1], style_hdr_left),
                Paragraph(row[2], style_hdr),
                Paragraph(row[3], style_hdr), Paragraph(row[4], style_hdr),
                Paragraph(row[5], style_hdr), Paragraph(row[6], style_hdr),
                Paragraph(row[7], style_hdr), Paragraph(row[8], style_hdr),
            ])

        elif m["type"] in ("section_begin", "section_finale"):
            # Abschnitts-Trennzeile: ganzer SPAN, grau, linksbündig
            table_rows.append([
                Paragraph(f"<b>{row[0]}</b>", style_section),
                "", "", "", "", "", "", "", ""
            ])

        elif m["type"] == "spacer":
            # Leerzeile zwischen Halbfinale 1 und 2
            table_rows.append(["", "", "", "", "", "", "", "", ""])

        else:  # starter
            withdrawn = m.get("withdrawn", False)

            def maybe_strike(text, sty):
                if not text:
                    return Paragraph("", sty)
                return Paragraph(f"<strike>{text}</strike>" if withdrawn else text, sty)

            start_val = row[0]
            if not isinstance(start_val, Table):
                start_val = Paragraph(str(start_val) if start_val else "", style_pos)

            nat_val = row[2]
            if not isinstance(nat_val, (Table, Image, Paragraph)):
                nat_val = Paragraph(str(nat_val) if nat_val else "", style_pos)

            content_val = row[1]
            if isinstance(content_val, str):
                content_para = maybe_strike(content_val, style_horse)
            else:
                content_para = content_val

            table_rows.append([
                start_val,
                content_para,
                nat_val,
                maybe_strike(row[3], style_pos), maybe_strike(row[4], style_pos),
                maybe_strike(row[5], style_pos), maybe_strike(row[6], style_pos),
                maybe_strike(row[7], style_pos), maybe_strike(row[8], style_pos),
            ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0),  colors.HexColor("#404040")),
        ("TEXTCOLOR",  (0,0), (-1,0),  colors.white),
        ("VALIGN",     (0,0), (-1,-1), "TOP"),
        ("ALIGN",      (0,0), (0,-1),  "CENTER"),
        ("ALIGN",      (1,0), (1,-1),  "LEFT"),
        ("ALIGN",      (2,0), (-1,-1), "CENTER"),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])

    # Abschnitts-Zeilen stylen + Zebra für Starter
    starter_count = 0
    for ri in range(1, len(table_rows)):
        if ri >= len(meta):
            continue
        m = meta[ri]

        if m["type"] in ("section_begin", "section_finale"):
            ts.add("SPAN",            (0, ri), (-1, ri))
            ts.add("BACKGROUND",      (0, ri), (-1, ri), colors.HexColor("#d0d0d0"))
            ts.add("ALIGN",           (0, ri), (-1, ri), "LEFT")
            ts.add("TEXTCOLOR",       (0, ri), (-1, ri), colors.black)
            ts.add("TOPPADDING",      (0, ri), (-1, ri), 8)
            ts.add("BOTTOMPADDING",   (0, ri), (-1, ri), 8)
            ts.add("LEFTPADDING",     (0, ri), (-1, ri), 6)
            starter_count = 0

        elif m["type"] == "spacer":
            ts.add("SPAN",            (0, ri), (-1, ri))
            ts.add("BACKGROUND",      (0, ri), (-1, ri), colors.white)
            ts.add("TOPPADDING",      (0, ri), (-1, ri), 3)
            ts.add("BOTTOMPADDING",   (0, ri), (-1, ri), 3)

        elif m["type"] == "starter":
            if starter_count % 2 == 1:
                ts.add("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#f5f5f5"))
            if m.get("withdrawn"):
                ts.add("TEXTCOLOR", (0, ri), (-1, ri), colors.darkgrey)
            starter_count += 1

    t.setStyle(ts)
    elements.append(t)

    # --- PDF bauen ---
    if show_banner or show_sponsor:
        doc.build(elements, canvasmaker=BannerCanvas)
    else:
        doc.build(elements)

    print(f"PDF Derby Cloud: {filename}")
