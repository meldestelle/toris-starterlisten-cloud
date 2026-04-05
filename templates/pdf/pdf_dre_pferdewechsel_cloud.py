# -*- coding: utf-8 -*-
# templates/pdf/pdf_dre_pferdewechsel.py - Pferdewechsel-Format: 3 Reiter tauschen Pferde
#
# Struktur:
#   3 Reiter reiten je alle 3 Pferde (9 Starts gesamt)
#   Kreuztabelle Reiter × Pferd mit Summen
#   Richtverfahren-abhängige Spalten: 402.A / 402.B / 402.C
#
# Pop-up-Eingabe:  Zuordnung Reiter ↔ eigenes Pferd + optionale Ergebnisse
# Richterzeile:    Inline-Format aus pdf_nat
# Starterliste:    Spaltenformat aus pdf_dre_5 (vollständige Pausenlogik)

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
        return f"{weekday}, {dt.day}. {month} {dt.year} um {dt.strftime('%H:%M')} Uhr"
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
    if not ioc_code:
        return ""
    names = {
        "DEU": "Deutschland", "AUT": "Österreich", "CHE": "Schweiz",
        "NLD": "Niederlande", "BEL": "Belgien", "FRA": "Frankreich",
        "ITA": "Italien", "ESP": "Spanien", "PRT": "Portugal",
        "GBR": "Großbritannien", "IRL": "Irland", "SWE": "Schweden",
        "NOR": "Norwegen", "DNK": "Dänemark", "FIN": "Finnland",
        "POL": "Polen", "CZE": "Tschechien", "HUN": "Ungarn",
        "ROU": "Rumänien", "HRV": "Kroatien", "SVN": "Slowenien",
        "GRC": "Griechenland", "TUR": "Türkei", "RUS": "Russland",
        "UKR": "Ukraine", "BLR": "Belarus", "LVA": "Lettland",
        "LTU": "Litauen", "EST": "Estland", "USA": "USA",
        "CAN": "Kanada", "AUS": "Australien", "NZL": "Neuseeland",
        "JPN": "Japan", "KOR": "Südkorea", "CHN": "China",
        "BRA": "Brasilien", "ARG": "Argentinien", "MEX": "Mexiko",
        "RSA": "Südafrika", "ZAF": "Südafrika", "ISR": "Israel",
        "ARE": "Vereinigte Arab. Emirate", "QAT": "Katar",
    }
    return names.get(ioc_code, ioc_code)


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
# Cloud-Version: Derby-Konfiguration aus starterlist-Dict lesen
# ---------------------------------------------------------------------------
def _ask_derby_config(starters, starterlist=None):
    """
    Cloud-Version ohne Tkinter.
    Liest die Konfiguration aus starterlist["derby_config"],
    das von app5_cloud.py befüllt wird.
    """
    # Standard-Zuordnung: Starter 1-3, eigene Pferde auf Diagonale
    entries = []
    for s in starters[:3]:
        ath   = s.get("athlete") or {}
        hrs   = s.get("horses") or []
        name  = ath.get("name", f"Reiter {s.get('startNumber', '?')}")
        pferd = hrs[0].get("name", f"Pferd {s.get('startNumber', '?')}") if hrs else "—"
        entries.append((name, pferd))

    cfg = (starterlist or {}).get("derby_config", {})

    own_results             = cfg.get("own_results", {})
    assignment              = cfg.get("assignment", list(entries))
    show_horse_details_from_4 = cfg.get("show_horse_details_from_4", False)

    return own_results, assignment, show_horse_details_from_4


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
def _build_starter_content(s, skip_horse_details=False):
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

        if not skip_horse_details:
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
            sex_de = SEX_MAP.get(str(sex).upper(), sex)
            if sex_de:
                details.append(sex_de)
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
# Kreuztabelle Reiter × Pferd aufbauen
# ---------------------------------------------------------------------------
def _build_cross_table(starters, own_results, assignment, page_width, judging_rule=""):
    """
    3×3 Kreuztabelle Reiter × Pferd.
    Eigenes-Pferd-Zelle: grau (#c8c8c8), kein Symbol.
    Spalten: Label | Pferd1 | Pferd2 | Pferd3 | Σ Reiter
    + Σ Pferd Zeile unten.
    """
    style_hdr   = ParagraphStyle("xhdr",  fontSize=8,  leading=10, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER, textColor=colors.white)
    style_hdr_l = ParagraphStyle("xhdrl", fontSize=8,  leading=10, fontName="Helvetica-Bold",
                                  alignment=TA_LEFT,   textColor=colors.white)
    style_cell  = ParagraphStyle("xcell", fontSize=9,  leading=11, fontName="Helvetica",
                                  alignment=TA_CENTER)
    style_own   = ParagraphStyle("xown",  fontSize=9,  leading=11, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER)
    style_sum   = ParagraphStyle("xsum",  fontSize=8,  leading=10, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER)
    style_lbl   = ParagraphStyle("xlbl",  fontSize=8,  leading=10, fontName="Helvetica-Bold",
                                  alignment=TA_LEFT,   textColor=colors.white)

    OWN_GREY = colors.HexColor("#c8c8c8")

    reiter_namen = [a[0] for a in assignment]
    pferd_namen  = [a[1] for a in assignment]

    # Spaltenbreiten: Label + 3 Pferd-Spalten + Σ
    col_label = 48 * mm
    col_sum   = 22 * mm
    col_pferd = (page_width - col_label - col_sum) / 3
    col_widths = [col_label, col_pferd, col_pferd, col_pferd, col_sum]

    # Header-Zeile
    hdr = [Paragraph("<b>Reiter \\ Pferd</b>", style_hdr_l)]
    for pname in pferd_namen:
        short = pname if len(pname) <= 16 else pname[:14] + "…"
        hdr.append(Paragraph(f"<b>{short}</b>", style_hdr))
    hdr.append(Paragraph("<b>Σ Reiter</b>", style_hdr))

    rows = [hdr]

    for ri, rname in enumerate(reiter_namen):
        short_r = rname if len(rname) <= 24 else rname[:22] + "…"
        row = [Paragraph(short_r, style_lbl)]
        for ci in range(3):
            is_own = (ri == ci)   # Diagonale = eigenes Pferd
            if is_own:
                score = own_results.get(ri, "")
                if score:
                    try:
                        val = float(score)
                        rule_up = str(judging_rule).upper().replace(" ","").replace(".","").replace(",","")
                        txt = f"{val:.2f}" if "402A" in rule_up else f"{val:.2f} %"
                    except:
                        txt = score
                    row.append(Paragraph(txt, style_own))
                else:
                    row.append(Paragraph("", style_own))
            else:
                row.append(Paragraph("", style_cell))
        row.append(Paragraph("", style_sum))
        rows.append(row)

    # Σ Pferd Zeile
    sum_row = [Paragraph("<b>Σ Pferd</b>", style_lbl)]
    for _ in range(4):
        sum_row.append(Paragraph("", style_sum))
    rows.append(sum_row)

    t = Table(rows, colWidths=col_widths)
    ts = TableStyle([
        ("BOX",           (0, 0), (-1, -1),  0.5,  colors.black),
        ("INNERGRID",     (0, 0), (-1, -1),  0.25, colors.grey),
        ("BACKGROUND",    (0, 0), (-1, 0),   colors.HexColor("#404040")),
        ("TEXTCOLOR",     (0, 0), (-1, 0),   colors.white),
        ("BACKGROUND",    (0, 1), (0, -1),   colors.HexColor("#404040")),
        ("TEXTCOLOR",     (0, 1), (0, -1),   colors.white),
        ("BACKGROUND",    (-1, 1), (-1, -1), colors.HexColor("#e8e8e8")),
        ("BACKGROUND",    (1, -1), (-1, -1), colors.HexColor("#e8e8e8")),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
    ])

    # Eigene-Pferd-Felder: grau (Diagonale)
    for ri in range(3):
        ci = ri + 1   # +1 wegen Label-Spalte
        actual_ri = ri + 1   # +1 wegen Header-Zeile
        ts.add("BACKGROUND", (ci, actual_ri), (ci, actual_ri), OWN_GREY)

    t.setStyle(ts)
    return t


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

    # Starter vorab laden für den Dialog
    starters = starterlist.get("starters") or []

    # Derby-Konfiguration per Pop-up erfragen
    own_results, assignment, show_horse_details_from_4 = _ask_derby_config(starters, starterlist)

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

    # --- Kreuztabelle Reiter × Pferd ---
    judging_rule = comp.get("judgingRule") or starterlist.get("judgingRule") or ""
    cross_table = _build_cross_table(starters, own_results, assignment, page_width, judging_rule)
    elements.append(cross_table)
    elements.append(Spacer(1, 4 * mm))

    # -----------------------------------------------------------------------
    # Starter-Tabelle aufbauen (vollständige dre5-Pausenlogik)
    # -----------------------------------------------------------------------
    judge_positions = _get_ordered_judge_positions(judges)

    # Breaks aufbereiten
    breaks = starterlist.get("breaks") or []
    breaks_by_after = {}
    for b in breaks:
        try:
            after_num = b.get("afterNumberInCompetition")
            k = -1 if after_num is None else int(after_num)
            breaks_by_after.setdefault(k, []).append(b)
        except:
            continue

    # Richtverfahren bestimmen
    judging_rule = comp.get("judgingRule") or starterlist.get("judgingRule") or ""
    rule = str(judging_rule).upper().replace(" ", "").replace(".", "").replace(",", "")

    # Spaltenbreiten + Header je Richtverfahren
    col1    = 22 * mm
    col_nat = 8  * mm

    if "402A" in rule:
        # 402.A: nur 1 Ergebnis-Spalte
        score_cols   = [30 * mm]
        score_labels = ["<b>Ergebnis</b>"]
        n_cols       = 1
    elif "402C" in rule:
        # 402.C: Aufgabe + Qualität + Total
        score_cols   = [20 * mm, 20 * mm, 20 * mm]
        score_labels = ["<b>Aufgabe</b>", "<b>Qualität</b>", "<b>Total</b>"]
        n_cols       = 3
    else:
        # 402.B / Standard: E H C M B + Total
        score_cols   = [12 * mm] * 6
        score_labels = [f"<b>{judge_positions[i]}</b>" for i in range(5)] + ["<b>Total</b>"]
        n_cols       = 6

    col2       = page_width - col1 - col_nat - sum(score_cols)
    col_widths = [col1, col2, col_nat] + score_cols
    n_total    = 3 + n_cols  # col1 + col2 + nat + score_cols

    def empty_row(txt=""):
        return [txt] + [""] * (n_total - 1)

    data_texts = []
    meta       = []

    # Tabellenkopf
    hdr = [
        "<b># / KNr</b>",
        "<b>Reiter / Pferd</b><br/><font size=7>Alter / Farbe / Geschlecht / Vater x Muttervater / Besitzer / Züchter</font>",
        "<b><font size=6>Nat</font></b>",
    ] + score_labels
    data_texts.append(hdr)
    meta.append({"type": "header"})

    # Pause VOR dem ersten Starter
    if 0 in breaks_by_after:
        for br in breaks_by_after[0]:
            pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
            data_texts.append(empty_row(pause_text))
            meta.append({"type": "pause"})

    # Starter-Schleife
    current_group    = None
    group_time_shown = False

    for s in starters:
        starter_group = s.get("groupNumber")

        if starter_group is not None and starter_group > 0 and starter_group != current_group:
            group_text = f"Abteilung {starter_group}"
            data_texts.append(empty_row(group_text))
            meta.append({"type": "group"})
            current_group    = starter_group
            group_time_shown = False

        nr            = str(s.get("startNumber", ""))
        hors_concours = bool(s.get("horsConcours", False))
        withdrawn     = bool(s.get("withdrawn", False))
        horses        = s.get("horses") or []
        cno           = str(horses[0].get("cno", "")) if horses else ""
        tstr          = _fmt_time(s.get("startTime"))

        if starter_group is None or starter_group == 0 or not group_time_shown:
            start_time_display = tstr
            if starter_group is not None and starter_group > 0:
                group_time_shown = True
        else:
            start_time_display = ""

        if withdrawn:
            nr_text  = f"<strike><b>{nr}</b></strike>"
            cno_text = f"<strike><font size=8>{cno}</font></strike>"
        else:
            nr_text  = f"<b>{nr}</b><br/><font size=7>AK</font>" if hors_concours else f"<b>{nr}</b>"
            cno_text = f"<font size=8>{cno}</font>"

        start_knr_data = []
        start_knr_data.append([Paragraph(nr_text, style_pos), Paragraph(cno_text, style_pos)])
        if start_time_display:
            time_text = f'<b><strike>{start_time_display}</strike></b>' if withdrawn else f'<b>{start_time_display}</b>'
            start_knr_data.append([Paragraph(time_text, style_pos)])

        if start_time_display:
            start_knr_table = Table(start_knr_data, colWidths=[11 * mm, 11 * mm])
            start_knr_table.setStyle(TableStyle([
                ("GRID",          (0,0), (1,0),  0.25, colors.grey),
                ("SPAN",          (0,1), (1,1)),
                ("VALIGN",        (0,0), (1,0),  "MIDDLE"),
                ("VALIGN",        (0,1), (1,1),  "BOTTOM"),
                ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                ("LEFTPADDING",   (0,0), (-1,-1), 1),
                ("RIGHTPADDING",  (0,0), (-1,-1), 1),
                ("TOPPADDING",    (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
                ("TOPPADDING",    (0,1), (1,1),   6),
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

        # Ab Starter 4 wiederholen sich die Pferde → Details ggf. weglassen
        try:
            nr_int = int(nr)
        except:
            nr_int = 0

        skip = (nr_int >= 4) and (not show_horse_details_from_4)
        combined_content, nationality = _build_starter_content(s, skip_horse_details=skip)
        nat_cell = _build_nat_cell(nationality, style_pos)

        row = [start_knr_table, combined_content, nat_cell] + [""] * n_cols
        data_texts.append(row)
        meta.append({"type": "starter", "withdrawn": withdrawn, "horsConcours": hors_concours, "start_number": nr_int})

        try:
            cur = int(nr)
        except:
            cur = None
        if cur is not None and cur in breaks_by_after:
            for br in breaks_by_after[cur]:
                pause_text = _format_pause_text(br.get("totalSeconds", 0), br.get("informationText", ""))
                data_texts.append(empty_row(pause_text))
                meta.append({"type": "pause"})

    # -----------------------------------------------------------------------
    # Tabelle rendern
    # -----------------------------------------------------------------------
    table_rows = []
    for i, row in enumerate(data_texts):
        if i >= len(meta):
            continue
        m = meta[i]

        empties = [Paragraph("", style_pos)] * (n_total - 3)

        if m["type"] == "header":
            table_rows.append(
                [Paragraph(row[j], style_hdr if j != 1 else style_hdr_left) for j in range(n_total)]
            )

        elif m["type"] == "group":
            table_rows.append(
                [Paragraph(f"<b>{row[0]}</b>", style_hdr_left)] +
                [Paragraph("", style_pos)] * (n_total - 1)
            )

        elif m["type"] == "pause":
            table_rows.append(
                [Paragraph(row[0], style_pause)] +
                [Paragraph("", style_pos)] * (n_total - 1)
            )

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
            content_para = maybe_strike(content_val, style_horse) if isinstance(content_val, str) else content_val

            score_cells = [maybe_strike(row[3 + j], style_pos) for j in range(n_cols)]

            table_rows.append([start_val, content_para, nat_val] + score_cells)

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

    # Zebra + Gruppen/Pausen-Styling (wie dre5)
    starter_count = 0
    for ri in range(1, len(table_rows)):
        if ri >= len(meta):
            continue
        m = meta[ri]

        if m["type"] == "group":
            ts.add("SPAN",       (0, ri), (-1, ri))
            ts.add("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#606060"))
            ts.add("TEXTCOLOR",  (0, ri), (-1, ri), colors.white)
            ts.add("ALIGN",      (0, ri), (-1, ri), "LEFT")
            starter_count = 0

        elif m["type"] == "pause":
            ts.add("SPAN",  (0, ri), (-1, ri))
            ts.add("ALIGN", (0, ri), (-1, ri), "CENTER")
            prev_was_gray = (starter_count - 1) % 2 == 1
            if not prev_was_gray:
                ts.add("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#f5f5f5"))
            starter_count -= 1

        elif m["type"] == "starter":
            if starter_count % 2 == 1:
                ts.add("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#f5f5f5"))
            if m.get("withdrawn"):
                ts.add("TEXTCOLOR", (0, ri), (-1, ri), colors.darkgrey)
            # Ab Starter 4 ohne Details: mehr Padding für bessere Lesbarkeit
            if m.get("start_number", 0) >= 4 and not show_horse_details_from_4:
                ts.add("TOPPADDING",    (0, ri), (-1, ri), 6)
                ts.add("BOTTOMPADDING", (0, ri), (-1, ri), 6)
            starter_count += 1

    t.setStyle(ts)
    elements.append(t)

    # --- PDF bauen ---
    if show_banner or show_sponsor:
        doc.build(elements, canvasmaker=BannerCanvas)
    else:
        doc.build(elements)

    print(f"PDF Pferdewechsel Cloud: {filename}")
