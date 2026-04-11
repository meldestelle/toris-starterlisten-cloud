# -*- coding: utf-8 -*-
# pdf_export.py
# 
# Available PDF Templates:
# - pdf_abstammung_logo.py - Standard with pedigree
# - pdf_int.py - International style with flags (English)
# - pdf_nat.py - National style with flags (German)
#
import os
import importlib.util
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

TEMPLATES_DIR = os.path.join("templates", "pdf")
OUTPUT_DIR = "Ausgabe"  # Einheitlicher Ausgabeordner

def _ensure_output_dir():
    """Stellt sicher, dass das Ausgabe-Verzeichnis existiert"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def _find_template_file(template_name: str) -> str:
    """
    Sucht im templates/pdf-Ordner nach einer passenden .py-Datei.
    - template_name kann "pdf_details" oder "pdf_details.py" oder "pdf_details.py.py" sein.
    - Gibt den Dateinamen (mit .py) zurück oder None wenn nicht gefunden.
    """
    if not os.path.isdir(TEMPLATES_DIR):
        raise FileNotFoundError(f"Templates-Ordner nicht gefunden: {TEMPLATES_DIR}")

    # Normalisiere: entferne doppelte Endungen und extrahiere Basisname
    base = os.path.splitext(os.path.basename(template_name or ""))[0]

    # Liste aller .py-Templates im Ordner
    candidates = [f for f in os.listdir(TEMPLATES_DIR) if f.endswith(".py")]

    # 1) exakte Übereinstimmung des Basenames
    for f in candidates:
        if os.path.splitext(f)[0] == base:
            return f

    # 2) falls kein basename angegeben, benutze pdf_default.py falls vorhanden
    if not base:
        if "pdf_default.py" in candidates:
            return "pdf_default.py"
        if candidates:
            return candidates[0]

    # 3) keine gefunden -> Fehlschlag mit Hinweis
    raise FileNotFoundError(
        f"Template {template_name} nicht gefunden in {TEMPLATES_DIR}. Verfügbare: {candidates}"
    )

def _find_logo_file(logo_dir: str, basename: str) -> str:
    """Sucht Logo in .png / .jpg / .jpeg"""
    for ext in [".png", ".jpg", ".jpeg"]:
        path = os.path.join(logo_dir, basename + ext)
        if os.path.exists(path):
            return path
    return None

def _get_competition_logo_path(starterlist: dict, username: str = None) -> str:
    """
    Ermittelt den Pfad zum prüfungsspezifischen Logo basierend auf XXY-Schema.
    Sucht zuerst im User-Unterordner, dann im gemeinsamen logos/ Ordner.
    """
    comp_number = starterlist.get("competitionNumber")
    div_number  = starterlist.get("divisionNumber")

    if username and username.strip() and username.strip().lower() != "standard":
        logo_dir = os.path.join("logos", username.strip())
    else:
        logo_dir = "logos"

    if not comp_number:
        fallback_path = _find_logo_file(logo_dir, "logo")
        if fallback_path:
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Standard-Logo gefunden, ohne Logo fortfahren")
            return None

    try:
        comp_formatted = f"{int(comp_number):02d}"
    except (ValueError, TypeError):
        comp_formatted = str(comp_number).zfill(2)

    if div_number:
        try:
            div_formatted = f"{int(div_number)}"
        except (ValueError, TypeError):
            div_formatted = str(div_number)
    else:
        div_formatted = "0"

    logo_basename = f"{comp_formatted}{div_formatted}"
    logo_path = _find_logo_file(logo_dir, logo_basename)

    print(f"PDF DEBUG: Suche Logo: {logo_dir}/{logo_basename}.*")

    if logo_path:
        print(f"PDF DEBUG: Prüfungsspezifisches Logo gefunden: {logo_path}")
        return logo_path
    else:
        fallback_path = _find_logo_file(logo_dir, "logo")
        if fallback_path:
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Logo gefunden, ohne Logo fortfahren")
            return None

def _get_banner_sponsor_paths(username: str = None) -> dict:
    """
    Ermittelt die Pfade für Banner und Sponsorenleiste.
    Sucht zuerst im User-Unterordner, dann im gemeinsamen logos/ Ordner.
    """
    if username and username.strip() and username.strip().lower() != "standard":
        user_dir = os.path.join("logos", username.strip())
    else:
        user_dir = None

    result = {}
    for key, basename in [("bannerPath", "banner"), ("sponsorPath", "sponsorenleiste")]:
        found = None
        if user_dir:
            found = _find_logo_file(user_dir, basename)
        if not found:
            found = _find_logo_file("logos", basename)
        if found:
            result[key] = found
            print(f"PDF DEBUG: {key} = {found}")
    return result
    """
    Ermittelt den Pfad zum prüfungsspezifischen Logo basierend auf XXY-Schema
    XX = Competition (zweistellig), Y = Division (einstellig, 0 wenn keine Abteilung)
    Gibt None zurück wenn kein Logo gefunden wird (kein Fehler)
    """
    comp_number = starterlist.get("competitionNumber")
    div_number  = starterlist.get("divisionNumber")

    if not comp_number:
        fallback_path = _find_logo_file("logos", "logo")
        if fallback_path:
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Standard-Logo gefunden, ohne Logo fortfahren")
            return None

    try:
        comp_formatted = f"{int(comp_number):02d}"
    except (ValueError, TypeError):
        comp_formatted = str(comp_number).zfill(2)

    if div_number:
        try:
            div_formatted = f"{int(div_number)}"
        except (ValueError, TypeError):
            div_formatted = str(div_number)
    else:
        div_formatted = "0"

    logo_basename = f"{comp_formatted}{div_formatted}"
    logo_path = _find_logo_file("logos", logo_basename)

    print(f"PDF DEBUG: Suche Logo: logos/{logo_basename}.*")

    if logo_path:
        print(f"PDF DEBUG: Prüfungsspezifisches Logo gefunden: {logo_path}")
        return logo_path
    else:
        fallback_path = _find_logo_file("logos", "logo")
        if fallback_path:
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Logo gefunden, ohne Logo fortfahren")
            return None

def create_pdf(starterlist: dict, filename: str, template_name: str, spacing_top_cm: float = 0, spacing_bottom_cm: float = 0, logo_max_width_cm: float = 5.0, print_options: dict = None, output_dir: str = None, username: str = None):
    """
    Lädt das angegebene Template-Modul aus templates/pdf und ruft dessen render(starterlist, filename) auf.
    Alle Dateien werden im 'Ausgabe'-Ordner erstellt.
    Erweitert starterlist um prüfungsspezifischen Logo-Pfad.
    
    Für Templates die mit "liste_" beginnen werden spacing_top_cm und spacing_bottom_cm
    als Abstände oben und unten verwendet.
    
    Für Templates mit Logo-Unterstützung wird logo_max_width_cm als maximale Logo-Breite verwendet.
    
    print_options enthält Druckoptionen:
        sponsor_top, sponsor_bottom, single_sided, show_banner, show_sponsor_bar, show_title
    """
    # Ausgabe-Ordner sicherstellen
    target_dir = output_dir if output_dir else OUTPUT_DIR
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    # Vollständigen Pfad für Ausgabedatei erstellen
    output_path = os.path.join(target_dir, filename)

    # Prüfungsspezifisches Logo ermitteln
    logo_path = _get_competition_logo_path(starterlist, username=username)

    # Starterlist um Logo-Information erweitern (nur wenn Logo vorhanden)
    enhanced_starterlist = starterlist.copy()
    if logo_path:
        enhanced_starterlist["logoPath"] = logo_path
        print(f"PDF DEBUG: Logo-Pfad hinzugefügt: {logo_path}")
    else:
        print("PDF DEBUG: Kein Logo verfügbar, ohne Logo fortfahren")

    # Banner und Sponsorenleiste Pfade ermitteln
    banner_sponsor = _get_banner_sponsor_paths(username=username)
    enhanced_starterlist.update(banner_sponsor)
    enhanced_starterlist["spacingTopCm"] = spacing_top_cm
    enhanced_starterlist["spacingBottomCm"] = spacing_bottom_cm
    if spacing_top_cm > 0 or spacing_bottom_cm > 0:
        print(f"PDF DEBUG: Abstände: Oben={spacing_top_cm}cm, Unten={spacing_bottom_cm}cm")

    # Druckoptionen in starterlist einfügen
    if print_options:
        enhanced_starterlist["printOptions"] = print_options
        print(f"PDF DEBUG: Druckoptionen: {print_options}")
    else:
        enhanced_starterlist["printOptions"] = {
            "sponsor_top": False,
            "sponsor_bottom": False,
            "single_sided": False,
            "show_banner": True,
            "show_sponsor_bar": True,
            "show_title": True,
            "show_header": True,
        }

    # Banner/Sponsor Pfade in printOptions eintragen (nach printOptions-Block!)
    enhanced_starterlist["printOptions"]["bannerPath"]  = banner_sponsor.get("bannerPath",  "")
    enhanced_starterlist["printOptions"]["sponsorPath"] = banner_sponsor.get("sponsorPath", "")
    
    try:
        template_file = _find_template_file(template_name)
    except Exception as e:
        raise

    template_path = os.path.join(TEMPLATES_DIR, template_file)

    # dynamisch importieren
    spec = importlib.util.spec_from_file_location("pdf_template_module", template_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    if not hasattr(module, "render"):
        raise AttributeError(f"Template {template_file} hat keine Funktion render(starterlist, filename)")

    # FooterCanvas patchen: banner/sponsorenleiste aus User-Verzeichnis nutzen
    if hasattr(module, "FooterCanvas"):
        banner_path  = enhanced_starterlist["printOptions"].get("bannerPath", "")
        sponsor_path = enhanced_starterlist["printOptions"].get("sponsorPath", "")

        if banner_path:
            # Monkey-patch: Banner-Pfad in __init__ überschreiben
            _orig_init = module.FooterCanvas.__init__
            def _patched_init(self, *args, _bp=banner_path, _orig=_orig_init, **kwargs):
                _orig(self, *args, **kwargs)
                # Banner-Pfad überschreiben wenn User-spezifisch
                if _bp and os.path.exists(_bp):
                    try:
                        pil_img = PILImage.open(_bp)
                        img_w, img_h = pil_img.size
                        dpi   = pil_img.info.get('dpi', (72, 72))
                        dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
                        if dpi_x > 150:
                            disp_w = img_w * 72.0 / dpi_x
                            disp_h = img_h * 72.0 / dpi_x
                        else:
                            disp_w, disp_h = img_w, img_h
                        self.banner_path   = _bp
                        self.banner_width  = A4[0]
                        self.banner_height = A4[0] * (disp_h / disp_w)
                    except:
                        pass
            module.FooterCanvas.__init__ = _patched_init

        if sponsor_path:
            # Monkey-patch: Sponsorenleiste-Pfad in draw_footer überschreiben
            _orig_footer = module.FooterCanvas.draw_footer
            def _patched_footer(self, _sp=sponsor_path, _orig=_orig_footer):
                if not self._print_options.get("show_sponsor_bar", True):
                    return
                if _sp and os.path.exists(_sp):
                    try:
                        from reportlab.lib.utils import ImageReader
                        from reportlab.lib.pagesizes import A4 as _A4
                        from reportlab.lib.units import mm as _mm
                        img    = ImageReader(_sp)
                        w_px, h_px = img.getSize()
                        img_w  = 190 * _mm
                        img_h  = img_w * (h_px / w_px)
                        x = (_A4[0] - img_w) / 2
                        self.drawImage(_sp, x, 4 * _mm, width=img_w, height=img_h,
                                       preserveAspectRatio=True, mask='auto')
                    except:
                        pass
                else:
                    _orig(self)
            module.FooterCanvas.draw_footer = _patched_footer

        # get_sponsor_bar_height patchen für korrekte Randberechnung
        if sponsor_path and hasattr(module, "get_sponsor_bar_height"):
            def _patched_height(_sp=sponsor_path):
                if os.path.exists(_sp):
                    try:
                        from reportlab.lib.utils import ImageReader
                        from reportlab.lib.units import mm as _mm
                        img = ImageReader(_sp)
                        w_px, h_px = img.getSize()
                        return (190 * _mm) * (h_px / w_px)
                    except:
                        return 25 * _mm
                return 0
            module.get_sponsor_bar_height = _patched_height

    # call render mit vollständigem Pfad und erweiterter starterlist
    # Versuche logo_max_width_cm zu übergeben, falls das Template es unterstützt
    try:
        module.render(enhanced_starterlist, output_path, logo_max_width_cm=logo_max_width_cm)
        print(f"PDF DEBUG: Template mit logo_max_width_cm={logo_max_width_cm}cm aufgerufen")
    except TypeError:
        module.render(enhanced_starterlist, output_path)
        print(f"PDF DEBUG: Template ohne logo_max_width_cm Parameter (alte Version)")

    return output_path