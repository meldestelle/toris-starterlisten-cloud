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

TEMPLATES_DIR = os.path.join("templates", "pdf")
OUTPUT_DIR = "Ausgabe"  # Standard Ausgabeordner (kann überschrieben werden)

def _ensure_output_dir(output_dir=None):
    """Stellt sicher, dass das Ausgabe-Verzeichnis existiert"""
    target_dir = output_dir if output_dir else OUTPUT_DIR
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
    return target_dir

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

def _get_competition_logo_path(starterlist: dict) -> str:
    """
    Ermittelt den Pfad zum prüfungsspezifischen Logo basierend auf XXY-Schema
    XX = Competition (zweistellig), Y = Division (einstellig, 0 wenn keine Abteilung)
    Gibt None zurück wenn kein Logo gefunden wird (kein Fehler)
    """
    comp_number = starterlist.get("competitionNumber")
    div_number = starterlist.get("divisionNumber")
    
    if not comp_number:
        # Fallback auf Standard-Logo prüfen
        fallback_path = os.path.join("logos", "logo.png")
        if os.path.exists(fallback_path):
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Standard-Logo gefunden, ohne Logo fortfahren")
            return None
    
    try:
        # Competition zweistellig formatieren
        comp_formatted = f"{int(comp_number):02d}"
    except (ValueError, TypeError):
        comp_formatted = str(comp_number).zfill(2)
    
    # Division einstellig formatieren (0 wenn keine Abteilung)
    if div_number:
        try:
            div_formatted = f"{int(div_number)}"
        except (ValueError, TypeError):
            div_formatted = str(div_number)
    else:
        div_formatted = "0"
    
    # Logo-Dateiname nach XXY-Schema
    logo_filename = f"{comp_formatted}{div_formatted}.png"
    logo_path = os.path.join("logos", logo_filename)
    
    print(f"PDF DEBUG: Suche Logo: {logo_path}")
    
    # Prüfen ob spezifisches Logo existiert
    if os.path.exists(logo_path):
        print(f"PDF DEBUG: Prüfungsspezifisches Logo gefunden: {logo_path}")
        return logo_path
    else:
        # Fallback auf Standard-Logo
        fallback_path = os.path.join("logos", "logo.png")
        if os.path.exists(fallback_path):
            print(f"PDF DEBUG: Verwende Standard-Logo: {fallback_path}")
            return fallback_path
        else:
            print("PDF DEBUG: Kein Logo gefunden, ohne Logo fortfahren")
            return None

def create_pdf(starterlist: dict, filename: str, template_name: str, spacing_top_cm: float = 0, spacing_bottom_cm: float = 0, logo_max_width_cm: float = 5.0, output_dir: str = None):
    """
    Lädt das angegebene Template-Modul aus templates/pdf und ruft dessen render(starterlist, filename) auf.
    Alle Dateien werden im output_dir (oder 'Ausgabe'-Ordner) erstellt.
    Erweitert starterlist um prüfungsspezifischen Logo-Pfad.
    
    Für Templates die mit "liste_" beginnen werden spacing_top_cm und spacing_bottom_cm
    als Abstände oben und unten verwendet.
    
    Für Templates mit Logo-Unterstützung wird logo_max_width_cm als maximale Logo-Breite verwendet.
    """
    # Ausgabe-Ordner sicherstellen
    target_output_dir = _ensure_output_dir(output_dir)
    
    # Vollständigen Pfad für Ausgabedatei erstellen
    output_path = os.path.join(target_output_dir, filename)
    
    # Prüfungsspezifisches Logo ermitteln
    logo_path = _get_competition_logo_path(starterlist)
    
    # Starterlist um Logo-Information erweitern (nur wenn Logo vorhanden)
    enhanced_starterlist = starterlist.copy()
    if logo_path:
        enhanced_starterlist["logoPath"] = logo_path
        print(f"PDF DEBUG: Logo-Pfad hinzugefügt: {logo_path}")
    else:
        print("PDF DEBUG: Kein Logo verfügbar, ohne Logo fortfahren")
    
    # Für liste_ Templates: Abstände hinzufügen
    if template_name.startswith("liste_"):
        enhanced_starterlist["spacingTopCm"] = spacing_top_cm
        enhanced_starterlist["spacingBottomCm"] = spacing_bottom_cm
        print(f"PDF DEBUG: Liste-Template erkannt - Abstände: Oben={spacing_top_cm}cm, Unten={spacing_bottom_cm}cm")
    
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

    # call render mit vollständigem Pfad und erweiterter starterlist
    # Versuche logo_max_width_cm zu übergeben, falls das Template es unterstützt
    try:
        module.render(enhanced_starterlist, output_path, logo_max_width_cm=logo_max_width_cm)
        print(f"PDF DEBUG: Template mit logo_max_width_cm={logo_max_width_cm}cm aufgerufen")
    except TypeError as te:
        # Fallback für alte Templates ohne logo_max_width_cm Parameter
        try:
            module.render(enhanced_starterlist, output_path)
            print(f"PDF DEBUG: Template ohne logo_max_width_cm Parameter (alte Version)")
        except Exception as e:
            print(f"PDF ERROR: Fehler im Template (ohne logo_max_width_cm): {e}")
            import traceback
            traceback.print_exc()
            raise
    except Exception as e:
        print(f"PDF ERROR: Fehler im Template: {e}")
        import traceback
        traceback.print_exc()
        raise
    
    # Prüfe ob Datei wirklich existiert
    if not os.path.exists(output_path):
        raise FileNotFoundError(f"PDF wurde nicht erstellt: {output_path}")
    
    # WICHTIG: Pfad zurückgeben für Download
    return output_path
