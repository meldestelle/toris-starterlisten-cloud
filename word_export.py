# word_export.py - Erweitert um prüfungsspezifisches Logo-System
import importlib
import os
import json

def determine_logo_path(starterlist):
    """
    Bestimmt den Logo-Pfad basierend auf XXY-Schema:
    XX = Prüfungsnummer (2-stellig mit führender 0)
    Y = Abteilungsnummer (1-stellig, 0 wenn keine Abteilung)
    
    Beispiele:
    - Prüfung 17, Abt. 2 → logos/172.png
    - Prüfung 10, keine Abt. → logos/100.png
    - Fallback → logos/logo.png
    """
    try:
        # Standard-Fallback
        default_logo = "logos/logo.png"
        
        # Prüfungsnummer ermitteln
        comp_number = starterlist.get("competitionNumber")
        if not comp_number:
            return default_logo if os.path.exists(default_logo) else None
            
        # Zu Integer konvertieren
        try:
            comp_num_int = int(comp_number)
        except (ValueError, TypeError):
            return default_logo if os.path.exists(default_logo) else None
        
        # Abteilungsnummer ermitteln
        division_num = 0  # Default: keine Abteilung
        division_number = starterlist.get("divisionNumber")
        
        if division_number:
            try:
                division_num = int(division_number)
            except (ValueError, TypeError):
                division_num = 0
        
        # XXY-Format erstellen: XX = Prüfung (2-stellig), Y = Abteilung (1-stellig)
        xxy_code = f"{comp_num_int:02d}{division_num}"
        specific_logo = f"logos/{xxy_code}.png"
        
        print(f"WORD EXPORT DEBUG: Prüfung {comp_num_int}, Abt. {division_num} → Logo: {specific_logo}")
        
        # Prüfungsspezifisches Logo prüfen
        if os.path.exists(specific_logo):
            print(f"WORD EXPORT DEBUG: Spezifisches Logo gefunden: {specific_logo}")
            return specific_logo
        
        # Fallback: Standard-Logo
        if os.path.exists(default_logo):
            print(f"WORD EXPORT DEBUG: Standard-Logo verwendet: {default_logo}")
            return default_logo
        
        print("WORD EXPORT DEBUG: Kein Logo gefunden")
        return None
        
    except Exception as e:
        print(f"WORD EXPORT DEBUG: Fehler bei Logo-Bestimmung: {e}")
        return "logos/logo.png" if os.path.exists("logos/logo.png") else None

def create_word(starterlist: dict, template_name: str, filename: str, logos_enabled: bool = True) -> str:
    """
    Erstellt ein Word-Dokument basierend auf der Starterliste und dem gewählten Template
    
    Args:
        starterlist: Dictionary mit allen Wettkampfdaten
        template_name: Name des zu verwendenden Templates
        filename: Zieldateiname
        logos_enabled: Ob Logos verwendet werden sollen (Standard: True)
        
    Returns:
        str: Pfad zur erstellten Datei
    """
    print(f"WORD EXPORT DEBUG: create_word aufgerufen mit template_name='{template_name}', logos_enabled={logos_enabled}")
    
    # Logo-Pfad bestimmen und in starterlist einfügen
    logo_path = None
    if logos_enabled:
        logo_path = determine_logo_path(starterlist)
        
    # Logo-Pfad in starterlist einfügen für Template-Zugriff
    starterlist_with_logo = starterlist.copy()
    starterlist_with_logo["logoPath"] = logo_path
    
    print(f"WORD EXPORT DEBUG: logoPath gesetzt auf: {logo_path}")
    
    # Template-Mapping erweitert um Logo-Varianten UND Stream-Templates
    template_mapping = {
        # Standard-Templates (ohne Logo)
        "word_standard": "templates.word.word_standard",
        "word_details_modern": "templates.word.word_standard",
        "word_abstammung": "templates.word.word_abstammung", 
        "word_dre_3": "templates.word.word_dre_3",
        "word_dre_5": "templates.word.word_dre_5",
        
        # Logo-Varianten
        "word_standard_logo": "templates.word.word_standard_logo",
        "word_details_modern_logo": "templates.word.word_standard_logo",
        "word_abstammung_logo": "templates.word.word_abstammung_logo",
        "word_dre_3_logo": "templates.word.word_dre_3_logo", 
        "word_dre_5_logo": "templates.word.word_dre_5_logo",
        "word_dre_402c_logo": "templates.word_dre_402c_logo",
        
        # International Style (mit Flaggen)
        "word_int": "templates.word.word_int",
        "word_int_abstammung": "templates.word.word_int_abstammung",
        
        # National Style (deutsche Texte)
        "word_nat": "templates.word.word_nat",
    }
    
    # Template-Modul bestimmen
    template_module_path = template_mapping.get(template_name)
    
    if not template_module_path:
        # Fallback: versuche Standard-Template
        print(f"WORD EXPORT WARNING: Template '{template_name}' nicht gefunden, verwende word_standard")
        template_module_path = "templates.word.word_standard"
    
    print(f"WORD EXPORT DEBUG: Verwende Template-Modul: {template_module_path}")
    
    try:
        # Template-Modul dynamisch laden - Alternativer Ansatz
        import sys
        import os
        
        # Template-Pfad zum Python-Pfad hinzufügen
        # Stream-Templates liegen in templates/stream
        if template_module_path.startswith("templates.stream"):
            template_dir = os.path.join(os.getcwd(), "templates", "stream")
        else:
            template_dir = os.path.join(os.getcwd(), "templates", "word")
        
        if template_dir not in sys.path:
            sys.path.insert(0, template_dir)
        
        # Template-Modul-Name extrahieren
        module_name = template_module_path.split('.')[-1]  # z.B. "word_standard_logo"
        
        # Direkter Import des Moduls
        template_module = importlib.import_module(module_name)
        
        # render-Funktion aufrufen
        result_path = template_module.render(starterlist_with_logo, filename)
        
        print(f"WORD EXPORT DEBUG: Word-Dokument erfolgreich erstellt: {result_path}")
        return result_path
        
    except Exception as e:
        print(f"WORD EXPORT ERROR: Fehler beim Erstellen des Word-Dokuments: {e}")
        raise

# Verfügbare Templates für UI-Auswahl
def get_available_word_templates():
    """Gibt eine Liste der verfügbaren Word-Templates zurück"""
    return [
        # Standard-Templates (ohne Logo) 
        ("word_standard", "Standard (6-Spalten kompakt)"),
        ("word_abstammung", "Abstammung (5-Spalten mit Details)"),
        ("word_dre_3", "Dressur 3-Richter (7-Spalten)"),
        ("word_dre_5", "Dressur 5-Richter (9-Spalten)"),
        
        # Logo-Varianten
        ("word_standard_logo", "Standard mit Logo (6-Spalten kompakt)"),
        ("word_abstammung_logo", "Abstammung mit Logo (5-Spalten mit Details)"), 
        ("word_dre_3_logo", "Dressur 3-Richter mit Logo (7-Spalten)"),
        ("word_dre_5_logo", "Dressur 5-Richter mit Logo (9-Spalten)"),
        ("word_dre_402c_logo", "Dressur 402,C mit Logo"),
        
        # International & National Styles
        ("word_int", "International Style (mit Flaggen)"),
        ("word_nat", "National Style (deutsche Texte mit Flaggen)"),
        ("word_int_abstammung", "International Abstammung (mit Flaggen)"),
    ]