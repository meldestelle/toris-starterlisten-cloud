# -*- coding: utf-8 -*-
"""
TORIS Starterlisten - Cloud Version V2
Multi-User Login mit Fallback
"""

import streamlit as st
import requests
import os
import shutil
from pathlib import Path
import tempfile
from datetime import datetime

# ============================================================================
# PAGE CONFIG - MUSS GANZ AM ANFANG SEIN!
# ============================================================================
st.set_page_config(
    page_title="TORIS Starterlisten Generator",
    page_icon="🏇",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# PASSWORTSCHUTZ - MULTI-USER
# ============================================================================

# Authentication State
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = "Standard"

# Login Screen
if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.title("🔒 TORIS Starterlisten")
        st.caption("Bitte einloggen")
        
        # Benutzername (optional)
        username = st.text_input(
            "Benutzername (optional)",
            placeholder="Leer lassen für Standard-Login",
            help="Optional: Benutzername für spezifisches Konto"
        )
        
        # Passwort
        password = st.text_input(
            "Passwort", 
            type="password",
            placeholder="Passwort eingeben"
        )
        
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🔓 Login", type="primary", use_container_width=True):
                # Standard-Passwort (ohne Benutzername)
                default_password = st.secrets.get("APP_PASSWORD", "")
                
                # User-spezifische Passwörter
                users = st.secrets.get("users", {})
                
                login_successful = False
                login_username = "Standard"
                
                # Fall 1: Kein Benutzername → Standard-Passwort prüfen
                if not username or username.strip() == "":
                    if password == default_password and default_password != "":
                        login_successful = True
                        login_username = "Standard"
                
                # Fall 2: Benutzername angegeben → User-Passwort prüfen
                else:
                    username_clean = username.strip().lower()
                    if username_clean in users:
                        if password == users[username_clean]:
                            login_successful = True
                            login_username = username.strip()
                
                # Login erfolgreich?
                if login_successful:
                    st.session_state.authenticated = True
                    st.session_state.username = login_username
                    st.success(f"✅ Login erfolgreich als {login_username}!")
                    st.rerun()
                else:
                    st.error("❌ Falsches Passwort oder Benutzername!")
        
        with col_b:
            if st.button("❌ Abbrechen", use_container_width=True):
                st.info("Zugriff verweigert")
    
    st.stop()

# ============================================================================
# AB HIER NUR FÜR EINGELOGGTE BENUTZER
# ============================================================================

# Session State Defaults
if "pdf_template" not in st.session_state:
    st.session_state.pdf_template = "pdf_standard_logo"
if "word_template" not in st.session_state:
    st.session_state.word_template = "word_standard_logo"
if "round_number" not in st.session_state:
    st.session_state.round_number = 1
if "spacing_top_cm" not in st.session_state:
    st.session_state.spacing_top_cm = 0.0
if "spacing_bottom_cm" not in st.session_state:
    st.session_state.spacing_bottom_cm = 0.0
if "logo_max_width_cm" not in st.session_state:
    st.session_state.logo_max_width_cm = 5.0
if "include_closed" not in st.session_state:
    st.session_state.include_closed = False
if "sponsor_top" not in st.session_state:
    st.session_state.sponsor_top = False
if "sponsor_bottom" not in st.session_state:
    st.session_state.sponsor_bottom = False
if "single_sided" not in st.session_state:
    st.session_state.single_sided = False
if "show_banner" not in st.session_state:
    st.session_state.show_banner = True
if "show_sponsor_bar" not in st.session_state:
    st.session_state.show_sponsor_bar = True
if "show_title" not in st.session_state:
    st.session_state.show_title = True
if "show_header" not in st.session_state:
    st.session_state.show_header = True

from pdf_export import create_pdf
from word_export import create_word

# API Configuration
API_BASE = st.secrets.get("API_BASE", "https://toris-test.portrix.net/api/results/v1")

# Verzeichnisse
BASE_DIR = Path(".")
TEMPLATES_DIR = BASE_DIR / "templates" / "pdf"
WORD_TEMPLATES_DIR = BASE_DIR / "templates" / "word"
LOGOS_DIR = BASE_DIR / "logos"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "toris_output"

TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
WORD_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
LOGOS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ============================================================================
# STYLING
# ============================================================================

def apply_custom_styles():
    st.markdown("""
    <style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: 1400px;
    }
    
    /* Button Styling - kein height forcing */
    .stButton > button {
        border-radius: 6px;
        border: 2px solid #3498db;
        font-weight: 600;
        background: white;
        color: #2c3e50;
        transition: all 0.2s ease;
        width: 100%;
    }
    
    .stButton > button:hover {
        background: #3498db;
        color: white;
        border-color: #2980b9;
        box-shadow: 0 2px 8px rgba(52, 152, 219, 0.3);
    }
    
    .status-badge {
        padding: 0.3rem 0.8rem;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
    }
    
    .status-provisional {
        background: #fff3cd;
        color: #856404;
        border: 1px solid #ffc107;
    }
    
    .status-published {
        background: #d4edda;
        color: #155724;
        border: 1px solid #28a745;
    }
    
    .status-unpublished {
        background: #f8d7da;
        color: #721c24;
        border: 1px solid #dc3545;
    }
    
    .section-header {
        font-size: 1.2rem;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #3498db;
    }
    </style>
    """, unsafe_allow_html=True)

# ============================================================================
# FILE MANAGEMENT FUNCTIONS
# ============================================================================

def get_available_templates():
    templates = []
    if TEMPLATES_DIR.exists():
        for f in TEMPLATES_DIR.glob("*.py"):
            if not f.name.startswith("__"):
                templates.append(f.stem)
    return sorted(templates)

def get_available_word_templates_from_folder():
    """Liest Word-Templates aus templates/word/ Ordner"""
    templates = []
    if WORD_TEMPLATES_DIR.exists():
        for f in WORD_TEMPLATES_DIR.glob("*.py"):
            if not f.name.startswith("__"):
                templates.append(f.stem)
    return sorted(templates)

def get_user_logos_dir():
    """Gibt das Logo-Verzeichnis für den aktuellen User zurück"""
    username = st.session_state.get("username", "Standard")
    if username and username.strip() and username.strip().lower() != "standard":
        user_dir = LOGOS_DIR / username.strip()
        user_dir.mkdir(parents=True, exist_ok=True)
        return user_dir
    return LOGOS_DIR

def _github_api_headers():
    token = st.secrets.get("GITHUB_TOKEN", "")
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

def _github_repo():
    return st.secrets.get("GITHUB_REPO", "meldestelle/toris-starterlisten-cloud")

def _github_logo_path(logo_name):
    """Gibt den Repo-Pfad für ein Logo zurück (z.B. logos/tom/logo.png)"""
    username = st.session_state.get("username", "Standard")
    if username and username.strip() and username.strip().lower() != "standard":
        return f"logos/{username.strip()}/{logo_name}"
    return f"logos/{logo_name}"

def github_upload_logo(logo_name: str, file_bytes: bytes) -> bool:
    """Lädt ein Logo via GitHub API ins Repository hoch"""
    import base64
    repo = _github_repo()
    path = _github_logo_path(logo_name)
    url  = f"https://api.github.com/repos/{repo}/contents/{path}"

    # Prüfe ob Datei schon existiert (für Update SHA benötigt)
    sha = None
    r = requests.get(url, headers=_github_api_headers())
    if r.status_code == 200:
        sha = r.json().get("sha")

    payload = {
        "message": f"Logo upload: {path}",
        "content": base64.b64encode(file_bytes).decode(),
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(url, headers=_github_api_headers(), json=payload)
    if r.status_code in (200, 201):
        print(f"GITHUB: Logo hochgeladen: {path}")
        return True
    else:
        print(f"GITHUB ERROR: Upload fehlgeschlagen: {r.status_code} {r.text}")
        return False

def github_delete_logo(logo_name: str) -> bool:
    """Löscht ein Logo via GitHub API aus dem Repository"""
    repo = _github_repo()
    path = _github_logo_path(logo_name)
    url  = f"https://api.github.com/repos/{repo}/contents/{path}"

    # SHA ermitteln
    r = requests.get(url, headers=_github_api_headers())
    if r.status_code != 200:
        print(f"GITHUB ERROR: Datei nicht gefunden: {path}")
        return False
    sha = r.json().get("sha")

    payload = {"message": f"Logo delete: {path}", "sha": sha}
    r = requests.delete(url, headers=_github_api_headers(), json=payload)
    if r.status_code == 200:
        print(f"GITHUB: Logo gelöscht: {path}")
        return True
    else:
        print(f"GITHUB ERROR: Löschen fehlgeschlagen: {r.status_code} {r.text}")
        return False

def get_available_logos():
    logos = []
    logos_dir = get_user_logos_dir()
    if logos_dir.exists():
        for f in logos_dir.glob("*"):
            if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
                logos.append(f.name)
    return sorted(logos)

def delete_logo(logo_name):
    """Löscht Logo lokal UND im GitHub-Repository"""
    # Lokal löschen
    logo_path = get_user_logos_dir() / logo_name
    if logo_path.exists():
        logo_path.unlink()
    # GitHub löschen
    return github_delete_logo(logo_name)

def delete_template(template_name):
    template_path = TEMPLATES_DIR / f"{template_name}.py"
    if template_path.exists():
        template_path.unlink()
        return True
    return False

def delete_word_template(template_name):
    """Löscht Word-Template aus templates/word/"""
    template_path = WORD_TEMPLATES_DIR / f"{template_name}.py"
    if template_path.exists():
        template_path.unlink()
        return True
    return False

def delete_logo(logo_name):
    logo_path = get_user_logos_dir() / logo_name
    if logo_path.exists():
        logo_path.unlink()
        return True
    return False

def save_uploaded_file(uploaded_file, target_dir):
    target_path = target_dir / uploaded_file.name
    with open(target_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return target_path

def render_file_manager():
    st.markdown('<div class="section-header">📁 Datei-Verwaltung</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📄 PDF Templates", "📝 Word Templates", "🖼️ Logos"])
    
    with tab1:
        st.subheader("Template hochladen")
        uploaded_template = st.file_uploader(
            "PDF Template (.py Datei)",
            type=['py'],
            key=f"template_uploader_{st.session_state.get('template_upload_key', 0)}",
            help="Python-Datei mit PDF-Template-Code hochladen"
        )
        
        if uploaded_template:
            if st.button("✅ Template speichern", key="save_template"):
                try:
                    save_uploaded_file(uploaded_template, TEMPLATES_DIR)
                    st.success(f"✅ Template '{uploaded_template.name}' gespeichert!")
                    # Clear uploader
                    if "template_upload_key" not in st.session_state:
                        st.session_state.template_upload_key = 0
                    st.session_state.template_upload_key += 1
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Fehler beim Speichern: {e}")
        
        st.subheader("Verfügbare Templates")
        templates = get_available_templates()
        
        if templates:
            for template in templates:
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.text(f"📄 {template}.py")
                with col2:
                    if st.session_state.pdf_template == template:
                        st.success("✓ Aktiv")
                    else:
                        if st.button("Aktivieren", key=f"activate_{template}"):
                            st.session_state.pdf_template = template
                            st.rerun()
                with col3:
                    if st.button("🗑️", key=f"del_template_{template}"):
                        if delete_template(template):
                            st.success(f"Template '{template}' gelöscht")
                            st.rerun()
        else:
            st.info("ℹ️ Keine Templates vorhanden.")
    
    with tab2:
        st.subheader("Template hochladen")
        uploaded_word_template = st.file_uploader(
            "Word Template (.py oder .docx Datei)",
            type=['py', 'docx'],
            key=f"word_template_uploader_{st.session_state.get('word_template_upload_key', 0)}",
            help="Python-Datei (.py) oder Word-Vorlage (.docx) hochladen"
        )
        
        if uploaded_word_template:
            if st.button("✅ Template speichern", key="save_word_template"):
                try:
                    save_uploaded_file(uploaded_word_template, WORD_TEMPLATES_DIR)
                    st.success(f"✅ Template '{uploaded_word_template.name}' gespeichert!")
                    # Clear uploader
                    if "word_template_upload_key" not in st.session_state:
                        st.session_state.word_template_upload_key = 0
                    st.session_state.word_template_upload_key += 1
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Fehler beim Speichern: {e}")
        
        st.subheader("Verfügbare Word Templates")
        word_templates = get_available_word_templates_from_folder()
        
        if word_templates:
            for template in word_templates:
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.text(f"📝 {template}.py")
                with col2:
                    if st.session_state.word_template == template:
                        st.success("✓ Aktiv")
                    else:
                        if st.button("Aktivieren", key=f"activate_word_{template}"):
                            st.session_state.word_template = template
                            st.rerun()
                with col3:
                    if st.button("🗑️", key=f"del_word_template_{template}"):
                        if delete_word_template(template):
                            st.success(f"Template '{template}' gelöscht")
                            st.rerun()
        else:
            st.info("ℹ️ Keine Word-Templates vorhanden.")
    
    with tab3:
        st.subheader("Logo hochladen")
        username = st.session_state.get("username", "Standard")
        if username and username.lower() != "standard":
            st.info(f"📁 Logos werden in **logos/{username}/** gespeichert")
        else:
            st.info("📁 Logos werden im gemeinsamen Ordner **logos/** gespeichert")
        uploaded_logo = st.file_uploader(
            "Logo-Datei",
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            key=f"logo_uploader_{st.session_state.get('logo_upload_key', 0)}",
            help="Logo-Bild hochladen"
        )
        
        if uploaded_logo:
            st.image(uploaded_logo, caption=f"Preview: {uploaded_logo.name}", width=200)
            
            if st.button("✅ Logo speichern", key="save_logo"):
                try:
                    file_bytes = uploaded_logo.getbuffer().tobytes()
                    # Debug: Token und Pfad anzeigen
                    token = st.secrets.get("GITHUB_TOKEN", "")
                    repo  = st.secrets.get("GITHUB_REPO", "")
                    path  = _github_logo_path(uploaded_logo.name)
                    st.info(f"🔍 Debug: repo={repo}, path={path}, token={'✅ vorhanden' if token else '❌ FEHLT'}")
                    # Lokal speichern
                    save_uploaded_file(uploaded_logo, get_user_logos_dir())
                    # GitHub speichern
                    import base64
                    url = f"https://api.github.com/repos/{repo}/contents/{path}"
                    sha = None
                    r = requests.get(url, headers=_github_api_headers())
                    if r.status_code == 200:
                        sha = r.json().get("sha")
                    payload = {"message": f"Logo upload: {path}", "content": base64.b64encode(file_bytes).decode()}
                    if sha:
                        payload["sha"] = sha
                    r = requests.put(url, headers=_github_api_headers(), json=payload)
                    if r.status_code in (200, 201):
                        st.success(f"✅ Logo '{uploaded_logo.name}' gespeichert und ins Repository übertragen!")
                    else:
                        st.error(f"❌ GitHub-Upload fehlgeschlagen: {r.status_code} - {r.json().get('message', r.text)}")
                    # Clear uploader
                    if "logo_upload_key" not in st.session_state:
                        st.session_state.logo_upload_key = 0
                    st.session_state.logo_upload_key += 1
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Fehler beim Speichern: {e}")
        
        st.subheader("Verfügbare Logos")
        logos = get_available_logos()
        
        if logos:
            cols = st.columns(4)
            for idx, logo in enumerate(logos):
                with cols[idx % 4]:
                    logo_path = get_user_logos_dir() / logo
                    try:
                        st.image(str(logo_path), caption=logo, width=150)
                        if st.button("🗑️", key=f"del_logo_{logo}"):
                            if delete_logo(logo):
                                st.success(f"Logo '{logo}' gelöscht")
                                st.rerun()
                            else:
                                st.error(f"Logo '{logo}' lokal gelöscht, GitHub-Löschen fehlgeschlagen")
                    except Exception as e:
                        st.error(f"Fehler: {e}")
        else:
            st.info("ℹ️ Keine Logos vorhanden.")

# ============================================================================
# API FUNCTIONS
# ============================================================================

def fetch_shows(api_key, include_closed=False):
    try:
        url = f"{API_BASE}/Shows"
        if include_closed:
            url += "?includeCompletedOrClosed=true"
        headers = {"X-API-Key": api_key} if api_key else {}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        st.error(f"❌ API Fehler: {e}")
        return []

def fetch_competitions(api_key, show_number):
    try:
        url = f"{API_BASE}/Shows/{show_number}/Competitions"
        headers = {"X-API-Key": api_key} if api_key else {}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        st.error(f"❌ Fehler beim Laden der Prüfungen: {e}")
        return []

def fetch_starterlist(api_key, show_number, comp_number, comp_div=None, round_number=1):
    headers = {"X-API-Key": api_key} if api_key else {}
    params = {}
    
    if round_number and round_number != 1:
        params["roundNumber"] = round_number
    
    # Mit Division
    if comp_div:
        url = f"{API_BASE}/Shows/{show_number}/Competitions/{comp_number}/{comp_div}/Starterlist"
        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 404 and round_number > 1:
                # Test ob Runde 1 existiert
                test_response = requests.get(url, headers=headers, timeout=10)
                if test_response.status_code == 200:
                    raise ValueError(f"Runde {round_number} existiert nicht für diese Prüfung")
        except ValueError:
            raise
        except Exception:
            pass
    
    # Ohne Division (Fallback)
    url = f"{API_BASE}/Shows/{show_number}/Competitions/{comp_number}/Starterlist"
    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 404 and round_number > 1:
            # Test ob Runde 1 existiert
            test_response = requests.get(url, headers=headers, timeout=10)
            if test_response.status_code == 200:
                raise ValueError(f"Runde {round_number} existiert nicht für diese Prüfung")
        response.raise_for_status()
        return response.json()
    except ValueError:
        raise
    except Exception as e:
        st.error(f"❌ Fehler beim Laden der Starterliste: {e}")
        return None

def fetch_competition_details(api_key, show_number, comp_number):
    try:
        headers = {"X-API-Key": api_key} if api_key else {}
        url = f"{API_BASE}/Shows/{show_number}/Competitions/{comp_number}"
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return None

def get_status_badge(publishing_status):
    """Erstellt Badge basierend auf Publishing Status"""
    if publishing_status == "PUBLISHED_PROVISIONAL" or publishing_status == 1:
        return '<span class="status-badge status-provisional">⚠️ Vorläufig</span>'
    elif publishing_status == "PUBLISHED_CONFIRMED" or publishing_status == 2:
        return '<span class="status-badge status-published">✅ Final</span>'
    else:
        return '<span class="status-badge status-unpublished">❌ Nicht veröffentlicht</span>'

def enhance_starterlist(starterlist, comp_obj, comp_details):
    if st.session_state.get("selected_show"):
        starterlist["showTitle"] = st.session_state["selected_show"].get("title")
        starterlist["showNumber"] = st.session_state["selected_show"].get("number")
    
    if comp_obj:
        starterlist["competitionTitle"] = comp_obj.get("title")
        starterlist["competitionNumber"] = comp_obj.get("number")
        starterlist["subtitle"] = comp_obj.get("subtitle")
        starterlist["informationText"] = comp_obj.get("informationText")
        starterlist["location"] = comp_obj.get("location")
        starterlist["start"] = comp_obj.get("start") or comp_obj.get("startTime")
    
    if comp_details:
        starterlist["subtitle"] = comp_details.get("subtitle", starterlist.get("subtitle"))
        starterlist["informationText"] = comp_details.get("informationText", starterlist.get("informationText"))
        starterlist["location"] = comp_details.get("location", starterlist.get("location"))
        starterlist["start"] = comp_details.get("start", starterlist.get("start"))
        starterlist["judgingRule"] = comp_details.get("judgingRule")
        
        if comp_details.get("judges"):
            starterlist["judges"] = comp_details.get("judges")
        
        # WICHTIG: dressageTests für Richter-Aufgaben-Zuordnung (402.C etc.)
        if comp_details.get("dressageTests"):
            starterlist["dressageTests"] = comp_details.get("dressageTests")
        
        if comp_details.get("divisions"):
            starterlist["divisions"] = comp_details.get("divisions")
    
    starterlist["roundNumber"] = st.session_state.round_number
    
    return starterlist

# ============================================================================
# MAIN APP
# ============================================================================

apply_custom_styles()

# Logo in Header
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image(str(BASE_DIR / "toris_logo.png"), width=80)
    except:
        st.markdown("🏇")
with col2:
    st.title("TORIS Starterlisten Generator")
    # Zeige angemeldeten Benutzer
    if st.session_state.get("username", "Standard") != "Standard":
        st.caption(f"PDF-Export mit Publishing Status • 👤 Angemeldet als: {st.session_state.username}")
    else:
        st.caption("PDF-Export mit Publishing Status")

# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.header("⚙️ Einstellungen")
    
    # API-Key
    api_key = st.text_input(
        "API Key",
        type="password",
        help="TORIS API Key für Zugriff"
    )
    
    if not api_key:
        st.warning("⚠️ Bitte API Key eingeben!")
        st.stop()
    
    st.markdown("---")
    
    # Template-Auswahl
    available_templates = get_available_templates()
    if available_templates:
        selected_template = st.selectbox(
            "PDF Template",
            available_templates,
            index=available_templates.index(st.session_state.pdf_template) 
                  if st.session_state.pdf_template in available_templates else 0
        )
        st.session_state.pdf_template = selected_template
    else:
        st.warning("⚠️ Keine PDF-Templates vorhanden!")
    
    # Word Template-Auswahl (aus templates/word/ Ordner)
    available_word_templates = get_available_word_templates_from_folder()
    if available_word_templates:
        selected_word_template = st.selectbox(
            "Word Template",
            available_word_templates,
            index=available_word_templates.index(st.session_state.word_template) 
                  if st.session_state.word_template in available_word_templates else 0
        )
        st.session_state.word_template = selected_word_template
    else:
        st.warning("⚠️ Keine Word-Templates vorhanden!")
        st.info("💡 Tipp: Templates in Tab 3 'Verwaltung' hochladen")
    
    # Logo-Breite
    st.subheader("🖼️ Logo")
    st.session_state.logo_max_width_cm = st.number_input(
        "Breite (cm)",
        min_value=1.0,
        max_value=10.0,
        value=st.session_state.logo_max_width_cm,
        step=0.5,
        format="%.1f",
        help="Maximale Breite des Logos (Standard: 5cm)"
    )

    # Druckoptionen
    st.markdown("---")
    st.subheader("🖨️ Druckoptionen")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.show_header = st.checkbox(
            "Prüfungskopf", value=st.session_state.show_header, key="show_header_cb"
        )
        st.session_state.show_banner = st.checkbox(
            "Banner", value=st.session_state.show_banner, key="show_banner_cb"
        )
        st.session_state.show_sponsor_bar = st.checkbox(
            "Sponsorenleiste", value=st.session_state.show_sponsor_bar, key="show_sponsor_bar_cb"
        )
        st.session_state.show_title = st.checkbox(
            "Show-Titel", value=st.session_state.show_title, key="show_title_cb"
        )
    with col2:
        st.session_state.sponsor_top = st.checkbox(
            "Sponsorenpapier OBEN", value=st.session_state.sponsor_top, key="sponsor_top_cb"
        )
        st.session_state.sponsor_bottom = st.checkbox(
            "Sponsorenpapier UNTEN", value=st.session_state.sponsor_bottom, key="sponsor_bottom_cb"
        )
        st.session_state.single_sided = st.checkbox(
            "Einseitiger Druck", value=st.session_state.single_sided, key="single_sided_cb"
        )

    # Abstände — sichtbar wenn Sponsorenpapier aktiv oder Prüfungskopf ausgeblendet
    needs_spacing = (
        st.session_state.sponsor_top
        or st.session_state.sponsor_bottom
        or not st.session_state.show_header
    )
    if needs_spacing:
        st.subheader("📏 Abstände")
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.spacing_top_cm = st.number_input(
                "Oben (cm)",
                min_value=0.0,
                max_value=10.0,
                value=st.session_state.spacing_top_cm,
                step=0.5,
                key="spacing_top_input"
            )
        with col2:
            st.session_state.spacing_bottom_cm = st.number_input(
                "Unten (cm)",
                min_value=0.0,
                max_value=10.0,
                value=st.session_state.spacing_bottom_cm,
                step=0.5,
                key="spacing_bottom_input"
            )
    else:
        # Standardwerte wenn keine Abstände nötig
        st.session_state.spacing_top_cm = 0.0
        st.session_state.spacing_bottom_cm = 0.0

# ============================================================================
# TAB LAYOUT
# ============================================================================

tab1, tab2, tab3 = st.tabs(["🎯 Veranstaltung", "📋 Prüfung & Export", "⚙️ Verwaltung"])

# ============================================================================
# TAB 1: VERANSTALTUNG
# ============================================================================

with tab1:
    st.markdown('<div class="section-header">🎯 Veranstaltung auswählen</div>', unsafe_allow_html=True)
    
    # Checkbox
    include_closed = st.checkbox(
        "Geschlossene Veranstaltungen anzeigen", 
        value=st.session_state.include_closed
    )
    
    # Button und Status - mehr Platz für Status
    col1, col2 = st.columns([1, 2])
    
    with col1:
        if st.button("🔄 Veranstaltungen laden", type="primary", use_container_width=True):
            with st.spinner("Lade Veranstaltungen..."):
                shows = fetch_shows(api_key, include_closed)
                if shows:
                    st.session_state.shows = shows
                    st.session_state.include_closed = include_closed
                    st.rerun()
                else:
                    st.error("❌ Keine Veranstaltungen verfügbar!")
    
    with col2:
        if "shows" in st.session_state and st.session_state.shows:
            st.success(f"✅ {len(st.session_state.shows)} Veranstaltungen geladen!")
        else:
            st.info("👆 Bitte erst 'Veranstaltungen laden' klicken")
    
    # Zeige Veranstaltungen wenn geladen
    if "shows" in st.session_state and st.session_state.shows:
        shows = st.session_state.shows
        
        # Veranstaltung auswählen und Gewählt-Status nebeneinander
        col1, col2 = st.columns([5, 1])
        
        with col1:
            show_options = {f"{s['title']} ({s['number']})": s for s in shows}
            selected_show_label = st.selectbox(
                "Veranstaltung auswählen",
                list(show_options.keys()),
                key="selected_show_box"
            )
        
        with col2:
            if selected_show_label:
                st.session_state.selected_show = show_options[selected_show_label]
                st.markdown("<br>", unsafe_allow_html=True)  # Spacing for alignment
                st.success("✅")
        
        if selected_show_label:
            # Prüfungen laden - Button und Status nebeneinander
            col1, col2 = st.columns([1, 2])
            
            with col1:
                if st.button("🔄 Prüfungen laden", type="primary", key="load_comps_tab1", use_container_width=True):
                    with st.spinner("Lade Prüfungen..."):
                        show_num = st.session_state.selected_show.get("number")
                        competitions = fetch_competitions(api_key, show_num)
                        if competitions:
                            st.session_state.competitions = competitions
                            st.rerun()
                        else:
                            st.error("❌ Keine Prüfungen gefunden!")
            
            with col2:
                if "competitions" in st.session_state and st.session_state.competitions:
                    st.success(f"✅ {len(st.session_state.competitions)} Prüfungen geladen und verfügbar")

# ============================================================================
# TAB 2: PRÜFUNG & EXPORT
# ============================================================================

with tab2:
    if "selected_show" not in st.session_state:
        st.warning("⚠️ Bitte erst Veranstaltung in Tab 1 auswählen!")
    else:
        show_number = st.session_state.selected_show.get("number")
        
        st.markdown('<div class="section-header">🎯 Prüfung auswählen</div>', unsafe_allow_html=True)
        
        # Zeige Prüfungen wenn geladen
        if "competitions" in st.session_state and st.session_state.competitions:
            competitions = st.session_state.competitions
            
            # Prüfungs-Dictionary mit Divisionen
            comp_options = {}
            for comp in competitions:
                comp_number = comp.get("number")
                comp_title = comp.get("title")
                divisions = comp.get("divisions", [])
                
                if divisions:
                    for div in divisions:
                        div_number = div.get("number")
                        key = f"{comp_number} - {comp_title} Abt. {div_number}"
                        comp_options[key] = (comp, div_number)
                else:
                    key = f"{comp_number} - {comp_title}"
                    comp_options[key] = (comp, None)
            
            # Prüfung und Umlauf nebeneinander
            col1, col2 = st.columns([4, 1])
            
            with col1:
                selected_comp_str = st.selectbox(
                    "Prüfung auswählen",
                    list(comp_options.keys()),
                    key="selected_comp_box"
                )
            
            with col2:
                st.session_state.round_number = st.selectbox(
                    "Umlauf",
                    options=[1, 2, 3, 4, 5],
                    index=st.session_state.round_number - 1
                )
            
            comp_tuple = comp_options[selected_comp_str]
            comp_obj = comp_tuple[0]
            comp_div = comp_tuple[1] if len(comp_tuple) > 1 else None
            
            # Prüfe ob sich die Auswahl geändert hat
            current_selection = f"{comp_obj.get('number')}_{comp_div}_{st.session_state.round_number}"
            if "last_selection" not in st.session_state:
                st.session_state.last_selection = current_selection
            
            # Wenn sich die Auswahl geändert hat, lösche die alte Starterliste
            if st.session_state.last_selection != current_selection:
                if "starterlist" in st.session_state:
                    del st.session_state.starterlist
                st.session_state.last_selection = current_selection
            
            # ================================================================
            # STARTERLISTE LADEN
            # ================================================================
            
            st.markdown('<div class="section-header">📋 Starterliste</div>', unsafe_allow_html=True)
            
            # Button und Status - mehr Platz für Status
            col1, col2 = st.columns([1, 2])
            
            with col1:
                if st.button("🔄 Starterliste laden", type="primary", use_container_width=True):
                    try:
                        with st.spinner("Lade Daten..."):
                            starterlist = fetch_starterlist(
                                api_key,
                                show_number,
                                comp_obj.get("number"),
                                comp_div,
                                st.session_state.round_number
                            )
                            
                            if starterlist:
                                comp_details = fetch_competition_details(
                                    api_key, 
                                    show_number, 
                                    comp_obj.get("number")
                                )
                                
                                starterlist = enhance_starterlist(starterlist, comp_obj, comp_details)
                                st.session_state.starterlist = starterlist
                                st.rerun()
                    except ValueError as ve:
                        st.error(f"❌ {str(ve)}")
                    except Exception as e:
                        st.error(f"❌ Fehler: {e}")
            
            with col2:
                if "starterlist" in st.session_state:
                    st.success("✅ Starterliste geladen!")
            
            # Zeige Starterliste
            if "starterlist" in st.session_state:
                starterlist = st.session_state.starterlist
                
                # Publishing Status Badge
                publishing_status = starterlist.get("publishingStatus", "NOT_PUBLISHED")
                status_badge = get_status_badge(publishing_status)
                
                col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 1, 1, 1, 2])
                with col1:
                    st.metric("Veranstaltung", starterlist.get("showNumber", "N/A"))
                with col2:
                    st.metric("Prüfung", starterlist.get("competitionNumber", "N/A"))
                with col3:
                    div_display = starterlist.get("divisionNumber", "-")
                    if div_display is None or div_display == 0:
                        div_display = "-"
                    st.metric("Abteilung", div_display)
                with col4:
                    st.metric("Umlauf", st.session_state.round_number)
                with col5:
                    st.metric("Starter", len(starterlist.get("starters", [])))
                with col6:
                    st.markdown(f"**Status:**<br>{status_badge}", unsafe_allow_html=True)
                
                # Starter-Tabelle in Expander (eingeklappt)
                entries = starterlist.get("starters", [])
                if entries:
                    total_count = len(entries)
                    preview_count = min(10, total_count)
                    
                    with st.expander(f"👁️ Starter-Vorschau (erste {preview_count} von {total_count})", expanded=False):
                        for entry in entries[:preview_count]:
                            athlete = entry.get("athlete", {})
                            horses = entry.get("horses", [])
                            horse_name = horses[0].get("name", "N/A") if horses else "N/A"
                            st.text(f"{entry.get('startNumber', 'N/A')} | {athlete.get('name', 'N/A')} - {horse_name}")
                        
                        if total_count > preview_count:
                            st.info(f"... und {total_count - preview_count} weitere Starter")
                
                # ============================================================
                # TEMPLATE-SPEZIFISCHE EINGABEN (Derby / Pferdewechsel)
                # ============================================================
                selected_tpl = st.session_state.pdf_template

                is_derby        = selected_tpl in ("pdf_dre_derby_cloud", "pdf_dre_derby_int_cloud")
                is_pfwechsel    = selected_tpl in ("pdf_dre_pferdewechsel_cloud", "pdf_dre_pferdewechsel_int_cloud")

                if is_derby:
                    st.markdown('<div class="section-header">🏇 Derby – Zeiten</div>', unsafe_allow_html=True)
                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        derby_begin = st.text_input(
                            "Prüfungsbeginn (z.B. 10:00 Uhr)",
                            value=st.session_state.get("derby_begin", ""),
                            key="derby_begin_input"
                        )
                        st.session_state.derby_begin = derby_begin
                    with col_d2:
                        derby_final = st.text_input(
                            "Finale (z.B. 11:30 Uhr)",
                            value=st.session_state.get("derby_final", ""),
                            key="derby_final_input"
                        )
                        st.session_state.derby_final = derby_final

                if is_pfwechsel:
                    st.markdown('<div class="section-header">🐴 Pferdewechsel – Konfiguration</div>', unsafe_allow_html=True)

                    # Richtverfahren aus Starterliste ermitteln
                    comp_data    = starterlist.get("competition") or {}
                    judging_rule = comp_data.get("judgingRule") or starterlist.get("judgingRule") or ""

                    # Die ersten 3 Starter für Namen
                    pw_starters  = starterlist.get("starters") or []
                    pw_entries   = []
                    for s in pw_starters[:3]:
                        ath   = s.get("athlete") or {}
                        hrs   = s.get("horses")  or []
                        rname = ath.get("name",  f"Reiter {s.get('startNumber','?')}")
                        pname = hrs[0].get("name", f"Pferd {s.get('startNumber','?')}") if hrs else "—"
                        pw_entries.append((rname, pname))

                    reiter_namen = [e[0] for e in pw_entries]
                    pferd_namen  = [e[1] for e in pw_entries]

                    # Frage 1: Ergebnis auf eigenem Pferd eintragen?
                    own_horse_ja = st.radio(
                        "Ergebnis auf eigenem Pferd eintragen?",
                        ["Nein", "Ja"],
                        horizontal=True,
                        key="pw_own_horse"
                    ) == "Ja"

                    own_results  = {}
                    assignment   = list(zip(reiter_namen, pferd_namen))  # Standard-Diagonale

                    if own_horse_ja and len(pw_entries) == 3:
                        st.markdown("**Zuordnung: Welches ist das eigene Pferd?**")
                        new_assignment = []
                        for i, rname in enumerate(reiter_namen):
                            chosen = st.radio(
                                rname,
                                pferd_namen,
                                index=i,
                                horizontal=True,
                                key=f"pw_assign_{i}"
                            )
                            new_assignment.append((rname, chosen))
                        assignment = new_assignment

                        st.markdown("**Ergebnis auf eigenem Pferd:**")
                        rule_up = str(judging_rule).upper().replace(" ","").replace(".","").replace(",","")
                        unit    = "" if "402A" in rule_up else " %"
                        for i, (rname, pname) in enumerate(assignment):
                            val = st.text_input(
                                f"{rname} auf {pname}{unit}",
                                value=st.session_state.get(f"pw_score_{i}", ""),
                                key=f"pw_score_input_{i}"
                            )
                            st.session_state[f"pw_score_{i}"] = val
                            if val.strip():
                                own_results[i] = val.strip().replace(",", ".")

                    # Pferdedetails ab Starter 4?
                    show_details = st.radio(
                        "Pferdedetails ab Starter 4 anzeigen?",
                        ["Nein – nur Pferdename", "Ja – vollständige Details"],
                        horizontal=True,
                        key="pw_show_details"
                    ) == "Ja – vollständige Details"

                    # Konfiguration in session_state speichern
                    st.session_state.pw_config = {
                        "own_results":              own_results,
                        "assignment":               assignment,
                        "show_horse_details_from_4": show_details,
                    }

                # ============================================================
                # PDF EXPORT
                # ============================================================
                
                st.markdown('<div class="section-header">📄 PDF Export</div>', unsafe_allow_html=True)
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    if st.button("🎨 PDF erstellen", type="primary", use_container_width=True):
                        if not available_templates:
                            st.error("❌ Keine Templates vorhanden!")
                        else:
                            try:
                                with st.spinner("Erstelle PDF..."):
                                    comp_number = starterlist.get('competitionNumber', '00')
                                    div_number = starterlist.get('divisionNumber', 0)
                                    
                                    try:
                                        comp_formatted = f"{int(comp_number):02d}"
                                    except (ValueError, TypeError):
                                        comp_formatted = str(comp_number).zfill(2)
                                    
                                    div_formatted = f"{int(div_number)}" if div_number else "0"
                                    
                                    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
                                    pdf_filename = f"{comp_formatted}{div_formatted}_start_{timestamp}.pdf"
                                    
                                    print_options = {
                                        "sponsor_top":      st.session_state.get("sponsor_top", False),
                                        "sponsor_bottom":   st.session_state.get("sponsor_bottom", False),
                                        "single_sided":     st.session_state.get("single_sided", False),
                                        "show_banner":      st.session_state.get("show_banner", True),
                                        "show_sponsor_bar": st.session_state.get("show_sponsor_bar", True),
                                        "show_title":       st.session_state.get("show_title", True),
                                        "show_header":      st.session_state.get("show_header", True),
                                    }

                                    # Template-spezifische Konfigurationen eintragen
                                    selected_tpl = st.session_state.pdf_template
                                    if selected_tpl in ("pdf_dre_derby_cloud", "pdf_dre_derby_int_cloud"):
                                        starterlist["derby_config"] = {
                                            "begin_time": st.session_state.get("derby_begin", ""),
                                            "final_time": st.session_state.get("derby_final", ""),
                                        }
                                    if selected_tpl in ("pdf_dre_pferdewechsel_cloud", "pdf_dre_pferdewechsel_int_cloud"):
                                        starterlist["derby_config"] = st.session_state.get("pw_config", {})

                                    pdf_path = create_pdf(
                                        starterlist,
                                        pdf_filename,
                                        st.session_state.pdf_template,
                                        st.session_state.spacing_top_cm,
                                        st.session_state.spacing_bottom_cm,
                                        st.session_state.logo_max_width_cm,
                                        output_dir=str(OUTPUT_DIR),
                                        print_options=print_options,
                                        username=st.session_state.get("username")
                                    )
                                    
                                    if not pdf_path or not os.path.exists(pdf_path):
                                        st.error("❌ PDF konnte nicht erstellt werden!")
                                    else:
                                        st.success("✅ PDF erfolgreich erstellt!")
                                        
                                        with open(pdf_path, "rb") as f:
                                            st.download_button(
                                                label="📥 PDF herunterladen",
                                                data=f,
                                                file_name=pdf_filename,
                                                mime="application/pdf",
                                                type="primary",
                                                use_container_width=True
                                            )
                            
                            except Exception as e:
                                st.error(f"❌ Fehler: {e}")
                                import traceback
                                st.code(traceback.format_exc())
                
                with col2:
                    st.markdown(f"**Template:** {st.session_state.pdf_template}")
                    st.markdown(f"**Umlauf:** {st.session_state.round_number}")
                
                # ============================================================
                # WORD EXPORT
                # ============================================================
                
                st.markdown('<div class="section-header">📝 Word Export</div>', unsafe_allow_html=True)
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    if st.button("📝 Word erstellen", type="secondary", use_container_width=True):
                        try:
                            with st.spinner("Erstelle Word-Dokument..."):
                                comp_number = starterlist.get('competitionNumber', '00')
                                div_number = starterlist.get('divisionNumber', 0)
                                
                                try:
                                    comp_formatted = f"{int(comp_number):02d}"
                                except (ValueError, TypeError):
                                    comp_formatted = str(comp_number).zfill(2)
                                
                                div_formatted = f"{int(div_number)}" if div_number else "0"
                                
                                timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
                                word_filename = f"{comp_formatted}{div_formatted}_start_{timestamp}.docx"
                                word_path = str(OUTPUT_DIR / word_filename)
                                
                                word_print_options = {
                                    "sponsor_top":        st.session_state.get("sponsor_top", False),
                                    "sponsor_bottom":     st.session_state.get("sponsor_bottom", False),
                                    "single_sided":       st.session_state.get("single_sided", False),
                                    "show_banner":        st.session_state.get("show_banner", True),
                                    "show_sponsor_bar":   st.session_state.get("show_sponsor_bar", True),
                                    "show_title":         st.session_state.get("show_title", True),
                                    "show_header":        st.session_state.get("show_header", True),
                                    "spacing_top_cm":     st.session_state.get("spacing_top_cm", 3.0),
                                    "spacing_bottom_cm":  st.session_state.get("spacing_bottom_cm", 2.0),
                                }
                                word_path = create_word(
                                    starterlist,
                                    st.session_state.word_template,
                                    word_path,
                                    logos_enabled=True,
                                    print_options=word_print_options,
                                    logo_max_width_cm=st.session_state.get("logo_max_width_cm", 5.0),
                                    username=st.session_state.get("username")
                                )
                                
                                if not word_path or not os.path.exists(word_path):
                                    st.error("❌ Word-Dokument konnte nicht erstellt werden!")
                                else:
                                    st.success("✅ Word-Dokument erfolgreich erstellt!")
                                    
                                    with open(word_path, "rb") as f:
                                        st.download_button(
                                            label="📥 Word herunterladen",
                                            data=f,
                                            file_name=word_filename,
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                            type="secondary",
                                            use_container_width=True
                                        )
                        
                        except Exception as e:
                            st.error(f"❌ Fehler: {e}")
                            import traceback
                            st.code(traceback.format_exc())
                
                with col2:
                    st.markdown(f"**Template:** {st.session_state.word_template}")
                    st.markdown(f"**Umlauf:** {st.session_state.round_number}")
            
            else:
                st.info("👆 Bitte erst Starterliste laden")
        
        else:
            st.info("👆 Bitte erst Prüfungen in Tab 1 laden")

# ============================================================================
# TAB 3: VERWALTUNG
# ============================================================================

with tab3:
    render_file_manager()

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #7f8c8d; font-size: 0.85rem;">
    <strong>TORIS Starterlisten Generator</strong> - Cloud Version V2
</div>
""", unsafe_allow_html=True)
