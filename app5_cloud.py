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

from pdf_export import create_pdf

# API Configuration
API_BASE = st.secrets.get("API_BASE", "https://toris-test.portrix.net/api/results/v1")

# Verzeichnisse
BASE_DIR = Path(".")
TEMPLATES_DIR = BASE_DIR / "templates" / "pdf"
LOGOS_DIR = BASE_DIR / "logos"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "toris_output"

TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
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

def get_available_logos():
    logos = []
    if LOGOS_DIR.exists():
        for f in LOGOS_DIR.glob("*"):
            if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
                logos.append(f.name)
    return sorted(logos)

def delete_template(template_name):
    template_path = TEMPLATES_DIR / f"{template_name}.py"
    if template_path.exists():
        template_path.unlink()
        return True
    return False

def delete_logo(logo_name):
    logo_path = LOGOS_DIR / logo_name
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
    
    tab1, tab2 = st.tabs(["📄 PDF Templates", "🖼️ Logos"])
    
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
        st.subheader("Logo hochladen")
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
                    save_uploaded_file(uploaded_logo, LOGOS_DIR)
                    st.success(f"✅ Logo '{uploaded_logo.name}' gespeichert!")
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
                    logo_path = LOGOS_DIR / logo
                    try:
                        st.image(str(logo_path), caption=logo, width=150)
                        if st.button("🗑️", key=f"del_logo_{logo}"):
                            if delete_logo(logo):
                                st.success(f"Logo '{logo}' gelöscht")
                                st.rerun()
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
        st.warning("⚠️ Keine Templates vorhanden!")
    
    # Abstände nur bei liste_ Templates
    if st.session_state.pdf_template.startswith("liste_"):
        st.subheader("📏 Abstände")
        st.session_state.spacing_top_cm = st.number_input(
            "Oben (cm)",
            min_value=0.0,
            max_value=10.0,
            value=st.session_state.spacing_top_cm,
            step=0.5
        )
        
        st.session_state.spacing_bottom_cm = st.number_input(
            "Unten (cm)",
            min_value=0.0,
            max_value=10.0,
            value=st.session_state.spacing_bottom_cm,
            step=0.5
        )
    
    st.subheader("🖼️ Logo")
    st.session_state.logo_max_width_cm = st.number_input(
        "Breite (cm)",
        min_value=1.0,
        max_value=10.0,
        value=st.session_state.logo_max_width_cm,
        step=0.5
    )

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
                                    
                                    pdf_path = create_pdf(
                                        starterlist,
                                        pdf_filename,
                                        st.session_state.pdf_template,
                                        st.session_state.spacing_top_cm,
                                        st.session_state.spacing_bottom_cm,
                                        st.session_state.logo_max_width_cm,
                                        output_dir=str(OUTPUT_DIR)
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
