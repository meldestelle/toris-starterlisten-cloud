# templates/word/word_standard_logo.py
import os
import json
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from docx.enum.table import WD_TABLE_ALIGNMENT

# mapping English -> German for sex
SEX_MAP = {
    "MARE": "Stute",
    "GELDING": "Wallach", 
    "STALLION": "Hengst",
    "M": "Wallach",
    "F": "Stute",
    "": ""
}

WEEKDAY_MAP = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag",
    "Sunday": "Sonntag"
}

FIXED_WORD_TEMPLATE = "word_details_modern_simple.docx"
OUTPUT_DIR = "Ausgabe"

def _ensure_output_dir():
    """Stellt sicher, dass das Ausgabe-Verzeichnis existiert"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def _safe_get(d, key, default=""):
    if not d:
        return default
    return d.get(key, default) if isinstance(d, dict) else default

def _format_time(iso):
    """Formatiert ISO-Zeit zu HH:MM:SS"""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        return dt.strftime("%H:%M:%S")
    except Exception:
        # Fallback: erste 8 Zeichen für vollständige Zeit
        return str(iso)[:8] if len(str(iso)) >= 8 else str(iso)

def _format_header_datetime(iso):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", ""))
        weekday = WEEKDAY_MAP.get(dt.strftime("%A"), dt.strftime("%A"))
        return f"{weekday}, {dt.strftime('%d.%m.%Y um %H:%M Uhr')}"
    except Exception:
        return str(iso)

def _process_information_text(text):
    """Verarbeitet informationText - entfernt führende \n und macht Text fett"""
    if not text:
        return ""
    text = text.lstrip('\n')
    return text

def _set_header_gray_background(row):
    """Set gray background for header row"""
    tr = row._tr
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'D9D9D9')  # Light gray
        tcPr.append(shd)

def _set_table_borders(table):
    """Add borders to all table cells"""
    tbl = table._tbl
    for row in tbl.tr_lst:
        for cell in row.tc_lst:
            tcPr = cell.tcPr
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                cell.append(tcPr)
            
            tcBorders = OxmlElement('w:tcBorders')
            
            for border_name in ['w:top', 'w:left', 'w:bottom', 'w:right']:
                border = OxmlElement(border_name)
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '4')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), '000000')
                tcBorders.append(border)
            
            tcPr.append(tcBorders)


def _add_sponsorship_footer(doc):
    """Fügt Sponsorenleiste im Footer jeder Seite hinzu"""
    sponsor_logo_path = os.path.join("logos", "sponsorenleiste.png")
    if os.path.exists(sponsor_logo_path):
        try:
            for section in doc.sections:
                footer = section.footer
                footer_para = footer.paragraphs[0]
                footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                footer_para.clear()
                
                footer_run = footer_para.add_run()
                footer_run.add_picture(sponsor_logo_path, width=Inches(7.5))
                
                footer_para.paragraph_format.space_after = Pt(0)
                footer_para.paragraph_format.space_before = Pt(0)
            
            print(f"WORD DRE 3 LOGO DEBUG: Sponsorenleiste im Footer hinzugefügt: {sponsor_logo_path}")
        except Exception as e:
            print(f"WORD DRE 3 LOGO DEBUG: Sponsorenleiste Footer-Fehler: {e}")
    else:
        print(f"WORD DRE 3 LOGO DEBUG: Sponsorenleiste nicht gefunden: {sponsor_logo_path}")



def _set_row_keep_together(row):
    """Prevent table row from breaking across pages"""
    tr = row._tr
    trPr = tr.trPr
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        tr.insert(0, trPr)
    
    # Add cantSplit property to prevent row from breaking across pages
    cantSplit = trPr.find(qn('w:cantSplit'))
    if cantSplit is None:
        cantSplit = OxmlElement('w:cantSplit')
        cantSplit.set(qn('w:val'), 'true')
        trPr.append(cantSplit)
    else:
        cantSplit.set(qn('w:val'), 'true')

def _set_compact_cell_format(paragraph):
    """Set compact formatting for table cells"""
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.0

def _set_header_row_repeat(table):
    """Enable header row to repeat on each new page"""
    tbl = table._tbl
    
    # Get the first row (header row)
    if len(tbl.tr_lst) > 0:
        header_row = tbl.tr_lst[0]
        
        # Set the tblHeader property to repeat the header row on new pages
        trPr = header_row.trPr
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            header_row.insert(0, trPr)
        
        # Add or update tblHeader element
        tblHeader = trPr.find(qn('w:tblHeader'))
        if tblHeader is None:
            tblHeader = OxmlElement('w:tblHeader')
            tblHeader.set(qn('w:val'), 'true')
            trPr.append(tblHeader)
        else:
            tblHeader.set(qn('w:val'), 'true')
        
        print("WORD DEBUG: Header row repeat enabled")

def _set_judges_table_widths(table):
    """Set fixed column widths for judges table"""
    tbl = table._tbl
    
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    tblStyle = tblPr.find(qn('w:tblStyle'))
    if tblStyle is not None:
        tblPr.remove(tblStyle)
    
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is not None:
        tblPr.remove(tblW)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '0')
    tblW.set(qn('w:type'), 'auto')
    tblPr.append(tblW)
    
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    
    tblGrid = OxmlElement('w:tblGrid')
    judges_column_widths_twips = [360, 7200]
    
    for width in judges_column_widths_twips:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(width))
        tblGrid.append(gridCol)
    
    if tblPr.getparent() is not None:
        tblPr.getparent().insert(tblPr.getparent().index(tblPr) + 1, tblGrid)
    else:
        tbl.insert(0, tblGrid)
    
    for row in tbl.tr_lst:
        for i, cell in enumerate(row.tc_lst):
            if i < len(judges_column_widths_twips):
                tcPr = cell.tcPr
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    cell.insert(0, tcPr)
                
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    tcPr.remove(tcW)
                
                tcW = OxmlElement('w:tcW')
                tcW.set(qn('w:w'), str(judges_column_widths_twips[i]))
                tcW.set(qn('w:type'), 'dxa')
                tcPr.append(tcW)

def _set_kompakt_column_widths(table):
    """Set table to use exact column widths for 6-column kompakt layout (Start|Zeit|KoNr|Reiter|Pferd|Ergeb)"""
    tbl = table._tbl
    
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    tblStyle = tblPr.find(qn('w:tblStyle'))
    if tblStyle is not None:
        tblPr.remove(tblStyle)
    
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is not None:
        tblPr.remove(tblW)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '0')
    tblW.set(qn('w:type'), 'auto')
    tblPr.append(tblW)
    
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    
    tblGrid = OxmlElement('w:tblGrid')
    # 6-column widths: Start | Zeit | KoNr. (breiter) | Reiter (breiter) | Pferd (breiter) | Ergeb. (breiter) - NUTZT GANZE SEITE
    kompakt_column_widths_twips = [600, 600, 900, 3600, 3600, 1500]  # KoNr von 600 auf 900 erweitert
    
    for width in kompakt_column_widths_twips:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(width))
        tblGrid.append(gridCol)
    
    if tblPr.getparent() is not None:
        tblPr.getparent().insert(tblPr.getparent().index(tblPr) + 1, tblGrid)
    else:
        tbl.insert(0, tblGrid)
    
    for row in tbl.tr_lst:
        for i, cell in enumerate(row.tc_lst):
            if i < len(kompakt_column_widths_twips):
                tcPr = cell.tcPr
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    cell.insert(0, tcPr)
                
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    tcPr.remove(tcW)
                
                tcW = OxmlElement('w:tcW')
                tcW.set(qn('w:w'), str(kompakt_column_widths_twips[i]))
                tcW.set(qn('w:type'), 'dxa')
                tcPr.append(tcW)

def _get_judge_data_for_display_kompakt(judges):
    """
    Prozessiert Richter für die Anzeige - alle Positionen für die Richterliste in korrekter Reihenfolge
    """
    pos_map = {
        0: "E", 1: "H", 2: "C", 3: "M", 4: "B", 5: "K", 6: "V", 
        7: "S", 8: "R", 9: "P", 10: "F", 11: "A",
        "WARM_UP_AREA": "Aufsicht", "WATER_JUMP": "Wasser"
    }
    
    dressage_positions = ["E", "H", "C", "M", "B"]
    judges_by_position = {}  # Dictionary of Lists!
    other_judges = []
    
    for judge in judges:
        position = judge.get("position", "")
        if isinstance(position, int):
            pos_label = pos_map.get(position, str(position))
        else:
            pos_label = pos_map.get(str(position), str(position))
        
        judge_copy = judge.copy()
        judge_copy["pos_label"] = pos_label
        
        if pos_label in dressage_positions:
            # Füge zu Liste hinzu statt zu überschreiben
            if pos_label not in judges_by_position:
                judges_by_position[pos_label] = []
            judges_by_position[pos_label].append(judge_copy)
        else:
            other_judges.append(judge_copy)
    
    # Return judges in E H C M B order first, then others
    result = []
    for pos in dressage_positions:
        if pos in judges_by_position:
            # Füge ALLE Richter dieser Position hinzu
            result.extend(judges_by_position[pos])
    
    # Add other judges sorted by position label
    other_judges.sort(key=lambda j: j["pos_label"])
    result.extend(other_judges)
    
    return result
    
    
def render(starterlist: dict, filename: str):
    """
    Creates a Word document with 6-column kompakt layout (Start|Zeit|KoNr|Reiter|Pferd|Ergeb) - Standard version WITH LOGO
    """
    print("WORD STANDARD LOGO DEBUG: Starting render function")
    
    if starterlist is None:
        raise ValueError("starterlist is None")

    if isinstance(starterlist, str):
        try:
            starterlist = json.loads(starterlist)
        except Exception as e:
            raise ValueError(f"starterlist is a string but not valid JSON: {e}")

    _ensure_output_dir()
    
    # Use the fixed Word template as base
    template_path = os.path.join("templates", "word", FIXED_WORD_TEMPLATE)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Fixed Word template not found: {template_path}")
    
    # Output path
    output_path = os.path.join(OUTPUT_DIR, filename)
    
    # Copy template to output
    shutil.copy2(template_path, output_path)
    
    # Open the copied template document
    print("WORD STANDARD LOGO DEBUG: Loading document from template")
    doc = Document(output_path)
    
    # Find and replace STARTER_TABLE placeholder
    print("WORD DEBUG: Looking for STARTER_TABLE placeholder")
    starter_table_found = False
    placeholder_paragraph = None
    
    for paragraph in doc.paragraphs:
        if "STARTER_TABLE" in paragraph.text:
            print("WORD DEBUG: Found STARTER_TABLE placeholder")
            placeholder_paragraph = paragraph
            starter_table_found = True
            break
    
    if not starter_table_found:
        print("WORD DEBUG: STARTER_TABLE placeholder not found, creating at end of document")
        placeholder_paragraph = doc.add_paragraph("STARTER_TABLE")

    # Extract data from starterlist for header generation
    show_title = starterlist.get("showTitle", "")
    comp_title = starterlist.get("competitionTitle", "")
    comp_number = starterlist.get("competitionNumber", "")
    comp_subtitle = starterlist.get("subtitle", "")
    comp_info = starterlist.get("informationText", "")
    location = starterlist.get("location", "")
    start_time = starterlist.get("start", "")
    division_number = starterlist.get("divisionNumber")

    # Handle division start times
    start_raw = None
    divisions = starterlist.get('divisions', [])
    
    if divisions and division_number is not None:
        try:
            div_num = int(division_number)
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
        start_raw = start_time
    
    if start_raw:
        start_time = start_raw

    # Remove the placeholder paragraph first
    placeholder_paragraph._element.getparent().remove(placeholder_paragraph._element)

    # LOGO DIREKT AM ANFANG - Prüfungsspezifisches Logo-System
    logo_path = starterlist.get("logoPath")
    if logo_path and os.path.exists(logo_path):
        try:
            # Logo-Paragraph ganz oben, rechtsbündig
            logo_para = doc.add_paragraph()
            logo_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            logo_run = logo_para.add_run()
            logo_run.add_picture(logo_path)  # Größeres Logo
            logo_para.paragraph_format.space_after = Pt(0)
            logo_para.paragraph_format.space_before = Pt(0)
            
            # Negativer Spacer um das Logo nach oben zu "ziehen"
            spacer_para = doc.add_paragraph()
            spacer_para.paragraph_format.space_before = Pt(-30)  # Zieht nachfolgende Elemente nach oben
            spacer_para.paragraph_format.space_after = Pt(0)
            
            print(f"WORD STANDARD LOGO DEBUG: Logo positioned at top right: {logo_path}")
        except Exception as e:
            print(f"WORD STANDARD LOGO DEBUG: Logo error: {e}")

    # KOMPLETTER HEADER ABER KOMPAKT - NO LOGO in standard version
    print("WORD DEBUG: Creating COMPLETE but COMPACT header")

    if show_title:
        # Add title
        title_para = doc.add_paragraph()
        title_run = title_para.add_run(show_title)
        title_run.font.size = Pt(18)
        title_run.font.bold = True
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        title_para.paragraph_format.space_after = Pt(0)
        title_para.paragraph_format.space_before = Pt(0)
        title_para.paragraph_format.line_spacing = 1.0

    # Add competition info - COMPACT
    if comp_info:
        processed_info_text = _process_information_text(comp_info)
        info_para = doc.add_paragraph()
        run = info_para.add_run(processed_info_text)
        run.font.size = Pt(14)
        run.font.bold = True
        info_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        info_para.paragraph_format.space_before = Pt(0)
        info_para.paragraph_format.space_after = Pt(0)
        info_para.paragraph_format.line_spacing = 1.0

    # Add competition line - COMPACT
    if comp_number or comp_title:
        comp_line = f"Prüfung {comp_number}"
        if comp_title:
            comp_line += f" – {comp_title}"
        if division_number:
            try:
                div_num = int(division_number)
                if div_num > 0:
                    comp_line += f" - {div_num}. Abt." if "abt" not in comp_title.lower() else ""
            except:
                pass
        
        comp_para = doc.add_paragraph()
        run = comp_para.add_run(comp_line)
        run.font.size = Pt(12)
        run.font.bold = True
        comp_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        comp_para.paragraph_format.space_before = Pt(0)
        comp_para.paragraph_format.space_after = Pt(0)
        comp_para.paragraph_format.line_spacing = 1.0

    # Add subtitle - COMPACT
    if comp_subtitle:
        sub_para = doc.add_paragraph()
        run = sub_para.add_run(comp_subtitle)
        run.font.size = Pt(11)
        sub_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        sub_para.paragraph_format.space_before = Pt(0)
        sub_para.paragraph_format.space_after = Pt(0)
        sub_para.paragraph_format.line_spacing = 1.0

    # Add date/location - MINIMAL spacing to table
    if start_time:
        date_text = _format_header_datetime(start_time)
        if location:
            date_text += f" - {location}"
        date_para = doc.add_paragraph()
        run = date_para.add_run(date_text)
        run.font.size = Pt(11)
        date_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        date_para.paragraph_format.space_before = Pt(0)
        date_para.paragraph_format.space_after = Pt(3)  # MINIMAL spacing to table
        date_para.paragraph_format.line_spacing = 1.0

    # Create table DIREKT nach kompaktem Header - 6 COLUMNS for kompakt layout (Start|Zeit|KoNr|Reiter|Pferd|Ergeb)
    print("WORD DEBUG: Creating table after compact header - 6 columns kompakt")
    table = doc.add_table(rows=0, cols=6)  # 6 columns: Start, Zeit, KoNr, Reiter, Pferd, Ergeb.
    
    # Try to set table style, but continue without if not available
    try:
        table.style = 'Table Grid'
        print("WORD DEBUG: Applied Table Grid style")
    except Exception as e:
        print(f"WORD DEBUG: Table Grid style not available: {e}, continuing without style")
    
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    # Add header row
    hdr_row = table.add_row()
    hdr_cells = hdr_row.cells
    
    headers = ["Start", "Zeit", "KoNr.", "Reiter", "Pferd", "Ergeb."]
    
    for i, header_text in enumerate(headers):
        if i >= len(hdr_cells):
            break
            
        p = hdr_cells[i].paragraphs[0]
        p.clear()
        
        run = p.add_run(str(header_text))
        run.font.bold = True
        run.font.size = Pt(9)
        if i == 3 or i == 4:  # Reiter and Pferd columns - linksbündig
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        else:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Apply gray background to header row
    _set_header_gray_background(hdr_row)

    _set_kompakt_column_widths(table)
    
    # Enable header row repeat on new pages
    _set_header_row_repeat(table)

    # Process starters and breaks mit Abteilungslogik
    starters = starterlist.get("starters", [])
    breaks = starterlist.get("breaks", [])
    
    print(f"WORD DEBUG: Found {len(starters)} starters in starterlist")
    print(f"WORD DEBUG: Found {len(breaks)} breaks in starterlist")
    
    breaks_map = {}
    for br in breaks:
        try:
            after_num = int(br.get("afterNumberInCompetition", -1))
            breaks_map[after_num] = br
        except:
            continue

    # Gruppierung nach groupNumber für Abteilungen
    starters_by_group = {}
    for starter in starters:
        group_num = starter.get("groupNumber", 0)
        if group_num not in starters_by_group:
            starters_by_group[group_num] = []
        starters_by_group[group_num].append(starter)
    
    # Gruppen sortiert verarbeiten
    for group_num in sorted(starters_by_group.keys()):
        group_starters = starters_by_group[group_num]
        
        # Abteilungsheader hinzufügen (wenn mehr als eine Gruppe)
        if len(starters_by_group) > 1 and group_num > 0:
            header_row = table.add_row()
            header_cells = header_row.cells
            
            # Abteilungstext linksbündig
            p_header = header_cells[0].paragraphs[0]
            p_header.clear()
            run = p_header.add_run(f"{group_num}. Abteilung")
            run.font.size = Pt(11)
            run.font.bold = True
            p_header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Linksbündig
            
            # Merge über alle Spalten
            cell_0 = header_cells[0]
            for i in range(1, 6):
                header_cells[i].paragraphs[0].clear()
                cell_0.merge(header_cells[i])
            
            # Grauer Hintergrund für Abteilungsheader
            tcPr = cell_0._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), 'E6E6E6')  # Hellgrau
            tcPr.append(shd)
        
        # Erste Startzeit pro Gruppe bestimmen
        group_start_time = None
        for starter in group_starters:
            start_time_candidate = starter.get("startTime")
            if start_time_candidate:
                group_start_time = start_time_candidate
                break
        # Starter der Gruppe verarbeiten
        for starter_idx, starter in enumerate(group_starters):
            row = table.add_row()
            row_cells = row.cells
            
            # Prevent this row from breaking across pages and set compact height
            _set_row_keep_together(row)
            
            # Set row height to be more compact
            tr = row._tr
            trPr = tr.trPr
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                tr.insert(0, trPr)
            
            # Set exact row height for compactness
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), '280')  # Kompakte Zeilenhöhe in twips
            trHeight.set(qn('w:hRule'), 'exact')
            trPr.append(trHeight)
            
            # Column 0: Start number (NUR NUMMER)
            start_num = starter.get("startNumber", "")
            
            p0 = row_cells[0].paragraphs[0]
            p0.clear()
            p0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            if start_num:
                run = p0.add_run(str(start_num))
                run.font.bold = True
                run.font.size = Pt(9)
            
            _set_compact_cell_format(p0)

            # Column 1: Zeit (immer anzeigen wenn keine Abteilungen, sonst nur erste Zeit pro Gruppe)
            p1 = row_cells[1].paragraphs[0]
            p1.clear()
            p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Zeit anzeigen: Entweder bei jedem Starter (wenn keine Abteilungen) oder nur beim ersten der Gruppe
            should_show_time = False
            if len(starters_by_group) == 1:  # Keine Abteilungen - zeige immer die Zeit
                individual_start_time = starter.get("startTime")
                if individual_start_time:
                    start_time_str = _format_time(individual_start_time)
                    should_show_time = True
            else:  # Abteilungen vorhanden - nur erste Zeit pro Gruppe
                if starter_idx == 0 and group_start_time:
                    start_time_str = _format_time(group_start_time)
                    should_show_time = True
            
            if should_show_time and start_time_str:
                run = p1.add_run(start_time_str)
                run.font.size = Pt(9)  # Zeit-Schriftgröße konsistent
            
            _set_compact_cell_format(p1)

            # Column 2: KoNr.
            horses = starter.get("horses", [])
            p2 = row_cells[2].paragraphs[0]
            p2.clear()
            p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            if horses:
                horse = horses[0]
                cno = _safe_get(horse, "cno", "")
                if cno:
                    run = p2.add_run(str(cno))
                    run.font.size = Pt(9)
            
            _set_compact_cell_format(p2)

            # Column 3: Reiter (NUR NAME - KOMPAKT, ABER BREITER)
            athlete = starter.get("athlete", {})
            rider_name = _safe_get(athlete, "name", "")
            
            p3 = row_cells[3].paragraphs[0]
            p3.clear()
            p3.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
            # Nur Reitername (KEINE Großschreibung)
            if rider_name:
                run = p3.add_run(rider_name)
                run.font.bold = True
                run.font.size = Pt(9)
            
            _set_compact_cell_format(p3)

            # Column 4: Pferd (NUR NAME - KOMPAKT, ABER BREITER)
            p4 = row_cells[4].paragraphs[0]
            p4.clear()
            p4.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
            if horses:
                horse = horses[0]
                horse_name = _safe_get(horse, "name", "")
                
                # Nur Pferdename (KEINE Großschreibung)
                if horse_name:
                    run = p4.add_run(horse_name)
                    run.font.bold = True
                    run.font.size = Pt(9)
            
            _set_compact_cell_format(p4)

            # Column 5: Ergeb. (BREITER für Ergebnisse)
            p5 = row_cells[5].paragraphs[0]
            p5.clear()
            p5.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            _set_compact_cell_format(p5)

            # Handle withdrawn entries
            if starter.get("withdrawn", False):
                for cell in row_cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                    
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(96, 96, 96)
                            run.font.strike = True

            # Check for breaks
            try:
                starter_num = int(start_num)
                if starter_num in breaks_map:
                    break_info = breaks_map[starter_num]
                    
                    break_row = table.add_row()
                    break_cells = break_row.cells
                    
                    # Prevent break row from splitting across pages
                    _set_row_keep_together(break_row)
                    
                    # Set row height to be compact - AUCH FÜR PAUSEN
                    tr = break_row._tr
                    trPr = tr.trPr
                    if trPr is None:
                        trPr = OxmlElement('w:trPr')
                        tr.insert(0, trPr)
                    
                    # Set exact row height for compactness
                    trHeight = OxmlElement('w:trHeight')
                    trHeight.set(qn('w:val'), '280')  # Kompakte Zeilenhöhe auch für Pausen
                    trHeight.set(qn('w:hRule'), 'exact')
                    trPr.append(trHeight)
                    
                    total_seconds = int(break_info.get("totalSeconds", 0) or 0)
                    info_text = break_info.get("informationText", "")
                    
                    if total_seconds > 0:
                        if total_seconds >= 3600:
                            hours = total_seconds // 3600
                            mins = (total_seconds % 3600) // 60
                            time_text = f"{hours} Std. {mins:02d} Min. Pause"
                        else:
                            mins = total_seconds // 60
                            time_text = f"{mins} Minuten Pause"
                        
                        break_text = f"{time_text} – {info_text}" if info_text else time_text
                    else:
                        break_text = info_text or "Pause"
                    
                    p_break = break_cells[0].paragraphs[0]
                    p_break.clear()
                    run = p_break.add_run(break_text)
                    run.font.size = Pt(10)
                    run.font.bold = True
                    p_break.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    
                    # Kompakte Formatierung auch für Pausen-Text
                    _set_compact_cell_format(p_break)
                    
                    # Clear other cells and merge (6 columns for kompakt)
                    for i in range(1, 6):
                        break_cells[i].paragraphs[0].clear()
                    
                    cell_0 = break_cells[0]
                    for i in range(1, 6):
                        cell_0.merge(break_cells[i])
                    
                    tcPr = cell_0._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                    
            except (ValueError, TypeError):
                pass

    _set_table_borders(table)
    _add_sponsorship_footer(doc)

    # Add judges section - WICHTIG: Mit korrekter Reihenfolge E H C M B
    judges = starterlist.get("judges", [])
    if judges:
        doc.add_paragraph()
        
        p = doc.add_paragraph()
        run = p.add_run("Richter:")
        run.font.bold = True
        run.font.size = Pt(11)
        
        # Get judge data for display - KORREKTE REIHENFOLGE E H C M B
        judge_display_data = _get_judge_data_for_display_kompakt(judges)
        
        if judge_display_data:
            judges_table = doc.add_table(rows=1 + len(judge_display_data), cols=2)
            
            # Try to set table style, but continue without if not available
            try:
                judges_table.style = 'Table Grid'
                print("WORD DEBUG: Applied Table Grid style to judges table")
            except Exception as e:
                print(f"WORD DEBUG: Table Grid style not available for judges table: {e}, continuing without style")
            
            judges_table.alignment = WD_TABLE_ALIGNMENT.LEFT
            
            # Header row
            hdr_cells = judges_table.rows[0].cells
            headers = ["Pos", "Richter"]
            
            for i, header_text in enumerate(headers):
                p = hdr_cells[i].paragraphs[0]
                p.clear()
                run = p.add_run(header_text)
                run.font.bold = True
                run.font.size = Pt(9)
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Apply gray background to judges header row
            _set_header_gray_background(judges_table.rows[0])
            
            # Judge rows - KORREKTE REIHENFOLGE E H C M B
            for i, judge_data in enumerate(judge_display_data):
                row = judges_table.rows[i + 1]
                cells = row.cells
                
                # Position
                p = cells[0].paragraphs[0]
                p.clear()
                run = p.add_run(judge_data.get("pos_label", ""))
                run.font.size = Pt(9)
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                
                # Name
                p = cells[1].paragraphs[0]
                p.clear()
                run = p.add_run(judge_data.get("name", ""))
                run.font.size = Pt(9)

            _set_table_borders(judges_table)
            _set_judges_table_widths(judges_table)

    # Save document
    print("WORD DEBUG: Starting document save")
    try:
        doc.save(output_path)
        print(f"WORD STANDARD LOGO DEBUG: Document saved to {output_path}")
        return output_path
    except Exception as e:
        print(f"WORD DEBUG: Error saving document: {e}")
        alt_output_path = filename
        try:
            doc.save(alt_output_path)
            print(f"WORD STANDARD LOGO DEBUG: Document saved to alternative location: {alt_output_path}")
            return alt_output_path
        except Exception as e2:
            print(f"WORD DEBUG: Failed to save to alternative location: {e2}")
            raise