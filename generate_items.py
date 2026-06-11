import io
import os
import re
import csv
import json
import time
from datetime import datetime
import pandas as pd
import openpyxl
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from mistralai.client import Mistral

from curriculum_analyzer import load_curriculum_framework, get_next_sequence_number

# ─── MISTRAL CLIENT SETUP ──────────────────────────────────────────────────────
MODEL = "mistral-large-latest"

def get_mistral_client(api_key: str):
    if api_key.lower() in ("mock", "test", ""):
        return None
    try:
        return Mistral(api_key=api_key)
    except Exception:
        return None

# ─── XML HELPERS FOR DOCX FORMATTING ──────────────────────────────────────────
def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing shading if any
    existing_shd = tcPr.find(qn('w:shd'))
    if existing_shd is not None:
        tcPr.remove(existing_shd)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_borders_all(cell, color_hex, sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    existing_borders = tcPr.find(qn('w:tcBorders'))
    if existing_borders is not None:
        tcPr.remove(existing_borders)
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), sz)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color_hex)
        tcBorders.append(element)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Remove existing cell margins if any
    existing_margins = tcPr.find(qn('w:tcMar'))
    if existing_margins is not None:
        tcPr.remove(existing_margins)
    tcMar = OxmlElement('w:tcMar')
    for name, w in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{name}')
        node.set(qn('w:w'), str(w))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_width_dxa(cell, w_dxa):
    cell.width = Pt(int(w_dxa) / 20)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(w_dxa))
    tcW.set(qn('w:type'), 'dxa')
    existing_tcW = tcPr.find(qn('w:tcW'))
    if existing_tcW is not None:
        tcPr.remove(existing_tcW)
    tcPr.append(tcW)

def create_styled_table(doc, num_rows, widths_dxa, border_color="cccccc"):
    num_cols = len(widths_dxa)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Table-level borders setup
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{b_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), border_color)
        tblBorders.append(b)
    tblPr.append(tblBorders)
    
    # Set total table width
    total_w = sum(widths_dxa)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(total_w))
    tblW.set(qn('w:type'), 'dxa')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)
    
    # Set column grid (tblGrid)
    tblGrid = OxmlElement('w:tblGrid')
    for w in widths_dxa:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(w))
        tblGrid.append(gridCol)
    existing_tblGrid = table._tbl.find(qn('w:tblGrid'))
    if existing_tblGrid is not None:
        table._tbl.remove(existing_tblGrid)
    table._tbl.insert(1, tblGrid)
    
    # Format each cell
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))
        for idx, cell in enumerate(row.cells):
            set_cell_width_dxa(cell, widths_dxa[idx])
            set_cell_margins(cell, top=60, bottom=60, left=120, right=120)
            set_cell_borders_all(cell, border_color)
            
    return table

# ─── PARAGRAPH HELPERS ────────────────────────────────────────────────────────
def add_section_header(doc, title, desc=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    
    r1 = p.add_run(title)
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0, 0, 0)
    
    if desc:
        r2 = p.add_run(desc)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)
        r2.font.bold = False
        r2.font.color.rgb = RGBColor(0, 0, 0)
    return p

def set_cell_text(cell, text, bold=False, italic=False, color_rgb="000000", alignment=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    if alignment is not None:
        p.alignment = alignment
        
    if not text:
        # Style empty run
        run = p.add_run("")
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        return
        
    lines = text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            if alignment is not None:
                p.alignment = alignment
        run = p.add_run(line)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        run.font.bold = bold
        run.font.italic = italic
        if color_rgb:
            run.font.color.rgb = RGBColor.from_string(color_rgb)

# ─── JSON CLEANUP AND PARSE ───────────────────────────────────────────────────
def clean_and_parse_json(text: str):
    cleaned = text.strip()
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines[-1].startswith("```"):
            lines = lines[:-1]
        cleaned = "\n".join(lines).strip()
    return json.loads(cleaned)

# ─── MISTRAL API AND RETRY SYSTEM ──────────────────────────────────────────────
def call_mistral(client, system_prompt, user_message, max_tokens=600, use_json=True, retries=3):
    if client is None:
        raise ValueError("Mistral client not initialized (Mock Mode)")
        
    last_err = None
    for attempt in range(retries):
        try:
            response = client.chat.complete(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message}
                ],
                temperature=0.1,
                max_tokens=max_tokens,
                response_format={"type": "json_object"} if use_json else None
            )
            content = response.choices[0].message.content
            if use_json:
                return clean_and_parse_json(content)
            else:
                return content.strip()
        except Exception as e:
            last_err = e
            time.sleep(1.5 * (attempt + 1))
            
    raise last_err

# ─── HIGH-FIDELITY MOCK FALLBACKS ──────────────────────────────────────────────
def generate_mock_mastery(row, inferred_type):
    # Try to determine from excel
    excel_m = str(row.get("Mastery_Level") or "").strip().upper()
    if excel_m in ("M1", "M2", "M3", "M4"):
        return {"mastery_level": excel_m, "reasoning": f"Validated mastery level {excel_m} from spreadsheet.", "operator_override": False}
        
    # Otherwise check type
    if inferred_type == "MCQ":
        m_level = "M2"
    elif inferred_type == "Extended Response":
        m_level = "M4"
    else:
        m_level = "M2"
        
    return {
        "mastery_level": m_level,
        "reasoning": f"Classified as {m_level} based on question structure (mock inference).",
        "operator_override": False
    }

def generate_mock_format_fields(row, item_type):
    stimulus = str(row.get("Stimulus_Text") or "").lower()
    stem = str(row.get("Item_Stem") or "").lower()
    combined = stimulus + " " + stem
    
    has_image = "Yes" if any(w in combined for w in ["image", "diagram", "figure", "picture", "photo", "saucer", "spoon"]) else "No"
    has_table = "Yes" if any(w in combined for w in ["table", "data table", "grid", "rows", "columns"]) else "No"
    has_equation = "Yes" if any(w in combined for w in ["h2o", "co2", "formula", "equation", "°c", "m/s", "f=ma"]) else "No"
    eq_format = "LaTeX" if has_equation == "Yes" else "N/A"
    
    return {
        "has_image": has_image,
        "has_table": has_table,
        "has_equation": has_equation,
        "equation_format": eq_format
    }

def generate_mock_rationales(row, correct_letter):
    stem = str(row.get("Item_Stem") or "")
    opts = {
        "A": str(row.get("Option_A") or ""),
        "B": str(row.get("Option_B") or ""),
        "C": str(row.get("Option_C") or ""),
        "D": str(row.get("Option_D") or "")
    }
    
    rationales = {}
    for letter, text in opts.items():
        if letter == correct_letter:
            rationales[f"rationale_{letter.lower()}"] = f"CORRECT ANSWER: {text} is correct because it represents the accurate scientific property of the material."
        else:
            rationales[f"rationale_{letter.lower()}"] = f"Within-category opposite or cross-category confusion: The learner selected '{text}' because they confused this property with another sensory property of materials."
            
    return rationales

def generate_mock_explanation(row, item_type, correct_letter, rationales):
    stem = str(row.get("Item_Stem") or "")
    opts = {
        "A": str(row.get("Option_A") or ""),
        "B": str(row.get("Option_B") or ""),
        "C": str(row.get("Option_C") or ""),
        "D": str(row.get("Option_D") or "")
    }
    
    if item_type == "MCQ":
        correct_text = opts.get(correct_letter, "")
        explanation = f"The correct answer is {correct_letter}. {correct_text}.\n\nWhen we test this material, we observe that it matches the properties required by the question stem. Dissolving, transparency, and lustre are distinct categories of material properties.\n\nWhy other options are incorrect:\n\n"
        for letter in ["A", "B", "C", "D"]:
            if letter != correct_letter:
                text = opts.get(letter, "")
                explanation += f"Option {letter} ({text}) is incorrect because it describes a different property of materials that does not solve the task described in the stem.\n\n"
        return explanation.strip()
    else:
        ans = str(row.get("Correct_Answer") or "")
        return f"The model answer is: {ans}.\n\nThis is correct because the material properties match the requirements of the scenario. Students who answer partially are likely to identify only one relevant property (e.g. strength) while ignoring others (e.g. water resistance)."

def generate_mock_rubric(row, max_score):
    stem = str(row.get("Item_Stem") or "")
    ans = str(row.get("Correct_Answer") or "")
    
    rows = []
    if max_score == 2:
        rows = [
            {"score": 2, "label": "Full marks", "criteria": "Student correctly identifies both material properties and explains their relevance.", "sample": "This material is suitable because it is both strong and waterproof."},
            {"score": 1, "label": "Partial marks", "criteria": "Student identifies only one property or fails to explain relevance.", "sample": "It is suitable because it is strong."},
            {"score": 0, "label": "Zero", "criteria": "Student gives an incorrect or off-topic explanation.", "sample": "It is suitable because it looks nice."}
        ]
    elif max_score == 3:
        rows = [
            {"score": 3, "label": "Full marks", "criteria": "Student applies the concept to the scenario with a fully justified explanation and correct terminology.", "sample": "The bag broke because dried leaves lack tensile strength and are not water-resistant, making them unsuitable for wet grocery items."},
            {"score": 2, "label": "Substantial", "criteria": "Student provides the correct outcome with mostly complete reasoning.", "sample": "It is flawed because dried leaves tear easily under heavy weight and disintegrate when wet."},
            {"score": 1, "label": "Basic", "criteria": "Student recognizes the bag failed but cannot describe properties correctly.", "sample": "It is flawed because the watermelon is heavy and leaves are weak."},
            {"score": 0, "label": "Zero", "criteria": "Student gives a misconception-based answer.", "sample": "It is flawed because watermelons shouldn't be in bags."}
        ]
    else: # max_score == 4 or fallback
        rows = [
            {"score": 4, "label": "Full marks", "criteria": "Student evaluates the flaw, identifies the error in reasoning, and corrects it with two properties.", "sample": "Boojho's claim is flawed. A grocery bag requires strength to hold heavy weights and water resistance to stay intact when wet. Dried leaves lack both, so it will tear."},
            {"score": 3, "label": "Strong partial", "criteria": "Student identifies the reasoning flaw and corrects it with one property justified.", "sample": "It is flawed because shopping bags need strength to hold watermelons, and leaves are too weak."},
            {"score": 2, "label": "Moderate partial", "criteria": "Student states what failed practically without framing as a reasoning flaw.", "sample": "The bag broke because it got wet and was too heavy."},
            {"score": 1, "label": "Minimal", "criteria": "Student names a property but has no explanation.", "sample": "It is flawed because leaves are not strong."},
            {"score": 0, "label": "Zero", "criteria": "Student defends the claim or gives irrelevant responses.", "sample": "The leaves are good because they are organic."}
        ]
        
    return {"rows": rows}

def get_matching_chapter_in_science_excel(selected_chap, excel_chapters):
    mapping = {
        "materials around us": "Materials Around Us",
        "living creatures": "Living Creatures: Exploring  their Characteristics",
        "diversity in the living world": "Diversity in the Living World",
        "exploring magnets": "Exploring Magnets",
        "mindful eating": "Mindful Eating: A Path  to a Healthy Body",
        "measurement of length and motion": "Measurement of Length and Motion",
        "temperature and its measurement": "Temperature and its Measurement",
        "a journey through states of water": "A Journey through States of Water",
        "methods of separation": "Method of Separation In Everyday Life",
        "nature's treasures": "Nature’s Treasures",
        "nature’s treasures": "Nature’s Treasures",
        "beyond earth": "Beyond Earth"
    }
    
    sel_norm = str(selected_chap).lower().strip().replace("’", "'").replace("`", "'")
    if sel_norm in mapping:
        mapped_name = mapping[sel_norm]
        for ec in excel_chapters:
            if ec.lower().strip().replace("’", "'").replace("`", "'") == mapped_name.lower().replace("’", "'"):
                return ec
    
    for ec in excel_chapters:
        ec_norm = ec.lower().strip().replace("’", "'").replace("`", "'")
        if sel_norm in ec_norm or ec_norm in sel_norm:
            return ec
            
    sel_words = set(sel_norm.split())
    for ec in excel_chapters:
        ec_words = set(ec.lower().strip().replace("’", "'").replace("`", "'").split())
        if sel_words.intersection(ec_words):
            return ec
            
    return selected_chap

def match_topic_heuristically(row_topic_or_stem, predefined_topics):
    if not predefined_topics:
        return None
    
    row_topic_norm = str(row_topic_or_stem).lower().strip()
    
    best_match = None
    best_score = -1
    
    for pt in predefined_topics:
        topic_name = pt['Topic']
        topic_norm = str(topic_name).lower().strip()
        
        if topic_norm == row_topic_norm:
            return pt
            
        if topic_norm in row_topic_norm or row_topic_norm in topic_norm:
            score = len(topic_norm)
            if score > best_score:
                best_score = score
                best_match = pt
                
    if best_match:
        return best_match
        
    for pt in predefined_topics:
        topic_name = pt['Topic']
        topic_words = set(re.findall(r'\w+', str(topic_name).lower()))
        input_words = set(re.findall(r'\w+', str(row_topic_or_stem).lower()))
        overlap = len(topic_words.intersection(input_words))
        if overlap > best_score:
            best_score = overlap
            best_match = pt
            
    return best_match or predefined_topics[0]

def classify_topic_with_ai(client, item_stem, stimulus_text, predefined_topics, row_topic=None):
    if not predefined_topics:
        return None
        
    if client is None:
        query = row_topic if row_topic else item_stem
        return match_topic_heuristically(query, predefined_topics)
        
    topics_list_str = ""
    for pt in predefined_topics:
        topics_list_str += f"- {pt['Topic']}\n"
        
    system_prompt = (
        "You are an educational curriculum mapping AI.\n"
        "Your task is to classify a science assessment item into exactly one topic from the predefined list of topics below.\n"
        "You must choose the topic that best matches the scientific concept tested in the question.\n\n"
        "PREDEFINED TOPICS:\n"
        f"{topics_list_str}\n"
        "INSTRUCTIONS:\n"
        "1. Choose the topic from the predefined list that is the best match for the question.\n"
        "2. Your output MUST be a JSON object containing exactly the key 'selected_topic' with the exact name of the topic from the list.\n"
        "3. Do not invent any new topics. Only select from the provided list."
    )
    
    user_message = (
        f"Item Stem: {item_stem}\n"
        f"Stimulus: {stimulus_text if stimulus_text else 'None'}\n"
    )
    if row_topic:
        user_message += f"Original Topic suggestion from Excel: {row_topic}\n"
        
    try:
        res = call_mistral(client, system_prompt, user_message, max_tokens=150, use_json=True)
        selected_name = res.get("selected_topic", "").strip()
        
        for pt in predefined_topics:
            if pt['Topic'].lower().strip() == selected_name.lower().strip():
                return pt
                
        for pt in predefined_topics:
            if selected_name.lower() in pt['Topic'].lower() or pt['Topic'].lower() in selected_name.lower():
                return pt
    except Exception as e:
        print(f"Topic classification AI call failed: {e}")
        
    query = row_topic if row_topic else item_stem
    return match_topic_heuristically(query, predefined_topics)

def get_framework_alignment_for_topic(topic_name, chapter_name, unique_combos):
    norm_topic = str(topic_name).lower().strip().replace("’", "'").replace("`", "'")
    norm_chap = str(chapter_name).lower().strip().replace("’", "'").replace("`", "'")
    
    for combo in unique_combos:
        c_topic = str(combo["topic_name"]).lower().strip().replace("’", "'").replace("`", "'")
        c_chap = str(combo["chapter_name"]).lower().strip().replace("’", "'").replace("`", "'")
        if c_topic == norm_topic and (norm_chap in c_chap or c_chap in norm_chap):
            return combo
            
    for combo in unique_combos:
        c_topic = str(combo["topic_name"]).lower().strip().replace("’", "'").replace("`", "'")
        if c_topic == norm_topic:
            return combo
            
    for combo in unique_combos:
        c_topic = str(combo["topic_name"]).lower().strip().replace("’", "'").replace("`", "'")
        if norm_topic in c_topic or c_topic in norm_topic:
            return combo
            
    for combo in unique_combos:
        c_chap = str(combo["chapter_name"]).lower().strip().replace("’", "'").replace("`", "'")
        if norm_chap in c_chap or c_chap in norm_chap:
            return combo
            
    return None

def compute_lo_id(combo, unique_combos):
    comp = combo.get("competency", "C-1.1")
    lo = combo.get("learning_outcome", "")
    
    comp_num = "1.1"
    m_comp = re.search(r"C-(\d+\.\d+)", comp)
    if m_comp:
        comp_num = m_comp.group(1)
    else:
        m_num = re.search(r"(\d+\.\d+)", comp)
        if m_num:
            comp_num = m_num.group(1)
            
    comp_los = []
    seen_lo = set()
    for entry in unique_combos:
        if entry["competency"] == comp:
            lo_clean = entry["learning_outcome"].strip()
            if lo_clean not in seen_lo:
                seen_lo.add(lo_clean)
                comp_los.append(lo_clean)
                
    try:
        curr_lo = lo.strip()
        match_lo_idx = 0
        for idx, entry_lo in enumerate(comp_los):
            if entry_lo.lower().strip() == curr_lo.lower().strip():
                match_lo_idx = idx
                break
        suffix_letter = chr(ord('a') + match_lo_idx)
    except Exception:
        suffix_letter = 'a'
        
    return f"LO-{comp_num}.{suffix_letter}"

# ─── CORE PIPELINE RUNNER ─────────────────────────────────────────────────────
def run_pipeline(excel_file, api_key: str, batch_meta: dict, progress_callback):
    """
    Ingests, validates, enriches and generates DOCX and CSV log.
    Returns: docx_bytes, csv_bytes, summary_dict
    """
    # 1. Load Excel
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        raise ValueError(f"Could not load Excel file: {e}")
        
    if "Questions" not in wb.sheetnames:
        raise ValueError("Excel file must contain a worksheet named 'Questions'")
        
    sheet = wb["Questions"]
    
    # Load Science Class 6 mapping
    parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if '__file__' in locals() else os.path.dirname(os.getcwd())
    science_path = os.path.join(parent_dir, "Science Class 6.xlsx")
    if not os.path.exists(science_path):
        science_path = os.path.join(os.path.dirname(os.getcwd()), "Science Class 6.xlsx")
    if not os.path.exists(science_path):
        science_path = "Science Class 6.xlsx"
        
    try:
        df_science = pd.read_excel(science_path)
    except Exception as e:
        raise ValueError(f"Could not load Science Class 6.xlsx: {e}")
        
    # Get the list of all unique chapters in the Science Class 6 sheet
    excel_chapters = df_science['Chapter'].dropna().unique().tolist()
    
    # Map batch_meta["chapter_name"] to the exact name in the excel
    matched_chapter = get_matching_chapter_in_science_excel(batch_meta["chapter_name"], excel_chapters)
    
    # Filter the Science Class 6 sheet to get predefined topics for this chapter
    df_chap = df_science[df_science['Chapter'] == matched_chapter]
    if df_chap.empty:
        df_chap = df_science[df_science['Chapter'].str.contains(matched_chapter[:10], case=False, na=False)]
        
    if df_chap.empty:
        raise ValueError(f"No topics found for chapter '{matched_chapter}' in Science Class 6.xlsx")
        
    predefined_topics = df_chap[['Topic ID', 'Topic']].drop_duplicates().to_dict('records')
    subject_id = str(df_chap.iloc[0]['Subject ID']).strip() if not df_chap.empty else "65ba59aef233a16d51f7163d"
    
    # Load Curriculum Framework dynamically for mapping NCF CG, Competency, etc.
    framework_path = os.path.join(parent_dir, "Final_Class 6_Science_Mastery Framework_ (1).xlsx")
    if not os.path.exists(framework_path):
        framework_path = os.path.join(os.path.dirname(os.getcwd()), "Final_Class 6_Science_Mastery Framework_ (1).xlsx")
    if not os.path.exists(framework_path):
        framework_path = "Final_Class 6_Science_Mastery Framework_ (1).xlsx"
        
    unique_combos = load_curriculum_framework(framework_path)
    
    # Track sequence number per topic ID
    current_seq = {}

    
    # Check headers
    expected_headers = [
        "Mastery_Level", "Topic", "Stimulus_Text", "Item_Stem",
        "Option_A", "Option_B", "Option_C", "Option_D",
        "Correct_Option", "Correct_Answer"
    ]
    
    sheet_headers = [cell.value for cell in sheet[1]]
    # Handle minor trailing whitespaces or casing in header validation
    clean_headers = [str(h).strip() for h in sheet_headers if h is not None]
    
    # Check if all expected headers exist
    for eh in expected_headers:
        if eh not in clean_headers:
            raise ValueError(f"Required column '{eh}' not found in 'Questions' sheet headers. Found: {clean_headers}")
            
    # Map headers to indices
    header_map = {str(cell.value).strip(): col_idx for col_idx, cell in enumerate(sheet[1], start=1) if cell.value is not None}
    
    # Read rows
    raw_rows = []
    for r_idx in range(2, sheet.max_row + 1):
        row_dict = {}
        row_empty = True
        for h_name, col_idx in header_map.items():
            val = sheet.cell(row=r_idx, column=col_idx).value
            if val is not None:
                row_empty = False
            row_dict[h_name] = val
        if not row_empty:
            row_dict["_excel_row_num"] = r_idx
            raw_rows.append(row_dict)
            
    total_raw = len(raw_rows)
    
    # 2. Ingest and Validate
    valid_items = []
    skipped_items = 0
    override_count = 0
    
    audit_logs = []
    
    client = get_mistral_client(api_key)
    is_mock = client is None
    
    for idx, row in enumerate(raw_rows):
        excel_row_num = row["_excel_row_num"]
        stem = str(row.get("Item_Stem") or "").strip()
        
        # Validation A: Stem must not be blank
        if not stem:
            skipped_items += 1
            audit_logs.append({
                "Row_Number": excel_row_num,
                "Item_ID": "N/A",
                "Status": "ERROR",
                "Message": "Skipped: Item Stem is blank.",
                "Mastery_Level": "N/A",
                "Item_Type": "N/A",
                "Overrides": ""
            })
            continue
            
        opt_a = str(row.get("Option_A") or "").strip()
        opt_b = str(row.get("Option_B") or "").strip()
        opt_c = str(row.get("Option_C") or "").strip()
        opt_d = str(row.get("Option_D") or "").strip()
        correct_opt = str(row.get("Correct_Option") or "").strip().upper()
        correct_ans = str(row.get("Correct_Answer") or "").strip()
        mastery_excel = str(row.get("Mastery_Level") or "").strip().upper()
        
        # Validate mastery level syntax if provided
        if mastery_excel and mastery_excel not in ("M1", "M2", "M3", "M4"):
            skipped_items += 1
            audit_logs.append({
                "Row_Number": excel_row_num,
                "Item_ID": "N/A",
                "Status": "ERROR",
                "Message": f"Skipped: Invalid Mastery Level '{mastery_excel}' (must be M1, M2, M3, or M4).",
                "Mastery_Level": mastery_excel,
                "Item_Type": "N/A",
                "Overrides": ""
            })
            continue
            
        # Inferred Item Type logic (Rule 1.3)
        has_options = bool(opt_a or opt_b or opt_c or opt_d)
        is_mcq_opt = correct_opt in ("A", "B", "C", "D")
        
        inferred_item_type = None
        if has_options and is_mcq_opt:
            inferred_item_type = "MCQ"
        elif not has_options:
            if mastery_excel == "M4":
                inferred_item_type = "Extended Response"
            elif mastery_excel == "M1":
                # M1 must be MCQ (Rule 1.3 constraint error). We override and log a warning.
                inferred_item_type = "MCQ"
            else: # M2, M3 or empty
                inferred_item_type = "Short Answer"
        else:
            # Partially filled options or bad correct option
            inferred_item_type = "MCQ"
            
        # Validation B: MCQ options check
        if inferred_item_type == "MCQ":
            if not (opt_a and opt_b and opt_c and opt_d):
                skipped_items += 1
                audit_logs.append({
                    "Row_Number": excel_row_num,
                    "Item_ID": "N/A",
                    "Status": "ERROR",
                    "Message": "Skipped: MCQ item is missing one or more options.",
                    "Mastery_Level": mastery_excel or "N/A",
                    "Item_Type": "MCQ",
                    "Overrides": ""
                })
                continue
            if not is_mcq_opt:
                skipped_items += 1
                audit_logs.append({
                    "Row_Number": excel_row_num,
                    "Item_ID": "N/A",
                    "Status": "ERROR",
                    "Message": f"Skipped: MCQ item has invalid Correct Option '{correct_opt}' (must be A, B, C, or D).",
                    "Mastery_Level": mastery_excel or "N/A",
                    "Item_Type": "MCQ",
                    "Overrides": ""
                })
                continue
                
        # All basic load validations passed!
        valid_items.append({
            "row": row,
            "inferred_item_type": inferred_item_type,
            "mastery_excel": mastery_excel
        })
        
    # Process valid items
    total_valid = len(valid_items)
    processed_items = []
    
    # Sequence numbers
    seq_num = batch_meta["start_seq"]
    
    for i, item in enumerate(valid_items):
        row = item["row"]
        excel_row_num = row["_excel_row_num"]
        inferred_item_type = item["inferred_item_type"]
        mastery_excel = item["mastery_excel"]
        topic = str(row.get("Topic") or "").strip()
        stimulus = str(row.get("Stimulus_Text") or "").strip()
        stem = str(row.get("Item_Stem") or "").strip()
        opt_a = str(row.get("Option_A") or "").strip()
        opt_b = str(row.get("Option_B") or "").strip()
        opt_c = str(row.get("Option_C") or "").strip()
        opt_d = str(row.get("Option_D") or "").strip()
        correct_opt = str(row.get("Correct_Option") or "").strip().upper()
        correct_ans = str(row.get("Correct_Answer") or "").strip()
        
        # ─── TOPIC CLASSIFICATION & ALIGNMENT ─────────────────────────────────
        pt = classify_topic_with_ai(client, stem, stimulus, predefined_topics, row_topic=topic)
        item_topic = pt['Topic']
        item_topic_id = pt['Topic ID']
        
        combo = get_framework_alignment_for_topic(item_topic, matched_chapter, unique_combos)
        if combo:
            item_ncf_cg = combo['ncf_cg']
            item_competency = combo['competency']
            item_learning_outcome = combo['learning_outcome']
            item_lo_id = compute_lo_id(combo, unique_combos)
        else:
            item_ncf_cg = "CG-1"
            item_competency = "C-1.1"
            item_learning_outcome = "Identifies and explains different properties of materials and relates them to their uses."
            item_lo_id = "LO-1.1.a"

        # Temp ID for display
        temp_id = f"G6-{batch_meta['chapter_code']}-{item_topic_id}-{mastery_excel or 'MX'}-XXX"
        
        # Log update to Streamlit UI
        if progress_callback:
            progress_callback(i + 1, total_valid, temp_id, mastery_excel or "MX", "Classifying and validating cognitive demand...")
            
        # ─── CALL 1: MASTERY LEVEL CLASSIFICATION ──────────────────────────────
        confirmed_mastery = "M2" # Fallback default
        override_msg = ""
        
        if is_mock:
            c1_res = generate_mock_mastery(row, inferred_item_type)
        else:
            try:
                system_prompt_c1 = (
                    "You are an expert assessment designer for the NCF 2023 Grade 6 Science curriculum in India.\n"
                    "You classify assessment items into one of four Mastery Levels based strictly on the cognitive\n"
                    "demand required to answer correctly — not on topic difficulty.\n\n"
                    "MASTERY LEVEL DEFINITIONS:\n\n"
                    "M1 (Remembering, DoK 1):\n"
                    "  The learner identifies, names, or recognises from a given set. No explanation required.\n"
                    "  Direct recall or recognition. ALWAYS MCQ or matching format.\n\n"
                    "M2 (Understanding, DoK 2):\n"
                    "  The learner explains, classifies, or describes using scientific reasoning in a familiar,\n"
                    "  structured context. Requires understanding, not just recall. MCQ with justification or Short Answer.\n\n"
                    "M3 (Applying, DoK 3):\n"
                    "  The learner solves an unfamiliar scenario, applies knowledge to a novel situation, or reasons\n"
                    "  through a non-obvious case. The stimulus introduces context not directly taught.\n"
                    "  Can be scenario-based MCQ or Short Answer / Extended Response.\n\n"
                    "M4 (Evaluating, DoK 3+):\n"
                    "  The learner evaluates a flawed claim, incorrect reasoning, or competing explanation in the stimulus,\n"
                    "  identifies the specific error in REASONING (not just facts), and constructs a justified correction.\n"
                    "  ALWAYS Extended Response. The flaw must be in the reasoning, not just the outcome.\n\n"
                    "CLASSIFICATION RULES:\n"
                    "1. Base classification on the cognitive demand of STEM + STIMULUS together, not topic difficulty.\n"
                    "2. An item answerable purely from memory → M1 or M2, never M3 or M4.\n"
                    "3. M3 requires an unfamiliar scenario; applying knowledge in a familiar context → M2.\n"
                    "4. M4 REQUIRES a flawed reasoning claim in the stimulus. No flaw → cannot be M4.\n"
                    "5. If operator supplied a Mastery_Level, validate it. If incorrect, override and explain.\n\n"
                    "OUTPUT: JSON only, no other text:\n"
                    '{"mastery_level": "M1", "reasoning": "One sentence justifying the classification.", "operator_override": false}'
                )
                
                correct_ans_sum = row["Correct_Option"] if inferred_item_type == "MCQ" else row["Correct_Answer"]
                user_msg_c1 = (
                    f"Item Type (inferred): {inferred_item_type}\n"
                    f"Operator-supplied Mastery Level: {mastery_excel if mastery_excel else 'Not provided — please classify.'}\n"
                    f"Stimulus: {stimulus if stimulus else 'None'}\n"
                    f"Item Stem: {stem}\n"
                    f"Correct Answer / Option: {correct_ans_sum}"
                )
                
                c1_res = call_mistral(client, system_prompt_c1, user_msg_c1, max_tokens=150, use_json=True)
            except Exception as exc:
                c1_res = {"mastery_level": "M2", "reasoning": f"Fallback to M2 due to API error: {exc}", "operator_override": False}
                
        confirmed_mastery = c1_res.get("mastery_level", "M2").upper()
        if confirmed_mastery not in ("M1", "M2", "M3", "M4"):
            confirmed_mastery = "M2"
            
        if mastery_excel and confirmed_mastery != mastery_excel:
            override_msg = f"Override: Operator suggested {mastery_excel}, AI classified as {confirmed_mastery}."
            override_count += 1
            
        # ─── CONSTRAINT ENFORCEMENT & OVERRIDES (Rule 3.12) ───────────────────
        item_type = inferred_item_type
        
        if confirmed_mastery == "M4" and item_type != "Extended Response":
            item_type = "Extended Response"
            msg = "M4 items must be Extended Response. Inferred item type overridden to Extended Response."
            override_msg = (override_msg + " " + msg).strip()
            override_count += 1
            
        elif confirmed_mastery == "M1" and item_type != "MCQ":
            item_type = "MCQ"
            msg = "M1 items must be MCQ. Inferred item type overridden to MCQ."
            override_msg = (override_msg + " " + msg).strip()
            override_count += 1
            
        elif confirmed_mastery == "M2" and item_type == "Extended Response":
            item_type = "Short Answer"
            msg = "M2 items can only be MCQ or Short Answer. Inferred item type overridden to Short Answer."
            override_msg = (override_msg + " " + msg).strip()
            override_count += 1
            
        # Warn for missing stimulus on M3/M4 items
        if confirmed_mastery in ("M3", "M4") and not stimulus:
            msg = f"WARNING: {confirmed_mastery} item is missing a stimulus."
            override_msg = (override_msg + " " + msg).strip()
            
        # Resolve sequence number for this topic
        if item_topic_id not in current_seq:
            current_seq[item_topic_id] = get_next_sequence_number(batch_meta['chapter_code'], item_topic_id)
        item_seq = current_seq[item_topic_id]
        current_seq[item_topic_id] += 1
        
        # Finalized Item ID
        item_id = f"G6-{batch_meta['chapter_code']}-{item_topic_id}-{confirmed_mastery}-{item_seq:03d}"
        
        # ─── CALL 2: FORMAT FIELD INFERENCE ──────────────────────────────────
        if progress_callback:
            progress_callback(i + 1, total_valid, item_id, confirmed_mastery, "Inferring rendering properties...")
            
        if is_mock:
            c2_res = generate_mock_format_fields(row, item_type)
        else:
            try:
                system_prompt_c2 = (
                    "You are a technical data-entry assistant for an educational assessment platform.\n"
                    "Analyse the item content and determine its rendering properties.\n\n"
                    "Has_Image:\n"
                    "  YES if stimulus or stem explicitly references an image, photograph, diagram, figure, picture,\n"
                    "  or contains image description text (e.g., 'Image description: ...'). NO otherwise.\n\n"
                    "Has_Table:\n"
                    "  YES if stimulus or stem contains a data table or references tabular data explicitly. NO otherwise.\n\n"
                    "Has_Equation:\n"
                    "  YES if the item contains mathematical symbols, chemical formulae, physical quantities in symbolic form\n"
                    "  (e.g., H₂O, CO₂, 37°C, m/s², F=ma), or scientific notation requiring special rendering.\n"
                    "  NO for plain-text descriptions of formulae.\n\n"
                    "Equation_Format:\n"
                    "  If Has_Equation is YES → 'LaTeX' if expressible cleanly in LaTeX; 'Image' if it requires\n"
                    "  a diagram or complex visual. 'N/A' if Has_Equation is NO.\n\n"
                    "OUTPUT: JSON only, no other text:\n"
                    '{"has_image": "Yes", "has_table": "No", "has_equation": "No", "equation_format": "N/A"}'
                )
                user_msg_c2 = (
                    f"Stimulus: {stimulus if stimulus else 'None'}\n"
                    f"Item Stem: {stem}\n"
                    f"Option A: {row.get('Option_A') if item_type == 'MCQ' else 'N/A'}\n"
                    f"Option B: {row.get('Option_B') if item_type == 'MCQ' else 'N/A'}\n"
                    f"Option C: {row.get('Option_C') if item_type == 'MCQ' else 'N/A'}\n"
                    f"Option D: {row.get('Option_D') if item_type == 'MCQ' else 'N/A'}"
                )
                c2_res = call_mistral(client, system_prompt_c2, user_msg_c2, max_tokens=120, use_json=True)
            except Exception as exc:
                c2_res = {"has_image": "No", "has_table": "No", "has_equation": "No", "equation_format": "N/A"}
                
        has_image = str(c2_res.get("has_image", "No")).strip().capitalize()
        has_table = str(c2_res.get("has_table", "No")).strip().capitalize()
        has_equation = str(c2_res.get("has_equation", "No")).strip().capitalize()
        
        # Clean up eq_format to be LaTeX or Image if Has_Equation is Yes, else blank
        eq_format_raw = str(c2_res.get("equation_format", "")).strip()
        if has_equation == "Yes":
            if "latex" in eq_format_raw.lower():
                eq_format = "LaTeX"
            elif "image" in eq_format_raw.lower():
                eq_format = "Image"
            else:
                eq_format = "LaTeX" # default fallback
        else:
            eq_format = ""
        
        # ─── CALL 3: DISTRACTOR RATIONALE GENERATION (MCQ Only) ───────────────
        distractor_rationales = {}
        if item_type == "MCQ":
            if progress_callback:
                progress_callback(i + 1, total_valid, item_id, confirmed_mastery, "Generating distractor rationales...")
                
            if is_mock:
                distractor_rationales = generate_mock_rationales(row, correct_opt)
            else:
                try:
                    system_prompt_c3 = (
                        "You are an expert assessment designer specialising in Grade 6 Science for the NCF 2023\n"
                        "curriculum in India. You write distractor rationales that explain the specific misconception\n"
                        "a learner must hold in order to choose a wrong MCQ option.\n\n"
                        "A strong distractor rationale:\n"
                        "  - Names the specific misconception (e.g., 'cross-category confusion,' 'within-category\n"
                        "    opposite,' 'overgeneralisation of a rule')\n"
                        "  - Explains the cognitive error in concrete terms\n"
                        "  - Is 1–2 sentences maximum\n"
                        "  - Is written for the item author, not the student\n\n"
                        "For the CORRECT answer, write 1 sentence beginning with 'CORRECT ANSWER:' that briefly\n"
                        "explains why it is correct.\n\n"
                        "OUTPUT: JSON only, no other text:\n"
                        '{\n'
                        '  "rationale_a": "...",\n'
                        '  "rationale_b": "...",\n'
                        '  "rationale_c": "...",\n'
                        '  "rationale_d": "..."\n'
                        '}'
                    )
                    
                    user_msg_c3 = (
                        f"Topic: {topic}\n"
                        f"Item Stem: {stem}\n"
                        f"Option A: {row.get('Option_A')} (Correct: {'Yes' if correct_opt == 'A' else 'No'})\n"
                        f"Option B: {row.get('Option_B')} (Correct: {'Yes' if correct_opt == 'B' else 'No'})\n"
                        f"Option C: {row.get('Option_C')} (Correct: {'Yes' if correct_opt == 'C' else 'No'})\n"
                        f"Option D: {row.get('Option_D')} (Correct: {'Yes' if correct_opt == 'D' else 'No'})\n"
                    )
                    
                    distractor_rationales = call_mistral(client, system_prompt_c3, user_msg_c3, max_tokens=600, use_json=True)
                except Exception:
                    # Fallback manually
                    distractor_rationales = generate_mock_rationales(row, correct_opt)
                    
        # ─── CALL 4: ANSWER EXPLANATION GENERATION ────────────────────────────
        if progress_callback:
            progress_callback(i + 1, total_valid, item_id, confirmed_mastery, "Generating answer explanation...")
            
        if is_mock:
            explanation_text = generate_mock_explanation(row, item_type, correct_opt, distractor_rationales)
        else:
            try:
                system_prompt_c4 = (
                    "You are a science teacher for Grade 6 students (approximately 11-year-olds) in India,\n"
                    "writing feedback that will be shown on an educational platform.\n\n"
                    "Your explanation must be:\n"
                    "  - Clear and accessible for Grade 6 students (short sentences, everyday analogies where helpful)\n"
                    "  - Scientifically accurate and aligned with the NCF 2023 Grade 6 Science curriculum\n"
                    "  - Complete: it must explain the correct answer AND (for MCQ) why each wrong answer is wrong\n\n"
                    "FOR MCQ ITEMS — Required structure:\n"
                    "  Paragraph 1: Explain the correct answer clearly (2-4 sentences).\n"
                    "  Then add: \"Why other options are incorrect:\"\n"
                    "  Option [letter] ([text]) is incorrect because [specific reason — name the conceptual error].\n"
                    "  Repeat for each incorrect option.\n\n"
                    "FOR SHORT ANSWER / EXTENDED RESPONSE — Required structure:\n"
                    "  Paragraph explaining the model answer (2-4 sentences covering all key ideas).\n"
                    "  Then: what students who partially answer or miss the answer are likely confusing.\n\n"
                    "LANGUAGE: Grade 6–8 level. No jargon without explanation. Use analogies from daily life.\n\n"
                    "OUTPUT: Plain text only. No JSON. No headers. No markdown. Just the explanation text."
                )
                
                correct_ans_or_opt_text = row["Correct_Option"] + ". " + str(row.get(f"Option_{correct_opt}")) if item_type == "MCQ" else row["Correct_Answer"]
                
                user_msg_c4 = (
                    f"Item Type: {item_type}\n"
                    f"Topic: {topic}\n"
                    f"Stimulus: {stimulus if stimulus else 'None'}\n"
                    f"Item Stem: {stem}\n"
                    f"Correct Answer: {correct_ans_or_opt_text}\n"
                    + (f"Option A: {row.get('Option_A')}\nOption B: {row.get('Option_B')}\nOption C: {row.get('Option_C')}\nOption D: {row.get('Option_D')}\n" if item_type == "MCQ" else "")
                    + (f"Distractor rationales: {json.dumps(distractor_rationales)}" if item_type == "MCQ" else "")
                )
                
                explanation_text = call_mistral(client, system_prompt_c4, user_msg_c4, max_tokens=500, use_json=False)
            except Exception:
                explanation_text = generate_mock_explanation(row, item_type, correct_opt, distractor_rationales)
                
        # ─── DETERMINISTIC RULES DERIVATIONS ──────────────────────────────────
        blooms = {
            "M1": "Remembering",
            "M2": "Understanding",
            "M3": "Applying",
            "M4": "Evaluating"
        }.get(confirmed_mastery, "Understanding")
        
        dok = {
            "M1": "1",
            "M2": "2",
            "M3": "3",
            "M4": "3"
        }.get(confirmed_mastery, "2")
        
        # Est Time
        if item_type == "MCQ":
            est_time = "1 min" if confirmed_mastery == "M1" else "2 mins"
        elif item_type == "Short Answer":
            est_time = "2 mins"
        else: # Extended Response
            est_time = "5 mins"
            
        # Max Score
        if item_type == "MCQ":
            max_score = 1
        elif item_type == "Short Answer":
            max_score = 2 if confirmed_mastery == "M2" else 3
        else: # Extended Response
            max_score = 3 if confirmed_mastery == "M3" else 4
            
        # Scoring Type
        scoring_type = "Dichotomous (0/1)" if item_type == "MCQ" else f"Polytomous (0-{max_score})"
        
        # Scale
        if item_type == "MCQ":
            scale = ""
        else:
            scale = ", ".join(str(s) for s in range(max_score + 1))
            
        # Prerequisite Level
        prereq = {
            "M1": "",
            "M2": "M1",
            "M3": "M2",
            "M4": "M3"
        }.get(confirmed_mastery, "")
        
        # Options Count
        num_options = "4" if item_type == "MCQ" else ""
        
        # Feedback type
        feedback_type = "Immediate" if item_type == "MCQ" else "Delayed"
        
        # ─── CALL 5: MARKING RUBRIC GENERATION (Non-MCQ Only) ─────────────────
        rubric_data = {"rows": []}
        if item_type != "MCQ":
            if progress_callback:
                progress_callback(i + 1, total_valid, item_id, confirmed_mastery, "Generating marking rubric...")
                
            if is_mock:
                rubric_data = generate_mock_rubric(row, max_score)
            else:
                try:
                    system_prompt_c5 = (
                        "You are a marking rubric designer for Grade 6 Science assessments aligned to NCF 2023 India.\n"
                        "Your rubrics follow the partial-credit scoring model: binary (full or zero) marking is\n"
                        "NOT acceptable. Every rubric must have a score tier for each point from 0 to the max score.\n\n"
                        "RUBRIC TIER DEFINITIONS (apply based on max score):\n\n"
                        "FOR MAX SCORE = 2 (Short Answer, M2) — 3 rows:\n"
                        "  2 marks (Full): Student correctly identifies BOTH required scientific properties/concepts\n"
                        "    using appropriate terminology AND demonstrates understanding of the relationship or application.\n"
                        "  1 mark (Partial): Student correctly identifies ONE property/concept using appropriate\n"
                        "    terminology; OR identifies both but uses informal/incorrect terminology; OR demonstrates\n"
                        "    partial understanding of only one aspect.\n"
                        "  0 marks (Zero): Irrelevant answer, fundamental misconception, or purely observational\n"
                        "    response with no scientific reasoning.\n\n"
                        "FOR MAX SCORE = 3 (Short Answer or Extended Response, M3) — 4 rows:\n"
                        "  3 marks (Full): Student correctly applies the relevant concept(s) to the given scenario,\n"
                        "    accurately predicts/explains the outcome, uses precise scientific terminology, AND demonstrates\n"
                        "    clear logical reasoning with no conceptual errors.\n"
                        "  2 marks (Substantial): Correctly applies the main concept(s) with minor omissions; OR correct\n"
                        "    outcome stated with mostly complete reasoning; OR accurate in 2 out of 3 required elements.\n"
                        "  1 mark (Basic): Some relevant scientific understanding shown but concept applied incorrectly\n"
                        "    or partially; OR correct outcome stated without reasoning; OR reasoning stated without\n"
                        "    a correct outcome.\n"
                        "  0 marks (Zero): Irrelevant, fundamentally incorrect, or blank.\n\n"
                        "FOR MAX SCORE = 4 (Extended Constructed Response, M4) — 5 rows:\n"
                        "  4 marks (Full): Student evaluates the flawed claim comprehensively: identifies the specific\n"
                        "    error in reasoning, provides a scientifically accurate correction supported by at least two\n"
                        "    relevant properties, AND articulates why the original claim fails against the task's actual\n"
                        "    requirements.\n"
                        "  3 marks (Strong): Identifies the flawed reasoning AND provides accurate correction with one\n"
                        "    key property clearly justified. Minor gap in a second supporting property, but core argument\n"
                        "    is scientifically sound.\n"
                        "  2 marks (Moderate): States what went wrong practically (bag broke, material failed) but does\n"
                        "    not explicitly frame it as a flaw in scientific reasoning; OR provides a correction but\n"
                        "    cannot clearly explain why the original claim was wrong.\n"
                        "  1 mark (Minimal): Recognises the claim is incorrect and names at least one relevant property,\n"
                        "    but provides no justification or uses the property incorrectly.\n"
                        "  0 marks (Zero): Defends the flawed claim, provides entirely irrelevant content, or is blank.\n\n"
                        "INSTRUCTIONS:\n"
                        "  - Write rubric rows that are SPECIFIC to this question — not generic.\n"
                        "  - Each row must contain: Criteria (what must be present) + Sample Response (a realistic\n"
                        "    example of a student answer at that tier, 1–3 sentences).\n"
                        "  - Sample responses must sound like real Grade 6 students, not model adults.\n"
                        "  - The sample response at 0 marks must show a specific realistic misconception.\n\n"
                        "OUTPUT: JSON only:\n"
                        "{\n"
                        '  "rows": [\n'
                        '    {"score": 2, "label": "Full marks", "criteria": "...", "sample": "..."},\n'
                        '    {"score": 1, "label": "Partial marks", "criteria": "...", "sample": "..."},\n'
                        '    {"score": 0, "label": "Zero", "criteria": "...", "sample": "..."}\n'
                        '  ]\n'
                        "}\n"
                        "Score values must match: 0 through max_score (descending order).\n"
                        'Label values: "Full marks", "Strong partial" (M4 only), "Substantial" (M3 only),\n'
                        '  "Partial marks", "Basic", "Minimal" (M4 only), "Zero".'
                    )
                    
                    user_msg_c5 = (
                        f"Item Type: {item_type}\n"
                        f"Mastery Level: {confirmed_mastery}\n"
                        f"Max Score: {max_score}\n"
                        f"Topic: {topic}\n"
                        f"Item Stem: {stem}\n"
                        f"Stimulus: {stimulus if stimulus else 'None'}\n"
                        f"Model Answer / Correct Response: {row.get('Correct_Answer')}"
                    )
                    
                    rubric_data = call_mistral(client, system_prompt_c5, user_msg_c5, max_tokens=800, use_json=True)
                except Exception:
                    rubric_data = generate_mock_rubric(row, max_score)
                    
            # Rubric Quality Check (Rule 7.6)
            rows = rubric_data.get("rows", [])
            valid_rows = []
            expected_scores = list(range(max_score, -1, -1))
            
            # Rebuild clean rows based on expected scores to guarantee correct structure
            for exp_score in expected_scores:
                # Find matching row from LLM output
                matching_row = None
                for r in rows:
                    if int(r.get("score", -1)) == exp_score:
                        matching_row = r
                        break
                        
                if matching_row and matching_row.get("criteria") and matching_row.get("sample"):
                    valid_rows.append(matching_row)
                else:
                    # Fallback criteria row
                    lbl = {
                        4: "Full marks",
                        3: "Substantial" if max_score == 3 else "Strong partial",
                        2: "Moderate partial" if max_score == 4 else ("Full marks" if max_score == 2 else "Substantial"),
                        1: "Partial marks" if max_score == 2 else "Basic",
                        0: "Zero"
                    }.get(exp_score, "Criteria tier")
                    
                    valid_rows.append({
                        "score": exp_score,
                        "label": lbl,
                        "criteria": "[RUBRIC ROW INCOMPLETE — PLEASE REVIEW]",
                        "sample": "[Sample response required]"
                    })
                    override_msg = (override_msg + f" WARNING: Rubric row for score {exp_score} missing or invalid. Created fallback.").strip()
                    
            rubric_data["rows"] = valid_rows
            
        # Compile item dictionary
        enriched_item = {
            "Item_ID": item_id,
            "Mastery_Level": confirmed_mastery,
            "Blooms_Level": blooms,
            "DoK_Level": dok,
            "Estimated_Time": est_time,
            "Max_Score": max_score,
            "Scoring_Type": scoring_type,
            "If_Partial_Scale": scale,
            "Prerequisite_Level": prereq,
            "Item_Type": item_type,
            "No_Of_Options": num_options,
            "Has_Image": has_image,
            "Image_File_Name": f"{item_id}-img01.png" if has_image == "Yes" else "",
            "Has_Table": has_table,
            "Has_Equation": has_equation,
            "Equation_Format": eq_format if has_equation == "Yes" else "",
            "Feedback_Type": feedback_type,
            "Topic": item_topic,
            "Topic_ID": item_topic_id,
            "Subject_ID": subject_id,
            "NCF_CG": item_ncf_cg,
            "Competency": item_competency,
            "Learning_Outcome": item_learning_outcome,
            "LO_ID": item_lo_id,
            "Stimulus_Text": stimulus,
            "Item_Stem": stem,
            "Explanation": explanation_text,
            "Rubric": rubric_data,
            # Original excel info
            "Option_A": opt_a,
            "Option_B": opt_b,
            "Option_C": opt_c,
            "Option_D": opt_d,
            "Correct_Option": correct_opt,
            "Correct_Answer": correct_ans,
            "Distractor_Rationales": distractor_rationales,
            "Row_Number": excel_row_num
        }
        
        processed_items.append(enriched_item)
        
        # Log successful audit entry
        status_label = "WARNING" if "WARNING" in override_msg else ("NOTICE" if override_msg else "SUCCESS")
        audit_logs.append({
            "Row_Number": excel_row_num,
            "Item_ID": item_id,
            "Status": status_label,
            "Message": f"Successfully processed item. {override_msg}".strip(),
            "Mastery_Level": confirmed_mastery,
            "Item_Type": item_type,
            "Overrides": override_msg
        })
        
        pass
        
    # ─── 3. GENERATE DOCX FILE ────────────────────────────────────────────────
    if progress_callback:
        progress_callback(total_valid, total_valid, "Finalizing...", "N/A", "Compiling Microsoft Word document (.docx)...")
        
    doc = Document()
    
    # Section margins: 1.5875 cm = 0.625 inches
    section = doc.sections[0]
    section.top_margin = Inches(0.625)
    section.bottom_margin = Inches(0.625)
    section.left_margin = Inches(0.625)
    section.right_margin = Inches(0.625)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    
    # Configure Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    
    today_str = datetime.today().strftime("%d/%m/%Y")
    
    for item_idx, item in enumerate(processed_items):
        # Table 1: PART A Header
        t1 = create_styled_table(doc, 1, [9360], border_color="999999")
        set_cell_text(t1.rows[0].cells[0], "PART A — ITEM METADATA  (Tech: used for bulk import and adaptive routing)", bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Subsection Header
        add_section_header(doc, "A1.  Item Identity")
        
        # Table 2: Identity Metadata
        t2 = create_styled_table(doc, 3, [1600, 3080, 1600, 3080])
        labels_t2 = [
            [("Item ID *", True), (item["Item_ID"], False), ("Version", True), ("v1.0", False)],
            [("Status *", True), ("Draft", False), ("Date", True), (today_str, False)],
            [("Authored by *", True), ("TechCurators", False), ("Reviewed by", True), ("", False)]
        ]
        for r_idx, row_labels in enumerate(labels_t2):
            for c_idx, (text, is_label) in enumerate(row_labels):
                cell = t2.rows[r_idx].cells[c_idx]
                if is_label:
                    set_cell_shading(cell, "e2e8f0")
                elif r_idx == 2 and c_idx == 3: # Reviewed by empty value cell
                    set_cell_shading(cell, "f4f6f9")
                set_cell_text(cell, text, bold=is_label)
                
        # Subsection Header
        add_section_header(doc, "A2.  Curriculum Alignment")
        
        # Table 3: Alignment Metadata
        t3 = create_styled_table(doc, 5, [1600, 3080, 1600, 3080])
        labels_t3 = [
            [("Grade *", True), ("6", False), ("Subject ID *", True), (item["Subject_ID"], False)],
            [("Chapter *", True), (batch_meta["chapter_name"], False), ("Chapter Code *", True), (batch_meta["chapter_code"], False)],
            [("Topic *", True), (item["Topic"], False), ("Topic ID *", True), (item["Topic_ID"], False)],
            [("NCF CG #", True), (item["NCF_CG"], False), ("Competency", True), (item["Competency"], False)],
            [("Learning Outcome *", True), (item["Learning_Outcome"], False), ("LO ID", True), (item["LO_ID"], False)]
        ]
        for r_idx, row_labels in enumerate(labels_t3):
            for c_idx, (text, is_label) in enumerate(row_labels):
                cell = t3.rows[r_idx].cells[c_idx]
                if is_label:
                    set_cell_shading(cell, "e2e8f0")
                set_cell_text(cell, text, bold=is_label)
                
        # Subsection Header
        add_section_header(doc, "A3.  Adaptive Routing  ", "(Tech: drives mastery progression logic)")
        
        # Table 4: Adaptive Routing
        t4 = create_styled_table(doc, 4, [1600, 3080, 1600, 3080])
        labels_t4 = [
            [("Mastery Level *", True), (item["Mastery_Level"], False), ("Bloom's Level *", True), (item["Blooms_Level"], False)],
            [("DoK Level *", True), (item["DoK_Level"], False), ("Estimated Time *", True), (item["Estimated_Time"], False)],
            [("Max Score *", True), (str(item["Max_Score"]), False), ("Scoring Type *", True), (item["Scoring_Type"], False)],
            [("If Partial — Scale", True), (item["If_Partial_Scale"], False), ("Prerequisite Level", True), (item["Prerequisite_Level"], False)]
        ]
        for r_idx, row_labels in enumerate(labels_t4):
            for c_idx, (text, is_label) in enumerate(row_labels):
                cell = t4.rows[r_idx].cells[c_idx]
                if is_label:
                    set_cell_shading(cell, "e2e8f0")
                set_cell_text(cell, text, bold=is_label)
                
        # Subsection Header
        add_section_header(doc, "A4.  Item Format & Rendering  ", "(Tech: platform display and scoring engine)")
        
        # Table 5: Item Format
        t5 = create_styled_table(doc, 4, [1600, 3080, 1600, 3080])
        # Format Item Type name for table display (Rule 3.12 mapping)
        display_type = "Extended Constructed Response" if item["Item_Type"] == "Extended Response" else item["Item_Type"]
        labels_t5 = [
            [("Item Type *", True), (display_type, False), ("No. of Options", True), (item["No_Of_Options"], False)],
            [("Has Image? *", True), (item["Has_Image"], False), ("Image File Name", True), (item["Image_File_Name"], False)],
            [("Has Table? *", True), (item["Has_Table"], False), ("Has Equation? *", True), (item["Has_Equation"], False)],
            [("Equation Format", True), (item["Equation_Format"], False), ("Feedback Type *", True), (item["Feedback_Type"], False)]
        ]
        for r_idx, row_labels in enumerate(labels_t5):
            for c_idx, (text, is_label) in enumerate(row_labels):
                cell = t5.rows[r_idx].cells[c_idx]
                if is_label:
                    set_cell_shading(cell, "e2e8f0")
                set_cell_text(cell, text, bold=is_label)
                
        # Subsection Header
        add_section_header(doc, "A5.  Quality & Review Flags")
        
        # Table 6: Quality Flags
        t6 = create_styled_table(doc, 2, [1600, 3080, 1600, 3080])
        labels_t6 = [
            [("Reviewed for bias?", True), ("Pending", False), ("Reviewed for accuracy?", True), ("Pending", False)],
            [("NCF scope verified?", True), ("Pending", False), ("Cultural check?", True), ("Pending", False)]
        ]
        for r_idx, row_labels in enumerate(labels_t6):
            for c_idx, (text, is_label) in enumerate(row_labels):
                cell = t6.rows[r_idx].cells[c_idx]
                if is_label:
                    set_cell_shading(cell, "e2e8f0")
                set_cell_text(cell, text, bold=is_label)
                
        # Table 7: Quality flags guidance
        t7 = create_styled_table(doc, 1, [9360], border_color="2e5f8a")
        set_cell_shading(t7.rows[0].cells[0], "f4f6f9")
        set_cell_text(t7.rows[0].cells[0], "All four quality flags must show 'Yes' before an item can be marked Approved.", italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Spacing element
        doc.add_paragraph().paragraph_format.space_before = Pt(4)
        
        # Table 8: PART B Header
        t8 = create_styled_table(doc, 1, [9360], border_color="999999")
        set_cell_text(t8.rows[0].cells[0], "PART B — ITEM CONTENT", bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Subsection Header
        add_section_header(doc, "B1.  Stimulus  ", "(the scenario, image, data, or text the learner reads before the question)")
        
        # Table 9: Stimulus guidance
        t9 = create_styled_table(doc, 1, [9360], border_color="2e5f8a")
        set_cell_shading(t9.rows[0].cells[0], "f4f6f9")
        set_cell_text(t9.rows[0].cells[0], "Leave blank if no stimulus. For M1 items a stimulus is often not needed. For M3 and M4 a well-designed stimulus is mandatory.", italic=True)
        
        # Table 10: Stimulus text
        t10 = create_styled_table(doc, 1, [2000, 7360])
        set_cell_shading(t10.rows[0].cells[0], "e2e8f0")
        set_cell_text(t10.rows[0].cells[0], "Stimulus text", bold=True)
        set_cell_text(t10.rows[0].cells[1], item["Stimulus_Text"])
        
        # Subsection Header
        add_section_header(doc, "B2.  Item Stem  ", "(the question or instruction the learner must respond to)")
        
        # Table 11: Stem text
        t11 = create_styled_table(doc, 1, [2000, 7360])
        set_cell_shading(t11.rows[0].cells[0], "e2e8f0")
        set_cell_text(t11.rows[0].cells[0], "Item stem *", bold=True)
        set_cell_text(t11.rows[0].cells[1], item["Item_Stem"])
        
        # Subsection Header
        add_section_header(doc, "B3.  Response Options  ", "(for MCQ only — skip for Short Answer / Extended Response)")
        
        # Table 12: Options (Borders 999999 on all cells)
        t12 = create_styled_table(doc, 5, [877, 1059, 7424], border_color="999999")
        
        # Row 1 headers
        set_cell_text(t12.rows[0].cells[0], "Option", bold=True)
        set_cell_text(t12.rows[0].cells[1], "Correct?", bold=True)
        set_cell_text(t12.rows[0].cells[2], "Option Text", bold=True)
        
        if item["Item_Type"] == "MCQ":
            for r_idx, letter in enumerate(["A", "B", "C", "D"], start=1):
                is_correct = "Yes" if item["Correct_Option"] == letter else "No"
                opt_val = item[f"Option_{letter}"]
                
                # Combine option text and distractor rationale in cell
                rat = item["Distractor_Rationales"].get(f"rationale_{letter.lower()}", "")
                combined_opt_text = f"{opt_val}\n{rat}"
                
                set_cell_text(t12.rows[r_idx].cells[0], letter, bold=True)
                set_cell_text(t12.rows[r_idx].cells[1], is_correct)
                set_cell_text(t12.rows[r_idx].cells[2], combined_opt_text)
        else:
            # Render empty MCQ table for non-MCQ
            for r_idx, letter in enumerate(["A", "B", "C", "D"], start=1):
                set_cell_text(t12.rows[r_idx].cells[0], letter, bold=True)
                set_cell_text(t12.rows[r_idx].cells[1], "")
                set_cell_text(t12.rows[r_idx].cells[2], "")
                
        # Subsection Header
        add_section_header(doc, "B4.  Correct Answer / Expected Response *")
        
        # Table 13: Correct answer
        t13 = create_styled_table(doc, 1, [2000, 7360])
        set_cell_shading(t13.rows[0].cells[0], "e2e8f0")
        set_cell_text(t13.rows[0].cells[0], "Correct answer", bold=True)
        
        if item["Item_Type"] == "MCQ":
            correct_val = item[f"Option_{item['Correct_Option']}"]
            ans_text = f"{item['Correct_Option']}. {correct_val}"
        else:
            ans_text = item["Correct_Answer"]
            
        set_cell_text(t13.rows[0].cells[1], ans_text)
        
        # Subsection Header
        add_section_header(doc, "B5.  Answer Explanation  ", "(shown to learner as feedback on the platform)")
        
        # Table 14: Explanation
        t14 = create_styled_table(doc, 1, [2000, 7360])
        set_cell_shading(t14.rows[0].cells[0], "e2e8f0")
        set_cell_text(t14.rows[0].cells[0], "Explanation *", bold=True)
        set_cell_text(t14.rows[0].cells[1], item["Explanation"])
        
        # Subsection Header
        add_section_header(doc, "B6.  Marking Rubric  ", "(for Short Answer and Extended Response only)")
        
        # Table 15: Rubric guidance
        t15 = create_styled_table(doc, 1, [9360], border_color="2e5f8a")
        set_cell_shading(t15.rows[0].cells[0], "f4f6f9")
        set_cell_text(t15.rows[0].cells[0], "Skip B6 for MCQ items. For Short and Extended Answer: partial scoring is required — binary marking (full or zero) is not acceptable.", italic=True)
        
        # Table 16: Rubric Table (Borders 999999 on all cells)
        rubric_rows_data = item["Rubric"].get("rows", [])
        num_rubric_rows = len(rubric_rows_data) + 1 # +1 for header
        
        if item["Item_Type"] == "MCQ" or num_rubric_rows == 1:
            # MCQ requires 3 blank rows under the header row (4 rows total)
            t16 = create_styled_table(doc, 4, [838, 1562, 6960], border_color="999999")
            set_cell_text(t16.rows[0].cells[0], "Score", bold=True)
            set_cell_text(t16.rows[0].cells[1], "Label", bold=True)
            set_cell_text(t16.rows[0].cells[2], "Criteria — what must be present for this score + sample response", bold=True)
            for r_idx in range(1, 4):
                set_cell_text(t16.rows[r_idx].cells[0], "")
                set_cell_text(t16.rows[r_idx].cells[1], "")
                set_cell_text(t16.rows[r_idx].cells[2], "")
        else:
            t16 = create_styled_table(doc, num_rubric_rows, [838, 1562, 6960], border_color="999999")
            set_cell_text(t16.rows[0].cells[0], "Score", bold=True)
            set_cell_text(t16.rows[0].cells[1], "Label", bold=True)
            set_cell_text(t16.rows[0].cells[2], "Criteria — what must be present for this score + sample response", bold=True)
            for r_idx, r_val in enumerate(rubric_rows_data, start=1):
                criteria_combined = f"Criteria: {r_val['criteria']}\nSample Response: {r_val['sample']}"
                set_cell_text(t16.rows[r_idx].cells[0], str(r_val["score"]))
                set_cell_text(t16.rows[r_idx].cells[1], r_val["label"])
                set_cell_text(t16.rows[r_idx].cells[2], criteria_combined)
                
        # Subsection Header
        add_section_header(doc, "B7.  Reviewer Notes  ", "(internal — not shown to learner)")
        
        # Table 17: Notes
        t17 = create_styled_table(doc, 1, [2000, 7360])
        set_cell_shading(t17.rows[0].cells[0], "e2e8f0")
        set_cell_text(t17.rows[0].cells[0], "Notes", bold=True)
        set_cell_text(t17.rows[0].cells[1], "") # Always empty cell on generation
        
        # Table 18: END OF ITEM
        t18 = create_styled_table(doc, 1, [10298], border_color="2e5f8a")
        set_cell_shading(t18.rows[0].cells[0], "f4f6f9")
        set_cell_text(t18.rows[0].cells[0], "END OF ITEM  ·  Duplicate this page for the next item  ·  Do not modify template structure", italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Page break after item (except the last one)
        if item_idx < len(processed_items) - 1:
            p_break = doc.add_paragraph()
            p_break.paragraph_format.space_before = Pt(0)
            p_break.paragraph_format.space_after = Pt(0)
            r_break = p_break.add_run()
            r_break.add_break(WD_BREAK.PAGE)
            
    # Save DOCX to Bytes
    doc_io = io.BytesIO()
    doc.save(doc_io)
    docx_bytes = doc_io.getvalue()
    
    # ─── 4. GENERATE CSV AUDIT LOG ─────────────────────────────────────────────
    csv_io = io.StringIO()
    csv_writer = csv.DictWriter(csv_io, fieldnames=["Row_Number", "Item_ID", "Status", "Message", "Mastery_Level", "Item_Type", "Overrides"])
    csv_writer.writeheader()
    for log in audit_logs:
        csv_writer.writerow(log)
    csv_bytes = csv_io.getvalue().encode("utf-8")
    
    # ─── 5. LOCAL BACKUP SAVE ──────────────────────────────────────────────────
    # Save a local backup inside the workspace 'output' folder
    try:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(excel_file.name)) if hasattr(excel_file, "name") and os.path.exists(excel_file.name) else os.getcwd(), "output")
        os.makedirs(output_dir, exist_ok=True)
        today_file = datetime.today().strftime("%Y%m%d")
        
        backup_docx_path = os.path.join(output_dir, f"Science_PAL_G6_Item_Bank_{today_file}.docx")
        with open(backup_docx_path, "wb") as f:
            f.write(docx_bytes)
            
        backup_csv_path = os.path.join(output_dir, f"Science_PAL_G6_Item_Bank_{today_file}_log.csv")
        with open(backup_csv_path, "wb") as f:
            f.write(csv_bytes)
    except Exception:
        # Ignore local backup failures (e.g. if run purely from memory without filenames)
        pass
        
    summary = {
        "written": len(processed_items),
        "skipped": skipped_items,
        "overrides": override_count
    }
    
    return docx_bytes, csv_bytes, summary
