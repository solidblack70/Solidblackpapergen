import streamlit as st
import subprocess
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- 1. Page Config & Custom CSS (Premium UI) ---
st.set_page_config(page_title="Solid Black | Paper Generator", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Hind+Vadodara:wght@400;600;700&display=swap');
    
    /* આખી સાઇટમાં Hind Vadodara ફોન્ટ */
    html, body, [class*="css"], .stTextInput>label, .stSelectbox>label, .stNumberInput>label, .stTextArea>label, p, h1, h2, h3, h4, h5, h6, span {
        font-family: 'Hind Vadodara', sans-serif !important;
    }
    
    /* મેઈન ટાઈટલ અને ટેક્સ્ટ સ્ટાઈલિંગ */
    .main-title {
        text-align: center;
        font-weight: 700;
        font-size: 32px;
        margin-top: 10px;
        margin-bottom: 25px;
        color: #111111;
    }
    
    /* જનરેટ બટન માટે પ્રીમિયમ બ્લેક એન્ડ વ્હાઇટ હોવર ઇફેક્ટ */
    div.stButton > button:first-child {
        background-color: #000000;
        color: #ffffff;
        border: 2px solid #000000;
        border-radius: 6px;
        padding: 10px 24px;
        font-size: 18px;
        font-weight: 600;
        transition: all 0.3s ease-in-out;
        width: 100%;
        box-shadow: 0px 4px 6px rgba(0,0,0,0.1);
    }
    div.stButton > button:first-child:hover {
        background-color: #ffffff;
        color: #000000;
        border: 2px solid #000000;
        box-shadow: 0px 6px 12px rgba(0,0,0,0.2);
    }
    
    /* ટેક્સ્ટ એરિયા અને ઇનપુટ બોક્સ બોર્ડર */
    .stTextArea textarea {
        border-radius: 6px !important;
        border: 1px solid #444 !important;
        font-size: 16px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 1 કોલમ અને 2 કોલમ વચ્ચે જગ્યા પાડવા માટે ---
def add_continuous_section_break(paragraph, num_cols):
    pPr = paragraph._element.get_or_add_pPr()
    existing_sectPr = pPr.find(qn('w:sectPr'))
    if existing_sectPr is not None:
        pPr.remove(existing_sectPr)
        
    sectPr = OxmlElement('w:sectPr')
    type_el = OxmlElement('w:type')
    type_el.set(qn('w:val'), 'continuous')
    sectPr.append(type_el)
    
    cols_el = OxmlElement('w:cols')
    cols_el.set(qn('w:num'), str(num_cols))
    cols_el.set(qn('w:space'), '720')
    if num_cols == 2:
        cols_el.set(qn('w:sep'), '1')
    sectPr.append(cols_el)
    
    margin = OxmlElement('w:pgMar')
    margin.set(qn('w:top'), '720')
    margin.set(qn('w:bottom'), '720')
    margin.set(qn('w:left'), '720')
    margin.set(qn('w:right'), '720')
    sectPr.append(margin)
    
    pPr.append(sectPr)

# --- નવું: 3 કોલમ હેડર ડિઝાઇન ---
def insert_header_table(doc, header_left, header_center, font_name):
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    
    # ટેબલને ડોક્યુમેન્ટની એકદમ શરૂઆતમાં ખસેડો
    first_para = doc.paragraphs[0]._p
    first_para.addprevious(table._tbl)
    
    # કોલમની સાઇઝ (અંદાજિત)
    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(3.5)
    table.columns[2].width = Inches(1.5)
    
    # -- 1. ડાબી બાજુ (નાનો લોગો + ગ્રે બોક્સ) --
    left_cell = table.cell(0, 0)
    p_logo_left = left_cell.paragraphs[0]
    p_logo_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_logo_left.add_run().add_picture('logo.png', width=Inches(1.5))
    except:
        p_logo_left.add_run("[Small Logo]").bold = True
        
    p_gray = left_cell.add_paragraph()
    p_gray.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p_gray._element.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'D9D9D9') # ગ્રે કલર
    pPr.append(shd)
    
    run_gray = p_gray.add_run(header_left)
    run_gray.font.name = font_name
    run_gray.font.bold = True
    run_gray.font.size = Pt(12)
    
    # -- 2. વચ્ચેનો ભાગ (મેઈન ટાઇટલ) --
    center_cell = table.cell(0, 1)
    p_center = center_cell.paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lines = header_center.strip().split('\n')
    for i, line in enumerate(lines):
        r = p_center.add_run(line)
        r.font.name = font_name
        r.font.bold = True
        r.font.size = Pt(20) if i == 0 else Pt(14)
        if i < len(lines) - 1:
            p_center.add_run('\n')
            
    # -- 3. જમણી બાજુ (મોટો ગોળ લોગો) --
    right_cell = table.cell(0, 2)
    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    try:
        p_right.add_run().add_picture('Asset 345@4x-8.png', width=Inches(1.2))
    except:
        p_right.add_run("[Round Logo]").bold = True

    # હેડર પછી ખાલી જગ્યા અને 1-કોલમ બ્રેક મૂકો (જેથી હેડર ભેગું ના થાય)
    p_break = doc.add_paragraph()
    table._tbl.addnext(p_break._p)
    add_continuous_section_break(p_break, 1)


# --- 2. વર્ડ ફાઈલનું ફોર્મેટિંગ ---
def set_formatting_and_margins(docx_filename, font_size, font_name, header_left, header_center):
    doc = Document(docx_filename)
    
    # આખા ડોક્યુમેન્ટને બાય-ડિફોલ્ટ 2 કોલમમાં સેટ કરો
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        sectPr = section._sectPr
        cols = sectPr.find(qn('w:cols')) or OxmlElement('w:cols')
        if cols not in sectPr: sectPr.append(cols)
        cols.set(qn('w:num'), '2')       
        cols.set(qn('w:space'), '720')   
        cols.set(qn('w:sep'), '1')       
    
    # હેડર ટેબલ ઉમેરો
    insert_header_table(doc, header_left, header_center, font_name)

    # વધારાની ખાલી જગ્યા (Empty Paragraphs) રિમૂવર
    for paragraph in list(doc.paragraphs):
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)
            paragraph._p = paragraph._element = None
            continue
            
    paragraphs = doc.paragraphs
    for i, paragraph in enumerate(paragraphs):
        if paragraph.style.name.startswith('List'):
            paragraph.style = doc.styles['Normal']
            
        for run in paragraph.runs:
            if '‡' in run.text:
                run.text = run.text.replace('‡', '\t')
                
        text = paragraph.text.strip()
        if not text: continue
        
        # --- બગ ફિક્સ: હવે માત્ર સ્પેશિયલ ટેગ દ્વારા જ ટાઇટલ ઓળખાશે ---
        if '###HEADER###' in text:
            clean_title = text.replace('###HEADER###', '').strip()
            
            # ટાઇટલ પહેલાના ભાગને 2 કોલમમાં પૂરો કરો
            if i > 0:
                add_continuous_section_break(paragraphs[i-1], 2)
            
            # આ ટાઇટલને 1 સળંગ કોલમ બનાવો
            add_continuous_section_break(paragraph, 1)
            
            # જૂનો બધો ગોટાળો હટાવી નવેસરથી ટેક્સ્ટ ઉમેરો
            paragraph.text = ""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)
            
            # ડાર્ક બ્લુ બેકગ્રાઉન્ડ સેટ કરવા
            pPr = paragraph._element.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), '1F4E79')
            pPr.append(shd)
            
            run = paragraph.add_run(clean_title)
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.bold = True
            run.font.size = Pt(font_size + 4)
            run.font.name = font_name
            
            r = run._element
            rPr = r.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            continue

        # પ્રશ્ન માટેનું સેટિંગ
        if re.match(r'^Q\.\d+', text):
            paragraph.paragraph_format.left_indent = Inches(0.35)
            paragraph.paragraph_format.first_line_indent = Inches(-0.35)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(2) 
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            paragraph.paragraph_format.tab_stops.clear_all()
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.35), WD_TAB_ALIGNMENT.LEFT)
            
        # ઓપ્શન માટેનું સેટિંગ
        elif re.match(r'^\(?[A-D][\)\.]', text):
            paragraph.paragraph_format.left_indent = Inches(0.35)
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            is_last_option = True
            for j in range(i + 1, len(paragraphs)):
                next_text = paragraphs[j].text.strip()
                if not next_text: continue
                if re.match(r'^\(?[A-D][\)\.]', next_text):
                    is_last_option = False
                break
                
            paragraph.paragraph_format.space_after = Pt(8) if is_last_option else Pt(0)
            
            paragraph.paragraph_format.tab_stops.clear_all()
            tabs_count = paragraph.text.count('\t')
            
            if tabs_count == 3: 
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.8), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(2.4), WD_TAB_ALIGNMENT.LEFT)
            elif tabs_count == 1: 
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
        else:
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.name = font_name
            r = run._element
            rPr = r.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
                
    doc.save(docx_filename)

# --- 3. સ્માર્ટ માર્કડાઉન પાર્સર ---
def format_content(raw_text, is_continuous):
    raw_text = raw_text.replace('**', '')
    raw_text = re.sub(r'\n{3,}', '\n\n', raw_text)
    
    lines = raw_text.split('\n')
    questions = []
    current_q = []
    
    q_start_pattern = r'^[\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s+'
    
    for line in lines:
        if not line.strip(): continue
        
        if line.strip().startswith('#'):
            if current_q:
                questions.append("\n".join(current_q))
            questions.append(line.strip())
            current_q = []
        elif re.match(q_start_pattern, line):
            if current_q:
                questions.append("\n".join(current_q))
            current_q = [line]
        else:
            current_q.append(line)
            
    if current_q:
        questions.append("\n".join(current_q))
        
    formatted_md = ""
    q_num = 1
    
    q_prefix_pattern = r'^([\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s*)+'
    labels = ['A', 'B', 'C', 'D']
    
    for q_block in questions:
        if q_block.startswith('#'):
            # જો નંબર્સ રીસેટ કરવાના હોય તો
            if not is_continuous:
                q_num = 1
                
            clean_title = q_block.replace('#', '', 1).strip()
            formatted_md += f"###HEADER### {clean_title}\n\n"
            continue
            
        opt_pattern = r'\s*\(?[1-4A-Da-d][\)\.]\s*(.*?)(?=\s+\(?[1-4A-Da-d][\)\.]|$)'
        matches = list(re.finditer(opt_pattern, q_block, flags=re.DOTALL))
        
        if len(matches) >= 4:
            opts = matches[-4:]
            q_text = q_block[:opts[0].start()].strip()
            
            q_text = re.sub(q_prefix_pattern, '', q_text).strip()
            q_text = re.sub(r'\n\s*\n', '\n', q_text)
            
            q_md = f"**Q.{q_num}**‡{q_text}"
            q_num += 1
            
            clean_opts = []
            for i, m in enumerate(opts):
                opt_content = m.group(1).strip()
                opt_content = re.sub(r'\s+', ' ', opt_content) 
                clean_opts.append(f"\\({labels[i]}\\) {opt_content}")
                
            lens = [len(o) for o in clean_opts]
            max_len = max(lens)
            
            if max_len < 16:
                opts_md = "‡".join(clean_opts)
            elif max_len < 36:
                opts_md = f"{clean_opts[0]}‡{clean_opts[1]}\n\n{clean_opts[2]}‡{clean_opts[3]}"
            else:
                opts_md = "\n\n".join(clean_opts)
                
            formatted_md += q_md + "\n\n" + opts_md + "\n\n"
        else:
            clean_q = re.sub(q_prefix_pattern, '', q_block).strip()
            if clean_q != q_block.strip() and re.match(q_start_pattern, q_block.strip()):
                 formatted_md += f"**Q.{q_num}**‡{clean_q}\n\n"
                 q_num += 1
            else:
                 formatted_md += q_block + "\n\n"
                 
    return formatted_md

# --- 4. Streamlit UI (Clean & Professional) ---

st.markdown("<h1 class='main-title'>Question Paper Generator</h1>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; margin-top: -15px; margin-bottom: 25px;'><span style='background-color: #000000; color: #ffffff; padding: 6px 18px; border-radius: 20px; font-size: 15px; font-weight: 700; box-shadow: 0px 4px 6px rgba(0,0,0,0.2); letter-spacing: 0.5px;'>Made by Solid Black Institute</span></div>", unsafe_allow_html=True)

# --- નવું: પેપરનું હેડર સેટિંગ ---
st.markdown("### 📝 પેપરનું હેડર ડિઝાઇન (Header Customization)")
col_h1, col_h2 = st.columns(2)
with col_h1:
    header_left = st.text_area("ડાબી બાજુનું ગ્રે બોક્સ (Left Gray Box):", "MCQ\nGUJARATI MEDIUM", height=130)
with col_h2:
    header_center = st.text_area("વચ્ચેનું મેઈન ટાઈટલ (Center Title):", "STD 12 SCIENCE\nMATHS\n40 MARKS\nDate : 07/08/26", height=130)

st.divider()

# સેટિંગ્સના 3 ભાગ
st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ (Settings)")
col1, col2, col3 = st.columns(3)
with col1:
    file_name = st.text_input("ફાઈલનું નામ (File Name):", "Solid_Black_Paper")
    # નવું ચેકબોક્સ: પ્રશ્ન નંબર્સ કંટીન્યુ રાખવા માટે
    is_continuous = st.checkbox("પ્રશ્નોના નંબર સળંગ (Continuous) રાખવા છે?", value=True)
with col2:
    font_size = st.number_input("ફોન્ટ સાઈઝ (Font Size):", min_value=8, max_value=20, value=10)
with col3:
    font_name = st.selectbox("પેપરનો ફોન્ટ (Font):", ["Hind Vadodara", "Shruti", "Cambria Math", "Noto Serif", "Times New Roman", "Calibri", "Arial"])

# પ્રશ્નો નાખવાનું બોક્સ
st.markdown("### ✍️ પ્રશ્નો (Questions Data)")
user_input = st.text_area("અહીં પ્રશ્નો પેસ્ટ કરો (ટાઇટલ/સૂચના મૂકવા તેની આગળ # લખો, દા.ત. # Section B):", height=280)

# ફાઈલ જનરેટ કરવાનું પ્રોસેસિંગ
if st.button("વર્ડ ફાઇલ જનરેટ કરો (Generate Word File)"):
    if user_input.strip():
        with st.spinner("Processing your document..."):
            # is_continuous ની વેલ્યૂ પાસ કરી
            processed_md = format_content(user_input, is_continuous)
            with open("temp.md", "w", encoding="utf-8") as f:
                f.write(processed_md)
                
            try:
                subprocess.run(["pandoc", "temp.md", "-o", "temp.docx"], check=True)
                # હેડર ટેક્સ્ટ પાસ કરી
                set_formatting_and_margins("temp.docx", font_size, font_name, header_left, header_center)
                
                final_file = f"{file_name}.docx"
                import shutil
                shutil.move("temp.docx", final_file)
                
                st.success("✅ ફાઇલ સફળતાપૂર્વક બની ગઈ છે!")
                with open(final_file, "rb") as file:
                    st.download_button("📄 ડાઉનલોડ કરો (Download)", file, file_name=final_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Error: {e}")
    else:
        st.warning("કૃપા કરીને પ્રશ્નો પેસ્ટ કરો.")
