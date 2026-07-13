import streamlit as st
import subprocess
import re
from docx import Document
from docx.shared import Inches, Pt
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
        margin-top: -15px;
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

# --- 2. વર્ડ ફાઈલનું ફોર્મેટિંગ અને અલાઈનમેન્ટ ---
def set_formatting_and_margins(docx_filename, font_size, font_name):
    doc = Document(docx_filename)
    section = doc.sections[0]
    
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
        
        # પ્રશ્ન માટેનું સેટિંગ
        if re.match(r'^Q\.\d+', text):
            paragraph.paragraph_format.left_indent = Inches(0.25)
            paragraph.paragraph_format.first_line_indent = Inches(-0.25)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(2) 
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # જસ્ટિફાઈડ સેટિંગ ઉમેર્યું
            
            paragraph.paragraph_format.tab_stops.clear_all()
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.25), WD_TAB_ALIGNMENT.LEFT)
            
        # ઓપ્શન માટેનું સેટિંગ
        elif re.match(r'^\(?[A-D][\)\.]', text):
            paragraph.paragraph_format.left_indent = Inches(0.25)
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # જસ્ટિફાઈડ સેટિંગ ઉમેર્યું
            
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
            
            if tabs_count == 3: # 4 ઓપ્શન 1 લાઈનમાં
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.8), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(2.4), WD_TAB_ALIGNMENT.LEFT)
            elif tabs_count == 1: # 2x2 મેટ્રિક્સ
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
        else:
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # જસ્ટિફાઈડ સેટિંગ ઉમેર્યું
            
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
def format_content(raw_text):
    raw_text = raw_text.replace('**', '')
    raw_text = re.sub(r'\n{3,}', '\n\n', raw_text)
    
    lines = raw_text.split('\n')
    questions = []
    current_q = []
    
    q_start_pattern = r'^[\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s+'
    
    for line in lines:
        if not line.strip(): continue
        if re.match(q_start_pattern, line):
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
        opt_pattern = r'\([1-4A-Da-d]\)\s*(.*?)(?=\([1-4A-Da-d]\)|$)'
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

# કંપનીનો લોગો મૂકવા માટે (વચ્ચે સેટ કરેલો છે)
col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    try:
        st.image("26242.png", use_container_width=True)
    except Exception:
        pass # જો લોગો ન મળે તો એરર નહિ બતાવે, સાઈટ ચાલુ રહેશે

st.markdown("<h1 class='main-title'>Question Paper Generator</h1>", unsafe_allow_html=True)

# સેટિંગ્સના 3 ભાગ
col1, col2, col3 = st.columns(3)
with col1:
    file_name = st.text_input("ફાઈલનું નામ (File Name):", "Solid_Black_Paper")
with col2:
    font_size = st.number_input("ફોન્ટ સાઈઝ (Font Size):", min_value=8, max_value=20, value=10)
with col3:
    font_name = st.selectbox("પેપરનો ફોન્ટ (Font):", ["Hind Vadodara", "Shruti", "Cambria Math", "Noto Serif", "Times New Roman", "Calibri", "Arial"])

# પ્રશ્નો નાખવાનું બોક્સ
user_input = st.text_area("અહીં પ્રશ્નો પેસ્ટ કરો (દરેક પ્રશ્ન વચ્ચે 1 ખાલી લાઈન હોવી જરૂરી છે):", height=280)

# ફાઈલ જનરેટ કરવાનું પ્રોસેસિંગ
if st.button("વર્ડ ફાઇલ જનરેટ કરો"):
    if user_input.strip():
        with st.spinner("Processing your document..."):
            processed_md = format_content(user_input)
            with open("temp.md", "w", encoding="utf-8") as f:
                f.write(processed_md)
                
            try:
                subprocess.run(["pandoc", "temp.md", "-o", "temp.docx"], check=True)
                set_formatting_and_margins("temp.docx", font_size, font_name)
                
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
