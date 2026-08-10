import streamlit as st
import subprocess
import re
import shutil
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

# --- 1. Page Config & Custom CSS ---
st.set_page_config(page_title="Solid Black | Paper Generator", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Hind+Vadodara:wght@400;600;700&display=swap');
    
    html, body, [class*="css"], .stTextInput>label, .stSelectbox>label, .stNumberInput>label, .stTextArea>label, p, h1, h2, h3, h4, h5, h6, span {
        font-family: 'Hind Vadodara', sans-serif !important;
    }
    
    .main-title {
        text-align: center;
        font-weight: 700;
        font-size: 32px;
        margin-top: 10px;
        margin-bottom: 25px;
        color: #111111;
    }
    
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
    
    .stTextArea textarea {
        border-radius: 6px !important;
        border: 1px solid #444 !important;
        font-size: 16px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 1-કોલમ અને 2-કોલમ અલગ પાડવા માટેનું XML લોજિક (Narrow માર્જિન સાથે) ---
def make_paragraph_a_section_break(paragraph, num_cols):
    p_pr = paragraph._element.get_or_add_pPr()
    sectPr = OxmlElement('w:sectPr')
    
    # ફરજિયાત Narrow માર્જિન સેટ કરવું (નહીંતર 2-કોલમમાં માર્જિન વિખેરાઈ જશે)
    pgMar = OxmlElement('w:pgMar')
    pgMar.set(qn('w:top'), '432')     # 0.3 inches
    pgMar.set(qn('w:bottom'), '432')  # 0.3 inches
    pgMar.set(qn('w:left'), '720')    # 0.5 inches (Narrow)
    pgMar.set(qn('w:right'), '720')   # 0.5 inches (Narrow)
    sectPr.append(pgMar)
    
    type_el = OxmlElement('w:type')
    type_el.set(qn('w:val'), 'continuous')
    sectPr.append(type_el)
    
    cols_el = OxmlElement('w:cols')
    cols_el.set(qn('w:num'), str(num_cols))
    if num_cols == 2:
        cols_el.set(qn('w:space'), '400') # કોલમ વચ્ચેની જગ્યા
        cols_el.set(qn('w:sep'), '1')
    sectPr.append(cols_el)
    
    p_pr.append(sectPr)

# ટેબલની બોર્ડર સેટ કરવા માટે ફંક્શન
def set_cell_border(border_el, val='single', sz='12', space='0', color='000000'):
    border_el.set(qn('w:val'), val)
    border_el.set(qn('w:sz'), sz)
    border_el.set(qn('w:space'), space)
    border_el.set(qn('w:color'), color)

# --- 3 કોલમ હેડર ડિઝાઇન ---
def insert_header_table(doc, header_left, header_center):
    h_font = "Times New Roman"
    
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    
    first_para = doc.paragraphs[0]._p
    first_para.addprevious(table._tbl)
    
    table.columns[0].width = Inches(2.4)
    table.columns[1].width = Inches(3.9) # વચ્ચેના ભાગ માટે વધુ જગ્યા
    table.columns[2].width = Inches(1.2)
    
    # 1. ડાબી બાજુ
    left_cell = table.cell(0, 0)
    p_logo_left = left_cell.paragraphs[0]
    p_logo_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists('logo.png'):
        p_logo_left.add_run().add_picture('logo.png', width=Inches(2.4))
    else:
        p_logo_left.add_run("[logo.png missing]").bold = True
        
    lines = header_left.strip().split('\n')
    if not lines: lines = ["MCQ", "GUJARATI MEDIUM"]
    if len(lines) == 1: lines.append(" ")
    
    # ગ્રે બોક્સ અને બોર્ડર 
    nested_table = left_cell.add_table(rows=2, cols=1)
    tblPr = nested_table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    
    for b_name in ['top', 'left', 'bottom', 'right', 'insideH']:
        b_el = OxmlElement(f'w:{b_name}')
        sz = '24' if b_name == 'insideH' else '12' 
        set_cell_border(b_el, sz=sz)
        tblBorders.append(b_el)
    tblPr.append(tblBorders)
    
    for r_idx in range(2):
        n_cell = nested_table.cell(r_idx, 0)
        tcPr = n_cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'D9D9D9') 
        tcPr.append(shd)
        
        np = n_cell.paragraphs[0]
        np.alignment = WD_ALIGN_PARAGRAPH.CENTER
        n_run = np.add_run(lines[r_idx])
        n_run.font.name = h_font
        n_run.font.bold = True
        n_run.font.size = Pt(14)
    
    # 2. વચ્ચેનો ભાગ (ટાઇટલ એકદમ પ્રોફેશનલ અને મોટા અક્ષરે)
    center_cell = table.cell(0, 1)
    p_center = center_cell.paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_lines = header_center.strip().split('\n')
    for i, line in enumerate(c_lines):
        r = p_center.add_run(line)
        r.font.name = h_font
        r.font.bold = True
        if i == 0:
            r.font.size = Pt(22) # સૌથી મોટું ટાઈટલ
        elif i == 1:
            r.font.size = Pt(16)
        else:
            r.font.size = Pt(14)
            
        if i < len(c_lines) - 1:
            p_center.add_run('\n')
            
    # 3. જમણી બાજુ
    right_cell = table.cell(0, 2)
    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if os.path.exists('sblogo.png'):
        p_right.add_run().add_picture('sblogo.png', width=Inches(1.2))
    else:
        p_right.add_run("[sblogo.png missing]").bold = True

    p_break = doc.add_paragraph()
    table._tbl.addnext(p_break._p)
    make_paragraph_a_section_break(p_break, 1)

# --- 2. વર્ડ ફાઈલનું ફોર્મેટિંગ ---
def set_formatting_and_margins(docx_filename, font_size, font_name, header_left, header_center):
    doc = Document(docx_filename)
    
    for section in doc.sections:
        # ફરજિયાત Narrow માર્જિન અહીં પણ સેટ કર્યું છે
        section.top_margin = Inches(0.3) 
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.5) 
        section.right_margin = Inches(0.5) 
        section.header_distance = Inches(0.1)
        section.footer_distance = Inches(0.1)
        
        # વોટરમાર્ક (sblogo.png)
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists('sblogo.png'):
            image_part, rel_id = header.part.get_or_add_image('sblogo.png')
            watermark_xml = f'''
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                 xmlns:v="urn:schemas-microsoft-com:vml" 
                 xmlns:o="urn:schemas-microsoft-com:office:office" 
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <w:pict>
                    <v:shape id="Watermark" style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:450pt;height:450pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" stroked="f">
                        <v:imagedata r:id="{rel_id}"/>
                    </v:shape>
                </w:pict>
            </w:r>
            '''
            header_para._p.append(parse_xml(watermark_xml))
            
        # ફૂટર (FOTTER@4x-8.png)
        footer = section.footer
        footer.is_linked_to_previous = False
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists('FOTTER@4x-8.png'):
            run = footer_para.add_run()
            run.add_picture('FOTTER@4x-8.png', width=Inches(7.5))

    insert_header_table(doc, header_left, header_center)

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
        
        # ### વાળું ટાઇટલ
        if '###HEADER###' in text:
            clean_title = text.replace('###HEADER###', '').strip()
            
            if i > 0:
                make_paragraph_a_section_break(paragraphs[i-1], 2)
            make_paragraph_a_section_break(paragraph, 1)
            
            paragraph.text = ""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)
            
            pPr = paragraph._element.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), '000080')
            pPr.append(shd)
            
            run = paragraph.add_run(clean_title)
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.bold = True
            run.font.size = Pt(font_size + 2)
            run.font.name = font_name
            continue

        # પ્રશ્નો
        if re.match(r'^Q\.\d+', text):
            paragraph.paragraph_format.left_indent = Inches(0.35)
            paragraph.paragraph_format.first_line_indent = Inches(-0.35)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(2) 
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.tab_stops.clear_all()
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.35), WD_TAB_ALIGNMENT.LEFT)
            
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
            
    # છેલ્લે આખું પેપર 2-કોલમમાં સેટ કરી દો
    final_section = doc.sections[-1]
    sectPr = final_section._sectPr
    cols = sectPr.find(qn('w:cols')) or OxmlElement('w:cols')
    if cols not in sectPr: sectPr.append(cols)
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '400')
    cols.set(qn('w:sep'), '1')
                
    doc.save(docx_filename)

# --- 3. સ્માર્ટ માર્કડાઉન પાર્સર ---
def format_content(raw_text, is_continuous, start_num, end_num):
    raw_text = raw_text.replace('**', '')
    raw_text = re.sub(r'\n{3,}', '\n\n', raw_text)
    
    lines = raw_text.split('\n')
    questions = []
    current_q = []
    
    q_start_pattern = r'^[\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s+'
    
    for line in lines:
        if not line.strip(): continue
        if line.strip().startswith('### '):
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
    q_num = start_num 
    
    q_prefix_pattern = r'^([\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s*)+'
    labels = ['A', 'B', 'C', 'D']
    
    for q_block in questions:
        if end_num > 0 and q_num > end_num and not q_block.startswith('### '):
            continue 
            
        if q_block.startswith('### '):
            if not is_continuous:
                q_num = start_num
            clean_title = q_block.replace('###', '', 1).strip()
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

# --- 4. Streamlit UI ---

# ⚠️ ઈમેજ ફાઈલોનું ચેકિંગ (અહીં જ ખબર પડી જશે કે કોઈ ફાઈલ મિસિંગ છે કે નહિ)
missing_assets = [f for f in ['logo.png', 'sblogo.png', 'FOTTER@4x-8.png'] if not os.path.exists(f)]
if missing_assets:
    st.error(f"⚠️ ચેતવણી: આ ઈમેજ ફાઈલો સિસ્ટમમાં નથી મળી: **{', '.join(missing_assets)}**. પ્લીઝ અપલોડ કરો, નહીંતર પેપરમાં તે નહિ દેખાય!")

col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    if os.path.exists('logo.png'):
        st.image("logo.png", use_container_width=True)

st.markdown("<h1 class='main-title'>Question Paper Generator</h1>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; margin-top: -15px; margin-bottom: 25px;'><span style='background-color: #000000; color: #ffffff; padding: 6px 18px; border-radius: 20px; font-size: 15px; font-weight: 700; box-shadow: 0px 4px 6px rgba(0,0,0,0.2); letter-spacing: 0.5px;'>Made by Yug Ghanshyam Padmani</span></div>", unsafe_allow_html=True)

st.markdown("### 📝 પેપરનું હેડર ડિઝાઇન")
col_h1, col_h2 = st.columns(2)
with col_h1:
    header_left = st.text_area("ડાબી બાજુનું ગ્રે બોક્સ (બે લાઈન ફરજિયાત):", "MCQ\nGUJARATI MEDIUM", height=100)
with col_h2:
    header_center = st.text_area("વચ્ચેનું મેઈન ટાઈટલ:", "STD 12 SCIENCE\nMATHS\n40 MARKS\nDate : 07/08/26", height=100)

st.divider()

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ (આ ભરવું ફરજિયાત છે!)")
col1, col2, col3 = st.columns(3)
with col1:
    file_name = st.text_input("ફાઈલનું નામ (File Name) [ફરજિયાત]:", value="", placeholder="દા.ત. Physics_Test_1")
    font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=10)
    font_name = st.selectbox("પેપરનો ફોન્ટ:", ["Hind Vadodara", "Shruti", "Cambria Math", "Noto Serif", "Times New Roman", "Calibri", "Arial"])
with col2:
    start_num = st.number_input("પ્રશ્ન ક્યાંથી શરૂ કરવો છે? (Start):", min_value=1, value=1)
    end_num = st.number_input("ક્યાં પૂરા કરવા છે? (End - 0 એટલે બધા જ):", min_value=0, value=0)
with col3:
    st.markdown("<br>", unsafe_allow_html=True)
    is_continuous = st.checkbox("પ્રશ્નોના નંબર સળંગ (Continuous) રાખવા છે?", value=True)
    st.info("જો સળંગ નંબર ના રાખવા હોય, તો ટીક માર્ક હટાવી દો.")

st.markdown("### ✍️ પ્રશ્નો (ફરજિયાત)")
user_input = st.text_area("અહીં પ્રશ્નો પેસ્ટ કરો (ડાર્ક બ્લુ ટાઇટલ મૂકવા તેની આગળ ફરજિયાત ### લખો, દા.ત. ### Section B):", height=280)

if st.button("વર્ડ અને PDF ફાઇલ જનરેટ કરો"):
    if not file_name.strip():
        st.error("⚠️ ભૂલ: 'ફાઈલનું નામ' ખાલી છે! કૃપા કરીને ફાઈલનું નામ લખો.")
    elif not user_input.strip():
        st.error("⚠️ ભૂલ: ડેટા ખાલી છે! કૃપા કરીને ઉપરના બોક્સમાં પ્રશ્નો પેસ્ટ કરો.")
    else:
        with st.spinner("Processing your document..."):
            processed_md = format_content(user_input, is_continuous, start_num, end_num)
            with open("temp.md", "w", encoding="utf-8") as f:
                f.write(processed_md)
                
            try:
                subprocess.run(["pandoc", "temp.md", "-o", "temp.docx"], check=True)
                set_formatting_and_margins("temp.docx", font_size, font_name, header_left, header_center)
                
                final_file = f"{file_name}.docx"
                shutil.move("temp.docx", final_file)
                
                st.success("✅ ફાઇલ સફળતાપૂર્વક બની ગઈ છે!")
                
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    with open(final_file, "rb") as file:
                        st.download_button("📄 Word (DOCX) ડાઉનલોડ કરો", file, file_name=final_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                with col_d2:
                    try:
                        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", final_file], check=True)
                        pdf_file = final_file.replace('.docx', '.pdf')
                        with open(pdf_file, "rb") as p_file:
                            st.download_button("📕 PDF ડાઉનલોડ કરો", p_file, file_name=pdf_file, mime="application/pdf")
                    except Exception:
                        st.error("⚠️ PDF બનાવવા માટે સિસ્ટમમાં LibreOffice નથી.")
                        st.info("Terminal માં આ રન કરો: `sudo apt-get update && sudo apt-get install libreoffice -y`")
                        
            except Exception as e:
                st.error(f"Error: {e}")
