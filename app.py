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
    .main-title { text-align: center; font-weight: 700; font-size: 32px; margin-top: 10px; margin-bottom: 25px; color: #111111; }
    div.stButton > button:first-child {
        background-color: #000000; color: #ffffff; border: 2px solid #000000; border-radius: 6px;
        padding: 10px 24px; font-size: 18px; font-weight: 600; transition: all 0.3s ease-in-out; width: 100%;
        box-shadow: 0px 4px 6px rgba(0,0,0,0.1);
    }
    div.stButton > button:first-child:hover {
        background-color: #ffffff; color: #000000; border: 2px solid #000000; box-shadow: 0px 6px 12px rgba(0,0,0,0.2);
    }
    .stTextArea textarea { border-radius: 6px !important; border: 1px solid #444 !important; font-size: 16px !important; }
    </style>
    """, unsafe_allow_html=True)

# --- XML Tools ---
def set_cell_border(border_el, val='single', sz='12', space='0', color='000000'):
    border_el.set(qn('w:val'), val)
    border_el.set(qn('w:sz'), sz)
    border_el.set(qn('w:space'), space)
    border_el.set(qn('w:color'), color)

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
    cols_el.set(qn('w:space'), '400')
    if num_cols == 2:
        cols_el.set(qn('w:sep'), '1')
    sectPr.append(cols_el)
    
    margin = OxmlElement('w:pgMar')
    margin.set(qn('w:top'), '432')
    margin.set(qn('w:bottom'), '432')
    margin.set(qn('w:left'), '720')
    margin.set(qn('w:right'), '720')
    sectPr.append(margin)
    pPr.append(sectPr)

# --- 3 કોલમ હેડર ડિઝાઇન (પર્ફેક્ટ માપ અને સાઇઝ) ---
def insert_header_table(doc, header_left, header_center, font_name):
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    
    first_para = doc.paragraphs[0]._p
    first_para.addprevious(table._tbl)
    
    # EXACT માપ (લોગો જમણા ખૂણે જાય અને બોક્સ ન કપાય)
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(4.2)
    table.columns[2].width = Inches(1.2)
    
    # -- 1. ડાબી બાજુ --
    left_cell = table.cell(0, 0)
    p_logo_left = left_cell.paragraphs[0]
    p_logo_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists('logo.png'):
        p_logo_left.add_run().add_picture('logo.png', width=Inches(1.7))
        
    lines = header_left.strip().split('\n')
    if not lines: lines = ["JEE MAIN", "GUJARATI MEDIUM"]
    if len(lines) == 1: lines.append(" ")
    
    nested_table = left_cell.add_table(rows=2, cols=1)
    tblBorders = OxmlElement('w:tblBorders')
    for b_name in ['top', 'left', 'bottom', 'right', 'insideH']:
        b_el = OxmlElement(f'w:{b_name}')
        sz = '24' if b_name == 'insideH' else '12' 
        set_cell_border(b_el, sz=sz)
        tblBorders.append(b_el)
    nested_table._tbl.tblPr.append(tblBorders)
    
    for r_idx in range(2):
        n_cell = nested_table.cell(r_idx, 0)
        tcPr = n_cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd', {qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): 'D9D9D9'})
        tcPr.append(shd)
        
        tcMar = OxmlElement('w:tcMar')
        for m in ['top', 'left', 'bottom', 'right']:
            el = OxmlElement(f'w:{m}', {qn('w:w'): '0', qn('w:type'): 'dxa'})
            tcMar.append(el)
        tcPr.append(tcMar)
        
        np = n_cell.paragraphs[0]
        np.alignment = WD_ALIGN_PARAGRAPH.CENTER
        np.paragraph_format.space_before = Pt(0)
        np.paragraph_format.space_after = Pt(0)
        
        n_run = np.add_run(lines[r_idx])
        n_run.font.name = "Arial"
        n_run.font.bold = True
        n_run.font.size = Pt(14) if r_idx == 0 else Pt(10)
        
    # -- 2. વચ્ચેનો ભાગ (નાની સાઈઝ) --
    center_cell = table.cell(0, 1)
    p_center = center_cell.paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_lines = header_center.strip().split('\n')
    for i, line in enumerate(c_lines):
        r = p_center.add_run(line)
        r.font.name = "Times New Roman"
        r.font.bold = True
        # ફોન્ટ સાઇઝ ઘટાડી જેથી મોટું ના દેખાય
        if i == 0: r.font.size = Pt(16)
        elif i == 1: r.font.size = Pt(14)
        else: r.font.size = Pt(12)
        if i < len(c_lines) - 1: p_center.add_run('\n')
            
    # -- 3. જમણી બાજુ (ખૂણામાં) --
    right_cell = table.cell(0, 2)
    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # ખૂણામાં ફિક્સ કરવા માટેનું સેટીંગ
    tcPr_right = right_cell._tc.get_or_add_tcPr()
    tcMar_right = OxmlElement('w:tcMar')
    for m in ['top', 'left', 'bottom', 'right']:
        el = OxmlElement(f'w:{m}', {qn('w:w'): '0', qn('w:type'): 'dxa'})
        tcMar_right.append(el)
    tcPr_right.append(tcMar_right)
    
    if os.path.exists('sblogo.png'):
        p_right.add_run().add_picture('sblogo.png', width=Inches(1.1))

    # હેડરની નીચેની બોર્ડર લાઈન
    p_line = doc.add_paragraph()
    table._tbl.addnext(p_line._p)
    pPr = p_line._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom', {qn('w:val'): 'single', qn('w:sz'): '6', qn('w:space'): '1', qn('w:color'): '000000'})
    pBdr.append(bottom)
    pPr.append(pBdr)

    add_continuous_section_break(p_line, 1)

# --- 2. વર્ડ ફાઈલનું ફોર્મેટિંગ ---
def set_formatting_and_margins(docx_filename, font_size, font_name, header_left, header_center):
    doc = Document(docx_filename)
    
    for section in doc.sections:
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.header_distance = Inches(0.1)
        section.footer_distance = Inches(0.1)
        
        sectPr = section._sectPr
        cols = sectPr.find(qn('w:cols')) or OxmlElement('w:cols')
        if cols not in sectPr: sectPr.append(cols)
        cols.set(qn('w:num'), '2')       
        cols.set(qn('w:space'), '400')   
        cols.set(qn('w:sep'), '1')       

        # વોટરમાર્ક અને ફૂટર 
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists('sblogo.png'):
            try:
                image_part, rel_id = header.part.get_or_add_image('sblogo.png')
                watermark_xml = '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pict><v:shape id="Watermark" style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:350pt;height:350pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" stroked="f"><v:imagedata r:id="' + rel_id + '" gain="35000f" blacklevel="15000f"/></v:shape></w:pict></w:r>'
                header_para._p.append(parse_xml(watermark_xml))
            except Exception: pass
            
        footer = section.footer
        footer.is_linked_to_previous = False
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists('footer.png'):
            try: run = footer_para.add_run(); run.add_picture('footer.png', width=Inches(7.5))
            except Exception: pass

    insert_header_table(doc, header_left, header_center, font_name)

    for paragraph in list(doc.paragraphs):
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)
            paragraph._p = paragraph._element = None
            continue
            
    paragraphs = doc.paragraphs
    for i, paragraph in enumerate(paragraphs):
        
        # બ્રહ્માસ્ત્ર: વર્ડની બુલેટિન (Auto-List) સિસ્ટમ કાયમ માટે ડિલીટ કરો!
        paragraph.style = doc.styles['Normal']
        pPr = paragraph._element.get_or_add_pPr()
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            pPr.remove(numPr)
            
        for run in paragraph.runs:
            if '‡' in run.text:
                run.text = run.text.replace('‡', '\t')
                
        text = paragraph.text.strip()
        if not text: continue
        
        if '###HEADER###' in text:
            clean_title = text.replace('###HEADER###', '').strip()
            if i > 0: add_continuous_section_break(paragraphs[i-1], 2)
            add_continuous_section_break(paragraph, 1)
            
            paragraph.text = ""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(8)
            paragraph.paragraph_format.space_after = Pt(8)
            
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), '000080')
            pPr.append(shd)
            
            run = paragraph.add_run(clean_title)
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.bold = True
            run.font.size = Pt(font_size + 2)
            run.font.name = "Times New Roman"
            continue

        if re.match(r'^Q\.\d+', text):
            paragraph.paragraph_format.left_indent = Inches(0.35)
            paragraph.paragraph_format.first_line_indent = Inches(-0.35)
            paragraph.paragraph_format.space_before = Pt(8)
            paragraph.paragraph_format.space_after = Pt(4)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.tab_stops.clear_all()
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.35), WD_TAB_ALIGNMENT.LEFT)
            
        elif re.match(r'^\(?[A-D][\)\.]', text) and '\t' in paragraph.text:
            paragraph.paragraph_format.left_indent = Inches(0.35)
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            is_last_option = True
            for j in range(i + 1, len(paragraphs)):
                next_text = paragraphs[j].text.strip()
                if not next_text: continue
                if re.match(r'^\(?[A-D][\)\.]', next_text) and '\t' in paragraphs[j].text:
                    is_last_option = False
                break
                
            paragraph.paragraph_format.space_after = Pt(8) if is_last_option else Pt(2)
            paragraph.paragraph_format.tab_stops.clear_all()
            tabs_count = paragraph.text.count('\t')
            
            # તમારા ઓરિજિનલ કોડનું Tab સેટિંગ
            if tabs_count == 3: 
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.8), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(2.4), WD_TAB_ALIGNMENT.LEFT)
            elif tabs_count == 1: 
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.6), WD_TAB_ALIGNMENT.LEFT)
        else:
            # નોર્મલ લખાણ (સોલ્યુશનના ફકરા)
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(4)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            if paragraph.style.name != "Times New Roman":
                run.font.name = font_name
                
    doc.save(docx_filename)

# --- 3. તમારા જૂના કોડ વાળો સ્માર્ટ માર્કડાઉન પાર્સર ---
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
            if not is_continuous:
                q_num = 1
                
            clean_title = q_block.replace('#', '', 1).strip()
            # ડૂચો દૂર કરવા માટે \n\n
            formatted_md += f"###HEADER### {clean_title}\n\n"
            continue
            
        # 1,2,3,4 અથવા A,B,C,D પકડશે
        opt_pattern = r'\s*\(?[1-4A-Da-d][\)\.]\s*(.*?)(?=\s+\(?[1-4A-Da-d][\)\.]|$)'
        matches = list(re.finditer(opt_pattern, q_block, flags=re.DOTALL))
        
        if len(matches) >= 4 and not "સ્ટેપ" in q_block and not "ઉકેલ" in q_block:
            opts = matches[-4:]
            q_text = q_block[:opts[0].start()].strip()
            
            q_text = re.sub(q_prefix_pattern, '', q_text).strip()
            # સવાલમાં ડૂચો અટકાવવા ફકરા સાચવો
            q_text = q_text.replace('\n', '\n\n')
            
            q_md = f"**Q.{q_num}** {q_text}"
            q_num += 1
            
            clean_opts = []
            for i, m in enumerate(opts):
                # નકામા Tabs/Enters સાફ કરો
                opt_content = re.sub(r'[\t\n\r]+', ' ', m.group(1).strip())
                # ફરજિયાત (A), (B), (C), (D) માં કન્વર્ટ
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
                 clean_q = clean_q.replace('\n', '\n\n')
                 formatted_md += f"**Q.{q_num}** {clean_q}\n\n"
                 q_num += 1
            else:
                 # સોલ્યુશનનો ડૂચો અટકાવવાની જડીબુટ્ટી
                 clean_q = q_block.replace('\n', '\n\n')
                 formatted_md += clean_q + "\n\n"
                 
    return formatted_md

# --- 4. Streamlit UI ---
st.markdown("<h1 class='main-title'>Question Paper Generator</h1>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; margin-top: -15px; margin-bottom: 25px;'><span style='background-color: #000000; color: #ffffff; padding: 6px 18px; border-radius: 20px; font-size: 15px; font-weight: 700; box-shadow: 0px 4px 6px rgba(0,0,0,0.2); letter-spacing: 0.5px;'>Made by Solid Black Institute</span></div>", unsafe_allow_html=True)

st.markdown("### 📝 પેપરનું હેડર ડિઝાઇન")
col_h1, col_h2 = st.columns(2)
with col_h1:
    header_left = st.text_area("ડાબી બાજુનું ગ્રે બોક્સ:", "JEE MAIN\nGUJARATI MEDIUM", height=130)
with col_h2:
    header_center = st.text_area("વચ્ચેનું મેઈન ટાઈટલ (5 લાઈન):", "STD 11 SCIENCE\nMATHS\n40 MARKS\nJEE MAIN\nDate 13/08/26", height=130)

st.divider()

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ")
col1, col2, col3 = st.columns(3)
with col1:
    file_name = st.text_input("ફાઈલનું નામ:", "Solid_Black_Paper")
    is_continuous = st.checkbox("પ્રશ્નોના નંબર સળંગ (Continuous) રાખવા છે?", value=True)
with col2:
    font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=10)
with col3:
    font_name = st.selectbox("પેપરનો ફોન્ટ:", ["Hind Vadodara", "Shruti", "Cambria Math", "Times New Roman", "Arial"])

st.markdown("### ✍️ પ્રશ્નો (Questions Data)")
user_input = st.text_area("અહીં પ્રશ્નો પેસ્ટ કરો (ટાઇટલ/સૂચના મૂકવા તેની આગળ # લખો):", height=280)

if not shutil.which("libreoffice"):
    st.warning("⚠️ સિસ્ટમમાં PDF બનાવવાનું સોફ્ટવેર નથી.")
    if st.button("🔧 અત્યારે જ PDF સોફ્ટવેર ઇન્સ્ટોલ કરો (ફક્ત 1 મિનિટ લાગશે)"):
        with st.spinner("ઇન્સ્ટોલ થઈ રહ્યું છે... પ્લીઝ 1 મિનિટ રાહ જુઓ ⏳"):
            os.system("sudo apt-get update && sudo apt-get install libreoffice -y")
            st.success("✅ ઇન્સ્ટોલ થઈ ગયું! હવે તમે PDF બનાવી શકશો.")

if st.button("વર્ડ અને PDF ફાઇલ જનરેટ કરો"):
    if user_input.strip():
        with st.spinner("Processing your document..."):
            processed_md = format_content(user_input, is_continuous)
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
                        st.download_button("📄 Word ડાઉનલોડ કરો", file, file_name=final_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                with col_d2:
                    if shutil.which("libreoffice"):
                        try:
                            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", final_file], check=True)
                            pdf_file = final_file.replace('.docx', '.pdf')
                            with open(pdf_file, "rb") as p_file:
                                st.download_button("📕 PDF ડાઉનલોડ કરો", p_file, file_name=pdf_file, mime="application/pdf")
                        except Exception as e:
                            st.error(f"⚠️ PDF એરર: {e}")
            except Exception as e:
                st.error(f"Error: {e}")
    else:
        st.warning("કૃપા કરીને પ્રશ્નો પેસ્ટ કરો.")
