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
    pPr.append(sectPr)

# --- 3 કોલમ હેડર ડિઝાઇન (Times New Roman + Thick Line + Big Logo) ---
def insert_header_table(doc, header_left, header_center):
    h_font = "Times New Roman" # હેડર માટે ફોન્ટ ફિક્સ
    
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    
    # ટેબલને એકદમ ઉપર ખસેડ્યું
    first_para = doc.paragraphs[0]._p
    first_para.addprevious(table._tbl)
    
    table.columns[0].width = Inches(2.4)
    table.columns[1].width = Inches(3.5)
    table.columns[2].width = Inches(1.2)
    
    # 1. ડાબી બાજુ (મોટો લોગો અને ગ્રે બોક્સ જાડી લાઈન સાથે)
    left_cell = table.cell(0, 0)
    p_logo_left = left_cell.paragraphs[0]
    p_logo_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_logo_left.add_run().add_picture('small_logo.png', width=Inches(2.4))
    except:
        pass
        
    lines = header_left.strip().split('\n')
    for idx, line in enumerate(lines):
        p_gray = left_cell.add_paragraph()
        p_gray.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        pPr = p_gray._element.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'D9D9D9')
        pPr.append(shd)
        
        # પહેલી લાઈન (MCQ) ની નીચે જાડી લાઈન મૂકો
        if idx == 0 and len(lines) > 1:
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '24') # જાડી લાઈન (3 pt)
            bottom.set(qn('w:space'), '4')
            bottom.set(qn('w:color'), '000000')
            pBdr.append(bottom)
            pPr.append(pBdr)
        
        run_gray = p_gray.add_run(line)
        run_gray.font.name = h_font
        run_gray.font.bold = True
        run_gray.font.size = Pt(14)
    
    # 2. વચ્ચેનો ભાગ
    center_cell = table.cell(0, 1)
    p_center = center_cell.paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lines = header_center.strip().split('\n')
    for i, line in enumerate(lines):
        r = p_center.add_run(line)
        r.font.name = h_font
        r.font.bold = True
        r.font.size = Pt(20) if i == 0 else Pt(14)
        if i < len(lines) - 1:
            p_center.add_run('\n')
            
    # 3. જમણી બાજુ
    right_cell = table.cell(0, 2)
    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    try:
        p_right.add_run().add_picture('round_logo.png', width=Inches(1.2))
    except:
        pass

    # હેડર ટેબલ પછી 1-કોલમ નો બ્રેક મૂકો (જેથી ૨-કોલમ અંદર ઘૂસી ન જાય)
    p_break = doc.add_paragraph()
    table._tbl.addnext(p_break._p)
    add_continuous_section_break(p_break, 1)

# --- 2. વર્ડ ફાઈલનું ફોર્મેટિંગ ---
def set_formatting_and_margins(docx_filename, font_size, font_name, header_left, header_center):
    doc = Document(docx_filename)
    
    # આખા ડોક્યુમેન્ટને બાય-ડિફોલ્ટ 2 કોલમમાં સેટ કરો અને માર્જિન સેટ કરો
    for section in doc.sections:
        section.top_margin = Inches(0.3) # હેડરની ઉપર થોડીક જ જગ્યા 
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5) # Narrow
        section.right_margin = Inches(0.5) # Narrow
        section.header_distance = Inches(0.1)
        section.footer_distance = Inches(0.1)
        
        # --- વોટરમાર્ક (Watermark) સેટિંગ ---
        header = section.header
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        try:
            image_part, rel_id = header.part.get_or_add_image('Asset 345@4x-8.png')
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
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            header_para._p.append(parse_xml(watermark_xml))
        except Exception:
            pass 
            
        # --- ફૂટર (Footer) સેટિંગ ---
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            footer_para.add_run().add_picture('FOTTER@4x-8.png', width=Inches(7.5))
        except Exception:
            pass
        
        sectPr = section._sectPr
        cols = sectPr.find(qn('w:cols')) or OxmlElement('w:cols')
        if cols not in sectPr: sectPr.append(cols)
        cols.set(qn('w:num'), '2')       
        cols.set(qn('w:space'), '720')   
        cols.set(qn('w:sep'), '1')       
    
    # 2-કોલમ સેટ કર્યા પછી, હેડર ટેબલ 1-કોલમમાં ઇન્સર્ટ કરો
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
        
        if '###HEADER###' in text:
            clean_title = text.replace('###HEADER###', '').strip()
            if i > 0:
                add_continuous_section_break(paragraphs[i-1], 2)
            
            add_continuous_section_break(paragraph, 1)
            
            paragraph.text = ""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)
            
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
            run.font.name = "Times New Roman" # ટાઇટલ ના ફોન્ટ ફિક્સ
            
            r = run._element
            rPr = r.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:ascii'), "Times New Roman")
            rFonts.set(qn('w:hAnsi'), "Times New Roman")
            rFonts.set(qn('w:cs'), "Times New Roman")
            continue

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
def format_content(raw_text, is_continuous, start_num):
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
    q_num = start_num 
    
    q_prefix_pattern = r'^([\s]*([Qq]\.?\s*\d+[\.\-\)]*|\d+[\.\-\)]+)\s*)+'
    labels = ['A', 'B', 'C', 'D']
    
    for q_block in questions:
        if q_block.startswith('#'):
            if not is_continuous:
                q_num = start_num
                
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

# --- 4. Streamlit UI ---

col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass 

st.markdown("<h1 class='main-title'>Question Paper Generator</h1>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; margin-top: -15px; margin-bottom: 25px;'><span style='background-color: #000000; color: #ffffff; padding: 6px 18px; border-radius: 20px; font-size: 15px; font-weight: 700; box-shadow: 0px 4px 6px rgba(0,0,0,0.2); letter-spacing: 0.5px;'>Made by Yug Ghanshyam Padmani</span></div>", unsafe_allow_html=True)

st.markdown("### 📝 પેપરનું હેડર ડિઝાઇન")
col_h1, col_h2 = st.columns(2)
with col_h1:
    header_left = st.text_area("ડાબી બાજુનું ગ્રે બોક્સ:", "MCQ\nGUJARATI MEDIUM", height=130)
with col_h2:
    header_center = st.text_area("વચ્ચેનું મેઈન ટાઈટલ:", "STD 12 SCIENCE\nMATHS\n40 MARKS\nDate : 07/08/26", height=130)

st.divider()

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ")
col1, col2, col3, col4 = st.columns(4)
with col1:
    file_name = st.text_input("ફાઈલનું નામ:", "Solid_Black_Paper")
with col2:
    start_num = st.number_input("પ્રશ્ન નંબર ક્યાંથી શરૂ કરવો છે?", min_value=1, value=1)
with col3:
    font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=10)
with col4:
    font_name = st.selectbox("પેપરનો ફોન્ટ:", ["Hind Vadodara", "Shruti", "Cambria Math", "Noto Serif", "Times New Roman", "Calibri", "Arial"])

is_continuous = st.checkbox("પ્રશ્નોના નંબર સળંગ (Continuous) રાખવા છે?", value=True)

st.markdown("### ✍️ પ્રશ્નો")
user_input = st.text_area("અહીં પ્રશ્નો પેસ્ટ કરો (ટાઇટલ/સૂચના મૂકવા તેની આગળ # લખો):", height=280)

if st.button("વર્ડ અને PDF ફાઇલ જનરેટ કરો"):
    if user_input.strip():
        with st.spinner("Processing your document..."):
            processed_md = format_content(user_input, is_continuous, start_num)
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
    else:
        st.warning("કૃપા કરીને પ્રશ્નો પેસ્ટ કરો.")
