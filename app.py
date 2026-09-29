import streamlit as st
import re
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

# --- 1. Page Config & Custom CSS ---
st.set_page_config(page_title="AITS Advanced Paper Generator", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Hind+Vadodara:wght@400;600;700&display=swap');
    html, body, [class*="css"], p, h1, h2, h3, h4, span, label { font-family: 'Hind Vadodara', sans-serif !important; }
    .main-title { text-align: center; font-weight: 700; font-size: 32px; margin-top: 5px; margin-bottom: 5px; color: #1a1a1a; }
    div.stButton > button:first-child { background-color: #1F4E79; color: #ffffff; border-radius: 6px; padding: 10px 24px; font-size: 18px; font-weight: bold; width: 100%; border: none; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. સ્માર્ટ ફોન્ટ અને ટેક્સ્ટ પ્રોસેસિંગ ---
def is_gujarati(char):
    """ચેક કરે છે કે અક્ષર ગુજરાતી છે કે નહીં"""
    return '\u0A80' <= char <= '\u0AFF'

def add_run_with_fonts(paragraph, text, is_bold, eng_font="Times New Roman", guj_font="Shruti", size=11):
    """એક જ વાક્યમાં ગુજરાતી અને અંગ્રેજી ફોન્ટ્સને સ્માર્ટ રીતે મિક્સ કરે છે"""
    if not text: return
    current_script = 'guj' if is_gujarati(text[0]) else 'eng'
    current_text = ""
    
    for char in text:
        char_script = 'guj' if is_gujarati(char) else 'eng'
        if char_script == current_script:
            current_text += char
        else:
            run = paragraph.add_run(current_text)
            run.bold = is_bold
            run.font.name = guj_font if current_script == 'guj' else eng_font
            run.font.size = Pt(size)
            current_text = char
            current_script = char_script
            
    if current_text:
        run = paragraph.add_run(current_text)
        run.bold = is_bold
        run.font.name = guj_font if current_script == 'guj' else eng_font
        run.font.size = Pt(size)

def parse_markdown_to_cell(cell, text, eng_font, guj_font, size):
    """માર્કડાઉન (Bold અને Tables) ને Word ફોર્મેટમાં કન્વર્ટ કરે છે"""
    lines = text.split('\n')
    cell.paragraphs[0].text = "" # Clear default paragraph
    current_p = cell.paragraphs[0]
    first_p = True
    in_table = False
    table = None
    
    for line in lines:
        if line.strip().startswith('|') and line.strip().endswith('|'):
            if '---' in line: continue # Skip markdown separator
            cells_text = [c.strip() for c in line.strip('|').split('|')]
            if not in_table:
                table = cell.add_table(rows=1, cols=len(cells_text))
                table.style = 'Table Grid'
                row_cells = table.rows[0].cells
                in_table = True
            else:
                row_cells = table.add_row().cells
            
            for i, c_text in enumerate(cells_text):
                if i < len(row_cells):
                    p = row_cells[i].paragraphs[0]
                    parts = c_text.split('**')
                    for j, part in enumerate(parts):
                        is_bold = (j % 2 != 0)
                        add_run_with_fonts(p, part, is_bold, eng_font, guj_font, size)
        else:
            in_table = False
            if line.strip() == "":
                cell.add_paragraph("")
                continue
            
            if not first_p:
                current_p = cell.add_paragraph()
            first_p = False
            
            current_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            parts = line.split('**')
            for j, part in enumerate(parts):
                is_bold = (j % 2 != 0)
                add_run_with_fonts(current_p, part, is_bold, eng_font, guj_font, size)

def split_questions(text):
    """Q.1, 1., **1.** વગેરે કોઈપણ ફોર્મેટમાંથી પ્રશ્નો છૂટા પાડે છે"""
    if not text.strip(): return []
    # Regex ટુ મેચ: **1. અથવા 1. અથવા Q.1
    pattern = r'(?m)^(?:\*\*)?(?:Q\.)?\s*\d+\.\s*(?:\*\*)?'
    matches = list(re.finditer(pattern, text))
    
    if not matches: return [text.strip()]
    
    questions = []
    # પ્રશ્નો શરૂ થતાં પહેલાની સૂચનાઓ (Instructions)
    if matches[0].start() > 0:
        pre_text = text[:matches[0].start()].strip()
        if pre_text: questions.append(pre_text)
            
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(text)
        questions.append(text[start:end].strip())
        
    return questions

# --- 3. Word Document જનરેટર ---
def generate_aits_doc(eng_text, guj_text, font_size, font_name, use_watermark):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        # AITS હેડર (દરેક પેજ પર)
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = "AITS"
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header_para.runs[0].font.name = "Arial"
        header_para.runs[0].font.bold = True
        header_para.runs[0].font.size = Pt(12)
        
        # વોટરમાર્ક અને ફૂટર ઓપ્શનલ
        if use_watermark:
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
                try:
                    footer_para.add_run().add_picture('footer.png', width=Inches(7.5))
                except Exception: pass

    # --- અદ્રશ્ય ટેબલ (લેઆઉટ) ---
    eng_questions = split_questions(eng_text)
    guj_questions = split_questions(guj_text)
    
    max_len = max(len(eng_questions), len(guj_questions))
    
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(3.75)
    table.columns[1].width = Inches(3.75)

    for i in range(max_len):
        row = table.add_row()
        eng_q = eng_questions[i] if i < len(eng_questions) else ""
        guj_q = guj_questions[i] if i < len(guj_questions) else ""
        
        parse_markdown_to_cell(row.cells[0], eng_q, "Times New Roman", font_name, font_size)
        parse_markdown_to_cell(row.cells[1], guj_q, "Times New Roman", font_name, font_size)
        
        # પ્રશ્નો વચ્ચે ગેપ
        row.cells[0].add_paragraph("")
        row.cells[1].add_paragraph("")

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- 4. Streamlit UI ---
st.markdown("<h1 class='main-title'>AITS Advanced Paper Generator</h1>", unsafe_allow_html=True)

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ")
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1: file_name = st.text_input("ફાઈલનું નામ:", value="AITS_Paper")
with col_s2: font_name = st.selectbox("ગુજરાતી ફોન્ટ:", ["Shruti", "Arial Unicode MS", "Hind Vadodara"])
with col_s3: font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=11)

use_watermark = st.checkbox("🖌️ પેપરમાં ફૂટર (footer.png) અને વોટરમાર્ક (sblogo.png) ઉમેરો?", value=False)

st.markdown("### ✍️ પ્રશ્નો પેસ્ટ કરો (જેવા છે તેવા જ)")
col_q1, col_q2 = st.columns(2)
with col_q1:
    st.markdown("#### 🇺🇸 અંગ્રેજી પ્રશ્નો (ડાબી બાજુ)")
    eng_input = st.text_area("Paste English Questions here...", height=400)
with col_q2:
    st.markdown("#### 🇮🇳 ગુજરાતી પ્રશ્નો (જમણી બાજુ)")
    guj_input = st.text_area("અહીં ગુજરાતી પ્રશ્નો પેસ્ટ કરો...", height=400)

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 પેપર જનરેટ કરો"):
    if not eng_input.strip() and not guj_input.strip():
        st.error("⚠️ ભૂલ: કૃપા કરીને પ્રશ્નો દાખલ કરો!")
    else:
        with st.spinner("તમારું એડવાન્સ AITS પેપર બની રહ્યું છે... પ્લીઝ વેઇટ ⏳"):
            try:
                docx_file = generate_aits_doc(eng_input, guj_input, font_size, font_name, use_watermark)
                st.success("✅ તમારું પેપર સફળતાપૂર્વક બની ગયું છે!")
                
                st.download_button(
                    label="📄 Word ફાઇલ ડાઉનલોડ કરો",
                    data=docx_file,
                    file_name=f"{file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"⚠️ કંઈક ભૂલ થઈ: {e}")
