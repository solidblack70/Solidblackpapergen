import streamlit as st
import re
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

# --- 1. Page Config & Custom CSS ---
st.set_page_config(page_title="AITS Bilingual Paper Generator", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Hind+Vadodara:wght@400;600;700&display=swap');
    html, body, [class*="css"], p, h1, h2, h3, h4, span, label { font-family: 'Hind Vadodara', sans-serif !important; }
    .main-title { text-align: center; font-weight: 700; font-size: 32px; margin-top: 5px; margin-bottom: 5px; color: #1a1a1a; }
    .subtitle-badge { text-align: center; margin-bottom: 25px; }
    .subtitle-badge span { background-color: #1a1a1a; color: #ffffff; padding: 5px 15px; border-radius: 20px; font-size: 13px; letter-spacing: 1px; }
    div.stButton > button:first-child { background-color: #1F4E79; color: #ffffff; border-radius: 6px; padding: 10px 24px; font-size: 18px; font-weight: bold; width: 100%; border: none; transition: 0.3s; }
    div.stButton > button:first-child:hover { background-color: #112d47; box-shadow: 0px 4px 10px rgba(0,0,0,0.2); }
    .stTextArea textarea { border-radius: 6px !important; border: 1px solid #999 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. પ્રશ્નો છૂટા પાડવાનું ફંક્શન ---
def parse_questions(text):
    """Q.1, Q.2 અથવા Q1, Q2 ના આધારે પ્રશ્નોને લિસ્ટમાં અલગ કરે છે"""
    if not text.strip(): return []
    # Regular expression થી પ્રશ્નો અલગ કરો (Q. નંબર જળવાઈ રહેશે)
    pattern = r'(?m)^(?=Q\.\s*\d+|Q\d+)'
    questions = re.split(pattern, text.strip())
    return [q.strip() for q in questions if q.strip()]

# --- 3. Word Document જનરેટર ---
def generate_bilingual_doc(eng_text, guj_text, font_size, font_name, use_footer_watermark):
    doc = Document()
    
    # પેજ માર્જિન સેટિંગ્સ
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        # AITS હેડર સેટિંગ (પ્રથમ પેજથી જ)
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = "AITS"
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header_para.runs[0].font.name = "Arial"
        header_para.runs[0].font.bold = True
        header_para.runs[0].font.size = Pt(12)
        
        # વોટરમાર્ક અને ફૂટર લોજીક (જો યુઝરે ચેકબોક્સ ટીક કર્યું હોય તો જ)
        if use_footer_watermark:
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
                    run = footer_para.add_run()
                    run.add_picture('footer.png', width=Inches(7.5))
                except Exception: pass

    # --- પ્રશ્નપત્ર (અદ્રશ્ય ટેબલ દ્વારા સામસામે) ---
    eng_questions = parse_questions(eng_text)
    guj_questions = parse_questions(guj_text)
    
    max_len = max(len(eng_questions), len(guj_questions))
    
    # 2 કોલમનું ટેબલ (ટેબલની બોર્ડર અદ્રશ્ય રહેશે)
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(3.75)
    table.columns[1].width = Inches(3.75)

    for i in range(max_len):
        row = table.add_row()
        
        # અંગ્રેજી પ્રશ્ન (ડાબી બાજુ)
        eng_q = eng_questions[i] if i < len(eng_questions) else ""
        cell_eng = row.cells[0]
        p_eng = cell_eng.paragraphs[0]
        p_eng.text = eng_q
        p_eng.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in p_eng.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_size)

        # ગુજરાતી પ્રશ્ન (જમણી બાજુ)
        guj_q = guj_questions[i] if i < len(guj_questions) else ""
        cell_guj = row.cells[1]
        p_guj = cell_guj.paragraphs[0]
        p_guj.text = guj_q
        p_guj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in p_guj.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            
        # પ્રશ્નો વચ્ચે થોડી જગ્યા રાખવા માટે
        row.cells[0].add_paragraph("")
        row.cells[1].add_paragraph("")

    # ફાઇલને મેમરીમાં સેવ કરો
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- 4. Streamlit UI ---
st.markdown("<h1 class='main-title'>AITS Bilingual Paper Generator</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle-badge'><span>Made by Yug Ghanshyam Padmani</span></div>", unsafe_allow_html=True)

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ")
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1: file_name = st.text_input("ફાઈલનું નામ:", value="AITS_Paper")
with col_s2: font_name = st.selectbox("ગુજરાતી ફોન્ટ:", ["Shruti", "Arial Unicode MS", "Hind Vadodara"])
with col_s3: font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=11)

# વોટરમાર્ક અને ફૂટર માટેનો ઓપ્શન
use_footer_watermark = st.checkbox("🖌️ પેપરમાં ફૂટર (footer.png) અને વોટરમાર્ક (sblogo.png) ઉમેરો?", value=False)

st.markdown("### ✍️ પ્રશ્નો પેસ્ટ કરો (Q.1, Q.2 ફોર્મેટમાં)")
col_q1, col_q2 = st.columns(2)
with col_q1:
    st.markdown("#### 🇺🇸 અંગ્રેજી પ્રશ્નો (ડાબી બાજુ)")
    eng_input = st.text_area("Paste English Questions here...", height=350, placeholder="Q.1 What is Physics?\n(1) A\n(2) B\n(3) C\n(4) D")
with col_q2:
    st.markdown("#### 🇮🇳 ગુજરાતી પ્રશ્નો (જમણી બાજુ)")
    guj_input = st.text_area("અહીં ગુજરાતી પ્રશ્નો પેસ્ટ કરો...", height=350, placeholder="Q.1 ભૌતિકવિજ્ઞાન શું છે?\n(1) A\n(2) B\n(3) C\n(4) D")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 પેપર જનરેટ કરો"):
    if not eng_input.strip() and not guj_input.strip():
        st.error("⚠️ ભૂલ: કૃપા કરીને પ્રશ્નો દાખલ કરો!")
    else:
        with st.spinner("તમારું પેપર બની રહ્યું છે... પ્લીઝ વેઇટ ⏳"):
            try:
                # વર્ડ ફાઇલ જનરેટ કરો
                docx_file = generate_bilingual_doc(eng_input, guj_input, font_size, font_name, use_footer_watermark)
                
                st.success("✅ તમારું પેપર સફળતાપૂર્વક બની ગયું છે!")
                
                # ડાઉનલોડ બટન
                st.download_button(
                    label="📄 Word ફાઇલ ડાઉનલોડ કરો",
                    data=docx_file,
                    file_name=f"{file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"⚠️ કંઈક ભૂલ થઈ: {e}")
