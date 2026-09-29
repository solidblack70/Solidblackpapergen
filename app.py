import streamlit as st
import subprocess
import re
import shutil
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml

# --- 1. Page Config & Custom CSS ---
st.set_page_config(page_title="AITS Paper Generator", layout="wide", page_icon="📝")

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

# --- 2. તમારો ઓરિજિનલ સ્માર્ટ માર્કડાઉન પાર્સર કોડ (100% Copy-Paste) ---
def format_content(raw_text, is_continuous, start_num=1, end_num=0):
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
        if end_num > 0 and q_num > end_num:
            q_num = 1
            
        if q_block.startswith('#'):
            clean_title = re.sub(r'^#+', '', q_block).strip()
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
            
            # તમારું 100% ઓરિજિનલ ગણિત
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

# --- 3. તમારું ઓરિજિનલ ફોર્મેટિંગ લોજિક ---
def apply_custom_formatting(doc_path, font_size, font_name):
    doc = Document(doc_path)
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    
    for i, paragraph in enumerate(paragraphs):
        # બ્રહ્માસ્ત્ર: ઓટો-બુલેટ કાયમ માટે બંધ
        paragraph.style = doc.styles['Normal']
        pPr = paragraph._element.get_or_add_pPr()
        numPrs = pPr.findall(qn('w:numPr'))
        for n in numPrs: pPr.remove(n)
            
        for run in paragraph.runs:
            if '‡' in run.text:
                run.text = run.text.replace('‡', '\t')
                
        text = paragraph.text.strip()
        
        # MCQ ફોર્મેટિંગ
        if re.match(r'^(\*\*|__)?Q\.\d+', text):
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
                paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.8), WD_TAB_ALIGNMENT.LEFT)
        else:
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
        # ફોન્ટ એપ્લાય (Math સમીકરણો બગડે નહિ તેનું ધ્યાન)
        for run in paragraph.runs:
            if run.font.name != 'Cambria Math':
                run.font.size = Pt(font_size)
                run.font.name = font_name
                
    doc.save(doc_path)

# --- 4. પ્રશ્નો ભેગા કરવાનું લોજિક (Merge Logic) ---
def extract_blocks(doc):
    blocks = []
    current_block = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if re.match(r'^(\*\*|__)?Q\.\s*\d+', text, re.IGNORECASE):
            if current_block: blocks.append(current_block)
            current_block = [p]
        elif text or current_block:
            current_block.append(p)
    if current_block: blocks.append(current_block)
    return blocks

def create_bilingual_paper(eng_docx, guj_docx, final_name, use_watermark):
    doc_eng = Document(eng_docx)
    doc_guj = Document(guj_docx)
    doc_final = Document()
    
    for section in doc_final.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        # AITS હેડર
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = "AITS"
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header_para.runs[0].font.name = "Arial"
        header_para.runs[0].font.bold = True
        header_para.runs[0].font.size = Pt(12)
        
        # વોટરમાર્ક અને ફૂટર
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

    eng_blocks = extract_blocks(doc_eng)
    guj_blocks = extract_blocks(doc_guj)
    
    table = doc_final.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(3.75)
    table.columns[1].width = Inches(3.75)
    
    max_len = max(len(eng_blocks), len(guj_blocks))
    
    for i in range(max_len):
        row = table.add_row()
        
        # અંગ્રેજી બ્લોક કોપી
        if i < len(eng_blocks):
            cell_eng = row.cells[0]
            for p in cell_eng.paragraphs:
                p._element.getparent().remove(p._element) # Remove default empty line
            for p in eng_blocks[i]:
                cell_eng._element.append(p._element)
                
        # ગુજરાતી બ્લોક કોપી
        if i < len(guj_blocks):
            cell_guj = row.cells[1]
            for p in cell_guj.paragraphs:
                p._element.getparent().remove(p._element)
            for p in guj_blocks[i]:
                cell_guj._element.append(p._element)
                
        # પ્રશ્નો વચ્ચે ગેપ માટે
        table.add_row()
        
    doc_final.save(final_name)

# --- 5. Streamlit UI ---
st.markdown("<h1 class='main-title'>AITS Advanced Paper Generator</h1>", unsafe_allow_html=True)

if not shutil.which("pandoc"):
    st.warning("⚠️ સિસ્ટમમાં PDF/LaTeX કન્વર્ટ કરવાનું સોફ્ટવેર (Pandoc) નથી.")
    if st.button("🔧 અત્યારે જ ઇન્સ્ટોલ કરો (ફક્ત 1 મિનિટ લાગશે)"):
        with st.spinner("ઇન્સ્ટોલ થઈ રહ્યું છે... પ્લીઝ 1 મિનિટ રાહ જુઓ ⏳"):
            os.system("sudo apt-get update && sudo apt-get install pandoc -y")
            st.success("✅ ઇન્સ્ટોલ થઈ ગયું! હવે પેપર બની જશે.")

st.markdown("### ⚙️ ફાઇલ સેટિંગ્સ")
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1: file_name = st.text_input("ફાઈલનું નામ:", value="AITS_Paper")
with col_s2: font_name = st.selectbox("ગુજરાતી ફોન્ટ:", ["Shruti", "Arial", "Hind Vadodara"])
with col_s3: font_size = st.number_input("ફોન્ટ સાઈઝ:", min_value=8, max_value=20, value=11)

use_watermark = st.checkbox("🖌️ પેપરમાં ફૂટર (footer.png) અને વોટરમાર્ક (sblogo.png) ઉમેરો?", value=False)

st.markdown("### ✍️ પ્રશ્નો પેસ્ટ કરો (જેવા છે તેવા જ)")
col_q1, col_q2 = st.columns(2)
with col_q1:
    st.markdown("#### 🇺🇸 અંગ્રેજી પ્રશ્નો (ડાબી બાજુ)")
    eng_input = st.text_area("Paste English Questions here...", height=350)
with col_q2:
    st.markdown("#### 🇮🇳 ગુજરાતી પ્રશ્નો (જમણી બાજુ)")
    guj_input = st.text_area("અહીં ગુજરાતી પ્રશ્નો પેસ્ટ કરો...", height=350)

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 પેપર જનરેટ કરો"):
    if not eng_input.strip() and not guj_input.strip():
        st.error("⚠️ ભૂલ: કૃપા કરીને પ્રશ્નો દાખલ કરો!")
    else:
        with st.spinner("તમારું લેટેક્સ વાળું એડવાન્સ પેપર બની રહ્યું છે... પ્લીઝ વેઇટ ⏳"):
            try:
                # 1. માર્કડાઉનમાં કન્વર્ટ કરો (તમારું અસલી લોજીક)
                eng_md = format_content(eng_input, is_continuous=True)
                guj_md = format_content(guj_input, is_continuous=True)
                
                with open("temp_eng.md", "w", encoding="utf-8") as f: f.write(eng_md)
                with open("temp_guj.md", "w", encoding="utf-8") as f: f.write(guj_md)
                
                # 2. Pandoc થી વર્ડ ફાઈલ બનાવો (LaTeX કન્વર્ઝન માટે)
                subprocess.run(["pandoc", "temp_eng.md", "-o", "temp_eng.docx"], check=True)
                subprocess.run(["pandoc", "temp_guj.md", "-o", "temp_guj.docx"], check=True)
                
                # 3. ફોર્મેટિંગ એપ્લાય કરો
                apply_custom_formatting("temp_eng.docx", font_size, "Times New Roman")
                apply_custom_formatting("temp_guj.docx", font_size, font_name)
                
                # 4. બંનેને એક ફાઈલમાં મર્જ કરો
                final_file = f"{file_name}.docx"
                create_bilingual_paper("temp_eng.docx", "temp_guj.docx", final_file, use_watermark)
                
                st.success("✅ તમારું પેપર સફળતાપૂર્વક બની ગયું છે!")
                
                with open(final_file, "rb") as file:
                    st.download_button(
                        label="📄 Word ફાઇલ ડાઉનલોડ કરો",
                        data=file,
                        file_name=final_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"⚠️ કંઈક ભૂલ થઈ: {e}")
