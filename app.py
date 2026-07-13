import streamlit as st

# પેજનું સેટિંગ્સ
st.set_page_config(page_title="Solid Black Paper Generator", page_icon="📝", layout="wide")

# મુખ્ય હેડિંગ
st.markdown('<h1 style="text-align: center; color: #000000;">Solid Black Paper Generator</h1>', unsafe_allow_html=True)
st.markdown('---')

# સેમ્પલ પ્રશ્નોનો ડેટાબેઝ (તમારા ઓરીજનલ કોડ મુજબ ફેરફાર કરી શકો છો)
if 'questions' not in st.session_state:
    st.session_state.questions = [
        {
            "text": "નળાકારના પાયાની ત્રિજ્યા 7 સેમી અને ઊંચાઈ 10 સેમી હોય, તો તેનું વક્રસપાટીનું ક્ષેત્રફળ શોધો. આ પ્રશ્ન લાંબો હોઈ શકે છે જેથી જસ્ટિફિકેશન બરાબર દેખાય.",
            "options": ["440 ચોસેમી", "220 ચોસેમી", "154 ચોસેમી", "308 ચોસેમી"]
        },
        {
            "text": "જો દ્વિઘાત સમીકરણના બીજ સમાન હોય, તો વિવેચક D નું મૂલ્ય કેટલું થાય? આ પ્રશ્નને પણ જસ્ટિફાઈડ ફોર્મેટમાં ચેક કરવા માટે લાંબો લખાણ આપવામાં આવ્યો છે.",
            "options": ["D > 0", "D < 0", "D = 0", "D થઈ શકતું નથી"]
        }
    ]

# પ્રશ્નો ડિસ્પ્લે કરવાનો વિભાગ
st.markdown('### 📝 પ્રશ્નપત્ર (Question Paper)')

# લૂપ ચલાવીને બધા જ સવાલોને Justified ફોર્મેટમાં પ્રિન્ટ કરવા
for i, q in enumerate(st.session_state.questions, 1):
    # સવાલને Justified કરવા માટે <div> ટેગ નો ઉપયોગ
    st.markdown(
        f'<div style="text-align: justify; font-size: 18px; font-weight: 500; margin-bottom: 8px;">'
        f'<b>પ્રશ્ન {i}:</b> {q["text"]}'
        f'</div>', 
        unsafe_allow_html=True
    )
    
    # ઓપ્શન્સ (વિકલ્પો) ને પણ Justified કરવા માટે
    cols = st.columns(2)
    for idx, option in enumerate(q["options"]):
        with cols[idx % 2]:
            st.markdown(
                f'<div style="text-align: justify; font-size: 16px; margin-left: 20px; margin-bottom: 12px;">'
                f'({chr(65 + idx)}) {option}'
                f'</div>', 
                unsafe_allow_html=True
            )
    st.markdown('<br>', unsafe_allow_html=True)

st.markdown('---')
st.caption("Solid Black Educational Brand - 2026")
