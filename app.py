
import streamlit as st
import pdfplumber
import re
import pandas as pd
import matplotlib.pyplot as plt
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from datetime import datetime

st.set_page_config(page_title="Akıllı Tahlil Asistanı", page_icon="🧬", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    .list-row { background-color: #ffffff; padding: 15px 20px; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 10px; border: 1px solid #e2e8f0; display: flex; align-items: center; }
    .list-header { font-weight: 600; color: #475569; font-size: 14px; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0; margin-bottom: 15px; }
    .badge { padding: 6px 12px; border-radius: 20px; font-size: 13px; font-weight: 700; display: inline-block; text-align: center; min-width: 80px; }
    .badge-low { background-color: #dbeafe; color: #1e40af; }
    .badge-normal { background-color: #dcfce3; color: #166534; }
    .badge-high { background-color: #fee2e2; color: #991b1b; }
    </style>
""", unsafe_allow_html=True)

CLINICAL_DICT = {
    "1. Kalp ve Damar Sistemleri": {
        "Troponin I": {"aliases": ["Troponin I", "Trop I"], "min": 0, "max": 0.04, "unit": "ng/mL"},
        "Total Kolesterol": {"aliases": ["Kolesterol", "Total Kolesterol"], "min": 120, "max": 200, "unit": "mg/dL"},
        "HDL Kolesterol": {"aliases": ["HDL", "HDL Kolesterol"], "min": 40, "max": 100, "unit": "mg/dL"},
        "LDL Kolesterol": {"aliases": ["LDL", "LDL Kolesterol"], "min": 0, "max": 130, "unit": "mg/dL"},
        "Trigliserit": {"aliases": ["Trigliserit", "Triglyceride"], "min": 0, "max": 150, "unit": "mg/dL"}
    },
    "2. Karaciğer ve Safra": {
        "AST (SGOT)": {"aliases": ["AST", "SGOT"], "min": 0, "max": 40, "unit": "U/L"},
        "ALT (SGPT)": {"aliases": ["ALT", "SGPT"], "min": 0, "max": 41, "unit": "U/L"},
        "GGT": {"aliases": ["GGT"], "min": 0, "max": 60, "unit": "U/L"}
    },
    "3. Böbrek Fonksiyonları": {
        "Üre": {"aliases": ["Üre", "Urea"], "min": 17, "max": 43, "unit": "mg/dL"},
        "Kreatinin": {"aliases": ["Kreatinin", "Creatinine"], "min": 0.6, "max": 1.2, "unit": "mg/dL"}
    },
    "4. Diyabet Metabolizması": {
        "Açlık Kan Şekeri": {"aliases": ["Açlık Kan Şekeri", "AKŞ", "Glukoz (Açlık)"], "min": 70, "max": 100, "unit": "mg/dL"},
        "HbA1c": {"aliases": ["HbA1c"], "min": 4.0, "max": 5.7, "unit": "%"}
    },
    "5. Tiroid Fonksiyonları": {
        "TSH": {"aliases": ["TSH"], "min": 0.4, "max": 4.0, "unit": "µIU/mL"},
        "Serbest T4": {"aliases": ["Serbest T4", "sT4"], "min": 0.82, "max": 1.77, "unit": "ng/dL"}
    },
    "6. Mineraller": {
        "Demir (Fe)": {"aliases": ["Demir", "Fe "], "min": 60, "max": 170, "unit": "µg/dL"},
        "B12 Vitamini": {"aliases": ["B12", "Vitamin B12"], "min": 200, "max": 900, "unit": "pg/mL"},
        "D Vitamini": {"aliases": ["D Vitamini", "25-OH"], "min": 30, "max": 100, "unit": "ng/mL"}
    },
    "7. Hemogram (Tam Kan Sayımı)": {
        "WBC (Lökosit)": {"aliases": ["WBC", "Lökosit"], "min": 4.0, "max": 10.0, "unit": "10^3/µL"},
        "RBC (Eritrosit)": {"aliases": ["RBC", "Eritrosit"], "min": 4.0, "max": 6.0, "unit": "10^6/µL"},
        "Hemoglobin (HGB)": {"aliases": ["HGB", "Hemoglobin"], "min": 12.0, "max": 17.5, "unit": "g/dL"}
    }
}

def get_badge(value, min_val, max_val):
    if value < min_val: return "<span class='badge badge-low'>Düşük</span>", "Düşük"
    elif value > max_val: return "<span class='badge badge-high'>Yüksek</span>", "Yüksek"
    else: return "<span class='badge badge-normal'>Normal</span>", "Normal"

def extract_date_from_text(text):
    match = re.search(r'\b(\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4})\b', text)
    return match.group(1).replace('-', '.').replace('/', '.').replace(' ', '') if match else datetime.now().strftime("%d.%m.%Y")

def clean_date_time_from_line(line):
    line = re.sub(r'\b\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4}\b', ' ', line)
    line = re.sub(r'\b\d{4}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{1,2}\b', ' ', line)
    line = re.sub(r'\b\d{1,2}:\d{2}(?::\d{2})?\b', ' ', line)
    return line

def extract_test_value(line, aliases):
    for alias in aliases:
        pattern = r'\b' + re.escape(alias) + r'\b' if alias.replace(' ', '').isalnum() else re.escape(alias)
        match = re.search(pattern, line, re.IGNORECASE)
        if match:
            right_part = line[match.end():]
            val_match = re.search(r'[-+]?\d+[\.,]\d+|[-+]?\d+', right_part)
            if val_match:
                try: return float(val_match.group(0).replace(',', '.'))
                except: continue
    return None

def process_pdfs(files):
    all_results = []
    for file in files:
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted: text += extracted + "\n"
                
        report_date = extract_date_from_text(text)
        found_in_this_file = set()
        
        for line in text.split('\n'):
            cleaned_line = clean_date_time_from_line(line)
            for category, tests in CLINICAL_DICT.items():
                for test_name, test_info in tests.items():
                    if test_name in found_in_this_file: continue
                    value = extract_test_value(cleaned_line, test_info["aliases"])
                    if value is not None:
                        found_in_this_file.add(test_name)
                        all_results.append({"Tarih": report_date, "Kategori": category, "Tahlil": test_name, "Değer": value, "Min": test_info["min"], "Max": test_info["max"], "Birim": test_info["unit"]})
    return pd.DataFrame(all_results)

def generate_excel_with_chart(test_name, df, fig):
    wb = Workbook()
    ws = wb.active
    ws.title = test_name[:30]
    ws.append(["Tarih", "Değer", "Birim", "Referans Aralığı", "Durum"])
    for _, row in df.iterrows():
        ref_str = f"{row['Min']} - {row['Max']}"
        _, durum = get_badge(row['Değer'], row['Min'], row['Max'])
        ws.append([row['Tarih'], row['Değer'], row['Birim'], ref_str, durum])
    img_data = io.BytesIO()
    fig.savefig(img_data, format='png', bbox_inches='tight', dpi=100)
    img_data.seek(0)
    ws.add_image(OpenpyxlImage(img_data), "G2")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

@st.dialog("📊 Tahlil Analizi ve Zaman Çizelgesi", width="large")
def show_details_modal(test_name, df_test, min_val, max_val, unit):
    st.markdown(f"### {test_name} Değişim Trendi")
    df_test['Tarih_Obj'] = pd.to_datetime(df_test['Tarih'], format="%d.%m.%Y", errors='coerce')
    df_test = df_test.sort_values("Tarih_Obj").dropna(subset=['Tarih_Obj']).drop(columns=['Tarih_Obj'])
    
    fig, ax = plt.subplots(figsize=(8, 4))
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#ffffff')
    ax.grid(True, linestyle='--', alpha=0.6, color='#cbd5e1')
    ax.plot(df_test['Tarih'], df_test['Değer'], marker='o', color='#2563eb', linewidth=2, markersize=8)
    ax.axhline(min_val, color='#16a34a', linestyle='--', alpha=0.8, label='Alt Limit')
    ax.axhline(max_val, color='#dc2626', linestyle='--', alpha=0.8, label='Üst Limit')
    ax.set_ylabel(unit, fontsize=10, color='#475569')
    ax.tick_params(axis='x', rotation=45, colors='#475569')
    for spine in ax.spines.values(): spine.set_color('#cbd5e1')
    ax.legend(loc='best', frameon=True)
    st.pyplot(fig)
    
    display_df = df_test[['Tarih', 'Değer', 'Birim']].copy()
    display_df['Durum'] = display_df['Değer'].apply(lambda x: get_badge(x, min_val, max_val)[1])
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    st.download_button(label="📥 Excel Olarak İndir (Grafikli)", data=generate_excel_with_chart(test_name, df_test, fig), file_name=f"{test_name}_Rapor.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

st.title("🧬 Akıllı Tahlil Asistanı")
st.markdown("Geçmiş hastane (e-Nabız) PDF raporlarınızı yükleyin. Sistem tahlillerinizi hayati sıraya göre düzenleyip görselleştirsin.")

uploaded_files = st.file_uploader("Tahlil Sonuçları (PDF) Yükle", type="pdf", accept_multiple_files=True)
if uploaded_files:
    with st.spinner("Tahliller analiz ediliyor..."):
        df_results = process_pdfs(uploaded_files)
    if df_results.empty: st.warning("Desteklenen tahlil bulunamadı.")
    else:
        df_results['Tarih_Obj'] = pd.to_datetime(df_results['Tarih'], format="%d.%m.%Y", errors='coerce')
        st.markdown("<div class='list-header'>", unsafe_allow_html=True)
        h_cols = st.columns([3, 2, 2, 2, 2])
        h_cols[0].markdown("**Tahlil Adı**")
        h_cols[1].markdown("**Sonuç**")
        h_cols[2].markdown("**Referans Aralığı**")
        h_cols[3].markdown("**Durum**")
        h_cols[4].markdown("**Aksiyon**")
        st.markdown("</div>", unsafe_allow_html=True)

        for category in CLINICAL_DICT.keys():
            cat_df = df_results[df_results['Kategori'] == category]
            if not cat_df.empty:
                st.markdown(f"#### 🩺 {category}")
                for test in cat_df['Tahlil'].unique():
                    test_history = cat_df[cat_df['Tahlil'] == test].copy()
                    latest_row = test_history.loc[test_history['Tarih_Obj'].idxmax()]
                    badge_html, _ = get_badge(latest_row['Değer'], latest_row['Min'], latest_row['Max'])
                    
                    with st.container():
                        st.markdown("<div class='list-row'>", unsafe_allow_html=True)
                        cols = st.columns([3, 2, 2, 2, 2])
                        cols[0].markdown(f"**{test}**")
                        cols[1].markdown(f"<span style='font-size:16px; font-weight:600;'>{latest_row['Değer']}</span> <span style='color:gray; font-size:12px;'>{latest_row['Birim']}</span>", unsafe_allow_html=True)
                        cols[2].markdown(f"<span style='color:gray;'>{latest_row['Min']} - {latest_row['Max']}</span>", unsafe_allow_html=True)
                        cols[3].markdown(badge_html, unsafe_allow_html=True)
                        if cols[4].button("📊 İncele", key=f"btn_{test}_{category}"):
                            show_details_modal(test, test_history, latest_row['Min'], latest_row['Max'], latest_row['Birim'])
                        st.markdown("</div>", unsafe_allow_html=True)
