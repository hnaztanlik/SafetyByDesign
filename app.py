# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import json
import matplotlib.pyplot as plt
import feedparser
import pydeck as pdk
import plotly.graph_objects as go

# ReportLab (PDF Engine)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# OpenPyXL (Excel)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

LIVE_DATA_FILE = "live_accidents.json"

def clean_tr(text):
    tr_map = {
        'ş': 's', 'Ş': 'S', 'ı': 'i', 'İ': 'I',
        'ğ': 'g', 'Ğ': 'G', 'ü': 'u', 'Ü': 'U',
        'ö': 'o', 'Ö': 'O', 'ç': 'c', 'Ç': 'C'
    }
    for tr_char, en_char in tr_map.items():
        text = str(text).replace(tr_char, en_char)
    return text

def get_base_database():
    return [
        {
            "incident_id": "MAIB-2023-014",
            "incident_date": "14/08/2023",
            "vessel_name": "M/Y Blue Horizon",
            "vessel_category": "Yat & Süperyat",
            "vessel_type": "Motor Yacht (Planing Hull)",
            "flag": "United Kingdom",
            "year_built": 2018,
            "loa": 18.2, "beam": 5.1, "draft": 1.25, "gross_tonnage": 38,
            "lat": 36.85, "lon": 28.27,
            "propulsion": "Twin Diesel (MAN V8)",
            "operating_area": "Akdeniz / Kıyısal (Coastal)",
            "incident_type": "Makine Dairesi Yangını",
            "root_cause": "Yüksek basınçlı yakıt borusu yırtılması ve sıcak egzoz manifoldu teması",
            "critical_zone": "Makine Dairesi",
            "severity": "Yüksek",
            "source": "MAIB",
            "report_url": "https://www.gov.uk/maib-reports",
            "summary_report": "M/Y Blue Horizon seyir halindeyken sancak ana makine yüksek basınç borusu titreşimle çatlamış, egzoz manifolduna püsküren mazot infilak etmiştir."
        },
        {
            "incident_id": "GISIS-2022-882",
            "incident_date": "03/11/2022",
            "vessel_name": "M/Y Sirena Star",
            "vessel_category": "Yat & Süperyat",
            "vessel_type": "Superyacht",
            "flag": "Malta",
            "year_built": 2015,
            "loa": 32.5, "beam": 7.4, "draft": 1.9, "gross_tonnage": 195,
            "lat": 35.89, "lon": 14.51,
            "propulsion": "Hybrid / Electric",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Dumandan Boğulma / Personel Mahsur Kalması",
            "root_cause": "Tek kaçış rotasının yangınla kapanması / Dikey kaçış şaftı eksikliği",
            "critical_zone": "Acil Kaçış Şaftı (Escape Trunk)",
            "severity": "Kritik",
            "source": "IMO GISIS",
            "report_url": "https://gisis.imo.org/Public/MCI/Default.aspx",
            "summary_report": "Makine dairesinde panodaki ark nedeniyle başlayan duman tek çıkış merdivenini kapatmış, bağımsız kaçış bacası bulunmadığından personel mahsur kalmıştır."
        },
        {
            "incident_id": "NTSB-2024-009",
            "incident_date": "19/04/2024",
            "vessel_name": "M/Y Aegea Green",
            "vessel_category": "Yat & Süperyat",
            "vessel_type": "Hybrid Superyacht",
            "flag": "United States",
            "year_built": 2021,
            "loa": 42.0, "beam": 8.8, "draft": 2.3, "gross_tonnage": 480,
            "lat": 25.76, "lon": -80.19,
            "propulsion": "Hybrid / Electric",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Batarya Termal Kaçak & Patlama",
            "root_cause": "Lityum-iyon batarya hücresi aşırı ısınması ve zehirli gaz tahliye yetersizliği",
            "critical_zone": "Batarya Kompartımanı",
            "severity": "Kritik",
            "source": "NTSB Marine",
            "report_url": "https://www.ntsb.gov/investigations/accidentreports/Pages/marine.aspx",
            "summary_report": "Lityum bataryada termal kaçak oluşmuş, zehirli gaz tahliye bacası yetersizliğinden yaşam alanına yayılmıştır."
        },
        {
            "incident_id": "EMCIP-2023-412",
            "incident_date": "28/01/2023",
            "vessel_name": "M/T Nordic Spirit",
            "vessel_category": "Ticari Gemi (Tanker)",
            "vessel_type": "Chemical / Oil Tanker",
            "flag": "Liberia",
            "year_built": 2012,
            "loa": 145.0, "beam": 23.0, "draft": 8.5, "gross_tonnage": 11500,
            "lat": 36.14, "lon": -5.35,
            "propulsion": "Slow Speed Diesel (MAN B&W)",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Kargo Tankı Patlaması & Yangın",
            "root_cause": "İnert gaz (IGS) sistem basınç kaybı ve hidrokarbon buharı statik kıvılcımı",
            "critical_zone": "Kargo Güvertesi / Tanklar",
            "severity": "Kritik",
            "source": "EMSA EMCIP",
            "report_url": "https://www.emsa.europa.eu/emcip-public-marine-casualties.html",
            "summary_report": "İnert gaz jeneratörünün basınç kaybı sonucu tank içi oksijen oranı yükselmiş, statik kıvılcımla tank infilak etmiştir."
        }
    ]

def get_solutions_catalog():
    return {
        "Yüksek basınçlı yakıt borusu yırtılması ve sıcak egzoz manifoldu teması": [
            {
                "name": "Çift Cidarlı (Double-Walled) Yakıt Borulama + Sürekli Basınç Takibi",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Sızıntıyı dış zırhta hapsederek sıcak yüzey temasını engeller.", "SOLAS II-2 ve DNV Klas kurallarına tam uyum sağlar."],
                "riskler": ["Boru donatım maliyetinde %35-40 artış."],
                "maliyet_etkisi": "Orta - Yüksek", "agirlik_etkisi": "+80 kg", "cad_referans": "DWG_Double_Walled_Piping_Standard.dwg"
            },
            {
                "name": "A-60 Egzoz Termal İzolasyon Ceketi + Otomatik Su Sisi (Water Mist)",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 2, "sev_mitigated": 4,
                "firsatlar": ["Düşük ilk yatırım maliyeti.", "Mevcut hatlara kolay montaj."],
                "riskler": ["İzolasyonun yağlanması durumunda ikincil yangın riski."],
                "maliyet_etkisi": "Düşük - Orta", "agirlik_etkisi": "+25 kg", "cad_referans": "DWG_Thermal_Jacket_WaterMist.dwg"
            }
        ],
        "Tek kaçış rotasının yangınla kapanması / Dikey kaçış şaftı eksikliği": [
            {
                "name": "Bağımsız Dikey Kaçış Bacası (Escape Trunk 600x600 mm) + Çelik Merdiven",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 2,
                "firsatlar": ["Açık havaya kesintisiz tahliye sağlar.", "Personel ölüm riskini sıfırlar."],
                "riskler": ["Üst güvertede 600x600 mm yer kaybı."],
                "maliyet_etkisi": "Orta", "agirlik_etkisi": "+50 kg", "cad_referans": "DWG_Escape_Trunk_Structural_Section.dwg"
            }
        ],
        "Lityum-iyon batarya hücresi aşırı ısınması ve zehirli gaz tahliye yetersizliği": [
            {
                "name": "A-60 İzolasyonlu Bağımsız Oda + Doğrudan Güverteye Açılan HF Gaz Tahliye Bacası",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Hidrojen Florür gazını yaşam mahalline girmeden tahliye eder.", "DNV Battery Safety notunu yükseltir."],
                "riskler": ["Ayrı kompartıman ve paslanmaz çelik baca maliyeti."],
                "maliyet_etkisi": "Yüksek", "agirlik_etkisi": "+120 kg", "cad_referans": "DWG_Battery_Room_Safety_Exhaust.dwg"
            }
        ]
    }

def create_3d_parametric_vessel(loa, beam, draft, category, critical_zones):
    depth = max(draft * 1.85, 2.8)
    half_beam = beam / 2.0
    
    u = np.linspace(0, 1, 30)
    v = np.linspace(0, np.pi, 16)
    U, V = np.meshgrid(u, v)
    
    form_factor = np.sin(U * np.pi / 2) ** 0.8
    keel_slope = (1 - U) * 0.18

    X = U * loa
    Y = half_beam * np.sin(V) * form_factor
    Z = depth * np.cos(V/2) - (draft * keel_slope)

    fig = go.Figure()

    # Gövde Sancak / İskele
    fig.add_trace(go.Surface(x=X, y=Y, z=Z, colorscale=[[0, '#0F172A'], [1, '#475569']], opacity=0.45, showscale=False, name="Gövde (Sancak)"))
    fig.add_trace(go.Surface(x=X, y=-Y, z=Z, colorscale=[[0, '#0F172A'], [1, '#475569']], opacity=0.45, showscale=False, name="Gövde (İskele)"))

    # Su Hattı
    wl_x = np.array([[0, loa], [0, loa]])
    wl_y = np.array([[-half_beam*1.1, -half_beam*1.1], [half_beam*1.1, half_beam*1.1]])
    wl_z = np.array([[draft, draft], [draft, draft]])
    fig.add_trace(go.Surface(x=wl_x, y=wl_y, z=wl_z, colorscale=[[0, '#0284C7'], [1, '#0284C7']], opacity=0.25, showscale=False, name=f"Su Hattı (T={draft}m)"))

    # Risk Blokları Ekleme
    def add_risk_box(x_start, x_end, b_scale, z_top, color, name):
        bx = [x_start, x_end, x_end, x_start, x_start, x_end, x_end, x_start]
        hb = half_beam * b_scale
        by = [-hb, -hb, hb, hb, -hb, -hb, hb, hb]
        bz = [0.15, 0.15, 0.15, 0.15, z_top, z_top, z_top, z_top]
        fig.add_trace(go.Mesh3d(x=bx, y=by, z=bz, alphahull=0, color=color, opacity=0.7, name=name, hoverinfo="name"))

    add_risk_box(loa * 0.10, loa * 0.38, 0.85, depth * 0.78, "#DC2626", "Makine Dairesi (Yüksek Yangın Riski)")
    add_risk_box(loa * 0.38, loa * 0.72, 0.90, depth * 0.85, "#F59E0B", "Batarya / Kargo Kompartımanı")
    add_risk_box(loa * 0.72, loa * 0.94, 0.60, depth * 0.90, "#10B981", "Baş Pik / Çatışma Perdesi")

    fig.update_layout(
        title=f"3D Parametrik Model: LOA={loa}m | B={beam}m | T={draft}m",
        scene=dict(
            xaxis=dict(title="Boy (X) [m]"),
            yaxis=dict(title="Genişlik (Y) [m]"),
            zaxis=dict(title="Yükseklik (Z) [m]"),
            aspectmode='data',
            camera=dict(eye=dict(x=-1.5, y=-1.8, z=1.25))
        ),
        margin=dict(l=0, r=0, b=0, t=35),
        height=500
    )
    return fig

# ==========================================
# STREAMLIT ARAYÜZ
# ==========================================
st.set_page_config(page_title="Safety-by-Design Maritime Platform", page_icon="⚓", layout="wide")

df_all = pd.DataFrame(get_base_database())
solutions_catalog = get_solutions_catalog()

st.title("⚓ Safety-by-Design: Gemi Kaza Analitiği & Karar Destek Platformu")
st.write("---")

st.subheader("1. Gemi ve Tasarım Parametreleri")
col1, col2 = st.columns(2)

with col1:
    vessel_name = st.text_input("Gemi / Proje Adı", value="M/Y Boreas 34")
    vessel_category = st.selectbox("Deniz Aracı Ana Kategorisi", options=["Yat & Süperyat", "Ticari Gemi (Tanker)", "Ticari Gemi (Kargo/Konteyner)", "Hizmet & Çalışma Gemileri (Tug/Supply)"])
    loa = st.number_input("Tam Boy (LOA) [Metre]", min_value=3.0, max_value=450.0, value=34.0, step=0.5)
    beam = st.number_input("Genişlik (Beam) [Metre]", min_value=1.0, max_value=75.0, value=7.6, step=0.1)

with col2:
    draft = st.number_input("Su Çekimi (Draft) [Metre]", min_value=0.3, max_value=25.0, value=1.9, step=0.05)
    displacement = st.number_input("Deplasman [Ton]", min_value=1.0, max_value=400000.0, value=185.0, step=1.0)
    propulsion = st.selectbox("Sevk / Yakıt Tipi", options=["Hybrid / Electric", "Twin Diesel", "Lithium Battery", "Slow Speed Diesel"])
    operating_area = st.selectbox("Operasyon / Seyir Bölgesi", options=["Açık Deniz / Okyanus (Unrestricted)", "Kıyısal / Yakın Sefer (Coastal)"])

st.write("---")

# Kaza Filtresi (Dinamik)
matched = df_all[
    (df_all["vessel_category"] == vessel_category) |
    (df_all["propulsion"] == propulsion)
]
if matched.empty:
    matched = df_all.copy()

st.subheader("2. Parametrik 3D Gemi Modeli & Risk Zonları")
vessel_3d_fig = create_3d_parametric_vessel(loa, beam, draft, vessel_category, matched["critical_zone"].tolist())
st.plotly_chart(vessel_3d_fig, use_container_width=True)

st.write("---")
st.subheader("3. Eşleşen Emsal Kaza Kayıtları")

for idx, row in matched.iterrows():
    with st.expander(f"🚨 [{row['incident_id']}] {row['vessel_name']} ({row['loa']}m) - {row['incident_type']}"):
        st.write(f"**Kök Neden:** {row['root_cause']}")
        st.write(f"**Kritik Bölge:** `{row['critical_zone']}` | **Şiddet:** :red[{row['severity']}] | **Kaynak:** {row['source']}")
        st.caption(row['summary_report'])

st.write("---")
st.subheader("4. FMEA Karar Destek & Önlem Seçimi")

selected_solutions_data = []
for idx, (rc, options) in enumerate(solutions_catalog.items()):
    st.markdown(f"**Kritik Arıza Modu {idx+1}:** *{rc}*")
    opt_names = [opt["name"] for opt in options]
    chosen_name = st.radio(f"Seçim {idx+1}", options=opt_names, key=f"rad_{idx}", label_visibility="collapsed")
    chosen_opt = next(item for item in options if item["name"] == chosen_name)
    
    r1 = chosen_opt["prob_initial"] * chosen_opt["sev_initial"]
    r2 = chosen_opt["prob_mitigated"] * chosen_opt["sev_mitigated"]
    reduction = round(((r1 - r2) / r1) * 100, 1)

    c1, c2, c3 = st.columns([2, 2, 1])
    with c1: st.success("✅ **Kazanımlar:**\n- " + "\n- ".join(chosen_opt["firsatlar"]))
    with c2: st.warning("⚠️ **Riskler:**\n- " + "\n- ".join(chosen_opt["riskler"]))
    with c3:
        st.metric("Risk Düşüşü", f"{r1} ➔ {r2}", delta=f"-%{reduction}")
        st.caption(f"Maliyet: {chosen_opt['maliyet_etkisi']}\nAğırlık: {chosen_opt['agirlik_etkisi']}")
    selected_solutions_data.append({"problem": rc, "solution": chosen_opt, "r1": r1, "r2": r2, "reduction": reduction})
    st.write("")
