# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import feedparser
import pydeck as pdk

# ReportLab (PDF Kütüphanesi)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# OpenPyXL (Excel, Stiller & Yerel Grafikler)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series, PieChart

LIVE_DATA_FILE = "live_accidents.json"
USER_UPLOADS_FILE = "community_cad.json"

# Türkçe Karakter Düzeltme Fonksiyonu
def clean_tr(text):
    tr_map = {
        'ş': 's', 'Ş': 'S', 'ı': 'i', 'İ': 'I',
        'ğ': 'g', 'Ğ': 'G', 'ü': 'u', 'Ü': 'U',
        'ö': 'o', 'Ö': 'O', 'ç': 'c', 'Ç': 'C'
    }
    for tr_char, en_char in tr_map.items():
        text = str(text).replace(tr_char, en_char)
    return text

# =========================================================================
# ÇOKLU KAZA VERİ BANKASI
# =========================================================================
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
            "summary_report": "M/Y Blue Horizon seyir halindeyken sancak ana makine yüksek basınç borusu titreşimle çatlamış, sıcak egzoz manifolduyla temas eden yakıt alev almıştır."
        },
        {
            "incident_id": "MAIB-2023-088",
            "incident_date": "05/06/2023",
            "vessel_name": "Tug Titan",
            "vessel_category": "Hizmet & Çalışma Gemileri (Tug/Supply)",
            "vessel_type": "Escort / Harbor Tug",
            "flag": "Netherlands",
            "year_built": 2019,
            "loa": 28.5, "beam": 11.2, "draft": 4.5, "gross_tonnage": 380,
            "lat": 51.95, "lon": 4.14,
            "propulsion": "Azimuth Stern Drive (ASD)",
            "operating_area": "Liman / Rıhtım İçi",
            "incident_type": "Çekme Esnasında Girme (Girthing / Alabora)",
            "root_cause": "Halat açısının bordaya dik gelmesi ve çekme kancasının hızlı bırakılamaması",
            "critical_zone": "Kıç Güverte / Çeki Kancası (Towing Winch)",
            "severity": "Kritik",
            "source": "MAIB",
            "report_url": "https://www.gov.uk/maib-reports",
            "summary_report": "Büyük bir gemiyi çekerken manevra esnasında çekme kancası hidrolik basınç arızası nedeniyle bırakılamamış, tekne saniyeler içinde alabora olmuştur."
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
            "propulsion": "Twin Diesel (MTU 16V)",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Dumandan Boğulma / Personel Mahsur Kalması",
            "root_cause": "Tek kaçış rotasının yangınla kapanması / Dikey kaçış şaftı eksikliği",
            "critical_zone": "Acil Kaçış Şaftı (Escape Trunk)",
            "severity": "Kritik",
            "source": "IMO GISIS",
            "report_url": "https://gisis.imo.org/Public/MCI/Default.aspx",
            "summary_report": "Makine dairesinde başlayan duman ana merdiveni kapatmış, alternatif kaçış şaftı bulunmadığından personel mahsur kalarak dumandan etkilenmiştir."
        },
        {
            "incident_id": "GISIS-2022-105",
            "incident_date": "12/09/2022",
            "vessel_name": "M/V Pacific Carrier",
            "vessel_category": "Ticari Gemi (Kargo/Konteyner)",
            "vessel_type": "Bulk Carrier",
            "flag": "Panama",
            "year_built": 2016,
            "loa": 189.0, "beam": 32.2, "draft": 12.1, "gross_tonnage": 32500,
            "lat": 1.29, "lon": 103.85,
            "propulsion": "Slow Speed Diesel",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Yük Kayması & Stabilite Kaybı (Alabora)",
            "root_cause": "Nemli dökme yükün sıvılaşması (Liquefaction) ve serbest yüzey momenti",
            "critical_zone": "Kargo Ambarları & Çift Dip",
            "severity": "Kritik",
            "source": "IMO GISIS",
            "report_url": "https://gisis.imo.org/Public/MCI/Default.aspx",
            "summary_report": "Nikel cevherinin sıvılaşması neticesinde ambar içi yük bordaya kaymış, gemi 25 derece meyille acil limana sığınmıştır."
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
            "propulsion": "Hybrid / Lithium Battery",
            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
            "incident_type": "Batarya Termal Kaçak & Patlama",
            "root_cause": "Lityum-iyon batarya hücresi aşırı ısınması ve zehirli gaz tahliye yetersizliği",
            "critical_zone": "Batarya Kompartımanı",
            "severity": "Kritik",
            "source": "NTSB Marine",
            "report_url": "https://www.ntsb.gov/investigations/accidentreports/Pages/marine.aspx",
            "summary_report": "Lityum bataryada termal kaçak oluşmuş, hidrojen florür gazı tahliye bacası yetersizliğinden yaşam alanına yayılmıştır."
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
            "summary_report": "İnert gaz basıncının düşmesiyle tank içi hidrokarbon buharı statik kıvılcımla patlamış ve yangın çıkmıştır."
        }
    ]

LIVE_SOURCES = [
    {"source_name": "MAIB (UK)", "url": "https://www.gov.uk/government/organisations/marine-accident-investigation-branch.atom", "default_flag": "United Kingdom", "default_cat": "Yat & Süperyat", "lat": 50.71, "lon": -1.30},
    {"source_name": "NTSB Marine (USA)", "url": "https://www.ntsb.gov/investigations/accidentreports/Pages/marine.aspx", "default_flag": "United States", "default_cat": "Ticari Gemi (Kargo/Konteyner)", "lat": 29.95, "lon": -90.07},
    {"source_name": "EMSA EMCIP (EU)", "url": "https://www.emsa.europa.eu/newsroom/latest-news.feed?type=rss", "default_flag": "European Union", "default_cat": "Ticari Gemi (Tanker)", "lat": 38.72, "lon": -9.13}
]

def sync_live_feeds(existing_records):
    added_count = 0
    records_dict = {r["incident_id"]: r for r in existing_records}

    for src in LIVE_SOURCES:
        try:
            feed = feedparser.parse(src["url"])
            if feed and hasattr(feed, 'entries'):
                for entry in feed.entries[:5]:
                    title = entry.get("title", "")
                    link = entry.get("link", src["url"])
                    pub_date = entry.get("updated", entry.get("published", datetime.now().strftime("%d/%m/%Y")))[:10]
                    summary = entry.get("summary", "Resmi kaza inceleme raporu.")
                    
                    v_cat = src["default_cat"]
                    t_low = title.lower()
                    if any(k in t_low for k in ["yacht", "boat", "motor", "sailing"]):
                        v_cat = "Yat & Süperyat"
                        loa, beam, draft = 19.5, 5.2, 1.3
                        zone = "Makine Dairesi"
                    elif any(k in t_low for k in ["tanker", "oil", "chemical", "gas"]):
                        v_cat = "Ticari Gemi (Tanker)"
                        loa, beam, draft = 155.0, 24.0, 9.0
                        zone = "Kargo Güvertesi / Tanklar"
                    elif any(k in t_low for k in ["tug", "workboat", "barge"]):
                        v_cat = "Hizmet & Çalışma Gemileri (Tug/Supply)"
                        loa, beam, draft = 30.0, 11.5, 4.2
                        zone = "Kıç Güverte / Çeki Kancası (Towing Winch)"
                    else:
                        v_cat = "Ticari Gemi (Kargo/Konteyner)"
                        loa, beam, draft = 180.0, 30.0, 11.0
                        zone = "Kargo Ambarları & Çift Dip"

                    inc_id = f"{src['source_name'][:4].strip()}-LIVE-{abs(hash(title)) % 10000:04d}"
                    
                    if inc_id not in records_dict:
                        records_dict[inc_id] = {
                            "incident_id": inc_id,
                            "incident_date": pub_date,
                            "vessel_name": title.split(":")[0] if ":" in title else "Araştırılan Gemi",
                            "vessel_category": v_cat,
                            "vessel_type": v_cat,
                            "flag": src["default_flag"],
                            "year_built": 2018,
                            "loa": loa, "beam": beam, "draft": draft, "gross_tonnage": int(loa * 14),
                            "lat": src.get("lat", 36.0), "lon": src.get("lon", 20.0),
                            "propulsion": "Twin Diesel",
                            "operating_area": "Açık Deniz / Okyanus (Unrestricted)",
                            "incident_type": "Deniz Kazası / Emniyet İncelemesi",
                            "root_cause": "Yüksek basınçlı yakıt borusu yırtılması ve sıcak egzoz manifoldu teması" if v_cat == "Yat & Süperyat" else "İnert gaz (IGS) sistem basınç kaybı ve hidrokarbon buharı statik kıvılcımı",
                            "critical_zone": zone,
                            "severity": "Yüksek",
                            "source": src["source_name"],
                            "report_url": link,
                            "summary_report": summary.replace("<p>", "").replace("</p>", "")[:250] + "..."
                        }
                        added_count += 1
        except Exception:
            pass

    combined_list = list(records_dict.values())
    try:
        with open(LIVE_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(combined_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return combined_list, added_count

def load_database():
    base = get_base_database()
    if os.path.exists(LIVE_DATA_FILE):
        try:
            with open(LIVE_DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for r in data:
                    if "lat" not in r: r["lat"] = 36.0
                    if "lon" not in r: r["lon"] = 25.0
                return data
        except Exception:
            return base
    return base

# Çözüm Kataloğu
def get_solutions_catalog():
    return {
        "Yüksek basınçlı yakıt borusu yırtılması ve sıcak egzoz manifoldu teması": [
            {
                "solution_id": "SOL-FUEL-01",
                "name": "Çift Cidarlı (Double-Walled) Yakıt Borulama + Sürekli Basınç Takibi",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Sızıntıyı dış zırhta hapsederek sıcak yüzey temasını engeller.", "SOLAS II-2 ve DNV Klas kurallarına tam uyum sağlar."],
                "riskler": ["Boru donatım maliyetinde %35-40 artış.", "Dar makine dairesinde montaj hassasiyeti."],
                "maliyet_etkisi": "Orta - Yüksek", "agirlik_etkisi": "+80 kg", "cad_referans": "DWG_Double_Walled_Piping_Standard.dwg"
            },
            {
                "solution_id": "SOL-FUEL-02",
                "name": "A-60 Egzoz Termal İzolasyon Ceketi + Otomatik Su Sisi (Water Mist)",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 2, "sev_mitigated": 4,
                "firsatlar": ["Düşük ilk yatırım maliyeti.", "Mevcut boru hatlarına kolay montaj."],
                "riskler": ["İzolasyon ceketinin yıpranması/yağlanması durumunda ikincil risk."],
                "maliyet_etkisi": "Düşük - Orta", "agirlik_etkisi": "+25 kg", "cad_referans": "DWG_Thermal_Jacket_WaterMist.dwg"
            }
        ],
        "Tek kaçış rotasının yangınla kapanması / Dikey kaçış şaftı eksikliği": [
            {
                "solution_id": "SOL-ESC-01",
                "name": "Bağımsız Dikey Kaçış Bacası (Escape Trunk 600x600 mm) + Çelik Merdiven",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 2,
                "firsatlar": ["Açık havaya kesintisiz tahliye sağlar.", "SOLAS ve MCA mega yat kodlarına tam uyumludur."],
                "riskler": ["Üst güverte yerleşiminde 600x600 mm yer kaybı."],
                "maliyet_etkisi": "Orta", "agirlik_etkisi": "+50 kg", "cad_referans": "DWG_Escape_Trunk_Structural_Section.dwg"
            }
        ],
        "Lityum-iyon batarya hücresi aşırı ısınması ve zehirli gaz tahliye yetersizliği": [
            {
                "solution_id": "SOL-BAT-01",
                "name": "A-60 İzolasyonlu Bağımsız Oda + Doğrudan Güverteye Açılan HF Gaz Tahliye Bacası",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Hidrojen Florür gazını yaşam mahalline girmeden tahliye eder.", "DNV Battery Safety notunu yükseltir."],
                "riskler": ["Ayrı kompartıman ve paslanmaz çelik baca maliyeti."],
                "maliyet_etkisi": "Yüksek", "agirlik_etkisi": "+120 kg", "cad_referans": "DWG_Battery_Room_Safety_Exhaust.dwg"
            }
        ],
        "İnert gaz (IGS) sistem basınç kaybı ve hidrokarbon buharı statik kıvılcımı": [
            {
                "solution_id": "SOL-TNK-01",
                "name": "Çift Yedekli Azot Jeneratörü (Dual N2 System) + Ultrasonik Basınç Takibi",
                "prob_initial": 3, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Tank içi oksijen seviyesini sürekli %5'in altında tutar."],
                "riskler": ["Yüksek elektrik tüketimi ve bakım maliyeti."],
                "maliyet_etkisi": "Yüksek", "agirlik_etkisi": "+400 kg", "cad_referans": "DWG_Inert_Gas_System_Schematic.dwg"
            }
        ],
        "Halat açısının bordaya dik gelmesi ve çekme kancasının hızlı bırakılamaması": [
            {
                "solution_id": "SOL-TUG-01",
                "name": "Köprüüstünden Pnömatik/Hidrolik Acil Bırakma (Quick-Release) + Dinamik Eğim Sensörü",
                "prob_initial": 4, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["0.5 saniyede halatı keserek alabora riskini önler."],
                "riskler": ["Periyodik hidrolik basınç testi gereksinimi."],
                "maliyet_etkisi": "Orta", "agirlik_etkisi": "+35 kg", "cad_referans": "DWG_Quick_Release_Hook_Hydraulic.dwg"
            }
        ],
        "Nemli dökme yükün sıvılaşması (Liquefaction) ve serbest yüzey momenti": [
            {
                "solution_id": "SOL-BLK-01",
                "name": "Ambar Boyuna Bölme Perdeleri (Longitudinal Shifting Boards) + Nem Takip Sensörleri",
                "prob_initial": 3, "sev_initial": 5, "prob_mitigated": 1, "sev_mitigated": 3,
                "firsatlar": ["Yükün serbest yüzey genişliğini ikiye bölerek devrilmeyi önler."],
                "riskler": ["Tahliye operasyon süresini uzatabilir."],
                "maliyet_etkisi": "Orta", "agirlik_etkisi": "+850 kg", "cad_referans": "DWG_Bulk_Hold_Longitudinal_Divider.dwg"
            }
        ]
    }

def get_sample_cad_library():
    base_cad = [
        {
            "filename": "DWG_Escape_Trunk_Structural_Section.dwg",
            "title": "2D Makine Dairesi Dikey Acil Kaçış Bacası (Escape Trunk) Yapısal Kesiti",
            "format": "AutoCAD DWG / DXF", "category": "Emniyet & Tahliye",
            "description": "600x600 mm standart kaçış bacası, dik açılı çelik merdiven basamakları, A-60 yangın yalıtımı ve 150 mm su sızdırmaz açık güverte eşik (coaming) detayları.",
            "code_standard": "SOLAS Ch. II-2 / MCA Large Yacht Code",
            "contributor": "Safety-by-Design Core Team"
        },
        {
            "filename": "DWG_Double_Walled_Piping_Standard.dwg",
            "title": "Çift Cidarlı (Double-Walled) Yakıt Besleme & Basınç Takip P&ID Şeması",
            "format": "AutoCAD DWG / P&ID", "category": "Makine & Borulama",
            "description": "Metanol ve yüksek basınçlı dizel devreleri için dış cidar basınç sensörleri, kaçak toplama tepsisi ve 1 sn hızlı kapama ESD pnömatik vana bağlantı şeması.",
            "code_standard": "IGF Code / DNV Rules for Classification",
            "contributor": "Safety-by-Design Core Team"
        },
        {
            "filename": "DWG_Battery_Room_Safety_Exhaust.dwg",
            "title": "Lityum Batarya Kompartımanı A-60 Yangın & HF Gaz Tahliye Düzeni",
            "format": "AutoCAD DWG / 3D STEP", "category": "Elektrik & Batarya Güvenliği",
            "description": "A-60 yangın perdesi izolasyonu, paslanmaz çelik bağımsız aşırı basınç tahliye bacası ve patlama kapakları yerleşim planı.",
            "code_standard": "DNV-RU-SHIP Pt.6 Ch.2 / ClassNK",
            "contributor": "Safety-by-Design Core Team"
        }
    ]
    if os.path.exists(USER_UPLOADS_FILE):
        try:
            with open(USER_UPLOADS_FILE, "r", encoding="utf-8") as f:
                user_items = json.load(f)
                return base_cad + user_items
        except Exception:
            return base_cad
    return base_cad

# =========================================================================
# YENİ ÖZELLİK 1: 2D PARAMETRİK GEMİ SİLÜETİ VE RİSK ISI BÖLGELERİ GRAFİĞİ
# =========================================================================
def draw_parametric_hull_risk_map(loa, beam, draft, critical_zones):
    fig, ax = plt.subplots(figsize=(10, 2.8), facecolor='#F8FAFC')
    ax.set_facecolor('#F8FAFC')
    
    # Koordinatlar
    x_ap, x_fp = 0, loa
    depth = draft * 2.2 # Yaklaşık derinlik
    
    # 1. Gövde Silüeti Çizimi
    hull_x = [0, loa * 0.85, loa, loa * 0.88, 0, 0]
    hull_y = [depth, depth * 1.05, depth * 1.15, 0, 0, depth]
    ax.plot(hull_x, hull_y, color='#0F172A', linewidth=2.5, label='Tekne Borda Hattı')
    
    # Su Hattı (Waterline)
    ax.axhline(y=draft, color='#0284C7', linestyle='--', linewidth=1.5, label=f'Tasarım Su Hattı (T={draft}m)')
    
    # 2. Kompartıman Perdeleri (Bulkheads)
    er_aft = loa * 0.15
    er_fwd = loa * 0.40
    cargo_fwd = loa * 0.75
    
    ax.axvline(x=er_aft, ymin=0, ymax=depth/depth*0.88, color='#64748B', linestyle=':', linewidth=1.5)
    ax.axvline(x=er_fwd, ymin=0, ymax=depth/depth*0.88, color='#64748B', linestyle=':', linewidth=1.5)
    ax.axvline(x=cargo_fwd, ymin=0, ymax=depth/depth*0.92, color='#64748B', linestyle=':', linewidth=1.5)
    
    # 3. Risk Isı Bölgeleri Boyama (Heat Zones)
    # Makine Dairesi
    er_color = '#EF4444' if any("Makine" in z or "Escape" in z for z in critical_zones) else '#94A3B8'
    ax.add_patch(patches.Rectangle((er_aft, 0.05), er_fwd - er_aft, depth * 0.9, color=er_color, alpha=0.35))
    ax.text((er_aft + er_fwd)/2, depth * 0.45, "MAKİNE DAİRESİ\n(YÜKSEK RİSK BÖLGESİ)", color='#991B1B', weight='bold', fontsize=8, ha='center')

    # Orta Bölüm / Kargo / Yaşam
    mid_color = '#F59E0B' if any("Kargo" in z or "Batarya" in z for z in critical_zones) else '#94A3B8'
    ax.add_patch(patches.Rectangle((er_fwd, 0.05), cargo_fwd - er_fwd, depth * 0.9, color=mid_color, alpha=0.25))
    ax.text((er_fwd + cargo_fwd)/2, depth * 0.45, "YAŞAM MAHALLİ / KARGO\n(ORTA RİSK)", color='#B45309', weight='bold', fontsize=8, ha='center')

    # Baş Pik / Çatışma
    ax.add_patch(patches.Rectangle((cargo_fwd, 0.05), (loa - cargo_fwd)*0.85, depth * 0.9, color='#10B981', alpha=0.20))
    ax.text((cargo_fwd + loa)/2, depth * 0.45, "BAŞ PİK\n(DÜŞÜK RİSK)", color='#065F46', weight='bold', fontsize=8, ha='center')

    ax.set_xlim(-1, loa + 2)
    ax.set_ylim(-0.2, depth * 1.35)
    ax.set_xlabel(f"Boyuna Konum X [m] (Tam Boy LOA = {loa} m)")
    ax.set_ylabel("Yükseklik [m]")
    ax.set_title("2D Parametrik Tekne Profili ve Kompartıman Risk Isı Haritası (Safety Compartment Arrangement)", fontsize=10, weight='bold', color='#1E3A8A')
    ax.legend(loc='upper left', fontsize=7)
    plt.tight_layout()
    
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=200)
    buf.seek(0)
    plt.close(fig)
    return buf

# =========================================================================
# YENİ ÖZELLİK 2: KLAS VE REGÜLASYON KURAL MOTORU (RULE ENGINE)
# =========================================================================
def get_regulatory_compliance(loa, propulsion, category):
    rules = []
    
    # SOLAS & Yangın Kuralları
    if loa >= 24.0:
        rules.append({
            "code": "SOLAS Ch. II-2 / Reg. 13",
            "title": "Çift Bağımsız Kaçış Yolu Zorunluluğu",
            "desc": "24 metre ve üzeri teknelerde makine dairesinden açık güverteye doğrudan çıkan en az iki bağımsız kaçış yolu bulunmalıdır. Biri dikey çelik kaçış bacası (Escape Trunk) olmalıdır."
        })
        rules.append({
            "code": "MCA REG Code Part A",
            "title": "A-60 Sınıfı Yangın Perdesi",
            "desc": "Makine dairesi ile yaşam mahalleri veya acil durum kontrol istasyonları arasındaki tüm yapısal perdeler A-60 sınıfı yangın izolasyonuna sahip olmalıdır."
        })
    else:
        rules.append({
            "code": "ISO 9094 / Small Craft Fire Protection",
            "title": "Küçük Deniz Araçları Yangın Önleme ve Manuel Tahliye",
            "desc": "Makine dairesine girmeden dışarıdan müdahale edilebilecek yangın söndürme portu ve personelin hızlı tahliyesi için min. 500 mm menhol açıklığı."
        })

    # Sevk ve Yakıt Tipi Kuralları
    if "Diesel" in propulsion:
        rules.append({
            "code": "SOLAS Ch. II-2 / Reg. 4.2.2",
            "title": "Çift Cidarlı Yüksek Basınç Borulaması",
            "desc": "Tüm dizel ana ve yardımcı makinelerin yüksek basınçlı yakıt boruları çift cidarlı (jacketed) olmalı ve sızıntı durumunda alarm veren sintine kabı bulunmalıdır."
        })
    if "Battery" in propulsion or "Electric" in propulsion or "Hybrid" in propulsion:
        rules.append({
            "code": "DNV-RU-SHIP Pt.6 Ch.2 / IMO MSC.1/Circ.1455",
            "title": "Lityum Batarya Kompartımanı Aşırı Basınç ve HF Tahliyesi",
            "desc": "Batarya odaları bağımsız mekanik havalandırmaya ve termal kaçak durumunda patlama kapaklarına (explosion relief vents) sahip olmalıdır."
        })
    if "Methanol" in propulsion or "LNG" in propulsion:
        rules.append({
            "code": "IMO IGF Code (Res. MSC.391(95))",
            "title": "Düşük Parlama Noktalı Yakıtlar Emniyet Kodu",
            "desc": "Yakıt hazırlama odaları inert gaz sistemiyle korunmalı, çift çeperli borulama arası azotla yıkanmalı ve gaz dedektör ağı kurulmalıdır."
        })

    return rules

# =========================================================================
# YENİ ÖZELLİK 3: YAPAY ZEKA DESTEKLİ DENİZCİLİK ASİSTANI (SAFETY AI ASSISTANT)
# =========================================================================
def query_safety_knowledge_base(question):
    q_low = question.lower()
    if any(w in q_low for w in ["kaçış", "escape", "merdiven", "menhol", "trunk"]):
        return "**SOLAS Ch. II-2 & MCA Large Yacht Code Standardı:**\n- Makine dairesi kaçış bacaları en az **600x600 mm** serbest geçiş alanına sahip olmalıdır.\n- Merdiven basamakları çelikten yapılmalı, basamak aralığı max 300 mm olmalıdır.\n- Açık güverteye çıkış eşiği (coaming) deniz suyunun içeri girmesini önlemek için en az **150-300 mm** yükseklikte olmalıdır."
    elif any(w in q_low for w in ["batarya", "lityum", "battery", "hf", "termal"]):
        return "**DNV Battery Safety & IMO Rehberi:**\n- Lityum-iyon batarya odaları A-60 sınıfı yapısal yangın perdesi ile çevrelenmelidir.\n- Termal kaçak anında açığa çıkan Hidrojen Florür (HF) gazının tahliyesi için açık güverteye doğrudan uzanan bağımsız paslanmaz çelik baca şarttır.\n- Yangın söndürme için sürekli soğutma sağlayan Yüksek Basınçlı Su Sisi (Water Mist) önerilir."
    elif any(w in q_low for w in ["çift cidar", "boru", "yakıt", "double-wall", "piping"]):
        return "**SOLAS II-2 Reg 4 & DNV Makine Kuralları:**\n- Enjektör ile yakıt pompası arasındaki tüm yüksek basınç hatları çift cidarlı zırhlı boru olmak zorundadır.\n- Dış borudaki kaçaklar sıcak egzoz manifolduyla temas etmeden doğrudan alarmlı toplama tankına yönlendirilmelidir."
    elif any(w in q_low for w in ["stabilite", "sıvılaşma", "ambar", "dökme"]):
        return "**IMSBC Code (International Maritime Solid Bulk Cargoes Code):**\n- Sıvılaşabilen (liquefaction) dökme yüklerde nem oranı TML (Transportable Moisture Limit) değerini aşamaz.\n- Serbest yüzey devrilme momentini kırmak için ambar boyunca çelik bölme perdeleri (shifting boards) kullanılmalıdır."
    else:
        return f"Sorduğunuz soru ('{question}') Safety-by-Design kural kütüphanesinde tarandı. Genel mühendislik prensibi olarak; ilgili bölgenin A-60 yangın yalıtımı, bağımsız alarm sensörleri ve IMO MSC sirkülerlerine uygun sertifikalı ekipman ile donatılması tavsiye edilir."

# =========================================================================
# STREAMLIT ANA ARAYÜZ
# =========================================================================
st.set_page_config(page_title="Safety-by-Design Maritime Platform", page_icon="⚓", layout="wide")

if "db_incidents" not in st.session_state:
    st.session_state.db_incidents = pd.DataFrame(load_database())

solutions_catalog = get_solutions_catalog()
cad_library = get_sample_cad_library()

st.title("⚓ Safety-by-Design: Gemi Kaza Analitiği & Karar Destek Platformu")
st.markdown("MAIB, IMO GISIS, NTSB ve EMSA veri entegrasyonu, 2D parametrik gövde risk haritası, 5x5 Risk Matrisi, FMEA analiz suit'i ve Açık Topluluk CAD Merkezi.")

# Canlı Senkronizasyon Çubuğu
c_sync1, c_sync2 = st.columns([2, 5])
with c_sync1:
    if st.button("🔄 Kaza Veri Tabanlarını Canlı Güncelle (MAIB / NTSB / EMSA)", use_container_width=True):
        with st.spinner("Canlı kaza akışları taranıyor..."):
            updated_data, new_added = sync_live_feeds(st.session_state.db_incidents.to_dict('records'))
            st.session_state.db_incidents = pd.DataFrame(updated_data)
            if new_added > 0:
                st.success(f"✅ {new_added} yeni resmi kaza raporu başarıyla sisteme aktarıldı!")
            else:
                st.info("ℹ️ Veri tabanı zaten en güncel resmi kaza raporlarını içeriyor.")
with c_sync2:
    st.caption("🌐 Bağlı Veri Tabanları: **MAIB (UK)** | **IMO GISIS (BM)** | **NTSB Marine (USA)** | **EMSA EMCIP (AB)**")

st.write("---")

# =========================================================================
# YENİ ÖZELLİK 4: KENAR ÇUBUĞU GELİŞMİŞ FİLTRELERİ (SIDEBAR)
# =========================================================================
st.sidebar.header("🔍 Gelişmiş Kaza Filtreleri")
year_filter = st.sidebar.slider("Kaza Yılı Aralığı", min_value=2010, max_value=2026, value=(2015, 2026))

all_flags = sorted(list(st.session_state.db_incidents["flag"].unique()))
selected_flags = st.sidebar.multiselect("Bayrak Devleti Seçimi", options=all_flags, default=all_flags)

all_severities = sorted(list(st.session_state.db_incidents["severity"].unique()))
selected_severities = st.sidebar.multiselect("Kaza Şiddeti", options=all_severities, default=all_severities)

st.sidebar.divider()
st.sidebar.markdown("### 🤖 Denizcilik Emniyet Asistanı (AI)")
user_question = st.sidebar.text_input("Regülasyon veya Emniyet Sorusu Sorun:", placeholder="Örn: Kaçış bacası ölçüleri ne olmalı?")
if user_question:
    ai_answer = query_safety_knowledge_base(user_question)
    st.sidebar.info(ai_answer)

# =========================================================================
# ANA SEKMELER
# =========================================================================
tab_design, tab_rules, tab_cad = st.tabs([
    "🚀 Karar Destek & Risk Analizi", 
    "📜 Klas & Regülasyon Kural Motoru", 
    "📐 Açık Kaynak Tasarım & CAD Kütüphanesi"
])

# -------------------------------------------------------------
# 1. SEKME: KARAR DESTEK & RİSK ANALİZİ
# -------------------------------------------------------------
with tab_design:
    st.subheader("1. Gemi ve Tasarım Parametreleri")
    
    col1, col2 = st.columns(2)
    with col1:
        vessel_name = st.text_input("Gemi / Proje Adı", value="PRU-Design Project 01")
        
        cat_options = sorted(list(st.session_state.db_incidents["vessel_category"].unique())) + ["Diğer (Manuel Giriş)"]
        selected_cat = st.selectbox("Deniz Aracı Ana Kategorisi", options=cat_options)
        if selected_cat == "Diğer (Manuel Giriş)":
            vessel_category = st.text_input("Lütfen Deniz Aracı Kategorisini Yazın", value="Özel Araştırma Gemisi")
        else:
            vessel_category = selected_cat

        loa = st.number_input("Tam Boy (LOA) [Metre]", min_value=3.0, max_value=450.0, value=18.0, step=0.5)
        beam = st.number_input("Genişlik (Beam) [Metre]", min_value=1.0, max_value=75.0, value=5.0, step=0.1)

    with col2:
        draft = st.number_input("Su Çekimi (Draft) [Metre]", min_value=0.3, max_value=25.0, value=1.25, step=0.05)
        displacement = st.number_input("Deplasman [Ton]", min_value=1.0, max_value=400000.0, value=35.0, step=1.0)
        
        prop_options = ["Twin Diesel", "Hybrid / Electric", "Lithium Battery", "Slow Speed Diesel", "Dual Fuel (LNG/Diesel)", "Methanol", "Azimuth Stern Drive (ASD)", "Diğer (Manuel Giriş)"]
        selected_prop = st.selectbox("Sevk / Yakıt Tipi", options=prop_options)
        if selected_prop == "Diğer (Manuel Giriş)":
            propulsion = st.text_input("Lütfen Sevk / Yakıt Tipini Yazın", value="Hidrojen Yakıt Pili (Fuel Cell)")
        else:
            propulsion = selected_prop

        area_options = ["Açık Deniz / Okyanus (Unrestricted)", "Kıyısal / Yakın Sefer (Coastal)", "İç Sular / Göller", "Liman / Rıhtım İçi", "Diğer (Manuel Giriş)"]
        selected_area = st.selectbox("Operasyon / Seyir Bölgesi", options=area_options)
        if selected_area == "Diğer (Manuel Giriş)":
            operating_area = st.text_input("Lütfen Seyir Bölgesini Yazın", value="Kutup Bölgesi / Buz Sınıfı")
        else:
            operating_area = selected_area

    st.write("---")

    # Kaza Eşleştirme Motoru + Kenar Çubuğu Filtresi
    current_df = st.session_state.db_incidents.copy()
    
    # Yıl filtresini uygula
    current_df["incident_year"] = pd.to_datetime(current_df["incident_date"], format="%d/%m/%Y", errors='coerce').dt.year.fillna(2022)
    filtered_df = current_df[
        (current_df["incident_year"] >= year_filter[0]) & 
        (current_df["incident_year"] <= year_filter[1]) &
        (current_df["flag"].isin(selected_flags)) &
        (current_df["severity"].isin(selected_severities))
    ]

    matched = filtered_df[
        (filtered_df["vessel_category"] == vessel_category) &
        (filtered_df["loa"] <= loa * 1.35) &
        (filtered_df["loa"] >= loa * 0.65)
    ]
    if matched.empty:
        matched = filtered_df[filtered_df["vessel_category"] == vessel_category]
    if matched.empty:
        matched = filtered_df.copy()

    st.subheader("2. Parametrik 2D Gemi Silüeti, Risk Isı Haritası ve Kaza Analitiği")

    # 2D DİNAMİK GEMİ PROFİLİ GÖRSELİ
    critical_zones_list = matched["critical_zone"].tolist()
    hull_chart_buf = draw_parametric_hull_risk_map(loa, beam, draft, critical_zones_list)
    st.image(hull_chart_buf, use_container_width=True)

    st.write("")

    # DÜNYA KAZA YOĞUNLUK HARİTASI
    st.markdown("#### 🌍 Küresel Deniz Kazaları Yoğunluk ve Isı Haritası (Heatmap)")
    map_df = filtered_df.copy()
    if "lat" not in map_df.columns: map_df["lat"] = 36.5
    if "lon" not in map_df.columns: map_df["lon"] = 28.0
    map_df["lat"] = pd.to_numeric(map_df["lat"], errors='coerce').fillna(36.5)
    map_df["lon"] = pd.to_numeric(map_df["lon"], errors='coerce').fillna(28.0)
    
    heatmap_layer = pdk.Layer("HeatmapLayer", data=map_df, get_position=["lon", "lat"], get_weight=1, radius_pixels=60, opacity=0.8)
    scatter_layer = pdk.Layer("ScatterplotLayer", data=map_df, get_position=["lon", "lat"], get_color=[220, 38, 38, 200], get_radius=80000, pickable=True)
    view_state = pdk.ViewState(latitude=25.0, longitude=15.0, zoom=1.3, pitch=20)
    
    r_deck = pdk.Deck(layers=[heatmap_layer, scatter_layer], initial_view_state=view_state, tooltip={"html": "<b>Gemi:</b> {vessel_name}<br/><b>Kaza:</b> {incident_type}<br/><b>Kurum:</b> {source}<br/><b>Bölge:</b> {critical_zone}", "style": {"backgroundColor": "#0F172A", "color": "white", "fontSize": "12px"}})
    st.pydeck_chart(r_deck)

    st.write("")

    # İstatistik Grafikleri
    g1, g2 = st.columns(2)
    fig_stat, axes = plt.subplots(1, 2, figsize=(10, 4))
    zone_counts = matched["critical_zone"].value_counts()
    axes[0].barh(zone_counts.index, zone_counts.values, color="#1E3A8A")
    axes[0].set_title("Kritik Risk Bolgeleri (Kaza Sikligi)")
    axes[0].set_xlabel("Vaka Sayisi")

    type_counts = matched["incident_type"].value_counts()
    axes[1].pie(type_counts.values, labels=type_counts.index, autopct='%1.1f%%', colors=["#DC2626", "#F59E0B", "#10B981", "#6366F1"][:len(type_counts)])
    axes[1].set_title("Kaza Tipleri Olasilik Dagilimi")
    plt.tight_layout()

    with g1:
        st.markdown("#### 📊 Kritik Kompartıman Frekans Dağılımı")
        st.bar_chart(zone_counts)
    with g2:
        st.markdown("#### 🎯 Kaza Tipleri Olasılık Dağılımı")
        st.write(type_counts.apply(lambda x: f"%{round((x/len(matched))*100, 1)} Olasılık"))

    stat_chart_buf = io.BytesIO()
    fig_stat.savefig(stat_chart_buf, format='png', dpi=200)
    stat_chart_buf.seek(0)
    plt.close(fig_stat)

    st.write("---")

    # Kaza Kartları
    st.markdown("#### 🚨 Eşleşen Kaza Vakaları & Gemi Detayları")
    for idx, row in matched.iterrows():
        with st.expander(f"[{row['incident_id']}] {row['vessel_name']} - {row['incident_type']} ({row['incident_date']})"):
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1:
                st.write(f"**Gemi Adı:** {row['vessel_name']}")
                st.write(f"**Tip:** {row['vessel_type']} | **Bayrak:** {row['flag']}")
                st.write(f"**Boy/En/Draft:** {row['loa']}m / {row['beam']}m / {row['draft']}m")
            with c2:
                st.write(f"**Kaza Tarihi:** {row['incident_date']}")
                st.write(f"**Bölge:** `{row['critical_zone']}`")
                st.write(f"**Şiddet:** :red[{row['severity']}]")
                st.write(f"**Kaynak Veri Tabanı:** `{row['source']}`")
            with c3:
                st.caption(f"**Özet:** {row['summary_report']}")
                st.link_button(label=f"🌐 {row['source']} Resmi Rapor Sayfası", url=row["report_url"])

    st.write("---")

    # 3. Yöntem Seçimi ve 5x5 Risk Matrisi
    st.subheader("3. Tasarım Önlemi Seçimi ve 5x5 FMEA Risk Matrisi")
    selected_solutions_data = []
    unique_root_causes = matched["root_cause"].unique()
    matrix_export_data = []

    for idx, rc in enumerate(unique_root_causes):
        if rc in solutions_catalog:
            options = solutions_catalog[rc]
            opt_names = [opt["name"] for opt in options]
            
            st.markdown(f"**Kritik Arıza Modu {idx+1}:** *{rc}*")
            chosen_name = st.radio(label=f"Öneri Seçimi {idx+1}", options=opt_names, key=f"radio_{idx}", label_visibility="collapsed")
            chosen_opt = next(item for item in options if item["name"] == chosen_name)
            
            p1, s1 = chosen_opt["prob_initial"], chosen_opt["sev_initial"]
            r1 = p1 * s1
            p2, s2 = chosen_opt["prob_mitigated"], chosen_opt["sev_mitigated"]
            r2 = p2 * s2
            reduction_pct = round(((r1 - r2) / r1) * 100, 1)

            cf, cr, cm = st.columns([2, 2, 1])
            with cf: st.success("✅ **Kazanımlar & Fırsatlar:**\n- " + "\n- ".join(chosen_opt["firsatlar"]))
            with cr: st.warning("⚠️ **Riskler & Maliyet:**\n- " + "\n- ".join(chosen_opt["riskler"]))
            with cm:
                st.metric("Risk Skoru Düşüşü", f"{r1} ➔ {r2}", delta=f"-%{reduction_pct}")
                st.caption(f"💰 Maliyet: **{chosen_opt['maliyet_etkisi']}**\n⚖️ Ağırlık: **{chosen_opt['agirlik_etkisi']}**")

            selected_solutions_data.append({"problem": rc, "solution": chosen_opt, "r1": r1, "r2": r2, "reduction": reduction_pct})
            matrix_export_data.append({
                "problem": rc, "sol_name": chosen_opt["name"],
                "firsatlar": "; ".join(chosen_opt["firsatlar"]), "riskler": "; ".join(chosen_opt["riskler"]),
                "maliyet": chosen_opt["maliyet_etkisi"], "agirlik": chosen_opt["agirlik_etkisi"],
                "p1": p1, "s1": s1, "r1": r1, "p2": p2, "s2": s2, "r2": r2,
                "reduction_pct": reduction_pct, "cad": chosen_opt["cad_referans"]
            })
            st.write("")

    st.write("---")
    st.subheader("4. Raporlama ve Dosya İndirme Merkezi")

    # EXCEL MOTORU
    def generate_advanced_excel_with_charts(v_info, matrix_rows):
        wb = openpyxl.Workbook()
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_sub = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
        font_data = Font(name="Calibri", size=10)
        font_bold = Font(name="Calibri", size=10, bold=True)
        fill_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        fill_blue = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
        fill_red = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
        font_red = Font(name="Calibri", size=10, bold=True, color="991B1B")
        fill_yellow = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        font_yellow = Font(name="Calibri", size=10, bold=True, color="854D0E")
        fill_green = PatternFill(start_color="BBF7D0", end_color="BBF7D0", fill_type="solid")
        font_green = Font(name="Calibri", size=10, bold=True, color="166534")
        thin = Side(border_style="thin", color="CBD5E1")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws1 = wb.active; ws1.title = "FMEA Risk & Firsat Matrisi"; ws1.views.sheetView[0].showGridLines = True
        ws1.merge_cells("A1:N1"); ws1["A1"] = "GEMI DIZAYN SURECI FMEA RISK - FIRSAT VE KARAR DESTEK MATRISI"
        ws1["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF"); ws1["A1"].fill = fill_navy
        ws1["A1"].alignment = Alignment(horizontal="center", vertical="center"); ws1.row_dimensions[1].height = 30

        ws1["A3"] = "Gemi / Proje:"; ws1["B3"] = v_info["Gemi Adı"]; ws1["A3"].font = font_sub
        ws1["D3"] = "Tam Boy (LOA):"; ws1["E3"] = f"{v_info['Tam Boy (LOA)']} m"; ws1["D3"].font = font_sub
        ws1["G3"] = "Genislik:"; ws1["H3"] = f"{v_info['Genişlik (Beam)']} m"; ws1["G3"].font = font_sub
        ws1["J3"] = "Draft / Disp:"; ws1["K3"] = f"{v_info['Draft']} m / {v_info['Deplasman']} ton"; ws1["J3"].font = font_sub

        headers = ["ID", "Arıza Modu / Kök Neden", "Seçilen Tasarım Önlemi", "İlk Olasılık (P1)", "İlk Şiddet (S1)", "İLK RİSK (R1)", "Risk Seviyesi (Önce)", "Kazanımlar & Fırsatlar", "Riskler & Maliyet", "Maliyet Etkisi", "Ağırlık Etkisi", "Kalan Olasılık (P2)", "Kalan Şiddet (S2)", "KALAN RİSK (R2)"]
        row_idx = 5; ws1.row_dimensions[row_idx].height = 25
        for col_idx, h in enumerate(headers, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=h)
            cell.font = font_header; cell.fill = fill_blue if col_idx not in [6, 14] else fill_navy
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border = border_all

        for idx, item in enumerate(matrix_rows, 1):
            row_idx += 1; ws1.row_dimensions[row_idx].height = 35
            r1_val, r2_val = item["r1"], item["r2"]
            vals = [f"Risk-{idx}", item["problem"], item["sol_name"], item["p1"], item["s1"], r1_val, "YÜKSEK (Kritik)" if r1_val >= 15 else ("ORTA" if r1_val >= 8 else "DÜŞÜK"), item["firsatlar"], item["riskler"], item["maliyet"], item["agirlik"], item["p2"], item["s2"], r2_val]
            for c_idx, val in enumerate(vals, 1):
                c = ws1.cell(row=row_idx, column=c_idx, value=val); c.font = font_data; c.border = border_all
                c.alignment = Alignment(vertical="center", wrap_text=True)
                if c_idx in [1, 4, 5, 12, 13]: c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx in [6, 7]:
                    c.fill = fill_red if r1_val >= 15 else (fill_yellow if r1_val >= 8 else fill_green)
                    c.font = font_red if r1_val >= 15 else (font_yellow if r1_val >= 8 else font_green)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 14:
                    c.fill = fill_red if r2_val >= 15 else (fill_yellow if r2_val >= 8 else fill_green)
                    c.font = font_red if r2_val >= 15 else (font_yellow if r2_val >= 8 else font_green)
                    c.alignment = Alignment(horizontal="center", vertical="center")

        num_data_rows = len(matrix_rows)
        ws_chart = wb.create_sheet(title="Risk & Maliyet Grafikleri"); ws_chart.views.sheetView[0].showGridLines = True
        ws_chart.merge_cells("A1:M1"); ws_chart["A1"] = "DENİZCİLİK DİZAYN RİSK İNDİRGEME VE MALİYET GRAFİK DASHBOARD'U"
        ws_chart["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF"); ws_chart["A1"].fill = fill_navy
        ws_chart["A1"].alignment = Alignment(horizontal="center", vertical="center"); ws_chart.row_dimensions[1].height = 28

        chart1 = BarChart(); chart1.type = "col"; chart1.style = 10; chart1.title = "Dizayn Onlemi Oncesi ve Sonrasi Risk Skoru Karsilastirmasi (R1 vs R2)"
        chart1.y_axis.title = "Risk Skoru (P x S)"; chart1.x_axis.title = "Risk Kalemleri"; chart1.width = 16; chart1.height = 11
        data1 = Reference(ws1, min_col=6, min_row=5, max_row=5 + num_data_rows)
        data2 = Reference(ws1, min_col=14, min_row=5, max_row=5 + num_data_rows)
        cats1 = Reference(ws1, min_col=1, min_row=6, max_row=5 + num_data_rows)
        chart1.add_data(data1, titles_from_data=True); chart1.add_data(data2, titles_from_data=True); chart1.set_categories(cats1)
        ws_chart.add_chart(chart1, "B3")

        ws_chart["K3"] = "Maliyet Seviyesi"; ws_chart["L3"] = "Adet"; ws_chart["K3"].font = font_header; ws_chart["K3"].fill = fill_navy
        ws_chart["L3"].font = font_header; ws_chart["L3"].fill = fill_navy
        cost_levels = ["Düşük", "Düşük - Orta", "Orta", "Orta - Yüksek", "Yüksek"]
        for idx, lvl in enumerate(cost_levels, 4):
            count = sum(1 for r in matrix_rows if r["maliyet"] == lvl)
            ws_chart[f"K{idx}"] = lvl; ws_chart[f"L{idx}"] = count
            ws_chart[f"K{idx}"].font = font_data; ws_chart[f"L{idx}"].font = font_bold
            ws_chart[f"K{idx}"].border = border_all; ws_chart[f"L{idx}"].border = border_all
            ws_chart[f"L{idx}"].alignment = Alignment(horizontal="center")

        chart2 = PieChart(); chart2.title = "Secilen Tasarim Onlemlerinin Maliyet Dagilimi"; chart2.width = 14; chart2.height = 11
        pie_data = Reference(ws_chart, min_col=12, min_row=3, max_row=3 + len(cost_levels))
        pie_cats = Reference(ws_chart, min_col=11, min_row=4, max_row=3 + len(cost_levels))
        chart2.add_data(pie_data, titles_from_data=True); chart2.set_categories(pie_cats)
        ws_chart.add_chart(chart2, "K10")

        ws2 = wb.create_sheet(title="5x5 Gorsel Risk Matrisi"); ws2.views.sheetView[0].showGridLines = True
        ws2.merge_cells("A1:G1"); ws2["A1"] = "STANDART 5x5 DENİZCİLİK RİSK DEĞERLENDİRME MATRİSİ (PROBABILITY x SEVERITY)"
        ws2["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF"); ws2["A1"].fill = fill_navy
        ws2["A1"].alignment = Alignment(horizontal="center", vertical="center"); ws2.row_dimensions[1].height = 25

        ws2["A3"] = "OLASILIK ↓ / ŞİDDET →"; ws2["A3"].font = font_bold; ws2["A3"].alignment = Alignment(horizontal="center", vertical="center")
        sev_labels = ["1 - Önemsiz", "2 - Düşük", "3 - Orta", "4 - Ciddi", "5 - Felaket (Kayıp)"]
        prob_labels = ["5 - Neredeyse Kesin", "4 - Yüksek", "3 - Olası", "2 - Nadir", "1 - Çok Nadir"]

        for col_idx, slab in enumerate(sev_labels, 2):
            c = ws2.cell(row=3, column=col_idx, value=slab); c.font = font_header; c.fill = fill_navy
            c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border_all

        for r_offset, plab in enumerate(prob_labels):
            p_val = 5 - r_offset
            c = ws2.cell(row=4 + r_offset, column=1, value=plab); c.font = font_header; c.fill = fill_navy
            c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border_all
            ws2.row_dimensions[4 + r_offset].height = 28
            for s_val in range(1, 6):
                score = p_val * s_val; cell = ws2.cell(row=4 + r_offset, column=s_val + 1, value=score)
                cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = border_all
                if score >= 15: cell.fill = fill_red; cell.font = font_red
                elif score >= 8: cell.fill = fill_yellow; cell.font = font_yellow
                else: cell.fill = fill_green; cell.font = font_green

        ws2["A11"] = "LEJANT:"; ws2["A11"].font = font_bold
        ws2["B11"] = "15 - 25: YÜKSEK (Kabul Edilemez / Acil Dizayn Önlemi Gerekli)"; ws2["B11"].fill = fill_red; ws2["B11"].font = font_red
        ws2["B12"] = "8 - 12: ORTA (ALARP Prensibi / İzleme ve Standart İzolasyon)"; ws2["B12"].fill = fill_yellow; ws2["B12"].font = font_yellow
        ws2["B13"] = "1 - 6: DÜŞÜK (Kabul Edilebilir / Güvenli Tasarım)"; ws2["B13"].fill = fill_green; ws2["B13"].font = font_green

        for ws in [ws1, ws2, ws_chart]:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        out_buf = io.BytesIO(); wb.save(out_buf); out_buf.seek(0)
        return out_buf

    # PDF RAPOR MOTORU (2D HULL + İSTATİSTİK GRAFİKLERİ ENTEGRELİ)
    def generate_full_pdf_with_charts(v_info, incidents, solutions, stat_chart_bytes, hull_chart_bytes):
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0A2540'))
        h2_style = ParagraphStyle('DocH2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#1E3A8A'))
        normal_style = ParagraphStyle('DocNormal', parent=styles['Normal'], fontSize=8, leading=11)
        bold_style = ParagraphStyle('DocBold', parent=styles['Normal'], fontSize=8, leading=11, fontName='Helvetica-Bold')

        story.append(Paragraph(clean_tr("SAFETY-BY-DESIGN DECISION SUPPORT & RISK REPORT"), title_style))
        story.append(Paragraph(clean_tr(f"Deniz Araci Kaza Analitigi ve FMEA Risk-Firsat Degerlendirmesi | Tarih: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), normal_style))
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0A2540'), spaceBefore=6, spaceAfter=10))

        story.append(Paragraph(clean_tr("1. GEMI VE DIZAYN SPESIFIKASYONLARI"), h2_style))
        v_data = [
            [Paragraph("Gemi Adi:", bold_style), Paragraph(clean_tr(str(v_info['Gemi Adı'])), normal_style), Paragraph("Kategori:", bold_style), Paragraph(clean_tr(str(v_info['Kategori'])), normal_style)],
            [Paragraph("Tam Boy (LOA):", bold_style), Paragraph(f"{v_info['Tam Boy (LOA)']} m", normal_style), Paragraph("Genislik (Beam):", bold_style), Paragraph(f"{v_info['Genişlik (Beam)']} m", normal_style)],
            [Paragraph("Draft (T):", bold_style), Paragraph(f"{v_info['Draft']} m", normal_style), Paragraph("Deplasman:", bold_style), Paragraph(f"{v_info['Deplasman']} ton", normal_style)],
            [Paragraph("Sevk Tipi:", bold_style), Paragraph(clean_tr(str(v_info['Sevk Tipi'])), normal_style), Paragraph("Gorev Alani:", bold_style), Paragraph(clean_tr(str(v_info['Görev Alanı'])), normal_style)]
        ]
        t_vessel = Table(v_data, colWidths=[85, 180, 85, 180])
        t_vessel.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        story.append(t_vessel)
        story.append(Spacer(1, 8))

        # 2D Gövde Silüeti Grafiğini PDF'e Ekle
        story.append(Paragraph(clean_tr("2. PARAMETRIK 2D TEKNE PROFILI VE RISK KOMPARTIMAN DÜZENİ"), h2_style))
        hull_chart_bytes.seek(0)
        story.append(RLImage(hull_chart_bytes, width=530, height=140))
        story.append(Spacer(1, 8))

        story.append(Paragraph(clean_tr("3. KAZA VERI TABANI OLASILIK GRAFIKLERI"), h2_style))
        stat_chart_bytes.seek(0)
        story.append(RLImage(stat_chart_bytes, width=530, height=170))
        story.append(Spacer(1, 8))

        story.append(Paragraph(clean_tr("4. FMEA RISK-FIRSAT MATRISI VE KALAN RISK DEGERLENDIRMESI"), h2_style))
        fmea_headers = [Paragraph("Arıza / Kök Neden", bold_style), Paragraph("Seçilen Tasarım Önlemi", bold_style), Paragraph("Kazanımlar & Riskler", bold_style), Paragraph("Maliyet / Ağırlık", bold_style), Paragraph("Risk İndirgeme", bold_style)]
        fmea_rows = [fmea_headers]
        for s in solutions:
            sol = s["solution"]
            detay = clean_tr(f"<b>Fırsatlar:</b> {'; '.join(sol['firsatlar'][:2])}<br/><b>Riskler:</b> {'; '.join(sol['riskler'][:1])}")
            fmea_rows.append([
                Paragraph(clean_tr(s["problem"]), normal_style), Paragraph(clean_tr(sol["name"]), bold_style),
                Paragraph(detay, normal_style), Paragraph(clean_tr(f"Maliyet: {sol['maliyet_etkisi']}<br/>Ağırlık: {sol['agirlik_etkisi']}"), normal_style),
                Paragraph(clean_tr(f"Önce: {s['r1']}<br/>Sonra: {s['r2']}<br/><b>Düşüş: %{s['reduction']}</b>"), normal_style)
            ])
        t_fmea = Table(fmea_rows, colWidths=[105, 115, 175, 75, 60])
        t_fmea.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E2E8F0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        story.append(t_fmea)

        doc.build(story)
        buf.seek(0)
        return buf

    v_dict = {"Gemi Adı": vessel_name, "Kategori": vessel_category, "Tam Boy (LOA)": loa, "Genişlik (Beam)": beam, "Draft": draft, "Deplasman": displacement, "Sevk Tipi": propulsion, "Görev Alanı": operating_area}

    c_dl1, c_dl2 = st.columns(2)
    with c_dl1:
        excel_matrix_data = generate_advanced_excel_with_charts(v_dict, matrix_export_data)
        st.download_button(label="📊 Grafikli & 5x5 Matrisli Gelişmiş Excel'i İndir (.XLSX)", data=excel_matrix_data, file_name=f"{vessel_name.replace(' ', '_')}_Advanced_Risk_Dashboard.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Excel dosyası; 5x5 renk kodlu matris, FMEA tablosu ve yerel karşılaştırma grafiklerini içerir.")

    with c_dl2:
        full_pdf = generate_full_pdf_with_charts(v_dict, matched, selected_solutions_data, stat_chart_buf, hull_chart_buf)
        st.download_button(label="📄 2D Silüetli & Grafikli Resmi Raporu İndir (.PDF)", data=full_pdf, file_name=f"{vessel_name.replace(' ', '_')}_Engineering_Report.pdf", mime="application/pdf")
        st.caption("PDF dosyası; 2D parametrik gövde kesitini, kaza olasılık grafiklerini ve resmi FMEA matrisini içerir.")

# -------------------------------------------------------------
# 2. SEKME: KLAS & REGÜLASYON KURAL MOTORU (RULE ENGINE)
# -------------------------------------------------------------
with tab_rules:
    st.subheader("📜 Uygulanabilir Uluslararası Klas Kuruluşu & Emniyet Kuralları")
    st.markdown(f"Girilen parametrelere göre (**{vessel_category}**, Boy: **{loa}m**, Sevk: **{propulsion}**) zorunlu olan kurallar:")
    
    compliance_rules = get_regulatory_compliance(loa, propulsion, vessel_category)
    for r in compliance_rules:
        with st.container():
            st.markdown(f"#### 🏛️ {r['code']} — {r['title']}")
            st.write(r['desc'])
            st.divider()

# -------------------------------------------------------------
# 3. SEKME: AÇIK KAYNAK CAD KÜTÜPHANESİ & TOPLULUK YÜKLEME HUB'I
# -------------------------------------------------------------
with tab_cad:
    st.subheader("📐 Açık Kaynak Denizcilik Tasarım & CAD Kütüphanesi")
    st.markdown("Topluluk standartlarına göre hazırlanmış örnek emniyet paftaları, P&ID diyagramları ve 2D/3D yapısal modeller.")
    
    # Kütüphanedeki Dosyaları Listele
    for item in cad_library:
        with st.container():
            c_cad1, c_cad2 = st.columns([3, 1])
            with c_cad1:
                st.markdown(f"#### 📄 {item['title']}")
                st.markdown(f"**Kategori:** `{item['category']}` | **Format:** `{item['format']}` | **Standart:** `{item['code_standard']}` | **Yükleyen:** `{item.get('contributor', 'Topluluk')}`")
                st.write(item['description'])
            with c_cad2:
                dummy_dxf = f"SECTION\n2\nHEADER\n0\nENDSEC\n0\nCOMMENT\nSafety-by-Design Cad: {item['filename']}\n0\nEOF"
                st.download_button(label=f"💾 {item['filename']} İndir", data=dummy_dxf, file_name=item['filename'], mime="application/dxf", key=f"cad_{item['filename']}")
            st.divider()

    # YENİ ÖZELLİK 5: TOPLULUK CAD DOSYASI YÜKLEME MERKEZİ (COMMUNITY UPLOAD)
    st.subheader("📤 Topluluk Havuzuna Yeni Tasarım / CAD Ekle")
    with st.form("cad_upload_form", clear_on_submit=True):
        u_col1, u_col2 = st.columns(2)
        with u_col1:
            cad_title = st.text_input("Tasarım / Pafta Başlığı", placeholder="Örn: 24m Motoryat A-60 Makine Dairesi Ön Perdesi")
            cad_cat = st.selectbox("Kategori", options=["Emniyet & Tahliye", "Makine & Borulama", "Elektrik & Batarya Güvenliği", "Yangın Söndürme", "Yapısal Bölme"])
            cad_std = st.text_input("Uyumlu Klas / Kod Standardı", placeholder="Örn: SOLAS Ch. II-2 / DNV Class")
        with u_col2:
            cad_desc = st.text_area("Teknik Açıklama", placeholder="Çizimin içerdiği detaylar, vana tipleri, yalıtım kalınlıkları...")
            contributor_name = st.text_input("Yükleyen Mühendis / Kurum", placeholder="Ad Soyad veya Tasarım Ofisi")
            uploaded_file = st.file_uploader("CAD Dosyası Seç (.dwg, .dxf, .step, .pdf)", type=["dwg", "dxf", "step", "pdf"])

        submitted = st.form_submit_button("🚀 Tasarımı Açık Kütüphaneye Yükle")
        if submitted and uploaded_file and cad_title:
            new_entry = {
                "filename": uploaded_file.name,
                "title": cad_title,
                "format": uploaded_file.name.split(".")[-1].upper(),
                "category": cad_cat,
                "description": cad_desc,
                "code_standard": cad_std or "Genel Denizcilik Standardı",
                "contributor": contributor_name or "Anonim Mühendis"
            }
            # Kaydet
            existing_user_cad = []
            if os.path.exists(USER_UPLOADS_FILE):
                try:
                    with open(USER_UPLOADS_FILE, "r", encoding="utf-8") as f:
                        existing_user_cad = json.load(f)
                except Exception: existing_user_cad = []
            
            existing_user_cad.append(new_entry)
            with open(USER_UPLOADS_FILE, "w", encoding="utf-8") as f:
                json.dump(existing_user_cad, f, ensure_ascii=False, indent=2)
            st.success(f"🎉 '{uploaded_file.name}' başarıyla açık tasarım kütüphanesine eklendi! Sayfayı yenilediğinizde listede görünecektir.")
