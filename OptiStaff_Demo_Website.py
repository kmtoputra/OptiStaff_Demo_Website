import streamlit as str
import pandas as pd
from datetime import datetime
import math
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
import base64
from PIL import Image, ImageDraw
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import time

# --- GLOBAL UI & APPLE ENTERPRISE CSS THEME ---
str.set_page_config(page_title="OptiStaff - Demo", layout="wide", initial_sidebar_state="auto")

str.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
    }
    
    @keyframes fadeIn {
        0% { opacity: 0; transform: translateY(12px); }
        100% { opacity: 1; transform: translateY(0); }
    }
    .main .block-container {
        animation: fadeIn 0.6s ease-out forwards;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    div.stButton > button {
        border-radius: 8px;
        border: 1px solid rgba(0,0,0,0.05);
        font-weight: 500;
        background-color: #0071E3 !important;
        color: white !important;
        transition: all 0.2s cubic-bezier(0.25, 0.1, 0.25, 1);
        padding: 0.5rem 1rem;
    }
    div.stButton > button:hover {
        background-color: #0077ED !important;
        transform: scale(1.01);
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
    div.stButton > button:active { transform: scale(0.98); }

    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 8px !important;
        transition: border-color 0.2s, box-shadow 0.2s;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #0071E3 !important;
        box-shadow: 0 0 0 1px #0071E3 !important;
    }
    
    img { max-width: 100%; height: auto; border-radius: 8px; }
    .stAlert { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# --- SECURE MASTER STATE MEMORY INITIALIZATION ---
if 'broadcast_db' not in str.session_state: str.session_state.broadcast_db = {}
if 'logged_in' not in str.session_state: str.session_state.logged_in = False
if 'user_role' not in str.session_state: str.session_state.user_role = None
if 'user_name' not in str.session_state: str.session_state.user_name = None
if 'user_event_id' not in str.session_state: str.session_state.user_event_id = None
if 'reset_key' not in str.session_state: str.session_state.reset_key = 0 

# --- PERISAI HTTP PROXY AUTOMATIC RETRY CONTROL ---
http_session = requests.Session()
strategi_retry = Retry(total=4, backoff_factor=1, status_forcelist=[500, 502, 503, 504], raise_on_status=False)
http_session.mount("https://", HTTPAdapter(max_retries=strategi_retry))

API_URL = "https://script.google.com/macros/s/AKfycbyMyVBSKWB_zxJKUjX7shSvk9pQ0WxB7krb3JlH4JFodo-iIOhPeJeEEzN5aZILaLjh/exec"

# ==========================================
# FUNGSI PENDUKUNG (GLOBAL SCOPE)
# ==========================================
def hitung_jarak(lat1, lon1, lat2, lon2):
    R = 6371000 
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi, dlam = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2) * math.sin(dlam/2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1-a)))

def terjemahkan_koordinat_ke_jalan(event_id, nama_event_gsheet):
    ev_clean = f"{event_id}".upper().strip()
    nm_clean = f"{nama_event_gsheet}".upper().strip()
    
    if "EVT01" in ev_clean or "MUSIC" in nm_clean or "FEST" in nm_clean:
        return "Mall Citraland Grogol, Jl. Letjen S. Parman, Tanjung Duren, Jakarta Barat"
    elif "EVT02" in ev_clean or "PSSI" in nm_clean or "GBK" in nm_clean:
        return "Stadion Utama Gelora Bung Karno (GBK), Pintu 1 Senayan, Jakarta Pusat"
    elif "EVT03" in ev_clean or "PRJ" in nm_clean:
        return "JIExpo Kemayoran, Gedung Pusat Niaga lt. 1, Kemayoran, Jakarta Pusat"
    return "Zonasi Kompleks Operasional Event"

def bersihkan_koordinat_radar(val, is_latitude=True):
    try:
        s = f"{val}".replace(',', '').replace('.', '').strip()
        if not s: return -6.2181 if is_latitude else 106.8024
        is_negative = s.startswith('-')
        if is_negative: s = s[1:]
        
        num = float(s)
        if is_latitude:
            while num > 9.0: num /= 10.0
            return -num if is_negative else num
        else:
            while num > 180.0: num /= 10.0
            if num < 90.0:
                while num < 100.0: num *= 10.0
            return -num if is_negative else num
    except:
        return -6.2181 if is_latitude else 106.8024

# --- GENERATOR DOKUMEN LPJ ---
def buat_word_lpj_apple_style(df, event_id, info_event):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_m = p_meta.add_run("OPTI STAFF CLOUD SYSTEM\nExecutive Report Management")
    run_m.font.size = Pt(8)
    run_m.font.color.rgb = RGBColor(140, 140, 142) 
    
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("Laporan Pertanggungjawaban")
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(29, 29, 31) 
    
    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(f"Project ID: {event_id} — {info_event.get('nama_event', 'Exhibition Pro Master')}")
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = RGBColor(100, 100, 100)
    
    p_line = doc.add_paragraph()
    p_line.add_run("—" * 65).font.color.rgb = RGBColor(210, 210, 215)
    
    h1 = doc.add_heading(level=1)
    run_h1 = h1.add_run("1. Ringkasan Operasional Lapangan")
    run_h1.font.size = Pt(14)
    run_h1.font.bold = True
    
    p_summary = doc.add_paragraph()
    total_log = len(df)
    total_valid = len(df[df['status_geotag'].astype('str').str.contains("MATCH", na=False)])
    total_fraud = len(df[df['status_geotag'].astype('str').str.contains("OUT OF RANGE", na=False)])
    
    p_summary.add_run(f"Waktu Penyerahan Dokumen: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    raw_loc = info_event.get('lokasi_target', '-6.2181, 106.8024')
    try: lt_p, ln_p = map(float, raw_loc.split(","))
    except: lt_p, ln_p = -6.2181, 106.8024
    
    c_lat_clean = bersihkan_koordinat_radar(lt_p, is_latitude=True)
    c_lon_clean = bersihkan_koordinat_radar(ln_p, is_latitude=False)
    alamat_nyata_jalan = terjemahkan_koordinat_ke_jalan(event_id, info_event.get('nama_event',''))
    
    p_summary.add_run(f"Pusat Koordinat Geofence: {c_lat_clean:.6f}, {c_lon_clean:.6f} ({alamat_nyata_jalan})\n")
    p_summary.add_run(f"Total Aktivitas Personel Lapangan: {total_log} Transaksi\n")
    p_summary.add_run(f"Hasil Validasi Radar Satelit: {total_valid} Valid / ")
    
    r_fraud = p_summary.add_run(f"{total_fraud} Terindikasi Pelanggaran Lokasi")
    if total_fraud > 0:
        r_fraud.font.color.rgb = RGBColor(255, 59, 48)
        r_fraud.font.bold = True
        
    h2 = doc.add_heading(level=1)
    run_h2 = h2.add_run("2. Log Detail Presensi Masuk & Pulang")
    run_h2.font.size = Pt(14)
    run_h2.font.bold = True
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    titles = ["Nama Personel", "Waktu Kirim", "Tipe Shift", "Keamanan Geotag", "Laporan Progress Kerja"]
    for idx, text in enumerate(titles):
        hdr_cells[idx].text = text
        hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
        bg_shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F7"/>')
        hdr_cells[idx]._tc.get_or_add_tcPr().append(bg_shading)
        
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = f"{row.get('crew_name', '-')}"
        row_cells[1].text = f"{row.get('timestamp', '-')}"
        row_cells[2].text = f"{row.get('tipe_absen', '-')}"
        
        status_gps = f"{row.get('status_geotag', '-')}"
        row_cells[3].text = status_gps
        if "OUT OF RANGE" in status_gps:
            row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 59, 48)
            row_cells[3].paragraphs[0].runs[0].font.bold = True
            
        row_cells[4].text = f"{row.get('status_pekerjaan', '-')}"
        
    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io.getvalue()

# ----------------------------------------
# AUTHENTICATION GATEWAY
# ----------------------------------------
if not str.session_state.logged_in:
    str.title("OptiStaff - Autentikasi Sistem")
    str.caption("Sistem Manajemen Kehadiran Terintegrasi Geofencing")
    str.markdown("---")
    
    with str.container():
        with str.form("login_gate"):
            input_email = str.text_input("Email Pengguna")
            input_pass = str.text_input("Kata Sandi", type="password")
            
            if str.form_submit_button("Masuk ke Sistem", use_container_width=True):
                with str.spinner("Memvalidasi kredensial..."):
                    if input_email == "eo@optistaff.com" and input_pass == "adminpro":
                        str.session_state.update({"logged_in": True, "user_role": "EO (Organizer)", "user_name": "Administrator Pusat", "user_event_id": "ALL"})
                        time.sleep(0.5) 
                        str.rerun()
                    else:
                        try:
                            res_raw = http_session.get(f"{API_URL}?sheet=crews", timeout=25)
                            df_crews_db = pd.DataFrame(res_raw.json())
                            if not df_crews_db.empty and input_email in df_crews_db['email'].values:
                                u_info = df_crews_db[df_crews_db['email'] == input_email].iloc[0]
                                if f"{u_info.get('password', '')}" == input_pass:
                                    if f"{u_info.get('status', 'Aktif')}" == 'Aktif':
                                        str.session_state.update({"logged_in": True, "user_role": u_info['role'], "user_name": u_info['nama'], "user_event_id": u_info['event_id']})
                                        time.sleep(0.5)
                                        str.rerun()
                                    else: str.error("Akses ditolak. Masa aktif akun Anda telah berakhir.")
                                else: str.error("Kredensial tidak valid.")
                            else: str.error("Email tidak ditemukan dalam sistem.")
                        except: str.error("Gagal terhubung ke server utama. Periksa koneksi internet Anda.")
    str.stop()

# --- SIDEBAR CONTROL PANEL ---
str.sidebar.markdown("## OptiStaff - Demo")
str.sidebar.markdown(f"**Pengguna:** {str.session_state.user_name}")
str.sidebar.info(f"Hak Akses: {str.session_state.user_role}")
if str.session_state.user_event_id and str.session_state.user_event_id != "ALL":
    str.sidebar.success(f"ID Proyek: {str.session_state.user_event_id}")
else: 
    str.sidebar.warning("Mode: Akses Administrator")

if str.sidebar.button("Keluar Sistem"):
    with str.spinner("Mengakhiri sesi..."):
        time.sleep(0.5)
        str.session_state.logged_in = False
        str.rerun()

# ----------------------------------------
# INTERFACE ROLE: EO (ORGANIZER)
# ----------------------------------------
if str.session_state.user_role == "EO (Organizer)":
    str.title("Portal Manajemen Pusat")
    tab_ev, tab_cr, tab_mon = str.tabs(["Manajemen Proyek", "Registrasi Personel", "Penutupan Proyek"])
    
    with tab_ev:
        str.subheader("Pendaftaran Proyek Baru")
        try: df_events = pd.DataFrame(http_session.get(f"{API_URL}?sheet=events", timeout=25).json())
        except: df_events = pd.DataFrame(columns=["event_id", "nama_event", "lokasi_target", "radius_meter", "status", "status_lpj"])

        if not df_events.empty and 'status' in df_events.columns:
            df_events['status'] = df_events['status'].fillna('Aktif').replace('', 'Aktif')
            df_events_view = df_events[df_events['status'] == 'Aktif']
        else: df_events_view = df_events

        with str.form("add_event_form", clear_on_submit=True):
            ev_id = str.text_input("ID Proyek:", value=f"EVT0{len(df_events)+1}")
            ev_nama = str.text_input("Nama Proyek:")
            ev_loc = str.text_input("Koordinat Pusat (Lat, Lon):", value="-6.2181, 106.8024")
            ev_rad = str.slider("Radius Geofence (Meter):", 50, 500, 100)
            if str.form_submit_button("Aktifkan Proyek", use_container_width=True) and ev_nama:
                with str.spinner("Memproses aktivasi proyek..."):
                    payload = {"target_sheet":"events", "event_id":ev_id, "nama_event":ev_nama, "lokasi_target":ev_loc, "radius_meter":ev_rad, "status": "Aktif", "status_lpj": "Belum Diserahkan"}
                    http_session.post(API_URL, json=payload)
                    str.success(f"Proyek '{ev_nama}' berhasil diaktifkan.")
                    time.sleep(1)
                    str.rerun()
        
        str.markdown("**Daftar Proyek Aktif**")
        str.dataframe(df_events_view, use_container_width=True)

    with tab_cr:
        str.subheader("Distribusi Personel Lapangan")
        try: df_crews = pd.DataFrame(http_session.get(f"{API_URL}?sheet=crews", timeout=25).json())
        except: df_crews = pd.DataFrame(columns=["crew_id", "nama", "email", "role", "event_id", "password", "status"])
        
        try:
            df_ev_list = pd.DataFrame(http_session.get(f"{API_URL}?sheet=events", timeout=25).json())
            df_ev_list['status'] = df_ev_list['status'].fillna('Aktif').replace('', 'Aktif')
            list_event_id = df_ev_list[df_ev_list['status'] == 'Aktif']["event_id"].tolist()
        except: list_event_id = ["EVT01"]

        with str.form("add_crew_form", clear_on_submit=True):
            c_id = str.text_input("ID Personel:", value=f"CRW0{len(df_crews)+1}")
            c_nama = str.text_input("Nama Lengkap:")
            c_email = str.text_input("Email Login:")
            c_pass = str.text_input("Kata Sandi:", value="12345")
            c_role = str.selectbox("Tingkat Akses:", ["Team Leader", "Crew"])
            c_ev_target = str.selectbox("Penugasan Proyek:", list_event_id)
            if str.form_submit_button("Daftarkan Personel", use_container_width=True) and c_nama and c_email:
                with str.spinner("Mendaftarkan personel ke database..."):
                    payload = {"target_sheet": "crews", "crew_id": c_id, "nama": c_nama, "email": c_email, "role": c_role, "event_id": c_ev_target, "password": c_pass, "status": "Aktif"}
                    http_session.post(API_URL, json=payload)
                    str.success("Personel berhasil didistribusikan.")
                    time.sleep(1)
                    str.rerun()
                    
        str.markdown("**Direktori Personel**")
        str.dataframe(df_crews, use_container_width=True)

    with tab_mon:
        str.subheader("Penyelesaian & Arsip Proyek")
        try:
            df_all_att = pd.DataFrame(http_session.get(f"{API_URL}?sheet=attendances", timeout=25).json())
            df_events_check = pd.DataFrame(http_session.get(f"{API_URL}?sheet=events", timeout=25).json())
            
            if not df_events_check.empty:
                df_events_check['status'] = df_events_check['status'].fillna('Aktif').replace('', 'Aktif')
                df_events_check['status_lpj'] = df_events_check['status_lpj'].fillna('Belum Diserahkan').replace('', 'Belum Diserahkan')
                
                unique_events_aktif = df_events_check[df_events_check['status'] == 'Aktif']["event_id"].tolist()
                
                if unique_events_aktif:
                    c_close1, c_close2 = str.columns([3, 1])
                    with c_close1:
                        event_target_close = str.selectbox("Pilih Proyek untuk Ditutup:", unique_events_aktif, key="close_ev")
                    with c_close2:
                        str.markdown("<div style='margin-top: 28px;' class='hidden-mobile'></div>", unsafe_allow_html=True)
                        tombol_close = str.button("Tutup Proyek", type="primary", use_container_width=True)
                    
                    info_event = df_events_check[df_events_check['event_id'] == event_target_close].iloc[0]
                    status_lpj_saat_ini = info_event.get('status_lpj', 'Belum Diserahkan')
                    
                    str.info(f"Status Dokumen LPJ: {status_lpj_saat_ini}")
                    
                    if tombol_close:
                        if status_lpj_saat_ini != "Sudah Diserahkan":
                            str.error("Operasi ditolak. Menunggu pengesahan laporan dari Team Leader lapangan.")
                        else:
                            with str.spinner("Mengarsipkan proyek dan memutus akses..."):
                                res_close = http_session.post(API_URL, json={"action": "complete_event", "event_id": event_target_close})
                                if res_close.json().get("status") == "success":
                                    str.success(f"Proyek {event_target_close} berhasil diselesaikan.")
                                    time.sleep(1)
                                    str.rerun()
                else: str.info("Semua proyek komersial telah diarsipkan.")
                
                str.markdown("---")
                str.subheader("Riwayat Log Kehadiran")
                if not df_all_att.empty:
                    selected_event_filter = str.selectbox("Filter berdasarkan Proyek:", df_all_att['event_id'].unique().tolist())
                    df_filtered = df_all_att[df_all_att['event_id'] == selected_event_filter]
                    df_lpj_eo = df_filtered[['attendance_id', 'crew_name', 'tipe_absen', 'timestamp', 'status_geotag', 'status_pekerjaan']].copy()
                    
                    def color_eo_status(val):
                        if "OUT OF RANGE" in f"{val}": return 'background-color: #fde8e8; color: #9b1c1c;'
                        return 'background-color: #eafbf0; color: #155724;'
                    str.dataframe(df_lpj_eo.style.applymap(color_eo_status, subset=['status_geotag']), use_container_width=True)
        except Exception as e: str.error("Gagal memuat modul pemantauan.")

# ----------------------------------------
# INTERFACE ROLE: TEAM LEADER
# ----------------------------------------
elif str.session_state.user_role == "Team Leader":
    active_ev_id = str.session_state.user_event_id
    str.title("Dashboard Pengawas Operasional")
    str.caption(f"ID Proyek Aktif: {active_ev_id}")
    
    tab_tl_mon, tab_tl_add, tab_tl_bc = str.tabs(["Pemantauan Radar", "Manajemen Kru", "Pusat Instruksi"])
    
    with tab_tl_mon:
        try:
            df_attendance = pd.DataFrame(http_session.get(f"{API_URL}?sheet=attendances", timeout=25).json())
            if not df_attendance.empty: df_attendance = df_attendance[df_attendance['event_id'] == active_ev_id]
            df_ev_all = http_session.get(f"{API_URL}?sheet=events", timeout=25).json()
            info_ev = next((item for item in df_ev_all if item["event_id"] == active_ev_id), {"nama_event": "Exhibition"})
        except: df_attendance = pd.DataFrame(); info_ev = {"nama_event": "Exhibition"}
            
        str.subheader("Log Kehadiran Personel")
        if df_attendance.empty: str.info("Belum ada data kehadiran terekam.")
        else:
            df_display = df_attendance.copy()
            str.dataframe(df_display[[c for c in df_display.columns if c != 'foto_selfie_url']], use_container_width=True)
            
            str.markdown("---")
            str.subheader("Radar Geolocation")
            df_map = df_display[['latitude', 'longitude', 'crew_name', 'status_geotag']].copy()
            
            df_map['latitude'] = df_map['latitude'].apply(lambda x: bersihkan_koordinat_radar(x, is_latitude=True))
            df_map['longitude'] = df_map['longitude'].apply(lambda x: bersihkan_koordinat_radar(x, is_latitude=False))
            df_map = df_map.dropna(subset=['latitude', 'longitude'])
            
            if not df_map.empty: str.map(df_map, latitude='latitude', longitude='longitude', size=40)
            
            str.markdown("---")
            str.subheader("Pengesahan Laporan Akhir (LPJ)")
            apple_docx = buat_word_lpj_apple_style(df_display, active_ev_id, info_ev)
            
            c_doc1, c_doc2 = str.columns(2)
            with c_doc1:
                str.download_button(label="Unduh Dokumen Draf (.docx)", data=apple_docx, file_name=f"LPJ_{active_ev_id}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="secondary", use_container_width=True)
            with c_doc2:
                if str.button("Sahkan & Kirim ke Pusat", type="primary", use_container_width=True):
                    with str.spinner("Mentransmisikan pengesahan digital..."):
                        res_lpj = http_session.post(API_URL, json={"action": "submit_lpj", "event_id": active_ev_id, "target_sheet": "events"}).json()
                        if res_lpj.get("status") == "success":
                            str.success("Dokumen berhasil disahkan. Menunggu penutupan proyek oleh pusat.")
                        else: str.error("Terjadi kendala sinkronisasi.")

            str.markdown("---")
            str.subheader("Verifikasi Visual Kamera")
            df_display['pilihan_id'] = df_display['crew_name'] + " (" + df_display['attendance_id'] + ")"
            id_terpilih = str.selectbox("Pilih Personel:", df_display['pilihan_id'].tolist())
            row_data = df_display[df_display['pilihan_id'] == id_terpilih].iloc[0]
            if row_data['foto_selfie_url'] and "base64," in f"{row_data['foto_selfie_url']}":
                c_img, c_txt = str.columns([1, 2])
                with c_img: str.image(row_data['foto_selfie_url'], use_container_width=True)
                with c_txt: 
                    # --- BLOK PERINGATAN OUT OF RANGE BERWARNA MERAH ---
                    status_geo = f"{row_data.get('status_geotag', '')}"
                    if "OUT OF RANGE" in status_geo:
                        str.error(f"PELANGGARAN LOKASI: {status_geo}")
                    else:
                        str.success(f"LOKASI VALID: {status_geo}")
                        
                    str.info(f"Tipe Shift: {row_data.get('tipe_absen','-')}")
                    str.write(f"**Laporan Pekerjaan:** {row_data.get('status_pekerjaan','-')}")

    with tab_tl_add:
        str.subheader("Rekrutmen Personel Lokal")
        try: df_crews_tl = pd.DataFrame(http_session.get(f"{API_URL}?sheet=crews", timeout=25).json())
        except: df_crews_tl = pd.DataFrame(columns=["crew_id", "nama", "email", "role", "event_id", "password", "status"])
        
        with str.form("form_add_by_tl", clear_on_submit=True):
            n_id = str.text_input("ID Personel:", value=f"CRW0{len(df_crews_tl)+1}")
            n_nama = str.text_input("Nama Lengkap:")
            n_email = str.text_input("Email Login:")
            n_pass = str.text_input("Kata Sandi (Password):", value="12345")
            if str.form_submit_button("Daftarkan Personel", type="primary", use_container_width=True) and n_nama and n_email:
                with str.spinner("Mendaftarkan kredensial..."):
                    payload_tl = {"target_sheet": "crews", "crew_id": n_id, "nama": n_nama, "email": n_email, "role": "Crew", "event_id": active_ev_id, "password": n_pass, "status": "Aktif"}
                    http_session.post(API_URL, json=payload_tl)
                    str.success("Personel berhasil didaftarkan.")
                    time.sleep(1)
                    str.rerun()

        str.markdown("**Direktori Personel Lapangan Anda**")
        if not df_crews_tl.empty:
            df_kru_terfilter = df_crews_tl[(df_crews_tl['role'] == 'Crew') & (df_crews_tl['event_id'] == active_ev_id)]
            str.dataframe(df_kru_terfilter, use_container_width=True)
        else:
            str.info("Belum ada personel terdaftar di proyek ini.")

    with tab_tl_bc:
        str.subheader("Pusat Instruksi Komando")
        current_bc = str.session_state.broadcast_db.get(active_ev_id, "Tidak ada instruksi khusus.")
        str.info(f"Instruksi Aktif: {current_bc}")
        rk = str.session_state.reset_key
        pesan_baru = str.text_area("Pesan Pembaruan:", key=f"bc_{rk}")
        if str.button("Publikasikan Instruksi", type="primary", use_container_width=True) and pesan_baru:
            str.session_state.broadcast_db[active_ev_id] = pesan_baru
            str.success("Instruksi berhasil diperbarui.")
            time.sleep(1)
            str.session_state.reset_key += 1
            str.rerun()

# ----------------------------------------
# INTERFACE ROLE: CREW (LAPANGAN)
# ----------------------------------------
else:
    active_ev_id = str.session_state.user_event_id
    str.title("Portal Operasional Lapangan")
    
    current_bc = str.session_state.broadcast_db.get(active_ev_id, "Silakan jalankan tugas operasional sesuai standar.")
    str.info(f"Instruksi Komando ({active_ev_id}): {current_bc}")
    
    try:
        df_event_kru = pd.DataFrame(http_session.get(f"{API_URL}?sheet=events", timeout=25).json())
        df_event_kru = df_event_kru[df_event_kru['event_id'] == active_ev_id]
        list_event = df_event_kru["nama_event"].tolist() if not df_event_kru.empty else ["Operasional Reguler"]
    except: list_event = ["Operasional Reguler"]

    with str.container():
        rk = str.session_state.reset_key
        event_pilihan = str.selectbox("Pilih Modul Penugasan:", list_event, key=f"ev_{rk}")
        status_shift = str.radio("Tipe Kehadiran:", ["Absen Masuk", "Absen Pulang"], horizontal=True, key=f"sh_{rk}")
        foto_file = str.camera_input("Verifikasi Kamera", key=f"cam_{rk}")
        
        col1, col2 = str.columns(2)
        with col1: mock_lat = str.number_input("Koordinat Latitude:", value=-6.2181, format="%.5f", key=f"lat_{rk}")
        with col2: mock_lon = str.number_input("Koordinat Longitude:", value=106.8024, format="%.5f", key=f"lon_{rk}")
        
        laporan_kerja = str.text_area("Laporan Progres Pekerjaan:", key=f"lap_{rk}")

        if str.button("Kirim Validasi Kehadiran", type="primary", use_container_width=True) and foto_file:
            with str.spinner("Memproses sinkronisasi data lokasi dan citra..."):
                waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                img = Image.open(io.BytesIO(foto_file.getvalue()))
                canvas = ImageDraw.Draw(img)
                
                W, H = img.size
                tinggi_kotak = int(H * 0.07) if H * 0.07 > 35 else 35
                canvas.rectangle(((0, 0), (W, tinggi_kotak)), fill="black")
                
                text_segel = f"AUTH: {waktu_sekarang} | {status_shift.upper()} | GEO: {mock_lat:.4f}, {mock_lon:.4f} | PRJ: {active_ev_id}"
                canvas.text((int(W * 0.02), int(tinggi_kotak * 0.2)), text_segel, fill="#00ffcc")
                
                buffered_io = io.BytesIO()
                img.save(buffered_io, format="JPEG")
                data_url_foto = f"data:image/jpeg;base64,{base64.b64encode(buffered_io.getvalue()).decode('utf-8')}"

                try:
                    lat_t, lon_t = map(float, f"{df_event_kru.iloc[0]['lokasi_target']}".split(","))
                    rad_t = float(df_event_kru.iloc[0]['radius_meter'])
                except: lat_t, lon_t, rad_t = -6.2181, 106.8024, 100

                c_lat_final = bersihkan_koordinat_radar(mock_lat, is_latitude=True)
                c_lon_final = bersihkan_koordinat_radar(mock_lon, is_latitude=False)

                jarak = hitung_jarak(c_lat_final, c_lon_final, lat_t, lon_t)
                
                if jarak <= rad_t:
                    status_absen = "MATCH (Valid)"
                    pesan_hasil = f"Kehadiran berhasil divalidasi ke server. (Jarak: {jarak:.1f} m)"
                    is_valid = True
                else:
                    status_absen = "OUT OF RANGE (Fraud Suspect)"
                    pesan_hasil = f"PERINGATAN: Lokasi Anda ({jarak:.1f} m) di luar radius event ({rad_t} m). Data tetap dicatat sebagai 'Fraud Suspect'."
                    is_valid = False
                
                payload = {"target_sheet": "attendances", "attendance_id": f"ATT-{datetime.now().strftime('%H%M%S')}", "event_id": active_ev_id, "crew_name": str.session_state.user_name, "timestamp": waktu_sekarang, "latitude": c_lat_final, "longitude": c_lon_final, "status_geotag": status_absen, "status_pekerjaan": laporan_kerja, "foto_selfie_url": data_url_foto, "tipe_absen": status_shift}
                res_p = http_session.post(API_URL, json=payload, timeout=25)
                
                if res_p.json().get("status") == "success":
                    if is_valid:
                        str.success(pesan_hasil)
                    else:
                        str.warning(pesan_hasil)
                    
                    time.sleep(2.5) 
                    str.session_state.reset_key += 1 
                    str.rerun()