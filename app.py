import io
import math
import time
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# KONFIGURASI HALAMAN STREAMLIT & TEMA
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Tulip Edition",
    layout="wide",
    page_icon="🌷"
)

# Custom CSS: Memastikan SEMUA Teks Berwarna Gelap & Kontras (Termasuk Tab Pink Tua)
st.markdown("""
<style>
    /* Background Utama Gradasi Pastel Soft */
    .stApp {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe6f0 50%, #e8f4f8 100%);
        color: #2D3142;
        font-family: 'Poppins', sans-serif;
    }
    
    /* Sidebar Kiri - Pink Soft dengan Teks Ungu Tua Tegas */
    section[data-testid="stSidebar"] {
        background-color: rgba(255, 235, 243, 0.98) !important;
        border-right: 2px solid #F8BBD0;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #3A0056 !important;
        font-weight: 700 !important;
    }

    /* Tab Navigasi - Teks Pink Tua Pekat (#880E4F) agar Kontras & Tidak Nabrak Background */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF !important;
        border-radius: 15px;
        padding: 8px;
        border: 2px solid #F48FB1;
        box-shadow: 0 4px 12px rgba(136, 14, 79, 0.08);
    }
    .stTabs [data-baseweb="tab"],
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] span,
    .stTabs button[role="tab"],
    .stTabs button[role="tab"] p {
        color: #880E4F !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 10px;
        padding: 8px 14px;
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"]:hover,
    .stTabs button[role="tab"]:hover {
        background-color: #FCE4EC !important;
        color: #AD1457 !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs button[aria-selected="true"] {
        background: linear-gradient(45deg, #AD1457, #D81B60) !important;
        box-shadow: 0px 4px 12px rgba(173, 20, 87, 0.35);
    }
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs button[aria-selected="true"] p {
        color: #FFFFFF !important;
    }
    
    /* Input Box & Number Input */
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border-radius: 8px;
        color: #111111 !important;
        border: 1px solid #F48FB1 !important;
    }

    /* Tombol Utama */
    .stButton>button {
        background: linear-gradient(45deg, #C2185B, #E91E63);
        color: white !important;
        border: none;
        border-radius: 12px;
        font-weight: bold;
        box-shadow: 0px 4px 10px rgba(194, 24, 91, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(45deg, #880E4F, #C2185B);
        transform: translateY(-2px);
    }
    
    /* Card Keterangan Status Warna Mobil */
    .status-card {
        padding: 15px;
        border-radius: 15px;
        margin-bottom: 15px;
        font-weight: bold;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .status-oke {
        background-color: rgba(200, 230, 201, 0.98);
        color: #1B5E20;
        border: 2px solid #81C784;
    }
    .status-sedang {
        background-color: rgba(255, 245, 157, 0.98);
        color: #E65100;
        border: 2px solid #FFF176;
    }
    .status-bahaya {
        background-color: rgba(255, 205, 210, 0.98);
        color: #B71C1C;
        border: 2px solid #E57373;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 1. KEAMANAN & AUTHENTICATION
# ----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col_top1, col_top2 = st.columns([4, 1])
    with col_top2:
        st.markdown(
            "<p style='text-align: right; font-size: 13px; color: #880E4F; margin: 0;'>pw: "
            "<span style='font-family: monospace; font-weight: bold; color: #C2185B;'>cantikaaa27</span></p>",
            unsafe_allow_html=True
        )

    st.markdown("<h1 style='text-align: center; color: #3A0056;'>🌷 Perencanaan Geometrik Jalan Interaktif</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #880E4F;'>Standar TPGJAK 1997 / RAE 2021</h4>", unsafe_allow_html=True)
    
    st.image(
        "https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?q=80&w=1200&auto=format&fit=crop",
        caption="Sistem Geometri Jalan Interaktif (Tulip Edition)",
        use_container_width=True
    )
    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("💡 **Petunjuk:** Password akses adalah **cantikaaa27**")
        pwd_input = st.text_input("Masukkan Password Akses:", type="password", placeholder="Ketik password di sini...")

        if st.button("Masuk Aplikasi 🚀", use_container_width=True):
            if pwd_input.strip() == "cantikaaa27":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Password salah! Silakan periksa kembali tulisan di kanan atas.")

    st.stop()

# ----------------------------------------------------------------------------
# 2. LOGIKA MATEMATIS & GEOMETRIK (BM 1997 / 2021)
# ----------------------------------------------------------------------------
V_TAB = [20, 30, 40, 50, 60, 80, 100, 120]
R_FC_TAB = [60, 130, 250, 350, 500, 900, 1500, 2500]

def sta_str(s: float, d: int = 2) -> str:
    sign = "-" if s < 0 else ""
    s_abs = abs(s)
    return f"{sign}{int(s_abs // 1000)}+{s_abs % 1000:0{4 + d}.{d}f}"

def f_max(V: float) -> float:
    return -0.00065 * V + 0.192 if V < 80 else -0.00125 * V + 0.24

def r_min(V: float, emax: float) -> float:
    return V**2 / (127 * (emax + f_max(V)))

def hitung_tikungan(V: float, delta: float, emax: float, en: float, Rc: float, sta_pi: float):
    fm = f_max(V)
    Rmin = r_min(V, emax)
    Dmax = 181913.53 * (emax + fm) / (V**2)
    D = 1432.39 / Rc
    e = -emax * (D**2) / (Dmax**2) + 2 * emax * D / Dmax
    e = max(e, en)

    re = 0.035 if V <= 70 else 0.025
    Ls_waktu = V * 3 / 3.6
    Ls_shortt = 0.022 * (V**3) / (Rc * 0.4) - 2.727 * V * e / 0.4
    Ls_re = (e + en) * V / (3.6 * re)
    Ls_min = max(Ls_waktu, Ls_shortt, Ls_re)
    Ls_pakai = math.ceil(Ls_min / 5) * 5

    R_fc = float(np.interp(V, V_TAB, R_FC_TAB))
    d2 = math.radians(delta / 2)
    
    res = dict(
        V=V, delta=delta, emax=emax, en=en, Rc=Rc, fm=fm, Rmin=Rmin,
        Dmax=Dmax, D=D, e=e, Ls_waktu=Ls_waktu, Ls_shortt=Ls_shortt,
        Ls_re=Ls_re, Ls_min=Ls_min, R_fc=R_fc, sta_pi=sta_pi, Ls_pakai=Ls_pakai
    )

    if Rc >= R_fc:
        Tc = Rc * math.tan(d2)
        Ec = Tc * math.tan(d2 / 2)
        Lc = math.radians(delta) * Rc
        res.update(
            jenis="Full Circle (FC)", Ls=0.0, Ls_sup=Ls_pakai, ths=0.0,
            dc=delta, Lc=Lc, T=Tc, E=Ec, Xs=0.0, Ys=0.0, p=0.0, k=0.0
        )
    else:
        ths = math.degrees(Ls_pakai / (2 * Rc))
        dc = delta - 2 * ths
        Lc = math.radians(dc) * Rc
        Ls = float(Ls_pakai)
        jenis = "Spiral-Circle-Spiral (S-C-S)"
        
        if Lc < 25:
            jenis = "Spiral-Spiral (S-S)"
            ths = delta / 2
            Ls = math.radians(ths) * 2 * Rc
            Lc, dc = 0.0, 0.0

        Xs = Ls * (1 - Ls**2 / (40 * Rc**2))
        Ys = (Ls**2) / (6 * Rc)
        p = Ys - Rc * (1 - math.cos(math.radians(ths)))
        k = Xs - Rc * math.sin(math.radians(ths))
        Ts = (Rc + p) * math.tan(d2) + k
        Es = (Rc + p) / math.cos(d2) - Rc
        res.update(
            jenis=jenis, Ls=Ls, Ls_sup=Ls, ths=ths, dc=dc, Lc=Lc, T=Ts,
            E=Es, Xs=Xs, Ys=Ys, p=p, k=k
        )

    res["Ltot"] = res["Lc"] + 2 * res["Ls"]
    res["sta_awal"] = sta_pi - res["T"]
    res["sta_akhir"] = res["sta_awal"] + res["Ltot"]
    res["v_safe"] = math.sqrt(127 * Rc * (res["e"] + fm))
    return res

def heading(s, Ls, Lc, Rc, dl):
    Ltot = 2 * Ls + Lc
    th = np.zeros_like(s)
    if Ls > 0:
        ths = Ls / (2 * Rc)
        m1 = (s > 0) & (s <= Ls)
        th[m1] = (s[m1]**2) / (2 * Rc * Ls)
        m2 = (s > Ls) & (s < Ls + Lc)
        th[m2] = ths + (s[m2] - Ls) / Rc
        m3 = (s >= Ls + Lc) & (s <= Ltot)
        th[m3] = dl - ((Ltot - s[m3])**2) / (2 * Rc * Ls)
    else:
        m = (s > 0) & (s <= Ltot)
        th[m] = s[m] / Rc
    th[s > Ltot] = dl
    return th

def hitung_profil(pvi_df, V):
    d = pvi_df.dropna().sort_values(pvi_df.columns[0]).reset_index(drop=True)
    if len(d) < 2:
        return None
    sta = d.iloc[:, 0].to_numpy(float)
    el = d.iloc[:, 1].to_numpy(float)
    lvm = d.iloc[:, 2].to_numpy(float)
    g = np.diff(el) / np.diff(sta)
    Jh = 0.278 * V * 2.5 + (V**2) / (254 * 0.40)

    kurva = []
    for i in range(1, len(sta) - 1):
        A = (g[i] - g[i - 1]) * 100
        if abs(A) < 1e-9:
            continue
        a = abs(A)
        if A < 0:
            L = a * (Jh**2) / 399
            if L < Jh:
                L = 2 * Jh - 399 / a
        else:
            L = a * (Jh**2) / (120 + 3.5 * Jh)
            if L < Jh:
                L = 2 * Jh - (120 + 3.5 * Jh) / a
        
        L = math.ceil(max(L, 0.6 * V, V * 3 / 3.6) / 10) * 10
        if lvm[i] > 0:
            L = lvm[i]
            
        kurva.append(dict(
            i=i, sta=sta[i], el=el[i], A=A, Lv=L, Ev=a * L / 800,
            bvc=sta[i] - L / 2, evc=sta[i] + L / 2,
            jenis="Cembung" if A < 0 else "Cekung",
            g1=g[i - 1], g2=g[i]
        ))

    x = np.arange(sta[0], sta[-1] + 0.01, 1.0)
    y = np.interp(x, sta, el)
    for k in kurva:
        m = (x >= k["bvc"]) & (x <= k["evc"])
        dx = x[m] - k["bvc"]
        yb = k["el"] - k["g1"] * k["Lv"] / 2
        y[m] = yb + k["g1"] * dx + (k["g2"] - k["g1"]) / (2 * k["Lv"]) * (dx**2)
        
    return dict(sta=sta, el=el, g=g * 100, kurva=kurva, x=x, y=y, Jh=Jh)

# ----------------------------------------------------------------------------
# 3. SIDEBAR INPUT & CALCULATIONS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Parameter Rencana")
    if st.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.subheader("1. Kriteria Desain")
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_sim = st.number_input("Kecepatan Kendaraan Simulasi (km/jam)", 10, 150, 65, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 45.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    
    Rmin = r_min(V, emax)
    st.caption(f"ℹ️ *R minimum acuan TPGJAK = {Rmin:.2f} m*")

    mode = st.radio("Jari-jari Tikungan Rc", ["Otomatis (Dibulatkan)", "Manual Input"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 300.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kiri", "Kanan"], horizontal=True)
    B = st.number_input("Lebar Perkerasan Jalan B (m)", 3.0, 30.0, 7.0, 0.5)

r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
sg = 1 if arah == "Kiri" else -1
Ls, Lc = r["Ls"], r["Lc"]
dl = math.radians(r["delta"])
Ltot = r["Ltot"]
ext = math.ceil(max(120, 1.4 * r["T"]) / 10) * 10

s_vals = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 0.5), [0, Ls, Ls + Lc, Ltot]]))
th_vals = heading(s_vals, Ls, Lc, Rc, dl)
dS = np.diff(s_vals)
x_horiz = np.concatenate([[0], np.cumsum(dS * (np.cos(th_vals[:-1]) + np.cos(th_vals[1:])) / 2)])
y_horiz = np.concatenate([[0], np.cumsum(dS * (np.sin(th_vals[:-1]) + np.sin(th_vals[1:])) / 2)])
i0 = np.argmin(np.abs(s_vals))
x_horiz, y_horiz = x_horiz - x_horiz[i0], (y_horiz - y_horiz[i0]) * sg
ths_ = th_vals * sg
nx, ny = -np.sin(ths_), np.cos(ths_)

# State DataFrame Profil Vertikal & Tanah Asli
if "pvi_df" not in st.session_state:
    st.session_state.pvi_df = pd.DataFrame({
        "STA (m)": [0.0, 250.0, 500.0, 750.0],
        "Elevasi Rencana (m)": [100.0, 105.0, 101.0, 107.0],
        "Lv Manual (m)": [0.0, 0.0, 0.0, 0.0]
    })

if "tanah_df" not in st.session_state:
    st.session_state.tanah_df = pd.DataFrame({
        "STA (m)": [0.0, 125.0, 250.0, 375.0, 500.0, 625.0, 750.0],
        "Elevasi Tanah (m)": [101.0, 103.5, 104.0, 102.5, 101.8, 105.0, 106.5]
    })

P_global = hitung_profil(st.session_state.pvi_df, V)
if P_global is not None:
    z_elev = np.interp(s_vals, P_global["x"], P_global["y"])
else:
    z_elev = np.zeros_like(s_vals)

# ----------------------------------------------------------------------------
# 4. EXPORT ENGINE (CAD DXF & EXCEL BM 2021 - FORMAT ALINYEMEN VERTIKAL LENGKAP)
# ----------------------------------------------------------------------------
def generate_excel_bm2021(r, pvi_df):
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=14, bold=True, color="3A0056")
    font_sec = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=10, bold=True, color="3A0056")
    font_bold = Font(name="Calibri", size=10, bold=True, color="111111")
    font_reg = Font(name="Calibri", size=10, color="111111")

    fill_header = PatternFill(start_color="C2185B", end_color="C2185B", fill_type="solid")
    fill_sub = PatternFill(start_color="F8BBD0", end_color="F8BBD0", fill_type="solid")
    fill_input = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
    fill_card = PatternFill(start_color="FFF8FA", end_color="FFF8FA", fill_type="solid")
    fill_ok = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")

    border_thin = Border(
        left=Side(style='thin', color='D8B4C6'),
        right=Side(style='thin', color='D8B4C6'),
        top=Side(style='thin', color='D8B4C6'),
        bottom=Side(style='thin', color='D8B4C6')
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # ------------------------------------------------------------------------
    # SHEET 1: TIKUNGAN HORISONTAL
    # ------------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Tikungan Horisontal"
    ws1["A1"] = "REKAPITULASI KALKULASI GEOMETRIK JALAN (BINA MARGA 2021)"
    ws1["A1"].font = font_title

    headers = ["Parameter Geometrik", "Nilai", "Satuan", "Formula / Acuan BM 2021"]
    for col_i, h in enumerate(headers, start=1):
        cell = ws1.cell(row=3, column=col_i, value=h)
        cell.font = font_sec
        cell.fill = fill_header
        cell.alignment = align_center

    fc = r["jenis"].startswith("Full")
    t_start = "TC" if fc else "TS"
    t_end = "CT" if fc else "ST"

    params = [
        ("Jenis Tikungan", r["jenis"], "-", "Kriteria R_fc (BM 2021)"),
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam", "Input Desain"),
        ("Sudut Tikungan (Delta)", r["delta"], "°", "Input Desain"),
        ("Jari-Jari Tikungan (Rc)", r["Rc"], "m", "Input Desain"),
        ("Superelevasi Maksimum (e_max)", r["emax"] * 100, "%", "Standar Fisiologis"),
        ("Superelevasi Desain (e)", round(r["e"] * 100, 2), "%", "=-e_max*D^2/Dmax^2 + 2*e_max*D/Dmax"),
        ("Panjang Transisi (Ls)", r["Ls_sup"], "m", "=MAX(Ls_waktu, Ls_shortt, Ls_re)"),
        ("Panjang Tangen (T)", round(r["T"], 3), "m", "=(Rc+p)*TAN(RAD(Delta/2))+k"),
        ("Jarak Eksternal (E)", round(r["E"], 3), "m", "=(Rc+p)/COS(RAD(Delta/2))-Rc"),
        ("Panjang Total Tikungan (Ltot)", round(r["Ltot"], 3), "m", "=Lc + 2*Ls"),
        ("Kecepatan Aman Maksimum (V_safe)", round(r["v_safe"], 2), "km/jam", "=SQRT(127*Rc*(e + f_max))"),
        ("STA PI", sta_str(r["sta_pi"]), "m", "Input Desain"),
        (f"STA {t_start} (Awal Tikungan)", sta_str(r["sta_awal"]), "m", "=STA_PI - Tangen(T)"),
        (f"STA {t_end} (Akhir Tikungan)", sta_str(r["sta_akhir"]), "m", f"=STA_{t_start} + Ltot")
    ]

    for row_i, (p, v, u, f) in enumerate(params, start=4):
        ws1.cell(row=row_i, column=1, value=p).border = border_thin
        c_v = ws1.cell(row=row_i, column=2, value=v)
        c_v.border = border_thin
        c_v.fill = fill_input
        ws1.cell(row=row_i, column=3, value=u).border = border_thin
        ws1.cell(row=row_i, column=4, value=f).border = border_thin

    for col_letter, w in [("A", 36), ("B", 22), ("C", 14), ("D", 42)]:
        ws1.column_dimensions[col_letter].width = w

    # ------------------------------------------------------------------------
    # SHEET 2: ALINYEMEN VERTIKAL (FORMAT LENGKAP SEPERTI DATA ALL VERTIKAL)
    # ------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Alinyemen Vertikal")
    ws2["A1"] = "PERHITUNGAN ALINYEMEN VERTIKAL JALAN (BINA MARGA / TPGJAK)"
    ws2["A1"].font = font_title

    # --- BLOK A (KIRI ATAS): DATA PARAMETER & JARAK PANDANG ---
    ws2.merge_cells("A3:C3")
    ws2["A3"] = "DATA KRITERIA PERENCANAAN"
    ws2["A3"].font = font_sec
    ws2["A3"].fill = fill_header
    ws2["A3"].alignment = align_center

    kriteria_rows = [
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam", "0", True),
        ("Waktu Reaksi (t)", 2.5, "detik", "0.0", True),
        ("Perlambatan (a)", 3.4, "m/det²", "0.0", True),
        ("Tinggi Mata Pengemudi (h1)", 1.05, "m", "0.00", True),
        ("Tinggi Halangan (h2)", 0.15, "m", "0.00", True),
        ("Jarak Pandang Henti (JPH / Jh)", "=0.278*B4*B5+(B4^2)/(254*0.4)", "m", "0.00", False),
        ("T1 (Waktu Awal Mendahului)", "=2.21+(0.026*B4)", "detik", "0.00", False),
        ("T2 (Waktu Mendahului)", "=6.56+(0.048*B4)", "detik", "0.00", False),
        ("Percepatan Mendahului (a_m)", "=2.052+(0.0036*B4)", "km/jam/s", "0.000", False),
        ("d1 (Jarak Reaksi)", "=0.278*B10*(B4-10+((B12*B10)/2))", "m", "0.00", False),
        ("d2 (Jarak Mendahului)", "=0.278*B4*B11", "m", "0.00", False),
        ("d3 (Jarak Bebas)", 30.0, "m", "0.00", True),
        ("d4 (Jarak Kendaraan Berlawanan)", "=(2/3)*B14", "m", "0.00", False),
        ("Jarak Pandang Mendahului (JPM)", "=B13+B14+B15+B16", "m", "0.00", False),
    ]

    for idx_k, (lbl, val, unit, num_fmt, is_inp) in enumerate(kriteria_rows, start=4):
        c1 = ws2.cell(row=idx_k, column=1, value=lbl)
        c2 = ws2.cell(row=idx_k, column=2, value=val)
        c3 = ws2.cell(row=idx_k, column=3, value=unit)
        c1.font = font_bold if not is_inp else font_reg
        c2.font = font_bold
        c3.font = font_reg
        c2.number_format = num_fmt
        c2.alignment = align_right
        c3.alignment = align_center
        if is_inp:
            c2.fill = fill_input
        else:
            c2.fill = fill_card
        for c in (c1, c2, c3):
            c.border = border_thin

    # --- BLOK B (KANAN ATAS): TABEL TITIK PVI, STASIONING & KELANDAIAN ---
    d_pvi = pvi_df.dropna().sort_values(pvi_df.columns[0]).reset_index(drop=True)
    n_pvi = len(d_pvi)

    ws2.merge_cells("E3:K3")
    ws2["E3"] = "DATA TITIK PVI, STASIONING & KELANDAIAN RENCANA"
    ws2["E3"].font = font_sec
    ws2["E3"].fill = fill_header
    ws2["E3"].alignment = align_center

    pvi_headers = [
        "Titik", "STA (m)", "Stasioning", "Jarak Antar Titik (m)",
        "Elevasi Rencana (m)", "Kelandaian g (%)", "Lv Manual (m)"
    ]
    for col_offset, h_txt in enumerate(pvi_headers):
        cell = ws2.cell(row=4, column=5 + col_offset, value=h_txt)
        cell.font = font_sub
        cell.fill = fill_sub
        cell.alignment = align_center
        cell.border = border_thin

    pvi_start_row = 5
    for idx_p in range(n_pvi):
        r_curr = pvi_start_row + idx_p
        sta_val = float(d_pvi.iloc[idx_p, 0])
        el_val = float(d_pvi.iloc[idx_p, 1])
        lv_man_val = float(d_pvi.iloc[idx_p, 2])
        titik_lbl = "Titik Awal (A)" if idx_p == 0 else ("Titik Akhir (B)" if idx_p == n_pvi - 1 else f"PPV {idx_p}")

        c_titik = ws2.cell(row=r_curr, column=5, value=titik_lbl)
        c_sta = ws2.cell(row=r_curr, column=6, value=sta_val)
        c_stastr = ws2.cell(row=r_curr, column=7, value=f'=INT(F{r_curr}/1000)&"+"&TEXT(MOD(F{r_curr},1000),"000.00")')

        if idx_p == 0:
            c_jarak = ws2.cell(row=r_curr, column=8, value=0.0)
        else:
            c_jarak = ws2.cell(row=r_curr, column=8, value=f"=F{r_curr}-F{r_curr-1}")

        c_el = ws2.cell(row=r_curr, column=9, value=el_val)

        if idx_p < n_pvi - 1:
            c_grad = ws2.cell(row=r_curr, column=10, value=f"=((I{r_curr+1}-I{r_curr})/H{r_curr+1})*100")
        else:
            c_grad = ws2.cell(row=r_curr, column=10, value="-")

        c_lvman = ws2.cell(row=r_curr, column=11, value=lv_man_val)

        c_sta.fill = fill_input
        c_el.fill = fill_input
        c_lvman.fill = fill_input

        c_sta.number_format = "0.00"
        c_jarak.number_format = "0.00"
        c_el.number_format = "0.000"
        if idx_p < n_pvi - 1:
            c_grad.number_format = "0.000"
        c_lvman.number_format = "0.0"

        for c in (c_titik, c_sta, c_stastr, c_jarak, c_el, c_grad, c_lvman):
            c.font = font_reg
            c.border = border_thin
            if c in (c_titik, c_stastr):
                c.alignment = align_center
            else:
                c.alignment = align_right

    # --- BLOK C (BAWAH): PERHITUNGAN LENGKUNG VERTIKAL PER PPV ---
    calc_start_row = max(19, pvi_start_row + n_pvi + 2)
    ws2.merge_cells(start_row=calc_start_row, start_column=1, end_row=calc_start_row, end_column=max(4, 2 + max(1, n_pvi - 2)))
    ws2.cell(row=calc_start_row, column=1, value="PERHITUNGAN LENGKUNG VERTIKAL SETIAP TITIK PPV").font = font_sec
    ws2.cell(row=calc_start_row, column=1).fill = fill_header
    ws2.cell(row=calc_start_row, column=1).alignment = align_center

    row_labels = [
        ("Titik PPV", None),
        ("Stasioning PPV", None),
        ("Jenis Lengkung Vertikal", None),
        ("Kelandaian Masuk g1 (%)", "0.000"),
        ("Kelandaian Keluar g2 (%)", "0.000"),
        ("Perbedaan Aljabar Kelandaian A (%)", "0.000"),
        ("1. Syarat Jarak Pandang Henti - Kasus L > Jh (m)", "0.00"),
        ("   Kontrol Syarat L > Jh", None),
        ("2. Syarat Jarak Pandang Henti - Kasus L < Jh (m)", "0.00"),
        ("   Kontrol Syarat L < Jh", None),
        ("3. Berdasarkan Syarat Drainase (Lv = 50 * A) (m)", "0.00"),
        ("4. Syarat Kenyamanan Mengemudi (Vr * 3 / 3.6) (m)", "0.00"),
        ("5. Syarat Jarak Pandang Mendah
