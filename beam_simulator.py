import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import io
import time
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ── NEW: Excel Export ──
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import BarChart, Reference, LineChart
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

# ── NEW: 3D ──
try:
    from mpl_toolkits.mplot3d import Axes3D
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    THREED_OK = True
except ImportError:
    THREED_OK = False

st.set_page_config(page_title="Beam Simulator Pro", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@700;900&family=Rajdhani:wght@400;600&display=swap');

    /* ── Base ── */
    .stApp {
        background: #0e1117;
        color: #e0e0e0;
        font-family: 'Rajdhani', sans-serif;
    }

    /* ── Sidebar ── */
    .stSidebar {
        background: rgba(15, 12, 30, 0.95) !important;
        border-right: 1px solid rgba(99,102,241,0.2) !important;
    }
    .stSidebar .stMarkdown h2 {
        background: linear-gradient(90deg, #00bfff, #a855f7);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        font-family: 'Orbitron', monospace; font-size: 16px !important;
    }
    .stSidebar .stMarkdown h3 {
        color: #00d4aa !important;
        border-left: 3px solid #00d4aa;
        padding-left: 8px;
        font-size: 13px !important;
    }

    /* ── Metric Cards ── */
    [data-testid="stMetric"] {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(99,102,241,0.25);
        border-radius: 16px;
        padding: 12px 16px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.4);
        transition: all 0.3s;
        animation: floatCard 4s ease-in-out infinite alternate;
    }
    @keyframes floatCard {
        0%   { transform: translateY(0px); }
        100% { transform: translateY(-4px); }
    }
    [data-testid="stMetricValue"] {
        color: #ffffff !important;
        font-size: 1.3rem !important;
        font-family: 'Orbitron', monospace !important;
    }
    [data-testid="stMetricLabel"] {
        color: #00bfff !important;
        font-weight: 600 !important;
        letter-spacing: 1px;
    }

    /* ── Headings ── */
    h1 {
        font-family: 'Orbitron', monospace !important;
        background: linear-gradient(90deg, #00bfff 0%, #a855f7 50%, #00d4aa 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        font-size: 2rem !important;
        letter-spacing: 2px;
    }
    h2, h3 {
        background: linear-gradient(90deg, #00d4aa, #00bfff);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }

    /* ── Feature Badges ── */
    .feature-badge {
        background: rgba(0,191,255,0.08);
        border: 1px solid #00bfff55;
        border-radius: 20px;
        padding: 5px 14px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block; margin: 3px;
        color: #00d4aa;
    }

    /* ── Safe / Unsafe boxes ── */
    .safe-box {
        background: rgba(0,200,83,0.1);
        border: 2px solid #00c853;
        border-radius: 14px; padding: 14px 22px;
        font-size: 16px; font-weight: bold; color: #00c853;
    }
    .unsafe-box {
        background: rgba(255,71,87,0.1);
        border: 2px solid #ff4757;
        border-radius: 14px; padding: 14px 22px;
        font-size: 16px; font-weight: bold; color: #ff4757;
    }

    /* ── AI Box ── */
    .ai-box {
        background: rgba(124,77,255,0.08);
        border: 1.5px solid #7c4dff55;
        border-radius: 14px; padding: 18px 22px; margin: 8px 0;
    }
    .ai-title { color: #a855f7; font-size: 16px; font-weight: bold; margin-bottom: 12px; }
    .ai-suggestion { color: #c8d8e8; font-size: 14px; padding: 4px 0; }
    .ai-warning { color: #ff6b6b; font-size: 14px; font-weight: bold; padding: 4px 0; }
    .ai-good { color: #00e676; font-size: 14px; padding: 4px 0; }

    /* ── Dynamic box ── */
    .dynamic-box {
        background: rgba(255,103,0,0.08);
        border: 1.5px solid #ff6700;
        border-radius: 12px; padding: 12px 18px; margin: 6px 0;
    }

    /* ── Buttons ── */
    .stButton > button, .stDownloadButton > button {
        background: rgba(13,33,55,0.9) !important;
        border: 1.5px solid #00bfff66 !important;
        color: #00bfff !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        transition: all 0.3s !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: rgba(0,191,255,0.15) !important;
        border-color: #00bfff !important;
    }

    /* ── Zoom Controls ── */
    .zoom-bar {
        position: fixed; bottom: 24px; right: 20px; z-index: 99999;
        display: flex; flex-direction: column; align-items: center; gap: 6px;
        background: rgba(5,15,32,0.95);
        border: 1.5px solid #00bfff44;
        border-radius: 16px; padding: 12px 10px;
        box-shadow: 0 0 30px rgba(0,191,255,0.15);
    }
    .zoom-btn {
        background: rgba(13,33,55,0.9);
        border: 1.5px solid #00bfff66;
        color: #00bfff; font-size: 22px; font-weight: bold;
        width: 46px; height: 46px; border-radius: 12px;
        cursor: pointer; display: flex; align-items: center; justify-content: center;
        transition: all 0.15s;
        -webkit-user-select: none; user-select: none;
        -webkit-tap-highlight-color: transparent;
        min-width: 46px; min-height: 46px;
    }
    .zoom-btn:hover  { background: rgba(0,191,255,0.2); border-color:#00bfff; }
    .zoom-btn:active { transform: scale(0.88); }
    .zoom-label {
        color: #00d4aa; font-size: 12px; font-weight: bold;
        font-family: 'Orbitron', monospace; text-align: center;
    }
    .zoom-reset {
        background: rgba(124,77,255,0.2);
        border: 1.5px solid #a855f766;
        color: #a855f7; font-size: 10px; font-weight: bold;
        width: 46px; height: 30px; border-radius: 10px;
        cursor: pointer; display: flex; align-items: center; justify-content: center;
        transition: all 0.15s;
        -webkit-user-select: none; user-select: none;
        -webkit-tap-highlight-color: transparent;
        letter-spacing: 1px;
    }
    .zoom-reset:hover  { background: rgba(168,85,247,0.3); }
    .zoom-reset:active { transform: scale(0.88); }
</style>

<div class="zoom-bar" id="zoomBar">
    <button class="zoom-btn" id="btnIn"  title="Zoom In">＋</button>
    <div class="zoom-label" id="zoom-pct">100%</div>
    <button class="zoom-btn" id="btnOut" title="Zoom Out">－</button>
    <button class="zoom-reset" id="btnReset">RESET</button>
</div>

<script>
(function(){
    var scale = 1.0;
    var MIN = 0.4, MAX = 2.8, STEP = 0.15;

    function getTarget() {
        return document.querySelector('[data-testid="stAppViewBlockContainer"]')
            || document.querySelector('.main .block-container')
            || document.querySelector('[data-testid="block-container"]')
            || document.querySelector('.block-container')
            || document.querySelector('.main section')
            || document.querySelector('.main');
    }

    function applyZoom() {
        var el = getTarget();
        if (el) {
            el.style.transformOrigin = 'top left';
            el.style.transform = 'scale(' + scale.toFixed(2) + ')';
            el.style.width = (100 / scale) + '%';
        }
        var lbl = document.getElementById('zoom-pct');
        if (lbl) lbl.textContent = Math.round(scale * 100) + '%';
        try { localStorage.setItem('bsZoom', scale.toFixed(2)); } catch(e){}
    }

    function zoomIn()    { scale = Math.min(MAX, parseFloat((scale + STEP).toFixed(2))); applyZoom(); }
    function zoomOut()   { scale = Math.max(MIN, parseFloat((scale - STEP).toFixed(2))); applyZoom(); }
    function zoomReset() { scale = 1.0; applyZoom(); }

    try {
        var s = parseFloat(localStorage.getItem('bsZoom'));
        if (s && s >= MIN && s <= MAX) scale = s;
    } catch(e){}

    function wire(id, fn) {
        var btn = document.getElementById(id);
        if (!btn) return;
        btn.addEventListener('touchstart', function(e){ e.preventDefault(); fn(); }, {passive:false});
        btn.addEventListener('click', function(e){ e.preventDefault(); fn(); });
    }

    document.addEventListener('keydown', function(e){
        if (e.ctrlKey || e.metaKey) {
            if (e.key === '=' || e.key === '+') { e.preventDefault(); zoomIn(); }
            if (e.key === '-')                  { e.preventDefault(); zoomOut(); }
            if (e.key === '0')                  { e.preventDefault(); zoomReset(); }
        }
    });

    function init() {
        wire('btnIn',    zoomIn);
        wire('btnOut',   zoomOut);
        wire('btnReset', zoomReset);
        applyZoom();
    }

    setTimeout(init, 800);
    setTimeout(applyZoom, 1500);
    setTimeout(applyZoom, 3000);
})();
</script>

<style>
/* ── Attractive Sidebar Colors ── */
.stSidebar {
    background: linear-gradient(180deg, #0a0015 0%, #0d0020 50%, #080018 100%) !important;
    border-right: 1px solid #6366f155 !important;
}

/* Selectbox */
.stSidebar [data-testid="stSelectbox"] > div > div {
    background: linear-gradient(135deg, #1a0a2e, #0d0a2e) !important;
    border: 1px solid #6366f166 !important;
    border-radius: 10px !important;
    color: #e0e0e0 !important;
}

/* Sliders — neon purple track */
.stSidebar [data-testid="stSlider"] > div > div > div > div {
    background: linear-gradient(90deg, #6366f1, #a855f7) !important;
}
.stSidebar [data-testid="stSlider"] > div > div > div > div > div {
    background: #ffffff !important;
    box-shadow: 0 0 8px #a855f7 !important;
}

/* Number input */
.stSidebar [data-testid="stNumberInput"] input {
    background: #1a0a2e !important;
    border: 1px solid #6366f166 !important;
    color: #e0e0e0 !important;
    border-radius: 8px !important;
}

/* Slider value text — neon green */
.stSidebar [data-testid="stSlider"] p {
    color: #00ffaa !important;
    font-weight: bold !important;
    font-size: 15px !important;
}

/* Section headers */
.stSidebar h2 {
    background: linear-gradient(90deg, #a855f7, #6366f1, #00bfff) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    font-size: 18px !important;
    font-weight: 900 !important;
}
.stSidebar h3 {
    color: #00ffaa !important;
    border-left: 3px solid #a855f7 !important;
    padding-left: 8px !important;
    font-size: 14px !important;
}

/* Labels */
.stSidebar label, .stSidebar p {
    color: #c8b4ff !important;
}

/* Main background subtle glow */
.stApp {
    background: linear-gradient(135deg, #06000f 0%, #0e0118 40%, #050010 100%) !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("# 🏗️ 2D Beam Stress & Deflection Simulator Pro")
st.markdown("""
<div style="margin: 8px 0 4px 0;">
    <span class="feature-badge">✨ 3D View</span>
    <span class="feature-badge">🏃 Dynamic Loads</span>
    <span class="feature-badge">📊 Excel Export</span>
    <span class="feature-badge">🧱 10 Materials</span>
    <span class="feature-badge">📐 Multi-Units</span>
    <span class="feature-badge">🔍 Zoom</span>
    <span class="feature-badge">🎨 Pro UI</span>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════
st.sidebar.markdown("## ⚙️ Beam Parameters")

# ── Unit System ──
st.sidebar.markdown("### 📐 Unit System")
length_unit = st.sidebar.selectbox("Length Unit", ["m", "cm", "mm"], index=0)
force_unit  = st.sidebar.selectbox("Force Unit",  ["N", "kN"], index=0)

# Conversion factors → always compute in SI (m, N)
L_CONV = {"m": 1.0, "cm": 0.01, "mm": 0.001}[length_unit]
F_CONV = {"N": 1.0, "kN": 1000.0}[force_unit]

# Max length in display unit
L_max_display = {"m": 100.0, "cm": 10000.0, "mm": 100000.0}[length_unit]
L_default     = {"m": 4.0,   "cm": 400.0,   "mm": 4000.0}[length_unit]
L_min         = {"m": 0.1,   "cm": 10.0,    "mm": 100.0}[length_unit]
L_step        = {"m": 0.1,   "cm": 10.0,    "mm": 100.0}[length_unit]

L_display = st.sidebar.slider(f"📏 Beam Length ({length_unit})", L_min, L_max_display, L_default, L_step)
L = L_display * L_CONV  # always in metres internally

# ── Supports — any count, any position, any type (overhangs allowed) ──
st.sidebar.markdown("### 🏛️ Supports")
st.sidebar.caption("Place supports anywhere along the beam — including inside the "
                    "span, to model overhangs like a cantilevered tip beyond the last support.")
SUPPORT_TYPES = ["Fixed", "Pinned", "Roller"]
SUPPORT_ICON  = {"Fixed": "🧱", "Pinned": "📍", "Roller": "⚪"}
n_supports = st.sidebar.number_input("Number of Supports", 1, 6, 2)

supports = []  # list of (position_m, type)
for i in range(int(n_supports)):
    st.sidebar.markdown(f"**Support {i+1}**")
    if i == 0:
        default_pos = 0.0
    elif i == 1:
        default_pos = L_display
    else:
        default_pos = round(L_display*i/(n_supports), 1)
    default_pos = min(max(default_pos, 0.0), L_display)
    pos_disp = st.sidebar.slider(f"Position {i+1} ({length_unit})", 0.0, L_display, default_pos, L_step, key=f"sup_pos{i}")
    default_type_idx = 0 if (i == 0 and n_supports == 1) else 1
    typ = st.sidebar.selectbox(f"Type {i+1}", SUPPORT_TYPES, index=default_type_idx,
                                format_func=lambda s: f"{SUPPORT_ICON[s]} {s}", key=f"sup_type{i}")
    supports.append((pos_disp * L_CONV, typ))

def _describe_beam(sups, L_):
    s_ = sorted(sups)
    at_l = [t for p, t in s_ if p <= 1e-9]
    at_r = [t for p, t in s_ if p >= L_ - 1e-9]
    if len(s_) == 1 and s_[0][1] == "Fixed" and (at_l or at_r):
        return "Cantilever"
    if len(s_) == 2 and at_l and at_r:
        if at_l[0] == "Fixed" and at_r[0] == "Fixed": return "Fixed-Fixed"
        if "Fixed" in (at_l[0], at_r[0]):             return "Propped Cantilever"
        return "Simply Supported"
    if len(s_) == 2: return "Overhanging Beam"
    if len(s_) >= 3: return "Continuous Beam"
    return "Custom Beam"

beam_type = _describe_beam(supports, L)
support_desc = ", ".join(f"{t} @ {p/L_CONV:g}{length_unit}" for p, t in sorted(supports))
st.sidebar.caption(f"➡️ Configuration: **{beam_type}**")

st.sidebar.markdown("### ⬇️ Point Loads")
n_loads = st.sidebar.number_input("Number of Point Loads", 0, 8, 1)

# Force slider max in display unit
F_max = {"N": 10000, "kN": 100}[force_unit]
F_def = {"N": 500,   "kN": 5}[force_unit]
F_step= {"N": 100,   "kN": 1}[force_unit]

loads = []
for i in range(int(n_loads)):
    st.sidebar.markdown(f"**Load {i+1}**")
    p_disp = st.sidebar.slider(f"P{i+1} ({force_unit})", 0, F_max, F_def, F_step, key=f"p{i}")
    default_pos = round(L_display/(n_loads+1)*(i+1), 1)
    default_pos = min(default_pos, L_display)
    a_disp = st.sidebar.slider(f"Position {i+1} ({length_unit})", 0.0, L_display, default_pos, L_step, key=f"a{i}")
    loads.append((p_disp * F_CONV, a_disp * L_CONV))  # store in SI

# ── Applied Moments (couples) — magnitude, position, direction ──
st.sidebar.markdown("### 🔄 Applied Moments")
n_moments = st.sidebar.number_input("Number of Applied Moments", 0, 5, 0)
M_max_disp = {"N": 20000, "kN": 200}[force_unit]  # displayed as force_unit·length_unit
M_step_disp= {"N": 100,   "kN": 1}[force_unit]

moments = []  # list of (signed_M0_SI, position_m) — CW is defined as +M0 internally
for i in range(int(n_moments)):
    st.sidebar.markdown(f"**Moment {i+1}**")
    m_disp = st.sidebar.slider(f"M{i+1} ({force_unit}·{length_unit})", 0, M_max_disp, M_step_disp*5, M_step_disp, key=f"m{i}")
    default_pos = round(L_display/(n_moments+1)*(i+1), 1)
    default_pos = min(default_pos, L_display)
    ma_disp = st.sidebar.slider(f"Moment Position {i+1} ({length_unit})", 0.0, L_display, default_pos, L_step, key=f"ma{i}")
    direction = st.sidebar.selectbox(f"Direction {i+1}", ["Clockwise ↻", "Counter-clockwise ↺"], key=f"mdir{i}")
    m_si = m_disp * F_CONV * L_CONV
    signed = m_si if direction.startswith("Clockwise") else -m_si
    moments.append((signed, ma_disp * L_CONV))

# ── Distributed Load — Uniform or Trapezoidal/Triangular (linearly varying) ──
st.sidebar.markdown("### 〰️ Distributed Load")
udl_unit_label = f"{force_unit}/{length_unit}"
w_max  = {"N": 2000, "kN": 20}[force_unit]
w_step = {"N": 50,   "kN": 1}[force_unit]
udl_shape = st.sidebar.selectbox("Load Shape", ["None", "Uniform (Constant)", "Trapezoidal / Triangular (Linear)"])

if udl_shape == "None":
    w1_disp, w2_disp, udl_x1_disp, udl_x2_disp = 0, 0, 0.0, L_display
elif udl_shape == "Uniform (Constant)":
    w1_disp = st.sidebar.slider(f"Intensity ({udl_unit_label})", 0, w_max, 0, w_step)
    w2_disp = w1_disp
    udl_x1_disp, udl_x2_disp = st.sidebar.slider(
        f"Span ({length_unit}) — start to end", 0.0, L_display, (0.0, L_display), L_step)
else:
    w1_disp = st.sidebar.slider(f"Start Intensity w1 ({udl_unit_label})", 0, w_max, 0, w_step)
    w2_disp = st.sidebar.slider(f"End Intensity w2 ({udl_unit_label})", 0, w_max, w_max//2, w_step)
    udl_x1_disp, udl_x2_disp = st.sidebar.slider(
        f"Span ({length_unit}) — start to end", 0.0, L_display, (0.0, L_display), L_step)

w1 = w1_disp * F_CONV / L_CONV
w2 = w2_disp * F_CONV / L_CONV
udl_x1 = udl_x1_disp * L_CONV
udl_x2 = udl_x2_disp * L_CONV
has_udl = udl_shape != "None" and (w1 > 0 or w2 > 0) and udl_x2 > udl_x1
if not has_udl:
    w1 = w2 = 0.0
    udl_x1, udl_x2 = 0.0, L_display
w = max(w1, w2)  # peak intensity (N/m) — used for "is there a distributed load?" checks and labels
w_peak_disp = max(w1_disp, w2_disp)
if not has_udl:
    udl_desc = "None"
elif abs(w1 - w2) < 1e-12:
    udl_desc = f"Uniform {w1_disp:g} {udl_unit_label} over {udl_x1_disp:g}-{udl_x2_disp:g} {length_unit}"
else:
    udl_desc = f"Linear {w1_disp:g} → {w2_disp:g} {udl_unit_label} over {udl_x1_disp:g}-{udl_x2_disp:g} {length_unit}"
moment_desc = " | ".join(
    f"M{i+1}={abs(m)/(F_CONV*L_CONV):g} {force_unit}·{length_unit} {'CW' if m > 0 else 'CCW'} @ {a/L_CONV:g}{length_unit}"
    for i, (m, a) in enumerate(moments) if m != 0) or "None"

# ── NEW: Dynamic/Moving Load ──
st.sidebar.markdown("### 🚗 Dynamic / Moving Load")
enable_dynamic = st.sidebar.checkbox("Enable Moving Load Analysis", False)
if enable_dynamic:
    P_moving = st.sidebar.slider("Moving Load (N)", 100, 5000, 1000, 100)
    dynamic_positions = 20

# ── NEW: Extended Material Library ──
material_data = {
    "🔩 Structural Steel (200 GPa)":    {"E": 200e9, "yield": 250e6,  "density": 7850, "color": "#4fc3f7"},
    "🔩 High-Strength Steel (200 GPa)": {"E": 200e9, "yield": 690e6,  "density": 7850, "color": "#0288d1"},
    "✈️ Aluminum 6061 (69 GPa)":        {"E": 69e9,  "yield": 276e6,  "density": 2700, "color": "#b0bec5"},
    "✈️ Aluminum 7075 (72 GPa)":        {"E": 72e9,  "yield": 503e6,  "density": 2810, "color": "#90a4ae"},
    "🌳 Oak Wood (12 GPa)":             {"E": 12e9,  "yield": 40e6,   "density": 700,  "color": "#8d6e63"},
    "🌲 Pine Wood (9 GPa)":             {"E": 9e9,   "yield": 33e6,   "density": 550,  "color": "#a1887f"},
    "🏗️ Concrete (30 GPa)":            {"E": 30e9,  "yield": 30e6,   "density": 2400, "color": "#9e9e9e"},
    "⚡ Copper (110 GPa)":             {"E": 110e9, "yield": 210e6,  "density": 8900, "color": "#ff8f00"},
    "🔮 Carbon Fiber (150 GPa)":        {"E": 150e9, "yield": 600e6,  "density": 1600, "color": "#7c4dff"},
    "🪨 Titanium (116 GPa)":           {"E": 116e9, "yield": 880e6,  "density": 4500, "color": "#80cbc4"},
}
E_sel = st.sidebar.selectbox("🧱 Material", list(material_data.keys()))
E_val    = material_data[E_sel]["E"]
yield_s  = material_data[E_sel]["yield"]
density  = material_data[E_sel]["density"]
mat_color = material_data[E_sel]["color"]

st.sidebar.markdown("### 📐 Cross Section Type")
section_type = st.sidebar.selectbox("Section Shape", ["Rectangular","Circular","I-Beam","T-Beam"])

def get_section_properties(section_type, sidebar):
    if section_type == "Rectangular":
        b = sidebar.number_input("Width b (m)",  0.01, 0.5, 0.05)
        h = sidebar.number_input("Height h (m)", 0.01, 0.5, 0.10)
        return (b*h**3)/12, h/2, h, b, {"b":b,"h":h}
    elif section_type == "Circular":
        d = sidebar.number_input("Diameter d (m)", 0.01, 0.5, 0.08)
        return (np.pi*d**4)/64, d/2, d, d, {"d":d}
    elif section_type == "I-Beam":
        sidebar.markdown("*Flange & Web*")
        bf = sidebar.number_input("Flange Width bf (m)", 0.02, 0.4, 0.10)
        tf = sidebar.number_input("Flange Thick tf (m)", 0.005, 0.05, 0.01)
        hw = sidebar.number_input("Web Height hw (m)",   0.05, 0.4,  0.10)
        tw = sidebar.number_input("Web Thick tw (m)",    0.005, 0.05, 0.006)
        ht = hw+2*tf
        I  = (bf*ht**3)/12 - ((bf-tw)*hw**3)/12
        return I, ht/2, ht, bf, {"bf":bf,"tf":tf,"hw":hw,"tw":tw}
    elif section_type == "T-Beam":
        sidebar.markdown("*Flange & Web*")
        bf = sidebar.number_input("Flange Width bf (m)", 0.02, 0.4, 0.10)
        tf = sidebar.number_input("Flange Thick tf (m)", 0.005, 0.05, 0.01)
        hw = sidebar.number_input("Web Height hw (m)",   0.05, 0.4,  0.10)
        tw = sidebar.number_input("Web Thick tw (m)",    0.005, 0.05, 0.006)
        ht = hw+tf
        Af = bf*tf; Aw = tw*hw; At = Af+Aw
        yb = (Af*(ht-tf/2)+Aw*(hw/2))/At
        I  = (bf*tf**3)/12+Af*(ht-tf/2-yb)**2+(tw*hw**3)/12+Aw*(hw/2-yb)**2
        return I, max(yb,ht-yb), ht, bf, {"bf":bf,"tf":tf,"hw":hw,"tw":tw,"y_bar":yb}

I, c_dist, h_total, b_total, dims = get_section_properties(section_type, st.sidebar)

# ════════════════════════════════════════════
# GENERAL BEAM SOLVER — any number of supports at any position
# (Fixed/Pinned/Roller), point loads, applied moments (CW/CCW), and a
# uniform OR linearly-varying (trapezoidal/triangular) distributed load
# over any custom [x1,x2] span.
#
# Uses a 2-node Euler-Bernoulli beam finite-element (exact for point
# loads/moments placed at nodes + consistent-load distributed loads),
# meshed only at support/load/moment positions — this is exact, not an
# approximation, and naturally handles determinate AND indeterminate
# beams (propped cantilevers, overhangs, multi-span beams) without a
# separate closed-form formula for every configuration. Stability is
# checked via the reduced stiffness matrix's rank (a true mechanism
# leaves it singular) rather than by hand-counting support types.
# ════════════════════════════════════════════
from scipy.integrate import cumulative_trapezoid

x = np.linspace(0, L, 2000)
colors_load = ['#ff4757','#ffd700','#ff6b35','#7c4dff','#00d4aa','#00bfff','#ff69b4','#adff2f']

def _elem_stiffness(EI, Le):
    return EI/Le**3 * np.array([
        [12,     6*Le,   -12,     6*Le],
        [6*Le,  4*Le**2, -6*Le,  2*Le**2],
        [-12,   -6*Le,    12,    -6*Le],
        [6*Le,  2*Le**2, -6*Le,  4*Le**2]])

def _elem_trap_load(wa, wb, Le):
    """Consistent nodal load vector for a linearly-varying load: wa at the
    left node, wb at the right node (wa==wb reduces to a uniform load)."""
    uni = wa*Le/12 * np.array([6, Le, 6, -Le])
    dw  = wb - wa
    tri = np.array([3*Le*dw/20, Le**2*dw/30, 7*Le*dw/20, -Le**2*dw/20])
    return uni + tri

def solve_beam(supports, loads, moments, udl, L, E, I):
    """
    supports: list of (position_m, type)   type in {Fixed, Pinned, Roller}
    loads:    list of (P, a)               downward point loads
    moments:  list of (M0_signed, a)       +M0 = Clockwise, -M0 = Counter-clockwise
    udl:      (w1, w2, x1, x2) or None     linearly varying w1→w2 over [x1,x2]
    Returns: reactions = [(pos, Rv, Mreact, type), ...], nodes, d
    Raises ValueError if the support configuration is a mechanism (unstable).
    """
    EI = E*I
    pts = {0.0, round(L, 9)}
    for pos, _ in supports: pts.add(round(min(max(pos, 0), L), 9))
    for P, a in loads:
        if P != 0: pts.add(round(min(max(a, 0), L), 9))
    for M0, a in moments:
        if M0 != 0: pts.add(round(min(max(a, 0), L), 9))
    if udl is not None:
        w1u, w2u, x1u, x2u = udl
        if (w1u != 0 or w2u != 0) and x2u > x1u:
            pts.add(round(min(max(x1u, 0), L), 9)); pts.add(round(min(max(x2u, 0), L), 9))
    nodes = sorted(pts)
    n = len(nodes)
    ndof = 2*n
    K = np.zeros((ndof, ndof))
    F = np.zeros(ndof)

    for e in range(n-1):
        Le = nodes[e+1] - nodes[e]
        if Le <= 1e-9:
            continue
        ke = _elem_stiffness(EI, Le)
        dofs = [2*e, 2*e+1, 2*e+2, 2*e+3]
        for i in range(4):
            for j in range(4):
                K[dofs[i], dofs[j]] += ke[i, j]
        if udl is not None:
            w1u, w2u, x1u, x2u = udl
            if (w1u != 0 or w2u != 0) and x2u > x1u and nodes[e] >= x1u-1e-9 and nodes[e+1] <= x2u+1e-9:
                k_ = (w2u-w1u)/(x2u-x1u)
                wa = w1u + k_*(nodes[e]-x1u)
                wb = w1u + k_*(nodes[e+1]-x1u)
                fe = _elem_trap_load(wa, wb, Le)
                for i in range(4):
                    F[dofs[i]] += fe[i]

    for P, a in loads:
        if P != 0:
            idx = nodes.index(round(min(max(a, 0), L), 9))
            F[2*idx] += P
    for M0, a in moments:
        if M0 != 0:
            idx = nodes.index(round(min(max(a, 0), L), 9))
            F[2*idx+1] += M0

    fixed_dofs = []
    sup_idx = []
    for pos, typ in supports:
        idx = nodes.index(round(min(max(pos, 0), L), 9))
        sup_idx.append((idx, typ))
        fixed_dofs += ([2*idx, 2*idx+1] if typ == "Fixed" else [2*idx])

    free_dofs = [d for d in range(ndof) if d not in fixed_dofs]
    if len(free_dofs) > 0:
        Kff = K[np.ix_(free_dofs, free_dofs)]
        if np.linalg.matrix_rank(Kff) < Kff.shape[0]:
            raise ValueError("mechanism")
        d_free = np.linalg.solve(Kff, F[free_dofs])
    else:
        d_free = np.zeros(0)   # fully restrained (e.g. Fixed-Fixed, no interior nodes)
    d = np.zeros(ndof)
    for i, dof in enumerate(free_dofs):
        d[dof] = d_free[i]
    R = K @ d - F  # in the (downward-positive) FE sign convention

    reactions = []
    for idx, typ in sup_idx:
        Rv = -R[2*idx]
        Mr = -R[2*idx+1] if typ == "Fixed" else 0.0
        reactions.append((nodes[idx], Rv, Mr, typ))
    return reactions, nodes, d

def beam_V_M(x_arr, reactions, loads, moments, udl):
    """Direct-statics V(x), M(x) from known reactions/loads/moments — exact,
    continuous, and valid for any number/position of supports."""
    V = np.zeros_like(x_arr); M = np.zeros_like(x_arr)
    for pos, Rv, Mr, typ in reactions:
        mask = (x_arr >= pos - 1e-9)
        V += Rv*mask
        M += Rv*(x_arr-pos)*mask - Mr*mask
    for P, a in loads:
        mask = (x_arr >= a - 1e-9)
        V -= P*mask
        M -= P*np.maximum(x_arr-a, 0)
    for M0, a in moments:
        M += M0*(x_arr >= a - 1e-9)
    if udl is not None:
        w1u, w2u, x1u, x2u = udl
        if (w1u != 0 or w2u != 0) and x2u > x1u:
            k_ = (w2u-w1u)/(x2u-x1u)
            xu = np.clip(x_arr, x1u, x2u)
            u  = xu - x1u
            A  = x_arr - x1u
            Fq    = w1u*u + k_*u**2/2
            Marm  = w1u*A*u - w1u*u**2/2 + k_*A*u**2/2 - k_*u**3/3
            V -= Fq
            M -= Marm
    return V, M

def beam_deflection(x_arr, reactions, M_arr, E, I):
    """Numerically double-integrate M/EI; the 2 integration constants are
    fixed from whichever v=0 / theta=0 conditions the supports impose —
    solved by least squares since redundant supports give more than 2
    (but fully consistent) conditions."""
    phi = cumulative_trapezoid(-M_arr/(E*I), x_arr, initial=0)
    Y   = cumulative_trapezoid(phi, x_arr, initial=0)
    rows, rhs = [], []
    for pos, Rv, Mr, typ in reactions:
        Yi = np.interp(pos, x_arr, Y); phi_i = np.interp(pos, x_arr, phi)
        rows.append([1, pos]); rhs.append(-Yi)
        if typ == "Fixed":
            rows.append([0, 1]); rhs.append(-phi_i)
    y0, theta0 = np.linalg.lstsq(np.array(rows, dtype=float), np.array(rhs, dtype=float), rcond=None)[0]
    return y0 + theta0*x_arr + Y

udl_params = (w1, w2, udl_x1, udl_x2) if has_udl else None
try:
    reactions, _nodes, _d = solve_beam(supports, loads, moments, udl_params, L, E_val, I)
except ValueError:
    st.error("⚠️ This support configuration is a **mechanism** (unstable — not "
              "enough restraint, or all supports act along one line without "
              "resisting rotation). Add another support or change a support type.")
    st.stop()

V, M = beam_V_M(x, reactions, loads, moments, udl_params)
y = beam_deflection(x, reactions, M, E_val, I)

RA = reactions[0][1]   # kept for any leftover code that expects a single "RA"/"RB" pair
RB = reactions[-1][1] if len(reactions) > 1 else 0.0

M_max     = max(abs(M))
y_max     = max(abs(y))*1000
sigma_max = (M_max*c_dist)/I/1e6
yield_MPa = yield_s/1e6
FOS       = yield_MPa/sigma_max if sigma_max > 0 else 999

# ── Weight Calculation ──
if section_type == "Rectangular":
    area = dims['b']*dims['h']
elif section_type == "Circular":
    area = np.pi*(dims['d']/2)**2
elif section_type in ["I-Beam","T-Beam"]:
    area = I / (c_dist**2) * 1.2  # approximate
else:
    area = 0.01
beam_weight = density * area * L * 9.81  # N

# ════════════════════════════════════════════
# AI SUGGESTIONS
# ════════════════════════════════════════════
def ai_suggestions(beam_type,section_type,L,loads,w,E,dims,
                   FOS,sigma_max,yield_MPa,y_max,M_max,I,h_total):
    suggestions=[]; warnings=[]; good=[]
    if FOS < 1.0:
        warnings.append("🚨 CRITICAL: Beam will fail immediately! Redesign required.")
    elif FOS < 1.5:
        warnings.append(f"⚠️ FOS dangerously low ({FOS:.2f}). Minimum recommended: 2.0")
    elif FOS < 2.0:
        warnings.append(f"⚠️ FOS ({FOS:.2f}) below recommended minimum of 2.0")
    else:
        good.append(f"✅ Factor of Safety is adequate ({FOS:.2f} ≥ 2.0)")
    stress_ratio = sigma_max/yield_MPa if yield_MPa > 0 else 0
    if stress_ratio > 0.9:
        warnings.append(f"🔴 Max stress ({sigma_max:.1f} MPa) is {stress_ratio*100:.0f}% of yield — critical!")
    elif stress_ratio > 0.6:
        warnings.append(f"🟡 Max stress is {stress_ratio*100:.0f}% of yield — moderate risk")
    else:
        good.append(f"✅ Stress level is safe ({stress_ratio*100:.0f}% of yield strength)")
    if y_max > 0:
        defl_ratio = (y_max/1000)/L
        if defl_ratio > 1/200:
            warnings.append(f"⚠️ Excessive deflection: L/{int(L*1000/y_max)} (recommended ≤ L/200)")
            suggestions.append("💡 Increase section height to reduce deflection")
        else:
            good.append(f"✅ Deflection within limit: L/{int(L*1000/y_max)}")
    if section_type == "Rectangular":
        h = dims.get('h',0.1)
        if FOS < 2.0:
            new_h = h*(2.0/FOS)**0.5
            suggestions.append(f"💡 Increase height from {h*1000:.0f}mm to {new_h*1000:.0f}mm for FOS ≥ 2.0")
    elif section_type == "Circular":
        d = dims.get('d',0.08)
        if FOS < 2.0:
            new_d = d*(2.0/FOS)**0.333
            suggestions.append(f"💡 Increase diameter from {d*1000:.0f}mm to {new_d*1000:.0f}mm for FOS ≥ 2.0")
    elif section_type == "I-Beam":
        good.append("✅ I-Beam is efficient for bending — good section choice!")
    elif section_type == "T-Beam":
        good.append("✅ T-Beam is good for one-directional bending")
    if "Steel" in E and FOS > 10:
        suggestions.append("💡 FOS very high — consider Aluminum to reduce weight by ~65%")
    if "Aluminum" in E and FOS < 2.0:
        suggestions.append("💡 Switch to Steel for 3x higher yield strength")
    if "Wood" in E or "Pine" in E or "Oak" in E:
        suggestions.append("💡 Wood is orthotropic — verify grain direction matches load direction")
    if "Concrete" in E:
        warnings.append("⚠️ Concrete is weak in tension — add steel reinforcement (rebar)!")
    if "Carbon Fiber" in E:
        good.append("✅ Carbon fiber: excellent strength-to-weight ratio for aerospace/high-perf apps")
    if beam_type == "Simply Supported" and FOS < 2.0:
        suggestions.append("💡 Consider fixing one or both ends, or adding an intermediate support — reduces max moment significantly")
    if beam_type == "Cantilever" and FOS < 2.0:
        suggestions.append("💡 Cantilever has high moment at the fixed end — add a support near the free end (propped cantilever)")
    if len(loads)==1 and loads[0][0]>0:
        if abs(loads[0][1]-L/2) < 0.1:
            good.append("✅ Load at center — symmetric loading condition")
    if FOS >= 3.0 and stress_ratio < 0.4:
        good.append("🏆 Excellent design! Well within all safety limits.")
    elif FOS >= 2.0:
        good.append("👍 Good design — meets standard engineering requirements.")
    return warnings, suggestions, good

# ════════════════════════════════════════════
# BEAM VISUAL
# ════════════════════════════════════════════
st.markdown("### 🔎 Beam Diagram")
fig_beam, ax_b = plt.subplots(figsize=(10, 2.8))
fig_beam.patch.set_facecolor('#0a1628')
ax_b.set_facecolor('#0d2137')
ax_b.set_xlim(-0.8, L+0.8); ax_b.set_ylim(-1.2, 2.8); ax_b.axis('off')
ax_b.add_patch(plt.Rectangle((0,0.2),L,0.3,color='#1b3a5c',zorder=2))
ax_b.add_patch(plt.Rectangle((0,0.2),L,0.3,fill=False,edgecolor='#00bfff',lw=2,zorder=3))
def _draw_support(ax, kind, x_pos):
    if kind == "Fixed":
        rx = x_pos-0.3 if x_pos <= 1e-9 else (x_pos if x_pos >= L-1e-9 else x_pos-0.15)
        ax.add_patch(plt.Rectangle((rx,-0.1),0.3,0.9,color='#00d4aa',zorder=4))
    elif kind == "Pinned":
        ax.plot([x_pos],[0.2],'^',color='#00d4aa',markersize=14,zorder=4)
    elif kind == "Roller":
        ax.plot([x_pos],[0.2],'o',color='#00d4aa',markersize=10,zorder=4)
    ax.text(x_pos,-0.3,f"{kind}\n{x_pos/L_CONV:g}{length_unit}",color='#00d4aa',ha='center',va='top',fontsize=7)

for _sp, _st in supports:
    _draw_support(ax_b, _st, _sp)

for i,(P,a) in enumerate(loads):
    if P > 0:
        p_disp_label = P / F_CONV
        ax_b.annotate('',xy=(a,0.5),xytext=(a,1.6),
            arrowprops=dict(arrowstyle='->',color=colors_load[i],lw=2.5))
        ax_b.text(a,1.75,f'P{i+1}={p_disp_label:.0f}{force_unit}',color=colors_load[i],ha='center',fontsize=8,fontweight='bold')

# Applied moments — curved arrow, direction shown (CW / CCW)
_hw = 0.035*(L+1.6)
for i,(M0,a) in enumerate(moments):
    if M0 == 0: continue
    cw = M0 > 0
    pA, pB = ((a-_hw,0.62),(a+_hw,0.62)) if cw else ((a+_hw,0.62),(a-_hw,0.62))
    ax_b.add_patch(mpatches.FancyArrowPatch(pA, pB, connectionstyle=f"arc3,rad={-1.6 if cw else 1.6}",
                   arrowstyle='-|>', mutation_scale=14, color='#ff69b4', lw=2.2, zorder=5))
    ax_b.text(a,2.35,f"M{i+1}={abs(M0)/(F_CONV*L_CONV):g} {force_unit}·{length_unit} {'↻' if cw else '↺'}",
              color='#ff69b4',ha='center',fontsize=8,fontweight='bold')

# Distributed load — uniform or linearly varying (arrow length ∝ intensity)
if has_udl:
    _wmax = max(w1, w2)
    _h = lambda q: 0.5 + 0.65*q/_wmax
    for xi in np.linspace(udl_x1, udl_x2, max(4, int(14*(udl_x2-udl_x1)/L)+1)):
        q = w1 + (w2-w1)*(xi-udl_x1)/(udl_x2-udl_x1)
        if _h(q) - 0.5 > 0.03:
            ax_b.annotate('',xy=(xi,0.5),xytext=(xi,_h(q)),
                arrowprops=dict(arrowstyle='->',color='#00bfff',lw=1))
    ax_b.plot([udl_x1,udl_x2],[_h(w1),_h(w2)],color='#00bfff',lw=2)
    ax_b.text(-0.7,2.68,f'Distributed load: {udl_desc}',color='#00bfff',ha='left',va='center',fontsize=8)
ax_b.annotate('',xy=(L,-0.7),xytext=(0,-0.7),
    arrowprops=dict(arrowstyle='<->',color='gray',lw=1))
ax_b.text(L/2,-1.0,f'L = {L_display}{length_unit}',color='gray',ha='center',fontsize=8)
st.pyplot(fig_beam)
st.markdown("---")

# ════════════════════════════════════════════
# SAFETY CHECK
# ════════════════════════════════════════════
st.markdown("### 🛡️ Safety Check")
if sigma_max < yield_MPa:
    st.markdown(f"""<div class='safe-box'>✅ SAFE &nbsp;|&nbsp; σ_max: {sigma_max:.2f} MPa &nbsp;<&nbsp; Yield: {yield_MPa:.0f} MPa &nbsp;|&nbsp; FOS: {FOS:.2f} &nbsp;|&nbsp; Weight: {beam_weight:.1f} N</div>""", unsafe_allow_html=True)
else:
    st.markdown(f"""<div class='unsafe-box'>❌ UNSAFE &nbsp;|&nbsp; σ_max: {sigma_max:.2f} MPa &nbsp;>&nbsp; Yield: {yield_MPa:.0f} MPa &nbsp;|&nbsp; BEAM WILL FAIL!</div>""", unsafe_allow_html=True)
st.markdown("---")

# ════════════════════════════════════════════
# RESULTS
# ════════════════════════════════════════════
st.markdown("### 📊 Results")
c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
c1.metric("R_max", f"{max(abs(r[1]) for r in reactions):.1f} N")
c2.metric("Supports", f"{len(reactions)}")
c3.metric("M_max", f"{M_max:.1f} N·m")
c4.metric("y_max", f"{y_max:.3f} mm")
c5.metric("σ_max", f"{sigma_max:.2f} MPa")
c6.metric("FOS",   f"{FOS:.2f}")
c7.metric("I",     f"{I:.2e} m⁴")
c8.metric("Weight",f"{beam_weight:.0f} N")

st.markdown("**Support Reactions** (upward +; fixed-end moment: counter-clockwise +)")
_rc = st.columns(min(len(reactions), 6))
for _i, (_pos, _Rv, _Mr, _typ) in enumerate(reactions):
    _txt = f"{_Rv:.1f} N" + (f" | M={_Mr:.1f} N·m" if _typ == "Fixed" else "")
    _rc[_i % len(_rc)].metric(f"R{_i+1} — {_typ} @ {_pos/L_CONV:g}{length_unit}", _txt)
st.markdown("---")

# ════════════════════════════════════════════
# AI SUGGESTIONS
# ════════════════════════════════════════════
st.markdown("### 🤖 AI Engineering Suggestions")
warnings, suggestions, good = ai_suggestions(
    beam_type,section_type,L,loads,w,E_sel,dims,
    FOS,sigma_max,yield_MPa,y_max,M_max,I,h_total)
ai_html = "<div class='ai-box'><div class='ai-title'>🤖 Smart Engineering Analysis</div>"
for warn in warnings: ai_html += f"<div class='ai-warning'>{warn}</div>"
for sugg in suggestions: ai_html += f"<div class='ai-suggestion'>{sugg}</div>"
for g in good: ai_html += f"<div class='ai-good'>{g}</div>"
ai_html += "</div>"
st.markdown(ai_html, unsafe_allow_html=True)
st.markdown("---")

# ════════════════════════════════════════════
# PLOTS (SFD, BMD, Deflection)
# ════════════════════════════════════════════
st.markdown("### 📈 Engineering Diagrams")
fig,(ax1,ax2,ax3) = plt.subplots(3,1,figsize=(10,9))
fig.patch.set_facecolor('#0a1628')
for ax in [ax1,ax2,ax3]:
    ax.set_facecolor('#0d2137'); ax.tick_params(colors='#aaaaaa')
    ax.spines[:].set_color('#1b3a5c')
    ax.yaxis.label.set_color('#aaaaaa'); ax.xaxis.label.set_color('#aaaaaa')
    ax.title.set_color('#00bfff'); ax.grid(True,color='#1b3a5c',linewidth=0.7)
fig.suptitle(f"{beam_type} | {section_type} | L={L_display}{length_unit} | Dist. load: {udl_desc}",
             fontsize=11,fontweight='bold',color='white')
ax1.plot(x,V,color='#00bfff',lw=2); ax1.fill_between(x,V,alpha=0.25,color='#00bfff')
ax1.axhline(0,color='white',lw=0.8)
for i,(P,a) in enumerate(loads):
    ax1.axvline(a,color=colors_load[i],lw=1,linestyle='--',alpha=0.7,label=f'P{i+1}@{a/L_CONV:.1f}{length_unit}')
ax1.set_ylabel(f"Shear Force ({force_unit})"); ax1.set_title("SFD")
ax1.legend(facecolor='#0d2137',labelcolor='white',fontsize=7)
ax2.plot(x,M,color='#ff6b35',lw=2); ax2.fill_between(x,M,alpha=0.25,color='#ff6b35')
ax2.axhline(0,color='white',lw=0.8)
ax2.set_ylabel(f"Bending Moment ({force_unit}·{length_unit})"); ax2.set_title("BMD")
ax3.plot(x,y*1000,color='#00d4aa',lw=2); ax3.fill_between(x,y*1000,alpha=0.25,color='#00d4aa')
ax3.axhline(0,color='white',lw=0.8)
ax3.set_xlabel(f"Position ({length_unit})"); ax3.set_ylabel("Deflection (mm)"); ax3.set_title("Deflection Curve")
plt.tight_layout()
st.pyplot(fig)

# ════════════════════════════════════════════
# NEW FEATURE 1: 3D BEAM VISUALIZATION
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 🎯 3D Beam Visualization")

def plot_3d_beam():
    fig3d = plt.figure(figsize=(12, 5))
    fig3d.patch.set_facecolor('#0a1628')
    
    ax3d = fig3d.add_subplot(121, projection='3d')
    ax3d.set_facecolor('#0d2137')
    
    # Beam geometry
    n_x = 50
    x3 = np.linspace(0, L, n_x)
    
    # Cross-section width and height
    if section_type == "Rectangular":
        bw, bh = dims['b'], dims['h']
    elif section_type == "Circular":
        bw, bh = dims['d'], dims['d']
    else:
        bw = b_total
        bh = h_total
    
    # Deflection along beam
    y_defl = np.interp(x3, x, y)
    
    # Draw beam as 3D box with deflection
    X = np.array([x3, x3])
    Z = np.array([np.zeros(n_x)-bh/2, np.zeros(n_x)+bh/2])
    Y_top = np.array([y_defl, y_defl]) - bh/2
    Y_bot = np.array([y_defl, y_defl]) + bh/2
    
    # Top face
    ax3d.plot_surface(X, np.array([[-bw/2]*n_x, [bw/2]*n_x]),
                      np.array([y_defl, y_defl])*1000 + bh/2*500,
                      alpha=0.6, color=mat_color)
    # Bottom face
    ax3d.plot_surface(X, np.array([[-bw/2]*n_x, [bw/2]*n_x]),
                      np.array([y_defl, y_defl])*1000 - bh/2*500,
                      alpha=0.6, color='#1b3a5c')
    
    # Deflection line
    ax3d.plot(x3, [0]*n_x, y_defl*1000, color='#ff4757', lw=2, label='Neutral Axis')
    
    # Load arrows
    for i,(P,a) in enumerate(loads):
        if P > 0:
            y_at_a = np.interp(a, x, y)*1000
            ax3d.quiver(a, 0, y_at_a+200, 0, 0, -150,
                       color=colors_load[i], arrow_length_ratio=0.3, lw=2)
    
    ax3d.set_xlabel('Length (m)', color='#aaaaaa', fontsize=7)
    ax3d.set_ylabel('Width', color='#aaaaaa', fontsize=7)
    ax3d.set_zlabel('Deflection (mm)', color='#aaaaaa', fontsize=7)
    ax3d.set_title(f'3D Beam — {section_type}', color='#00bfff', fontsize=9)
    ax3d.tick_params(colors='#aaaaaa', labelsize=6)
    ax3d.xaxis.pane.fill = False
    ax3d.yaxis.pane.fill = False
    ax3d.zaxis.pane.fill = False
    ax3d.grid(True, color='#1b3a5c', alpha=0.3)
    
    # Stress color map on beam side view
    ax_stress = fig3d.add_subplot(122)
    ax_stress.set_facecolor('#0d2137')
    
    x_cs = np.linspace(0, L, 200)
    y_cs = np.linspace(-h_total/2, h_total/2, 60)
    M_cs = np.interp(x_cs, x, M)
    SIGMA_cs = np.outer(y_cs, M_cs)/I/1e6
    
    cmap3d = mcolors.LinearSegmentedColormap.from_list('stress3d',
        ['#0000ff','#00bfff','#00ff00','#ffff00','#ff0000'])
    vmax3d = sigma_max if sigma_max > 0 else 1
    im3d = ax_stress.imshow(SIGMA_cs, aspect='auto', origin='lower',
        extent=[0, L, -h_total/2*1000, h_total/2*1000],
        cmap=cmap3d, vmin=-vmax3d, vmax=vmax3d)
    
    # Plot deflected shape overlay
    ax_stress.plot(x, y*1000*5, color='white', lw=2, linestyle='--', label=f'Deflection ×5')
    ax_stress.legend(facecolor='#0d2137', labelcolor='white', fontsize=7)
    
    cbar3d = fig3d.colorbar(im3d, ax=ax_stress, pad=0.02)
    cbar3d.set_label('Stress (MPa)', color='white', fontsize=8)
    cbar3d.ax.yaxis.set_tick_params(color='white')
    plt.setp(cbar3d.ax.yaxis.get_ticklabels(), color='white')
    ax_stress.set_xlabel("Position (m)", color='#aaaaaa')
    ax_stress.set_ylabel("Height (mm)", color='#aaaaaa')
    ax_stress.set_title("Stress + Deflection Side View", color='#00bfff', fontsize=9)
    ax_stress.tick_params(colors='#aaaaaa')
    ax_stress.spines[:].set_color('#1b3a5c')
    
    plt.tight_layout()
    return fig3d

fig3d = plot_3d_beam()
st.pyplot(fig3d)

# ════════════════════════════════════════════
# NEW FEATURE 2: DYNAMIC / MOVING LOAD
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 🚗 Dynamic / Moving Load Analysis")

if enable_dynamic:
    st.markdown(f"""<div class='dynamic-box'>
    🚗 <b>Moving Load: {P_moving} N</b> — Analyzing influence lines across beam length
    </div>""", unsafe_allow_html=True)
    
    pos_arr = np.linspace(0.01, L-0.01, dynamic_positions)
    M_env_max = np.zeros_like(x)
    M_env_min = np.zeros_like(x)
    y_env_max = np.zeros_like(x)
    
    all_M = []
    all_y_defl = []
    
    for pos in pos_arr:
        _mv = [(P_moving, pos)]
        rx_d, _, _ = solve_beam(supports, _mv, [], None, L, E_val, I)
        _, M_d = beam_V_M(x, rx_d, _mv, [], None)
        y_d = beam_deflection(x, rx_d, M_d, E_val, I)

        all_M.append(M_d)
        all_y_defl.append(y_d*1000)
        M_env_max = np.maximum(M_env_max, M_d)
        M_env_min = np.minimum(M_env_min, M_d)
        y_env_max = np.maximum(y_env_max, np.abs(y_d*1000))
    
    fig_dyn, (ax_d1, ax_d2, ax_d3) = plt.subplots(3, 1, figsize=(10, 9))
    fig_dyn.patch.set_facecolor('#0a1628')
    for ax in [ax_d1, ax_d2, ax_d3]:
        ax.set_facecolor('#0d2137')
        ax.tick_params(colors='#aaaaaa')
        ax.spines[:].set_color('#1b3a5c')
        ax.grid(True, color='#1b3a5c', linewidth=0.5)
        ax.yaxis.label.set_color('#aaaaaa')
        ax.xaxis.label.set_color('#aaaaaa')
        ax.title.set_color('#00bfff')
    
    fig_dyn.suptitle(f"Moving Load Analysis — P={P_moving}N on {beam_type}", 
                     color='white', fontsize=11, fontweight='bold')
    
    # All BMD positions (influence lines)
    for i, M_d in enumerate(all_M):
        alpha = 0.15 + 0.5*(i/len(all_M))
        ax_d1.plot(x, M_d, color='#ff6b35', alpha=0.3, lw=0.8)
    ax_d1.fill_between(x, M_env_max, M_env_min, alpha=0.3, color='#ff6b35', label='Moment Envelope')
    ax_d1.plot(x, M_env_max, color='#ff4757', lw=2, label='Max BMD')
    ax_d1.plot(x, M_env_min, color='#ffd700', lw=2, label='Min BMD')
    ax_d1.axhline(0, color='white', lw=0.8)
    ax_d1.set_ylabel("Bending Moment (N·m)")
    ax_d1.set_title("BMD Envelope — All Load Positions")
    ax_d1.legend(facecolor='#0d2137', labelcolor='white', fontsize=7)
    
    # Deflection envelope
    for y_d in all_y_defl:
        ax_d2.plot(x, y_d, color='#00d4aa', alpha=0.2, lw=0.8)
    ax_d2.plot(x, y_env_max, color='#00d4aa', lw=2.5, label='Max Deflection Envelope')
    ax_d2.axhline(0, color='white', lw=0.8)
    ax_d2.set_ylabel("Deflection (mm)")
    ax_d2.set_title("Deflection Envelope")
    ax_d2.legend(facecolor='#0d2137', labelcolor='white', fontsize=7)
    
    # Max moment vs load position (influence line for center)
    mid_idx = len(x)//2
    M_at_mid = [M_d[mid_idx] for M_d in all_M]
    ax_d3.plot(pos_arr, M_at_mid, color='#7c4dff', lw=2.5, marker='o', markersize=4)
    ax_d3.fill_between(pos_arr, M_at_mid, alpha=0.25, color='#7c4dff')
    ax_d3.axhline(0, color='white', lw=0.8)
    ax_d3.set_xlabel("Load Position (m)")
    ax_d3.set_ylabel("Moment at Midspan (N·m)")
    ax_d3.set_title("Influence Line — Midspan Moment")
    
    plt.tight_layout()
    st.pyplot(fig_dyn)
    
    max_dyn_moment = np.max(M_env_max)
    max_dyn_defl   = np.max(y_env_max)
    st.info(f"🚗 Dynamic Analysis: Peak Moment = **{max_dyn_moment:.1f} N·m** | Peak Deflection = **{max_dyn_defl:.3f} mm** (worst case load position)")
else:
    st.info("☝️ Enable **Moving Load Analysis** in the sidebar to see influence lines & envelopes")

# ════════════════════════════════════════════
# STRESS HEATMAP
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 🌡️ Stress Heatmap (ANSYS Style)")
x_h = np.linspace(0,L,300)
y_h = np.linspace(-h_total/2,h_total/2,80)
M_h = np.interp(x_h,x,M)
SIGMA = np.outer(y_h,M_h)/I/1e6
fig_hm,ax_hm = plt.subplots(figsize=(12,3))
fig_hm.patch.set_facecolor('#0a1628'); ax_hm.set_facecolor('#0a1628')
cmap = mcolors.LinearSegmentedColormap.from_list('ansys',
    ['#0000ff','#00bfff','#00ff00','#ffff00','#ff6600','#ff0000'])
vmax = sigma_max if sigma_max > 0 else 1
im = ax_hm.imshow(SIGMA,aspect='auto',origin='lower',
    extent=[0,L,-h_total/2*1000,h_total/2*1000],cmap=cmap,vmin=-vmax,vmax=vmax)
cbar = fig_hm.colorbar(im,ax=ax_hm,pad=0.02)
cbar.set_label('Stress (MPa)',color='white',fontsize=9)
cbar.ax.yaxis.set_tick_params(color='white')
plt.setp(cbar.ax.yaxis.get_ticklabels(),color='white')
ax_hm.set_xlabel("Position along beam (m)",color='#aaaaaa')
ax_hm.set_ylabel("Height (mm)",color='#aaaaaa')
ax_hm.set_title(f"Bending Stress Heatmap — {section_type} | {E_sel}",color='#00bfff',fontsize=10)
ax_hm.tick_params(colors='#aaaaaa'); ax_hm.spines[:].set_color('#1b3a5c')
st.pyplot(fig_hm)

# ════════════════════════════════════════════
# CROSS SECTION VISUALIZER
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📐 Cross Section Visualizer")
fig_cs,(ax_cs,ax_sd) = plt.subplots(1,2,figsize=(10,4))
fig_cs.patch.set_facecolor('#0a1628')
for ax in [ax_cs,ax_sd]:
    ax.set_facecolor('#0d2137'); ax.tick_params(colors='#aaaaaa')
    ax.spines[:].set_color('#1b3a5c'); ax.grid(True,color='#1b3a5c')
ax_cs.set_title(f"{section_type} Cross Section",color='#00bfff')
if section_type == "Rectangular":
    b,h = dims['b'],dims['h']
    ax_cs.add_patch(mpatches.Rectangle((-b/2,-h/2),b,h,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-b,b); ax_cs.set_ylim(-h,h)
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "Circular":
    d = dims['d']
    ax_cs.add_patch(mpatches.Circle((0,0),d/2,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-d,d); ax_cs.set_ylim(-d,d); ax_cs.set_aspect('equal')
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "I-Beam":
    bf,tf,hw,tw = dims['bf'],dims['tf'],dims['hw'],dims['tw']
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,hw/2),bf,tf,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-tw/2,-hw/2),tw,hw,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,-hw/2-tf),bf,tf,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-bf,bf); ax_cs.set_ylim(-(hw+2*tf),(hw+2*tf))
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "T-Beam":
    bf,tf,hw,tw = dims['bf'],dims['tf'],dims['hw'],dims['tw']
    yb = dims['y_bar']; ht = hw+tf
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,hw-yb),bf,tf,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-tw/2,-yb),tw,hw,color=mat_color+'44',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-bf,bf); ax_cs.set_ylim(-ht,ht)
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
ax_cs.set_xlabel("Width (m)",color='#aaaaaa')
ax_cs.set_ylabel("Height (m)",color='#aaaaaa')
ax_cs.legend(facecolor='#0d2137',labelcolor='white',fontsize=7)
y_fib = np.linspace(-h_total/2,h_total/2,100)
sig_fib = (M_max*y_fib)/I/1e6
ax_sd.plot(sig_fib,y_fib,color='#ff6b35',lw=2.5)
ax_sd.fill_betweenx(y_fib,sig_fib,0,where=(sig_fib>0),color='#ff475733',label='Tension')
ax_sd.fill_betweenx(y_fib,sig_fib,0,where=(sig_fib<0),color='#00bfff33',label='Compression')
ax_sd.axvline(0,color='white',lw=0.8)
ax_sd.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
ax_sd.set_xlabel("Stress (MPa)",color='#aaaaaa')
ax_sd.set_ylabel("Height (m)",color='#aaaaaa')
ax_sd.set_title("Stress Distribution",color='#00bfff')
ax_sd.legend(facecolor='#0d2137',labelcolor='white',fontsize=8)
plt.tight_layout()
st.pyplot(fig_cs)

# ════════════════════════════════════════════
# NEW FEATURE 3: MATERIAL COMPARISON
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 🧪 Material Comparison Table")
mat_comparison = []
for mat_name, mat_props in material_data.items():
    E_m = mat_props["E"]
    y_m = mat_props["yield"]
    d_m = mat_props["density"]
    fos_m = (y_m/1e6) / sigma_max if sigma_max > 0 else 999
    defl_m = y_max * (E_val / E_m)  # scale deflection by E ratio
    weight_m = d_m * area * L * 9.81
    mat_comparison.append({
        "Material": mat_name,
        "E (GPa)": f"{E_m/1e9:.0f}",
        "Yield (MPa)": f"{y_m/1e6:.0f}",
        "FOS": f"{min(fos_m,999):.2f}",
        "Deflection (mm)": f"{defl_m:.3f}",
        "Weight (N)": f"{weight_m:.0f}",
        "Safe?": "✅" if fos_m > 2.0 else ("⚠️" if fos_m > 1.0 else "❌")
    })

import pandas as pd
df_mat = pd.DataFrame(mat_comparison)
st.dataframe(df_mat.set_index("Material"), width='stretch')

# ════════════════════════════════════════════
# NEW FEATURE 4: EXCEL EXPORT
# ════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Download Results")
col_d1, col_d2, col_d3 = st.columns(3)

# PNG Download
buf = io.BytesIO()
fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',facecolor='#0a1628')
buf.seek(0)
col_d1.download_button("📥 Download Diagrams (PNG)", buf, "beam_analysis.png", "image/png")

# Excel Export
def generate_excel():
    wb = openpyxl.Workbook()
    
    # ── Sheet 1: Summary ──
    ws1 = wb.active
    ws1.title = "Summary"
    
    hdr_fill  = PatternFill("solid", fgColor="0D2137")
    hdr_font  = Font(bold=True, color="FFFFFF", size=11)
    val_fill  = PatternFill("solid", fgColor="1B3A5C")
    val_font  = Font(color="E0E0E0", size=10)
    safe_font = Font(bold=True, color="00C853", size=11)
    fail_font = Font(bold=True, color="FF4757", size=11)
    thin = Border(
        left=Side(style='thin', color='2A4A6C'),
        right=Side(style='thin', color='2A4A6C'),
        top=Side(style='thin', color='2A4A6C'),
        bottom=Side(style='thin', color='2A4A6C')
    )
    center = Alignment(horizontal='center', vertical='center')
    
    def set_cell(ws, row, col, value, fill=None, font=None, align=None, border=None):
        c = ws.cell(row=row, column=col, value=value)
        if fill:   c.fill = fill
        if font:   c.font = font
        if align:  c.alignment = align
        if border: c.border = border
        return c
    
    # Title
    ws1.merge_cells('A1:G1')
    set_cell(ws1, 1, 1, "2D Beam Stress & Deflection Analysis Report",
             PatternFill("solid", fgColor="001A33"),
             Font(bold=True, color="00BFFF", size=14), center)
    ws1.row_dimensions[1].height = 30
    
    # Input Parameters
    ws1.merge_cells('A3:G3')
    set_cell(ws1, 3, 1, "INPUT PARAMETERS", hdr_fill, hdr_font, center)
    
    params = [
        ("Beam Type", beam_type), ("Length (m)", L),
        ("Material", E_sel), ("Section Type", section_type),
        ("Supports", support_desc), ("Distributed Load", udl_desc), ("Applied Moments", moment_desc),
        ("E (GPa)", f"{E_val/1e9:.0f}"),
        ("Yield Strength (MPa)", f"{yield_MPa:.0f}"),
        ("Moment of Inertia (m⁴)", f"{I:.4e}"),
        ("Beam Weight (N)", f"{beam_weight:.1f}"),
    ]
    for i, (k, v) in enumerate(params, start=4):
        set_cell(ws1, i, 1, k, val_fill, Font(color="00D4AA", bold=True, size=10), border=thin)
        set_cell(ws1, i, 2, v, PatternFill("solid", fgColor="0A1628"), val_font, border=thin)
        ws1.column_dimensions['A'].width = 28
        ws1.column_dimensions['B'].width = 22
    
    row = 4 + len(params) + 1
    
    # Loads
    ws1.merge_cells(f'A{row}:G{row}')
    set_cell(ws1, row, 1, "APPLIED LOADS", hdr_fill, hdr_font, center)
    row += 1
    set_cell(ws1, row, 1, "Load #", hdr_fill, hdr_font, center, thin)
    set_cell(ws1, row, 2, "Force (N)", hdr_fill, hdr_font, center, thin)
    set_cell(ws1, row, 3, "Position (m)", hdr_fill, hdr_font, center, thin)
    row += 1
    for i, (P, a) in enumerate(loads, 1):
        set_cell(ws1, row, 1, f"P{i}", val_fill, val_font, center, thin)
        set_cell(ws1, row, 2, P, val_fill, val_font, center, thin)
        set_cell(ws1, row, 3, a, val_fill, val_font, center, thin)
        row += 1
    row += 1
    
    # Results
    ws1.merge_cells(f'A{row}:G{row}')
    set_cell(ws1, row, 1, "RESULTS", hdr_fill, hdr_font, center)
    row += 1
    
    results = [
        (f"Reaction R{i+1} — {t} @ {p/L_CONV:g}{length_unit} (N)", f"{Rv:.2f}" + (f"  (M={Mr:.2f} N·m)" if t == "Fixed" else ""))
        for i, (p, Rv, Mr, t) in enumerate(reactions)
    ] + [
        ("Max Bending Moment (N·m)", f"{M_max:.2f}"),
        ("Max Deflection (mm)", f"{y_max:.4f}"),
        ("Max Bending Stress (MPa)", f"{sigma_max:.2f}"),
        ("Yield Strength (MPa)", f"{yield_MPa:.0f}"),
        ("Factor of Safety", f"{FOS:.2f}"),
        ("Status", "SAFE ✅" if sigma_max < yield_MPa else "UNSAFE ❌"),
    ]
    for k, v in results:
        set_cell(ws1, row, 1, k, val_fill, Font(color="00D4AA", bold=True, size=10), border=thin)
        cell = set_cell(ws1, row, 2, v, PatternFill("solid", fgColor="0A1628"),
                        safe_font if "SAFE ✅" in str(v) else (fail_font if "UNSAFE" in str(v) else val_font),
                        border=thin)
        row += 1
    
    # ── Sheet 2: Raw Data ──
    ws2 = wb.create_sheet("SFD_BMD_Data")
    ws2.sheet_properties.tabColor = "00BFFF"
    
    headers = ["Position (m)", "Shear Force (N)", "Bending Moment (N·m)", "Deflection (mm)"]
    for col, h in enumerate(headers, 1):
        set_cell(ws2, 1, col, h, hdr_fill, hdr_font, center, thin)
        ws2.column_dimensions[chr(64+col)].width = 20
    
    step = max(1, len(x)//200)
    for row_i, idx in enumerate(range(0, len(x), step), start=2):
        vals = [round(float(x[idx]),4), round(float(V[idx]),4),
                round(float(M[idx]),4), round(float(y[idx]*1000),6)]
        for col, v in enumerate(vals, 1):
            fill = PatternFill("solid", fgColor="0A1628" if row_i%2==0 else "0D1F35")
            set_cell(ws2, row_i, col, v, fill, val_font, center, thin)
    
    # ── Sheet 3: Material Comparison ──
    ws3 = wb.create_sheet("Material_Comparison")
    ws3.sheet_properties.tabColor = "7C4DFF"
    
    mat_headers = ["Material", "E (GPa)", "Yield (MPa)", "Density (kg/m³)", "FOS", "Deflection (mm)", "Weight (N)", "Safe?"]
    for col, h in enumerate(mat_headers, 1):
        set_cell(ws3, 1, col, h, hdr_fill, hdr_font, center, thin)
        ws3.column_dimensions[chr(64+col)].width = 22
    
    for row_i, row_data in enumerate(mat_comparison, start=2):
        vals = [row_data["Material"], row_data["E (GPa)"], row_data["Yield (MPa)"],
                material_data[row_data["Material"]]["density"],
                row_data["FOS"], row_data["Deflection (mm)"], row_data["Weight (N)"], row_data["Safe?"]]
        for col, v in enumerate(vals, 1):
            fill = PatternFill("solid", fgColor="0A1628" if row_i%2==0 else "0D1F35")
            fnt = (safe_font if v == "✅" else fail_font if v == "❌" else val_font)
            set_cell(ws3, row_i, col, v, fill, fnt, center, thin)
    
    # Add Line chart for SFD
    chart = LineChart()
    chart.title = "Shear Force Diagram"
    chart.style = 10
    chart.y_axis.title = "Shear Force (N)"
    chart.x_axis.title = "Position"
    
    n_pts = min(50, len(x)//step)
    data_ref = Reference(ws2, min_col=2, min_row=1, max_row=n_pts+1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.series[0].graphicalProperties.line.solidFill = "00BFFF"
    ws2.add_chart(chart, "F2")
    
    # BMD chart
    chart2 = LineChart()
    chart2.title = "Bending Moment Diagram"
    chart2.style = 10
    chart2.y_axis.title = "Moment (N·m)"
    data_ref2 = Reference(ws2, min_col=3, min_row=1, max_row=n_pts+1)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.series[0].graphicalProperties.line.solidFill = "FF6B35"
    ws2.add_chart(chart2, "F20")
    
    buf_xl = io.BytesIO()
    wb.save(buf_xl)
    buf_xl.seek(0)
    return buf_xl

if EXCEL_OK:
    if col_d2.button("📊 Generate Excel Report"):
        with st.spinner("Generating Excel..."):
            xl_buf = generate_excel()
        col_d2.download_button("📥 Download Excel", xl_buf,
                               "beam_analysis.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    col_d2.warning("Install openpyxl for Excel export")

# PDF Report
def generate_pdf(beam_type,L,loads,w,E,section_type,dims,I,c_dist,
                 RA,RB,M_max,y_max,sigma_max,yield_MPa,FOS,fig,warnings,suggestions,good):
    buf_pdf = io.BytesIO()
    doc = SimpleDocTemplate(buf_pdf,pagesize=A4,leftMargin=20*mm,rightMargin=20*mm,
                            topMargin=20*mm,bottomMargin=20*mm)
    title_style = ParagraphStyle('t',fontSize=18,textColor=colors.HexColor('#0D2137'),
                                  fontName='Helvetica-Bold',alignment=TA_CENTER,spaceAfter=6)
    sub_style   = ParagraphStyle('s',fontSize=11,textColor=colors.HexColor('#1B3A5C'),
                                  fontName='Helvetica',alignment=TA_CENTER,spaceAfter=12)
    head_style  = ParagraphStyle('h',fontSize=12,textColor=colors.HexColor('#0D2137'),
                                  fontName='Helvetica-Bold',spaceBefore=10,spaceAfter=4)
    body_style  = ParagraphStyle('b',fontSize=9,fontName='Helvetica',spaceAfter=3)
    story = []
    story.append(Paragraph("2D Beam Stress & Deflection Simulator Pro",title_style))
    story.append(Paragraph("Engineering Analysis Report",sub_style))
    story.append(Spacer(1,5*mm))
    story.append(Paragraph("1. Input Parameters",head_style))
    load_str = " | ".join([f"P{i+1}={p}N@{a}m" for i,(p,a) in enumerate(loads)])
    dim_str  = " | ".join([f"{k}={v}" for k,v in dims.items() if k!='y_bar'])
    params = [["Parameter","Value"],
              ["Beam Type",beam_type],["Length",f"{L}m"],
              ["Loads",load_str],["Distributed Load",udl_desc],["Applied Moments",moment_desc],["Supports",support_desc],
              ["Material",E],["Section",section_type],
              ["Dimensions",dim_str],["I",f"{I:.4e} m4"],
              ["Beam Weight",f"{beam_weight:.1f} N"]]
    t = Table(params,colWidths=[70*mm,100*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0D2137')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.HexColor('#F0F4FA'),colors.white]),
        ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#CBD5E0')),
        ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),6)]))
    story.append(t); story.append(Spacer(1,5*mm))
    story.append(Paragraph("2. Results",head_style))
    safe_str = "SAFE" if sigma_max < yield_MPa else "UNSAFE"
    results = [["Result","Value","Unit"],
               *[[f"R{i+1} ({t} @ {p/L_CONV:g}{length_unit})",f"{Rv:.2f}" + (f" / M={Mr:.2f}" if t == "Fixed" else ""),"N"]
                 for i,(p,Rv,Mr,t) in enumerate(reactions)],
               ["M_max",f"{M_max:.2f}","N.m"],["y_max",f"{y_max:.4f}","mm"],
               ["sigma_max",f"{sigma_max:.2f}","MPa"],["Yield",f"{yield_MPa:.0f}","MPa"],
               ["FOS",f"{FOS:.2f}","-"],["Status",safe_str,"-"]]
    t2 = Table(results,colWidths=[80*mm,60*mm,30*mm])
    sc = colors.HexColor('#00C853') if sigma_max < yield_MPa else colors.HexColor('#FF4757')
    t2.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0D2137')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.HexColor('#F0F4FA'),colors.white]),
        ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#CBD5E0')),
        ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),6),
        ('TEXTCOLOR',(1,-1),(1,-1),sc),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold')]))
    story.append(t2); story.append(Spacer(1,5*mm))
    story.append(Paragraph("3. AI Engineering Suggestions",head_style))
    for warn in warnings: story.append(Paragraph(warn,body_style))
    for sugg in suggestions: story.append(Paragraph(sugg,body_style))
    for g in good: story.append(Paragraph(g,body_style))
    story.append(Spacer(1,5*mm))
    story.append(Paragraph("4. Engineering Diagrams",head_style))
    buf_img = io.BytesIO()
    fig.savefig(buf_img,format='png',dpi=120,bbox_inches='tight',facecolor='#0a1628')
    buf_img.seek(0)
    story.append(Image(buf_img,width=170*mm,height=120*mm))
    doc.build(story)
    buf_pdf.seek(0)
    return buf_pdf

if col_d3.button("📄 Generate PDF Report"):
    with st.spinner("Generating PDF..."):
        pdf_buf = generate_pdf(beam_type,L,loads,w,E_sel,section_type,dims,I,c_dist,
                               RA,RB,M_max,y_max,sigma_max,yield_MPa,FOS,fig,
                               warnings,suggestions,good)
    col_d3.download_button("📥 Download PDF Report",pdf_buf,"beam_report.pdf","application/pdf")

with st.expander("📚 Theory & Formulas"):
    st.markdown(f"""
    **{beam_type} | {section_type} Section | {E_sel}**
    - `I` = **{I:.4e} m⁴** | `c` = **{c_dist*1000:.1f} mm**
    - Reactions: {", ".join(f"R{i+1}={r[1]:.1f} N" for i, r in enumerate(reactions))}
    - `M_max` = **{M_max:.1f} N·m** | `σ_max` = **{sigma_max:.2f} MPa**
    - `FOS` = **{FOS:.2f}** | `y_max` = **{y_max:.4f} mm**
    - `Beam Weight` = **{beam_weight:.1f} N** | `Density` = **{density} kg/m³**
    """)
