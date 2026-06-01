import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

st.set_page_config(page_title="Beam Simulator", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0a1628; color: #e0e0e0; }
    .stSidebar { background-color: #0d2137; }
    [data-testid="stMetricValue"] { color: #ffffff !important; font-size: 1.4rem !important; }
    [data-testid="stMetricLabel"] { color: #00bfff !important; }
    .safe-box { background: #00c85322; border: 2px solid #00c853; border-radius: 10px; padding: 12px 20px; font-size: 16px; font-weight: bold; color: #00c853; }
    .unsafe-box { background: #ff475722; border: 2px solid #ff4757; border-radius: 10px; padding: 12px 20px; font-size: 16px; font-weight: bold; color: #ff4757; }
    .ai-box { background: #7c4dff22; border: 2px solid #7c4dff; border-radius: 10px; padding: 15px 20px; margin: 5px 0; }
    .ai-title { color: #7c4dff; font-size: 18px; font-weight: bold; margin-bottom: 10px; }
    .ai-suggestion { color: #e0e0e0; font-size: 14px; padding: 4px 0; }
    .ai-warning { color: #ff4757; font-size: 14px; font-weight: bold; padding: 4px 0; }
    .ai-good { color: #00c853; font-size: 14px; padding: 4px 0; }
    h1 { color: #00bfff !important; } h2, h3 { color: #00d4aa !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("# 🏗️ 2D Beam Stress & Deflection Simulator")
st.markdown("---")

st.sidebar.markdown("## ⚙️ Beam Parameters")
beam_type = st.sidebar.selectbox("🏛️ Beam Type", ["Simply Supported","Cantilever","Fixed-Fixed"])
L = st.sidebar.slider("📏 Beam Length (m)", 0.5, 10.0, 4.0, 0.1)

st.sidebar.markdown("### ⬇️ Point Loads")
n_loads = st.sidebar.number_input("Number of Point Loads", 1, 8, 1)
loads = []
for i in range(int(n_loads)):
    st.sidebar.markdown(f"**Load {i+1}**")
    p = st.sidebar.slider(f"P{i+1} (N)", 0, 10000, 500, 100, key=f"p{i}")
    default_pos = round(L/(n_loads+1)*(i+1), 1)
    default_pos = min(default_pos, L)
    a = st.sidebar.slider(f"Position {i+1} (m)", 0.0, L, default_pos, 0.1, key=f"a{i}")
    loads.append((p, a))

st.sidebar.markdown("### 〰️ UDL")
w = st.sidebar.slider("UDL Intensity (N/m)", 0, 2000, 0, 50)

material_data = {
    "🔩 Steel (200 GPa)":   {"E": 200e9, "yield": 250e6},
    "✈️ Aluminum (69 GPa)": {"E": 69e9,  "yield": 276e6},
    "⚡ Copper (110 GPa)":  {"E": 110e9, "yield": 210e6},
}
E = st.sidebar.selectbox("🧱 Material", material_data.keys())
E_val   = material_data[E]["E"]
yield_s = material_data[E]["yield"]

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

# ── Calculations ──
x = np.linspace(0, L, 2000)
colors_load = ['#ff4757','#ffd700','#ff6b35','#7c4dff','#00d4aa','#00bfff','#ff69b4','#adff2f']

if beam_type == "Simply Supported":
    RA = sum((P*(L-a))/L for P,a in loads) + w*L/2
    RB = sum((P*a)/L     for P,a in loads) + w*L/2
    V  = np.full_like(x, RA)
    for P,a in loads: V -= P*(x >= a).astype(float)
    V -= w*x
    M  = RA*x - 0.5*w*x**2
    for P,a in loads: M -= P*np.maximum(x-a, 0)
    y  = (w*x)/(24*E_val*I)*(L**3-2*L*x**2+x**3)
    for P,a in loads:
        bv = L-a
        y += np.where(x<a,
            (P*bv*x)/(6*E_val*I*L)*(L**2-bv**2-x**2),
            (P*a*(L-x))/(6*E_val*I*L)*(2*L*(L-x)-a**2-(L-x)**2))

elif beam_type == "Cantilever":
    RA = sum(P for P,a in loads) + w*L
    MA_fix = sum(P*(L-a) for P,a in loads) + w*L**2/2
    RB = 0
    V  = np.full_like(x, RA)
    for P,a in loads: V -= P*(x >= a).astype(float)
    V -= w*x
    M  = -MA_fix + RA*x - 0.5*w*x**2
    for P,a in loads: M -= P*np.maximum(x-a, 0)
    y  = np.zeros_like(x)
    for P,a in loads:
        y += np.where(x<=a,
            (P*x**2)/(6*E_val*I)*(3*a-x),
            (P*a**2)/(6*E_val*I)*(3*x-a))
    y += (w*x**2)/(24*E_val*I)*(x**2-4*L*x+6*L**2)
    y  = y - y[0]

elif beam_type == "Fixed-Fixed":
    RA = sum((P*(L-a)**2*(2*a+L))/L**3 for P,a in loads) + w*L/2
    RB = sum((P*a**2*(3*L-2*a))/L**3   for P,a in loads) + w*L/2
    MA_fix = sum((P*a*(L-a)**2)/L**2   for P,a in loads) + w*L**2/12
    V  = np.full_like(x, RA)
    for P,a in loads: V -= P*(x >= a).astype(float)
    V -= w*x
    M  = -MA_fix + RA*x - 0.5*w*x**2
    for P,a in loads: M -= P*np.maximum(x-a, 0)
    y  = np.zeros_like(x)
    for P,a in loads:
        bv = L-a
        y += np.where(x<a,
            (P*bv**2*x**2)/(6*E_val*I*L**3)*(3*a*L-x*(3*a+bv)),
            (P*a**2*(L-x)**2)/(6*E_val*I*L**3)*(3*bv*L-(L-x)*(3*bv+a)))
    y += (w*x**2*(L-x)**2)/(24*E_val*I)

M_max     = max(abs(M))
y_max     = max(abs(y))*1000
sigma_max = (M_max*c_dist)/I/1e6
yield_MPa = yield_s/1e6
FOS       = yield_MPa/sigma_max if sigma_max > 0 else 999

# ── AI Suggestions ──
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
        suggestions.append("💡 Switch to Steel (200 GPa) for 3x higher yield strength")
    if beam_type == "Simply Supported" and FOS < 2.0:
        suggestions.append("💡 Consider Fixed-Fixed beam — reduces max moment by ~50%")
    if beam_type == "Cantilever" and FOS < 2.0:
        suggestions.append("💡 Cantilever has high moment at fixed end — add intermediate support")
    if len(loads)==1 and loads[0][0]>0:
        if abs(loads[0][1]-L/2) < 0.1:
            good.append("✅ Load at center — symmetric loading condition")
    if FOS >= 3.0 and stress_ratio < 0.4:
        good.append("🏆 Excellent design! Well within all safety limits.")
    elif FOS >= 2.0:
        good.append("👍 Good design — meets standard engineering requirements.")
    return warnings, suggestions, good

# ── Beam Visual ──
st.markdown("### 🔎 Beam Diagram")
fig_beam, ax_b = plt.subplots(figsize=(10, 2.8))
fig_beam.patch.set_facecolor('#0a1628')
ax_b.set_facecolor('#0d2137')
ax_b.set_xlim(-0.8, L+0.8); ax_b.set_ylim(-1.2, 2.8); ax_b.axis('off')
ax_b.add_patch(plt.Rectangle((0,0.2),L,0.3,color='#1b3a5c',zorder=2))
ax_b.add_patch(plt.Rectangle((0,0.2),L,0.3,fill=False,edgecolor='#00bfff',lw=2,zorder=3))
if beam_type == "Simply Supported":
    ax_b.plot([0],[0.2],'v',color='#00d4aa',markersize=14,zorder=4)
    ax_b.plot([L],[0.2],'o',color='#00d4aa',markersize=10,zorder=4)
    ax_b.text(0,-0.3,'A (Pin)',color='#00d4aa',ha='center',fontsize=8)
    ax_b.text(L,-0.3,'B (Roller)',color='#00d4aa',ha='center',fontsize=8)
elif beam_type == "Cantilever":
    ax_b.add_patch(plt.Rectangle((-0.3,-0.1),0.3,0.9,color='#00d4aa',zorder=4))
    ax_b.text(0,-0.3,'Fixed',color='#00d4aa',ha='center',fontsize=8)
    ax_b.text(L,-0.3,'Free End',color='#ff4757',ha='center',fontsize=8)
elif beam_type == "Fixed-Fixed":
    ax_b.add_patch(plt.Rectangle((-0.3,-0.1),0.3,0.9,color='#00d4aa',zorder=4))
    ax_b.add_patch(plt.Rectangle((L,-0.1),0.3,0.9,color='#00d4aa',zorder=4))
    ax_b.text(0,-0.3,'Fixed A',color='#00d4aa',ha='center',fontsize=8)
    ax_b.text(L,-0.3,'Fixed B',color='#00d4aa',ha='center',fontsize=8)
for i,(P,a) in enumerate(loads):
    if P > 0:
        ax_b.annotate('',xy=(a,0.5),xytext=(a,1.6),
            arrowprops=dict(arrowstyle='->',color=colors_load[i],lw=2.5))
        ax_b.text(a,1.75,f'P{i+1}={P}N',color=colors_load[i],ha='center',fontsize=8,fontweight='bold')
if w > 0:
    for xi in np.linspace(0.1,L-0.1,14):
        ax_b.annotate('',xy=(xi,0.5),xytext=(xi,1.1),
            arrowprops=dict(arrowstyle='->',color='#00bfff',lw=1))
    ax_b.plot([0,L],[1.15,1.15],color='#00bfff',lw=2)
    ax_b.text(L/2,1.28,f'w={w} N/m',color='#00bfff',ha='center',fontsize=8)
ax_b.annotate('',xy=(L,-0.7),xytext=(0,-0.7),
    arrowprops=dict(arrowstyle='<->',color='gray',lw=1))
ax_b.text(L/2,-1.0,f'L = {L}m',color='gray',ha='center',fontsize=8)
st.pyplot(fig_beam)
st.markdown("---")

# ── Safety Check ──
st.markdown("### 🛡️ Safety Check")
if sigma_max < yield_MPa:
    st.markdown(f"""<div class='safe-box'>✅ SAFE &nbsp;|&nbsp; σ_max: {sigma_max:.2f} MPa &nbsp;<&nbsp; Yield: {yield_MPa:.0f} MPa &nbsp;|&nbsp; FOS: {FOS:.2f}</div>""", unsafe_allow_html=True)
else:
    st.markdown(f"""<div class='unsafe-box'>❌ UNSAFE &nbsp;|&nbsp; σ_max: {sigma_max:.2f} MPa &nbsp;>&nbsp; Yield: {yield_MPa:.0f} MPa &nbsp;|&nbsp; BEAM WILL FAIL!</div>""", unsafe_allow_html=True)
st.markdown("---")

# ── Results ──
st.markdown("### 📊 Results")
c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric("RA", f"{RA:.1f} N")
c2.metric("RB", f"{RB:.1f} N")
c3.metric("M_max", f"{M_max:.1f} N·m")
c4.metric("y_max", f"{y_max:.3f} mm")
c5.metric("σ_max", f"{sigma_max:.2f} MPa")
c6.metric("FOS",   f"{FOS:.2f}")
c7.metric("I",     f"{I:.2e} m⁴")
st.markdown("---")

# ── AI Suggestions ──
st.markdown("### 🤖 AI Engineering Suggestions")
warnings, suggestions, good = ai_suggestions(
    beam_type,section_type,L,loads,w,E,dims,
    FOS,sigma_max,yield_MPa,y_max,M_max,I,h_total)
ai_html = "<div class='ai-box'><div class='ai-title'>🤖 Smart Engineering Analysis</div>"
for warn in warnings: ai_html += f"<div class='ai-warning'>{warn}</div>"
for sugg in suggestions: ai_html += f"<div class='ai-suggestion'>{sugg}</div>"
for g in good: ai_html += f"<div class='ai-good'>{g}</div>"
ai_html += "</div>"
st.markdown(ai_html, unsafe_allow_html=True)
st.markdown("---")

# ── Plots ──
st.markdown("### 📈 Engineering Diagrams")
fig,(ax1,ax2,ax3) = plt.subplots(3,1,figsize=(10,9))
fig.patch.set_facecolor('#0a1628')
for ax in [ax1,ax2,ax3]:
    ax.set_facecolor('#0d2137'); ax.tick_params(colors='#aaaaaa')
    ax.spines[:].set_color('#1b3a5c')
    ax.yaxis.label.set_color('#aaaaaa'); ax.xaxis.label.set_color('#aaaaaa')
    ax.title.set_color('#00bfff'); ax.grid(True,color='#1b3a5c',linewidth=0.7)
fig.suptitle(f"{beam_type} | {section_type} | L={L}m | UDL={w}N/m",
             fontsize=11,fontweight='bold',color='white')
ax1.plot(x,V,color='#00bfff',lw=2); ax1.fill_between(x,V,alpha=0.25,color='#00bfff')
ax1.axhline(0,color='white',lw=0.8)
for i,(P,a) in enumerate(loads):
    ax1.axvline(a,color=colors_load[i],lw=1,linestyle='--',alpha=0.7,label=f'P{i+1}@{a}m')
ax1.set_ylabel("Shear Force (N)"); ax1.set_title("SFD")
ax1.legend(facecolor='#0d2137',labelcolor='white',fontsize=7)
ax2.plot(x,M,color='#ff6b35',lw=2); ax2.fill_between(x,M,alpha=0.25,color='#ff6b35')
ax2.axhline(0,color='white',lw=0.8)
ax2.set_ylabel("Bending Moment (N·m)"); ax2.set_title("BMD")
ax3.plot(x,y*1000,color='#00d4aa',lw=2); ax3.fill_between(x,y*1000,alpha=0.25,color='#00d4aa')
ax3.axhline(0,color='white',lw=0.8)
ax3.set_xlabel("Position (m)"); ax3.set_ylabel("Deflection (mm)"); ax3.set_title("Deflection Curve")
plt.tight_layout()
st.pyplot(fig)

# ── Stress Heatmap ──
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
ax_hm.set_title(f"Bending Stress Heatmap — {section_type} Section",color='#00bfff',fontsize=10)
ax_hm.tick_params(colors='#aaaaaa'); ax_hm.spines[:].set_color('#1b3a5c')
st.pyplot(fig_hm)

# ── Cross Section Visualizer ──
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
    ax_cs.add_patch(mpatches.Rectangle((-b/2,-h/2),b,h,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-b,b); ax_cs.set_ylim(-h,h)
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "Circular":
    d = dims['d']
    ax_cs.add_patch(mpatches.Circle((0,0),d/2,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-d,d); ax_cs.set_ylim(-d,d); ax_cs.set_aspect('equal')
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "I-Beam":
    bf,tf,hw,tw = dims['bf'],dims['tf'],dims['hw'],dims['tw']
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,hw/2),bf,tf,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-tw/2,-hw/2),tw,hw,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,-hw/2-tf),bf,tf,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.set_xlim(-bf,bf); ax_cs.set_ylim(-(hw+2*tf),(hw+2*tf))
    ax_cs.axhline(0,color='#ffd700',lw=1.5,linestyle='--',label='Neutral Axis')
elif section_type == "T-Beam":
    bf,tf,hw,tw = dims['bf'],dims['tf'],dims['hw'],dims['tw']
    yb = dims['y_bar']; ht = hw+tf
    ax_cs.add_patch(mpatches.Rectangle((-bf/2,hw-yb),bf,tf,color='#1b3a5c',ec='#00bfff',lw=2))
    ax_cs.add_patch(mpatches.Rectangle((-tw/2,-yb),tw,hw,color='#1b3a5c',ec='#00bfff',lw=2))
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

# ── Download ──
st.markdown("---")
st.markdown("### 💾 Download Results")
col_d1,col_d2 = st.columns(2)
buf = io.BytesIO()
fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',facecolor='#0a1628')
buf.seek(0)
col_d1.download_button("📥 Download Diagrams (PNG)", buf, "beam_analysis.png", "image/png")

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
    story.append(Paragraph("2D Beam Stress & Deflection Simulator",title_style))
    story.append(Paragraph("Engineering Analysis Report",sub_style))
    story.append(Spacer(1,5*mm))
    story.append(Paragraph("1. Input Parameters",head_style))
    load_str = " | ".join([f"P{i+1}={p}N@{a}m" for i,(p,a) in enumerate(loads)])
    dim_str  = " | ".join([f"{k}={v}" for k,v in dims.items() if k!='y_bar'])
    params = [["Parameter","Value"],
              ["Beam Type",beam_type],["Length",f"{L}m"],
              ["Loads",load_str],["UDL",f"{w}N/m"],
              ["Material",E],["Section",section_type],
              ["Dimensions",dim_str],["I",f"{I:.4e} m4"]]
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
               ["RA",f"{RA:.2f}","N"],["RB",f"{RB:.2f}","N"],
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

if col_d2.button("📄 Generate PDF Report"):
    with st.spinner("Generating PDF..."):
        pdf_buf = generate_pdf(beam_type,L,loads,w,E,section_type,dims,I,c_dist,
                               RA,RB,M_max,y_max,sigma_max,yield_MPa,FOS,fig,
                               warnings,suggestions,good)
    col_d2.download_button("📥 Download PDF Report",pdf_buf,"beam_report.pdf","application/pdf")

with st.expander("📚 Theory & Formulas"):
    st.markdown(f"""
    **{beam_type} | {section_type} Section**
    - `I` = **{I:.4e} m⁴** | `c` = **{c_dist*1000:.1f} mm**
    - `RA` = **{RA:.1f} N** | `RB` = **{RB:.1f} N**
    - `M_max` = **{M_max:.1f} N·m** | `σ_max` = **{sigma_max:.2f} MPa**
    - `FOS` = **{FOS:.2f}** | `y_max` = **{y_max:.4f} mm**
    """)