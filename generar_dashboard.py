#!/usr/bin/env python3
"""
NetserGroup Dashboard Generator — Version 2.0 (UI Renovada)
Genera index.html con GSAP, AOS, Lucide Icons, glassmorphism y animaciones avanzadas.
Genera widget.html para vista movil.
"""
import os, sys, json
from datetime import datetime

try:
    import openpyxl
except ImportError:
    os.system('pip install openpyxl')
    import openpyxl

EXCEL = 'Reporte_NetserGroup_Final.xlsx'
CLIENTS = ['HP Comercial','HPE','Payless','Netapp','Lexmark','Lexmark Kit','CTDI','Monthly Fee','Lenovo']
BOTS = ['BackUp Mobility','Cierre POs','Cierre Alpha','Cierre HPCM','Tasas Cambio',
        'Encuestas Dell','Respaldo Invoice','Cierre Residencias','Receiving Lab',
        'Reporte Inv HP','HPCM Cenam','HPCM Chile','Licencias FSM','Regularizacion Mobility']
ICONS = {'HP Comercial':'&#128424;','HPE':'&#128424;','Payless':'&#128722;','Netapp':'&#9729;',
         'Lexmark':'&#128424;','Lexmark Kit':'&#128424;','CTDI':'&#128295;','Monthly Fee':'&#128178;','Lenovo':'&#128187;'}

def read_data():
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    ws = wb['Datos']
    h = [str(c.value).strip() if c.value else '' for c in ws[1]]
    v = [c.value for c in ws[2]]
    return dict(zip(h, v))

def get_clients(rec):
    d = {}
    for c in CLIENTS:
        try: d[c] = int(float(rec.get(c,0) or 0))
        except: d[c] = 0
    return d

def get_bots(rec):
    r = []
    for b in BOTS:
        v = str(rec.get(b,'')).strip() if rec.get(b) else ''
        ok = v in ['✔','✓','OK','True','1','1.0']
        r.append({'name':b,'ok':ok})
    return r

def build_html(cd, bd):
    total = sum(cd.values())
    active_cl = sum(1 for v in cd.values() if v > 0)
    bots_ok = sum(1 for b in bd if b['ok'])
    bots_fail = sum(1 for b in bd if not b['ok'])
    total_bots = len(bd)
    rate = round(bots_ok/total_bots*100) if total_bots else 0
    now = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    rate_col = '#00B894' if rate==100 else '#FDCB6E' if rate>=80 else '#E17055'

    # Sort clients by value desc
    cs = sorted(cd.items(), key=lambda x:-x[1])

    # Active clients for charts
    ac = [(n,v) for n,v in cs if v>0]
    if not ac: ac = cs
    cl = json.dumps([n for n,v in ac])
    cv = json.dumps([v for n,v in ac])
    num_ac = len(ac)

    # Client cards
    cards = ''
    for i,(n,v) in enumerate(cs):
        pct = round(v/total*100) if total else 0
        icon = ICONS.get(n,'&#128203;')
        act = ' active' if i==0 and v>0 else ''
        zero = ' zero' if v==0 else ''
        pct_html = f'<div class="client-pct">{pct}%</div>' if v>0 else ''
        delay = 50 + i * 50
        cards += f'''<div class="client-card{act}" data-aos="zoom-in" data-aos-delay="{delay}">
    <div class="client-icon">{icon}</div>
    <div class="client-value{zero} counter" data-target="{v}">0</div>
    <div class="client-name">{n}</div>{pct_html}
    <div class="client-bar"><div class="client-bar-fill" data-width="{pct}"></div></div>
  </div>\n'''

    # Bot rows
    rows = ''
    for i,b in enumerate(bd):
        if b['ok']:
            s = '<span class="status-ok"><i data-lucide="check-circle"></i> OK</span>'
        else:
            s = '<span class="status-fail"><i data-lucide="x-circle"></i> FAIL</span>'
        rows += f'<tr><td>{b["name"]}</td><td>{s}</td></tr>\n'

    html = f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>NetserGroup Dashboard</title>
<!-- Libraries -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"><\/script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"><\/script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.5/gsap.min.js"><\/script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.5/ScrollTrigger.min.js"><\/script>
<link href="https://unpkg.com/aos@2.3.4/dist/aos.css" rel="stylesheet">
<script src="https://unpkg.com/aos@2.3.4/dist/aos.js"><\/script>
<script src="https://unpkg.com/lucide@latest"><\/script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

/* ====== RESET & VARIABLES ====== */
*{{margin:0;padding:0;box-sizing:border-box}}
:root{{
  --bg:#060E1A;--bg2:#0A1628;--bg3:#0D1B2A;--bg-card:rgba(18,35,61,.65);
  --brd:#1E3A5F;--brd-hover:rgba(0,180,216,.35);
  --a:#00B4D8;--a2:#0077B6;--a3:#48CAE4;--a4:#90E0EF;
  --ok:#00B894;--ok2:#55EFC4;--fail:#E17055;--fail2:#FF7675;--warn:#FDCB6E;--warn2:#FFEAA7;
  --t1:#FFFFFF;--t2:#D1D9E6;--t3:#5A7A9A;--t4:#3D5A80;
  --glass:rgba(255,255,255,.03);--glass2:rgba(255,255,255,.06);
  --shadow:0 8px 32px rgba(0,0,0,.3);--shadow-hover:0 16px 48px rgba(0,180,216,.15);
  --radius:18px;--radius-sm:12px;--radius-xs:8px;
  --transition:all .4s cubic-bezier(.4,0,.2,1);
}}
html{{scroll-behavior:smooth}}
body{{font-family:'Inter',sans-serif;background:var(--bg);color:var(--t1);min-height:100vh;overflow-x:hidden;
  transition:background .5s,color .5s}}
::selection{{background:var(--a);color:var(--bg)}}
::-webkit-scrollbar{{width:6px}}
::-webkit-scrollbar-track{{background:var(--bg)}}
::-webkit-scrollbar-thumb{{background:var(--brd);border-radius:3px}}
::-webkit-scrollbar-thumb:hover{{background:var(--a2)}}

/* ====== PARTICLE CANVAS ====== */
#pc{{position:fixed;top:0;left:0;width:100%;height:100%;z-index:0;pointer-events:none;opacity:.3}}

/* ====== AMBIENT GLOW ====== */
.ambient{{position:fixed;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:0;overflow:hidden}}
.ambient .orb{{position:absolute;border-radius:50%;filter:blur(80px);opacity:.08;animation:orbFloat 20s ease-in-out infinite}}
.ambient .orb:nth-child(1){{width:600px;height:600px;background:var(--a);top:-200px;left:-100px;animation-delay:0s}}
.ambient .orb:nth-child(2){{width:500px;height:500px;background:var(--a2);bottom:-150px;right:-100px;animation-delay:-7s}}
.ambient .orb:nth-child(3){{width:400px;height:400px;background:#6C5CE7;top:50%;left:50%;animation-delay:-14s}}
@keyframes orbFloat{{
  0%,100%{{transform:translate(0,0) scale(1)}}
  25%{{transform:translate(80px,-60px) scale(1.1)}}
  50%{{transform:translate(-40px,80px) scale(.9)}}
  75%{{transform:translate(60px,40px) scale(1.05)}}
}}

/* ====== LAYOUT ====== */
.dashboard{{position:relative;z-index:1;max-width:1480px;margin:0 auto;padding:20px 28px 40px}}

/* ====== NAVIGATION BAR ====== */
.navbar{{display:flex;justify-content:space-between;align-items:center;padding:14px 24px;
  background:rgba(13,27,42,.8);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);
  border-radius:var(--radius);border:1px solid var(--brd);margin-bottom:20px;
  position:sticky;top:10px;z-index:100;transition:var(--transition)}}
.navbar::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--a),var(--a2),#6C5CE7,var(--a));
  background-size:300% 100%;animation:navGlow 6s ease infinite;border-radius:var(--radius) var(--radius) 0 0}}
@keyframes navGlow{{0%{{background-position:0 50%}}50%{{background-position:100% 50%}}100%{{background-position:0 50%}}}}
.navbar.scrolled{{background:rgba(6,14,26,.95);box-shadow:var(--shadow)}}
.nav-left{{display:flex;align-items:center;gap:14px}}
.nav-logo{{width:44px;height:44px;animation:logoFloat 4s ease-in-out infinite}}
@keyframes logoFloat{{0%,100%{{transform:translateY(0) rotate(0deg)}}50%{{transform:translateY(-5px) rotate(2deg)}}}}
.nav-title h1{{font-size:20px;font-weight:800;letter-spacing:-.5px;
  background:linear-gradient(135deg,#fff 0%,#B8C4D6 50%,var(--a3) 100%);
  background-size:200% 200%;animation:titleShimmer 4s ease infinite;
  -webkit-background-clip:text;-webkit-text-fill-color:transparent}}
@keyframes titleShimmer{{0%{{background-position:0 50%}}50%{{background-position:100% 50%}}100%{{background-position:0 50%}}}}
.nav-title p{{font-size:11px;color:var(--t3);margin-top:2px;letter-spacing:.5px}}
.nav-right{{display:flex;align-items:center;gap:12px}}
.nav-clock{{font-size:13px;font-weight:600;color:var(--a3);font-variant-numeric:tabular-nums;
  padding:6px 14px;background:rgba(0,180,216,.06);border:1px solid rgba(0,180,216,.12);
  border-radius:var(--radius-xs);letter-spacing:.5px}}
.nav-status{{display:flex;align-items:center;gap:8px;padding:8px 16px;
  background:rgba(0,184,148,.06);border:1px solid rgba(0,184,148,.15);
  border-radius:var(--radius-sm);font-size:11px;color:var(--ok);font-weight:500}}
.pulse-dot{{width:8px;height:8px;background:var(--ok);border-radius:50%;position:relative}}
.pulse-dot::after{{content:'';position:absolute;inset:-4px;border-radius:50%;
  border:2px solid var(--ok);animation:pulseRing 2s cubic-bezier(.4,0,.6,1) infinite;opacity:0}}
@keyframes pulseRing{{0%{{transform:scale(.8);opacity:.6}}100%{{transform:scale(1.8);opacity:0}}}}
.theme-toggle{{width:38px;height:38px;border-radius:10px;border:1px solid var(--brd);
  background:var(--glass);color:var(--t3);cursor:pointer;display:flex;align-items:center;
  justify-content:center;transition:var(--transition);font-size:16px}}
.theme-toggle:hover{{border-color:var(--a);color:var(--a);background:rgba(0,180,216,.08);transform:rotate(15deg)}}

/* ====== HEADER KPI STRIP ====== */
.kpi-strip{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}}
.kpi-mini{{background:var(--bg-card);backdrop-filter:blur(12px);border:1px solid var(--brd);
  border-radius:var(--radius);padding:20px;display:flex;align-items:center;gap:16px;
  transition:var(--transition);position:relative;overflow:hidden;cursor:pointer}}
.kpi-mini::before{{content:'';position:absolute;inset:0;
  background:linear-gradient(135deg,transparent 60%,var(--glass2));pointer-events:none}}
.kpi-mini:hover{{border-color:var(--brd-hover);transform:translateY(-4px);box-shadow:var(--shadow-hover)}}
.kpi-mini .kpi-icon{{width:48px;height:48px;border-radius:14px;display:flex;align-items:center;
  justify-content:center;flex-shrink:0;position:relative}}
.kpi-mini .kpi-icon::after{{content:'';position:absolute;inset:0;border-radius:14px;
  opacity:0;transition:opacity .3s}}
.kpi-mini:hover .kpi-icon::after{{opacity:1}}
.kpi-mini.success .kpi-icon{{background:rgba(0,184,148,.1);border:1px solid rgba(0,184,148,.2);color:var(--ok)}}
.kpi-mini.success .kpi-icon::after{{background:rgba(0,184,148,.15)}}
.kpi-mini.danger .kpi-icon{{background:rgba(225,112,85,.1);border:1px solid rgba(225,112,85,.2);color:var(--fail)}}
.kpi-mini.danger .kpi-icon::after{{background:rgba(225,112,85,.15)}}
.kpi-mini.info .kpi-icon{{background:rgba(0,180,216,.1);border:1px solid rgba(0,180,216,.2);color:var(--a)}}
.kpi-mini.info .kpi-icon::after{{background:rgba(0,180,216,.15)}}
.kpi-mini.warning .kpi-icon{{background:rgba(253,203,110,.1);border:1px solid rgba(253,203,110,.2);color:var(--warn)}}
.kpi-mini.warning .kpi-icon::after{{background:rgba(253,203,110,.15)}}
.kpi-data{{flex:1}}
.kpi-value{{font-size:30px;font-weight:800;line-height:1;font-variant-numeric:tabular-nums;margin-bottom:4px}}
.kpi-mini.success .kpi-value{{color:var(--ok)}}
.kpi-mini.danger .kpi-value{{color:var(--fail)}}
.kpi-mini.info .kpi-value{{color:var(--a)}}
.kpi-mini.warning .kpi-value{{color:var(--warn)}}
.kpi-label{{font-size:11px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.8px}}
.kpi-sub{{font-size:10px;color:var(--t4);margin-top:2px}}

/* ====== SECTION HEADERS ====== */
.section-header{{display:flex;align-items:center;justify-content:space-between;margin:28px 0 16px}}
.section-left{{display:flex;align-items:center;gap:12px}}
.section-bar{{width:4px;height:28px;border-radius:2px;
  background:linear-gradient(180deg,var(--a),var(--a2))}}
.section-header h2{{font-size:16px;font-weight:700;color:var(--t2);letter-spacing:.3px}}
.section-badge{{font-size:10px;padding:4px 10px;background:rgba(0,180,216,.08);
  border:1px solid rgba(0,180,216,.15);border-radius:20px;color:var(--a3);font-weight:600}}

/* ====== CLIENT CARDS ====== */
.clients-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:14px;margin-bottom:8px}}
.client-card{{background:var(--bg-card);backdrop-filter:blur(12px);border:1px solid var(--brd);
  border-radius:var(--radius);padding:22px 16px 30px;text-align:center;cursor:pointer;
  position:relative;overflow:hidden;transition:var(--transition)}}
.client-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:linear-gradient(90deg,transparent,var(--a3),transparent);
  transform:scaleX(0);transition:transform .4s cubic-bezier(.4,0,.2,1)}}
.client-card::after{{content:'';position:absolute;inset:0;
  background:radial-gradient(circle at var(--mx,50%) var(--my,50%),rgba(0,180,216,.08),transparent 60%);
  opacity:0;transition:opacity .3s}}
.client-card:hover{{border-color:var(--brd-hover);transform:translateY(-8px) scale(1.02);
  box-shadow:var(--shadow-hover)}}
.client-card:hover::before{{transform:scaleX(1)}}
.client-card:hover::after{{opacity:1}}
.client-card.active{{border-color:var(--a);background:linear-gradient(145deg,rgba(13,40,71,.7),rgba(22,52,86,.7))}}
.client-card.active::before{{transform:scaleX(1);background:linear-gradient(90deg,var(--a),var(--a3))}}
.client-icon{{font-size:24px;margin-bottom:8px;opacity:.6;transition:var(--transition)}}
.client-card:hover .client-icon{{opacity:1;transform:scale(1.3) rotate(-5deg)}}
.client-value{{font-size:38px;font-weight:800;line-height:1;margin-bottom:6px;
  font-variant-numeric:tabular-nums;transition:var(--transition)}}
.client-value.zero{{color:var(--t4)}}
.client-name{{font-size:11px;color:#7B8DA6;font-weight:600;letter-spacing:.3px}}
.client-pct{{position:absolute;top:10px;right:12px;font-size:10px;color:var(--t3);font-weight:700;
  padding:2px 8px;background:rgba(0,180,216,.06);border-radius:20px}}
.client-bar{{position:absolute;bottom:0;left:0;right:0;height:4px;background:rgba(30,58,95,.3)}}
.client-bar-fill{{height:100%;border-radius:0 0 var(--radius) var(--radius);
  background:linear-gradient(90deg,var(--a2),var(--a3));
  transition:width 2s cubic-bezier(.4,0,.2,1)}}

/* ====== PANELS / CARDS ====== */
.grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-bottom:8px}}
.grid-2{{display:grid;grid-template-columns:1fr 1.5fr;gap:20px;margin-bottom:8px}}
.grid-4{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}}
.panel{{background:var(--bg-card);backdrop-filter:blur(12px);border:1px solid var(--brd);
  border-radius:var(--radius);padding:28px;position:relative;overflow:hidden;
  transition:var(--transition)}}
.panel::before{{content:'';position:absolute;inset:0;
  background:linear-gradient(135deg,var(--glass) 0%,transparent 50%);pointer-events:none}}
.panel:hover{{border-color:rgba(0,180,216,.2)}}
.panel-title{{font-size:13px;font-weight:600;color:var(--t2);margin-bottom:16px;
  display:flex;align-items:center;gap:8px}}
.panel-title i{{width:16px;height:16px;color:var(--a3)}}

/* ====== TOTAL RING (SVG) ====== */
.ring-container{{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%}}
.ring-svg{{width:200px;height:200px;transform:rotate(-90deg)}}
.ring-bg{{fill:none;stroke:var(--brd);stroke-width:6}}
.ring-progress{{fill:none;stroke:url(#ringGrad);stroke-width:6;stroke-linecap:round;
  stroke-dasharray:565.48;stroke-dashoffset:565.48;transition:stroke-dashoffset 2.5s cubic-bezier(.4,0,.2,1)}}
.ring-center{{text-align:center;margin-top:-170px;margin-bottom:40px}}
.ring-value{{font-size:52px;font-weight:900;line-height:1;
  background:linear-gradient(135deg,var(--a),var(--a3));
  -webkit-background-clip:text;-webkit-text-fill-color:transparent}}
.ring-label{{font-size:12px;color:var(--t3);margin-top:4px;font-weight:500;letter-spacing:1px;text-transform:uppercase}}
.ring-sub{{font-size:12px;color:var(--t4);margin-top:8px}}
.ring-glow{{position:absolute;width:200px;height:200px;border-radius:50%;
  background:radial-gradient(circle,rgba(0,180,216,.1),transparent 70%);
  top:50%;left:50%;transform:translate(-50%,-55%);animation:ringPulse 3s ease-in-out infinite}}
@keyframes ringPulse{{0%,100%{{opacity:.3;transform:translate(-50%,-55%) scale(1)}}50%{{opacity:.6;transform:translate(-50%,-55%) scale(1.1)}}}}

/* ====== CHART CONTAINER ====== */
.chart-wrap{{position:relative;width:100%;height:270px}}

/* ====== TABLE ====== */
.table-container{{overflow:hidden}}
.table-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}}
.table-search{{display:flex;align-items:center;gap:8px;padding:8px 14px;
  background:var(--glass);border:1px solid var(--brd);border-radius:var(--radius-xs);
  transition:var(--transition)}}
.table-search:focus-within{{border-color:var(--a);box-shadow:0 0 0 3px rgba(0,180,216,.1)}}
.table-search input{{background:none;border:none;outline:none;color:var(--t2);font-family:Inter;font-size:12px;width:140px}}
.table-search input::placeholder{{color:var(--t4)}}
.table-search i{{width:14px;height:14px;color:var(--t4)}}
table{{width:100%;border-collapse:separate;border-spacing:0 6px}}
thead th{{text-align:left;font-size:10px;color:var(--t3);text-transform:uppercase;
  letter-spacing:1.2px;padding:10px 16px;font-weight:700}}
thead th:last-child{{text-align:center}}
tbody tr{{background:rgba(26,42,68,.3);transition:var(--transition);cursor:pointer}}
tbody tr:hover{{background:rgba(0,180,216,.06);transform:translateX(6px)}}
td{{padding:13px 16px;font-size:13px;color:var(--t2)}}
td:first-child{{border-radius:var(--radius-sm) 0 0 var(--radius-sm);font-weight:500}}
td:last-child{{border-radius:0 var(--radius-sm) var(--radius-sm) 0;text-align:center}}
.status-ok{{display:inline-flex;align-items:center;gap:6px;padding:5px 14px;
  background:rgba(0,184,148,.08);border:1px solid rgba(0,184,148,.18);border-radius:20px;
  font-size:11px;font-weight:600;color:var(--ok);transition:var(--transition)}}
.status-ok:hover{{background:rgba(0,184,148,.15);box-shadow:0 0 16px rgba(0,184,148,.15)}}
.status-ok i{{width:13px;height:13px}}
.status-fail{{display:inline-flex;align-items:center;gap:6px;padding:5px 14px;
  background:rgba(225,112,85,.08);border:1px solid rgba(225,112,85,.18);border-radius:20px;
  font-size:11px;font-weight:600;color:var(--fail);transition:var(--transition)}}
.status-fail i{{width:13px;height:13px}}

/* ====== BOT HEALTH CENTER ====== */
.bot-health{{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:20px}}
.bot-gauge{{position:relative;width:200px;height:200px}}
.bot-gauge-center{{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center}}
.bot-gauge-value{{font-size:36px;font-weight:900;color:{rate_col}}}
.bot-gauge-label{{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:1.5px;margin-top:2px}}
.bot-stats{{display:flex;gap:24px;margin-top:20px}}
.bot-stat{{text-align:center}}
.bot-stat-value{{font-size:18px;font-weight:700}}
.bot-stat-label{{font-size:9px;color:var(--t4);text-transform:uppercase;letter-spacing:1px;margin-top:2px}}

/* ====== OPERATIONAL KPIs ====== */
.op-kpi{{background:var(--bg-card);backdrop-filter:blur(12px);border:1px solid var(--brd);
  border-radius:var(--radius);padding:24px 16px;text-align:center;position:relative;
  overflow:hidden;transition:var(--transition);cursor:pointer}}
.op-kpi::before{{content:'';position:absolute;inset:0;
  background:linear-gradient(135deg,var(--glass),transparent 60%);pointer-events:none}}
.op-kpi:hover{{transform:translateY(-6px);box-shadow:var(--shadow-hover);border-color:var(--brd-hover)}}
.op-kpi .op-icon{{width:44px;height:44px;border-radius:12px;display:flex;align-items:center;
  justify-content:center;margin:0 auto 14px;font-size:18px;transition:var(--transition)}}
.op-kpi:hover .op-icon{{transform:scale(1.1) rotate(5deg)}}
.op-kpi.s .op-icon{{background:rgba(0,184,148,.1);border:1px solid rgba(0,184,148,.2)}}
.op-kpi.f .op-icon{{background:rgba(225,112,85,.1);border:1px solid rgba(225,112,85,.2)}}
.op-kpi.r .op-icon{{background:rgba(0,180,216,.1);border:1px solid rgba(0,180,216,.2)}}
.op-kpi.m .op-icon{{background:rgba(253,203,110,.1);border:1px solid rgba(253,203,110,.2)}}
.op-value{{font-size:38px;font-weight:800;line-height:1;margin-bottom:6px;font-variant-numeric:tabular-nums}}
.op-kpi.s .op-value{{color:var(--ok)}}.op-kpi.f .op-value{{color:var(--fail)}}
.op-kpi.r .op-value{{color:var(--a)}}.op-kpi.m .op-value{{color:var(--warn)}}
.op-label{{font-size:10px;color:#7B8DA6;font-weight:700;text-transform:uppercase;letter-spacing:1px}}
.op-sub{{font-size:10px;color:var(--t4);margin-top:4px}}

/* ====== WIDGET LINK ====== */
.widget-link{{display:flex;justify-content:center;margin:28px 0 8px}}
.widget-btn{{display:inline-flex;align-items:center;gap:10px;padding:14px 32px;
  background:linear-gradient(135deg,rgba(0,119,182,.12),rgba(0,180,216,.08));
  border:1px solid rgba(0,180,216,.25);border-radius:var(--radius);color:var(--a);
  font-size:13px;font-weight:600;text-decoration:none;transition:var(--transition);
  font-family:Inter,sans-serif;position:relative;overflow:hidden}}
.widget-btn::before{{content:'';position:absolute;top:0;left:-100%;width:100%;height:100%;
  background:linear-gradient(90deg,transparent,rgba(0,180,216,.1),transparent);
  transition:left .6s}}
.widget-btn:hover{{border-color:var(--a);box-shadow:0 4px 20px rgba(0,180,216,.15);transform:translateY(-2px)}}
.widget-btn:hover::before{{left:100%}}
.widget-btn i{{width:18px;height:18px}}

/* ====== FOOTER ====== */
footer{{margin-top:36px;text-align:center;padding:20px;
  border-top:1px solid rgba(30,58,95,.5)}}
footer p{{font-size:11px;color:var(--t4)}}
footer .footer-brand{{color:var(--a3);font-weight:600}}

/* ====== RESPONSIVE ====== */
@media(max-width:1100px){{
  .grid-2,.grid-3{{grid-template-columns:1fr}}
  .grid-4,.kpi-strip{{grid-template-columns:repeat(2,1fr)}}
  .navbar{{flex-direction:column;gap:14px;position:relative;top:0}}
  .nav-right{{justify-content:center;flex-wrap:wrap}}
}}
@media(max-width:600px){{
  .clients-grid{{grid-template-columns:repeat(2,1fr)}}
  .grid-4,.kpi-strip{{grid-template-columns:1fr 1fr}}
  .dashboard{{padding:12px 14px 30px}}
  .nav-title h1{{font-size:16px}}
  .kpi-value{{font-size:24px}}
  .client-value{{font-size:28px}}
}}

/* ====== TOOLTIP ====== */
.tooltip{{position:relative}}
.tooltip::after{{content:attr(data-tip);position:absolute;bottom:calc(100% + 8px);left:50%;
  transform:translateX(-50%) translateY(4px);padding:6px 12px;background:rgba(6,14,26,.95);
  border:1px solid var(--brd);border-radius:var(--radius-xs);font-size:10px;color:var(--t2);
  white-space:nowrap;opacity:0;pointer-events:none;transition:all .2s;z-index:50}}
.tooltip:hover::after{{opacity:1;transform:translateX(-50%) translateY(0)}}
</style>
</head>
<body>

<!-- Ambient background orbs -->
<div class="ambient">
  <div class="orb"></div>
  <div class="orb"></div>
  <div class="orb"></div>
</div>

<!-- Particle canvas -->
<canvas id="pc"></canvas>

<div class="dashboard">

<!-- ====== NAVBAR ====== -->
<nav class="navbar" id="navbar">
  <div class="nav-left">
    <svg class="nav-logo" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="logoGrad" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" style="stop-color:#0077B6"/>
          <stop offset="100%" style="stop-color:#00B4D8"/>
        </linearGradient>
        <filter id="glow"><feGaussianBlur stdDeviation="2" result="blur"/>
          <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
        </filter>
      </defs>
      <circle cx="50" cy="50" r="26" fill="url(#logoGrad)" opacity="0.85" filter="url(#glow)"/>
      <ellipse cx="50" cy="50" rx="34" ry="13" fill="none" stroke="#00B4D8" stroke-width="1.5" opacity="0.5" transform="rotate(-25 50 50)">
        <animateTransform attributeName="transform" type="rotate" values="-25 50 50;335 50 50" dur="30s" repeatCount="indefinite"/>
      </ellipse>
      <ellipse cx="50" cy="50" rx="37" ry="10" fill="none" stroke="#0077B6" stroke-width="1" opacity="0.35" transform="rotate(30 50 50)">
        <animateTransform attributeName="transform" type="rotate" values="30 50 50;-330 50 50" dur="25s" repeatCount="indefinite"/>
      </ellipse>
      <ellipse cx="50" cy="50" rx="30" ry="8" fill="none" stroke="#48CAE4" stroke-width="1.5" opacity="0.45" transform="rotate(75 50 50)">
        <animateTransform attributeName="transform" type="rotate" values="75 50 50;435 50 50" dur="20s" repeatCount="indefinite"/>
      </ellipse>
      <text x="50" y="88" text-anchor="middle" font-family="Inter,Arial" font-size="7" font-weight="800" fill="white">NETSER</text>
      <text x="50" y="96" text-anchor="middle" font-family="Inter,Arial" font-size="5" fill="#48CAE4">GROUP</text>
    </svg>
    <div class="nav-title">
      <h1>NETSERGROUP</h1>
      <p>Dashboard Operacional &mdash; Monitoreo en Tiempo Real</p>
    </div>
  </div>
  <div class="nav-right">
    <div class="nav-clock" id="liveClock">--:--:--</div>
    <div class="nav-status"><span class="pulse-dot"></span> Sistema Operativo</div>
    <button class="theme-toggle tooltip" data-tip="Cambiar tema" id="themeToggle" aria-label="Toggle theme">
      <i data-lucide="sun"></i>
    </button>
  </div>
</nav>

<!-- ====== KPI STRIP ====== -->
<div class="kpi-strip" data-aos="fade-up" data-aos-delay="100">
  <div class="kpi-mini success tooltip" data-tip="Bots funcionando correctamente">
    <div class="kpi-icon"><i data-lucide="bot"></i></div>
    <div class="kpi-data">
      <div class="kpi-value counter" data-target="{bots_ok}">0</div>
      <div class="kpi-label">Bots OK</div>
      <div class="kpi-sub">{bots_ok} / {total_bots} activos</div>
    </div>
  </div>
  <div class="kpi-mini danger tooltip" data-tip="Bots con errores">
    <div class="kpi-icon"><i data-lucide="alert-triangle"></i></div>
    <div class="kpi-data">
      <div class="kpi-value counter" data-target="{bots_fail}">0</div>
      <div class="kpi-label">Bots Fail</div>
      <div class="kpi-sub">{bots_fail} errores detectados</div>
    </div>
  </div>
  <div class="kpi-mini warning tooltip" data-tip="Alertas del sistema">
    <div class="kpi-icon"><i data-lucide="bell"></i></div>
    <div class="kpi-data">
      <div class="kpi-value counter" data-target="{bots_fail}">0</div>
      <div class="kpi-label">Alertas</div>
      <div class="kpi-sub">{"Sin alertas pendientes" if bots_fail == 0 else str(bots_fail) + " alertas activas"}</div>
    </div>
  </div>
  <div class="kpi-mini info tooltip" data-tip="Total de casos activos hoy">
    <div class="kpi-icon"><i data-lucide="briefcase"></i></div>
    <div class="kpi-data">
      <div class="kpi-value counter" data-target="{total}">0</div>
      <div class="kpi-label">Casos Activos</div>
      <div class="kpi-sub">{active_cl} clientes activos</div>
    </div>
  </div>
</div>

<!-- ====== CLIENTS ====== -->
<div class="section-header" data-aos="fade-right" data-aos-delay="50">
  <div class="section-left">
    <div class="section-bar"></div>
    <h2>Casos por Cliente</h2>
  </div>
  <span class="section-badge">{len(cd)} clientes</span>
</div>
<div class="clients-grid">
{cards}</div>

<!-- ====== RESUMEN GENERAL ====== -->
<div class="section-header" data-aos="fade-right">
  <div class="section-left">
    <div class="section-bar"></div>
    <h2>Resumen General</h2>
  </div>
  <span class="section-badge">Vista consolidada</span>
</div>
<div class="grid-3">
  <div class="panel" data-aos="fade-up" data-aos-delay="100">
    <div class="ring-container" style="position:relative">
      <div class="ring-glow"></div>
      <svg class="ring-svg" viewBox="0 0 200 200">
        <defs>
          <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" style="stop-color:#00B4D8"/>
            <stop offset="100%" style="stop-color:#48CAE4"/>
          </linearGradient>
        </defs>
        <circle class="ring-bg" cx="100" cy="100" r="90"/>
        <circle class="ring-progress" id="mainRing" cx="100" cy="100" r="90"/>
      </svg>
      <div class="ring-center">
        <div class="ring-value counter" data-target="{total}">0</div>
        <div class="ring-label">casos activos</div>
      </div>
      <div class="ring-sub">{active_cl} clientes activos de {len(cd)}</div>
    </div>
  </div>
  <div class="panel" data-aos="fade-up" data-aos-delay="200">
    <div class="panel-title"><i data-lucide="bar-chart-3"></i> Distribucion por Cliente</div>
    <div class="chart-wrap"><canvas id="barC"></canvas></div>
  </div>
  <div class="panel" data-aos="fade-up" data-aos-delay="300">
    <div class="panel-title"><i data-lucide="pie-chart"></i> Proporcion de Casos</div>
    <div class="chart-wrap"><canvas id="doC"></canvas></div>
  </div>
</div>

<!-- ====== BOTS ====== -->
<div class="section-header" data-aos="fade-right">
  <div class="section-left">
    <div class="section-bar"></div>
    <h2>Estado de Bots / Flujos</h2>
  </div>
  <span class="section-badge">{total_bots} flujos monitoreados</span>
</div>
<div class="grid-2">
  <div class="panel table-container" data-aos="fade-up" data-aos-delay="100">
    <div class="table-header">
      <div class="panel-title" style="margin-bottom:0"><i data-lucide="cpu"></i> Flujos Activos</div>
      <div class="table-search">
        <i data-lucide="search"></i>
        <input type="text" placeholder="Buscar bot..." id="botSearch">
      </div>
    </div>
    <table>
      <thead><tr><th>Bot / Flujo</th><th>Estado</th></tr></thead>
      <tbody id="botTable">
{rows}      </tbody>
    </table>
  </div>
  <div class="panel" data-aos="fade-up" data-aos-delay="200" style="display:flex;flex-direction:column;align-items:center;justify-content:center">
    <div class="panel-title"><i data-lucide="activity"></i> Estado Global de Bots</div>
    <div class="bot-gauge">
      <canvas id="botD"></canvas>
      <div class="bot-gauge-center">
        <div class="bot-gauge-value">{rate}%</div>
        <div class="bot-gauge-label">Tasa Exito</div>
      </div>
    </div>
    <div class="bot-stats">
      <div class="bot-stat">
        <div class="bot-stat-value" style="color:var(--ok)">{bots_ok}</div>
        <div class="bot-stat-label">Exitosos</div>
      </div>
      <div class="bot-stat">
        <div class="bot-stat-value" style="color:var(--fail)">{bots_fail}</div>
        <div class="bot-stat-label">Fallidos</div>
      </div>
      <div class="bot-stat">
        <div class="bot-stat-value" style="color:var(--a)">0</div>
        <div class="bot-stat-label">Alertas</div>
      </div>
    </div>
  </div>
</div>

<!-- ====== RESUMEN OPERACIONAL ====== -->
<div class="section-header" data-aos="fade-right">
  <div class="section-left">
    <div class="section-bar"></div>
    <h2>Resumen Operacional</h2>
  </div>
</div>
<div class="grid-4">
  <div class="op-kpi s" data-aos="flip-up" data-aos-delay="100">
    <div class="op-icon"><i data-lucide="check-circle" style="color:var(--ok)"></i></div>
    <div class="op-value counter" data-target="{bots_ok}">0</div>
    <div class="op-label">Exitosos</div>
    <div class="op-sub">{bots_ok} / {total_bots} bots</div>
  </div>
  <div class="op-kpi f" data-aos="flip-up" data-aos-delay="200">
    <div class="op-icon"><i data-lucide="x-circle" style="color:var(--fail)"></i></div>
    <div class="op-value counter" data-target="{bots_fail}">0</div>
    <div class="op-label">Fallidos</div>
    <div class="op-sub">{bots_fail} / {total_bots} bots</div>
  </div>
  <div class="op-kpi r" data-aos="flip-up" data-aos-delay="300">
    <div class="op-icon"><i data-lucide="star" style="color:var(--a)"></i></div>
    <div class="op-value counter" data-target="{rate}">0</div>
    <div class="op-label">Tasa de Exito</div>
    <div class="op-sub">Rendimiento global</div>
  </div>
  <div class="op-kpi m" data-aos="flip-up" data-aos-delay="400">
    <div class="op-icon"><i data-lucide="zap" style="color:var(--warn)"></i></div>
    <div class="op-value counter" data-target="{total}">0</div>
    <div class="op-label">Total Casos</div>
    <div class="op-sub">Todos los clientes</div>
  </div>
</div>

<!-- ====== WIDGET LINK ====== -->
<div class="widget-link" data-aos="fade-up">
  <a href="widget.html" class="widget-btn">
    <i data-lucide="smartphone"></i> Ver Widget Movil
  </a>
</div>

<footer>
  <p>Ultima Actualizacion: <span id="footerTime">{now}</span> &mdash;
  <span class="footer-brand">NetserGroup</span> &copy; 2026 &mdash; Humberto Henriquez</p>
</footer>

</div>

''' + '<' + 'script>' + '''
// ====== INITIALIZE LUCIDE ICONS ======
lucide.createIcons();

// ====== INITIALIZE AOS ======
AOS.init({
  duration: 800,
  easing: 'ease-out-cubic',
  once: true,
  offset: 40
});

// ====== LIVE CLOCK ======
function updateClock() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  document.getElementById('liveClock').textContent =
    pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
}
setInterval(updateClock, 1000);
updateClock();

// ====== NAVBAR SCROLL EFFECT ======
window.addEventListener('scroll', () => {
  document.getElementById('navbar').classList.toggle('scrolled', window.scrollY > 40);
});

// ====== PARTICLE NETWORK (ENHANCED) ======
const cv = document.getElementById('pc'), cx = cv.getContext('2d');
let mouseX = 0, mouseY = 0;
function resizeCanvas() { cv.width = innerWidth; cv.height = innerHeight; }
resizeCanvas();
window.addEventListener('resize', resizeCanvas);
document.addEventListener('mousemove', e => { mouseX = e.clientX; mouseY = e.clientY; });

const particles = [];
for (let i = 0; i < 65; i++) {
  particles.push({
    x: Math.random() * cv.width, y: Math.random() * cv.height,
    vx: (Math.random() - .5) * .35, vy: (Math.random() - .5) * .35,
    r: Math.random() * 2 + .5, o: Math.random() * .5 + .1,
    hue: 190 + Math.random() * 30
  });
}

function drawParticles() {
  cx.clearRect(0, 0, cv.width, cv.height);
  particles.forEach((p, i) => {
    const dmx = mouseX - p.x, dmy = mouseY - p.y;
    const dm = Math.hypot(dmx, dmy);
    if (dm < 200) { p.vx += dmx * 0.00003; p.vy += dmy * 0.00003; }
    p.vx *= 0.999; p.vy *= 0.999;
    p.x += p.vx; p.y += p.vy;
    if (p.x < 0) p.x = cv.width; if (p.x > cv.width) p.x = 0;
    if (p.y < 0) p.y = cv.height; if (p.y > cv.height) p.y = 0;
    const grad = cx.createRadialGradient(p.x, p.y, 0, p.x, p.y, p.r * 3);
    grad.addColorStop(0, 'hsla(' + p.hue + ',80%,60%,' + p.o + ')');
    grad.addColorStop(1, 'hsla(' + p.hue + ',80%,60%,0)');
    cx.fillStyle = grad;
    cx.beginPath(); cx.arc(p.x, p.y, p.r * 3, 0, Math.PI * 2); cx.fill();
    cx.fillStyle = 'hsla(' + p.hue + ',80%,70%,' + (p.o + .1) + ')';
    cx.beginPath(); cx.arc(p.x, p.y, p.r, 0, Math.PI * 2); cx.fill();
    particles.forEach((q, j) => {
      if (j <= i) return;
      const d = Math.hypot(p.x - q.x, p.y - q.y);
      if (d < 140) {
        cx.strokeStyle = 'hsla(195,80%,50%,' + (0.06 * (1 - d / 140)) + ')';
        cx.lineWidth = .6;
        cx.beginPath(); cx.moveTo(p.x, p.y); cx.lineTo(q.x, q.y); cx.stroke();
      }
    });
    if (dm < 200) {
      cx.strokeStyle = 'hsla(195,80%,60%,' + (0.08 * (1 - dm / 200)) + ')';
      cx.lineWidth = .8;
      cx.beginPath(); cx.moveTo(p.x, p.y); cx.lineTo(mouseX, mouseY); cx.stroke();
    }
  });
  requestAnimationFrame(drawParticles);
}
drawParticles();

// ====== GSAP COUNTER ANIMATION ======
gsap.registerPlugin(ScrollTrigger);

function animateCounters() {
  document.querySelectorAll('.counter').forEach(el => {
    const target = parseInt(el.dataset.target) || 0;
    if (el._animated) return;
    el._animated = true;
    gsap.fromTo(el, { innerText: 0 }, {
      innerText: target,
      duration: target > 50 ? 2 : 1.2,
      ease: 'power2.out',
      snap: { innerText: 1 },
      scrollTrigger: { trigger: el, start: 'top 85%', toggleActions: 'play none none none' },
      onUpdate: function() { el.textContent = Math.round(this.targets()[0].innerText); }
    });
  });
}
animateCounters();

// ====== SVG RING ANIMATION ======
ScrollTrigger.create({
  trigger: '#mainRing',
  start: 'top 80%',
  onEnter: () => {
    var pct = ''' + str(total) + ''' / Math.max(''' + str(total) + ''', 30);
    var circumference = 2 * Math.PI * 90;
    document.getElementById('mainRing').style.strokeDashoffset = circumference * (1 - pct);
  }
});

// ====== CLIENT CARDS 3D TILT + MOUSE GLOW ======
document.querySelectorAll('.client-card').forEach(card => {
  card.addEventListener('mousemove', e => {
    const rect = card.getBoundingClientRect();
    const x = (e.clientX - rect.left) / rect.width;
    const y = (e.clientY - rect.top) / rect.height;
    card.style.setProperty('--mx', (x * 100) + '%');
    card.style.setProperty('--my', (y * 100) + '%');
    card.style.transform = 'perspective(800px) rotateX(' + ((y - .5) * -12) + 'deg) rotateY(' + ((x - .5) * 12) + 'deg) translateY(-8px) scale(1.02)';
  });
  card.addEventListener('mouseleave', () => { card.style.transform = ''; });
  card.addEventListener('click', () => {
    document.querySelectorAll('.client-card').forEach(c => c.classList.remove('active'));
    card.classList.add('active');
  });
});

// ====== ANIMATE CLIENT BAR FILLS ======
setTimeout(() => {
  document.querySelectorAll('.client-bar-fill').forEach(bar => {
    bar.style.width = bar.dataset.width + '%';
  });
}, 500);

// ====== BOT TABLE SEARCH ======
document.getElementById('botSearch').addEventListener('input', function() {
  const q = this.value.toLowerCase();
  document.querySelectorAll('#botTable tr').forEach(row => {
    const name = row.querySelector('td').textContent.toLowerCase();
    row.style.display = name.includes(q) ? '' : 'none';
    if (name.includes(q) && q) {
      gsap.fromTo(row, { opacity: 0, x: -10 }, { opacity: 1, x: 0, duration: .3 });
    }
  });
});

// ====== TABLE ROW STAGGER ======
gsap.utils.toArray('#botTable tr').forEach((row, i) => {
  gsap.from(row, {
    opacity: 0, x: -30,
    duration: .5, delay: i * .06,
    ease: 'power2.out',
    scrollTrigger: { trigger: row, start: 'top 90%' }
  });
});

// ====== CHARTS ======
Chart.register(ChartDataLabels);
Chart.defaults.color = '#5A7A9A';
Chart.defaults.font.family = 'Inter';

function createGradient(ctx, colors) {
  var gradient = ctx.createLinearGradient(0, 0, 0, ctx.canvas.height);
  gradient.addColorStop(0, colors[0]);
  gradient.addColorStop(1, colors[1]);
  return gradient;
}

var barCtx = document.getElementById('barC').getContext('2d');
var barColors = [['#0D47A1','#1976D2'],['#1565C0','#2196F3'],['#1976D2','#42A5F5'],
  ['#1E88E5','#64B5F6'],['#2196F3','#90CAF9'],['#1A237E','#3F51B5'],
  ['#283593','#5C6BC0'],['#303F9F','#7986CB'],['#3949AB','#9FA8DA']];
var barBg = [];
for (var bi = 0; bi < ''' + str(num_ac) + '''; bi++) {
  barBg.push(createGradient(barCtx, barColors[bi] || barColors[0]));
}

new Chart(barCtx, {
  type: 'bar',
  data: {
    labels: ''' + cl + ''',
    datasets: [{
      data: ''' + cv + ''',
      backgroundColor: barBg,
      borderColor: ['#1E88E5','#42A5F5','#64B5F6','#90CAF9','#BBDEFB','#5C6BC0','#7986CB','#9FA8DA','#C5CAE9'].slice(0, ''' + str(num_ac) + '''),
      borderWidth: 2, borderRadius: 10, borderSkipped: false
    }]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    animation: { duration: 1800, easing: 'easeOutQuart', delay: ctx => ctx.dataIndex * 200 },
    plugins: {
      legend: { display: false },
      datalabels: { anchor: 'end', align: 'top', color: '#D1D9E6', font: { weight: 700, size: 14 },
        formatter: v => v > 0 ? v : '' },
      tooltip: { backgroundColor: 'rgba(6,14,26,.95)', borderColor: '#1E3A5F',
        borderWidth: 1, cornerRadius: 12, padding: 14, titleFont: { weight: 700 },
        callbacks: { label: ctx => ctx.parsed.y + ' casos (' + Math.round(ctx.parsed.y / ''' + str(max(total, 1)) + ''' * 100) + '%)' } }
    },
    scales: {
      x: { grid: { display: false }, ticks: { font: { size: 11, weight: 600 } } },
      y: { grid: { color: 'rgba(30,58,95,.2)', drawBorder: false }, beginAtZero: true,
        ticks: { stepSize: 5, font: { size: 10 } } }
    }
  },
  plugins: [ChartDataLabels]
});

new Chart(document.getElementById('doC'), {
  type: 'doughnut',
  data: {
    labels: ''' + cl + ''',
    datasets: [{
      data: ''' + cv + ''',
      backgroundColor: ['#0D47A1','#1565C0','#1976D2','#1E88E5','#2196F3','#1A237E','#283593','#303F9F','#3949AB'].slice(0, ''' + str(num_ac) + '''),
      borderColor: '#0A1628', borderWidth: 4, hoverOffset: 20,
      hoverBorderColor: '#48CAE4', hoverBorderWidth: 3
    }]
  },
  options: {
    responsive: true, maintainAspectRatio: false, cutout: '65%',
    animation: { animateRotate: true, duration: 2200, easing: 'easeOutQuart' },
    plugins: {
      legend: { position: 'bottom', labels: { padding: 16, usePointStyle: true, pointStyleWidth: 10,
          font: { size: 11, weight: 500 }, color: '#94A3B8' } },
      datalabels: { color: '#E2E8F0', font: { weight: 700, size: 14 },
        formatter: (v, ctx) => {
          const s = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
          return s > 0 ? Math.round(v / s * 100) + '%' : '';
        }, anchor: 'center', align: 'center' },
      tooltip: { backgroundColor: 'rgba(6,14,26,.95)', borderColor: '#1E3A5F',
        borderWidth: 1, cornerRadius: 12, padding: 14,
        callbacks: { label: ctx => ' ' + ctx.label + ': ' + ctx.parsed + ' casos' } }
    }
  },
  plugins: [ChartDataLabels]
});

new Chart(document.getElementById('botD'), {
  type: 'doughnut',
  data: {
    labels: ['Exitosos', 'Fallidos'],
    datasets: [{ data: [''' + str(bots_ok) + ''', ''' + str(bots_fail) + '''],
      backgroundColor: ['rgba(0,184,148,.85)', 'rgba(225,112,85,.85)'],
      borderWidth: 0, hoverOffset: 10 }]
  },
  options: {
    responsive: true, maintainAspectRatio: false, cutout: '75%',
    animation: { animateRotate: true, duration: 2000, easing: 'easeOutQuart' },
    plugins: { legend: { display: false }, datalabels: { display: false },
      tooltip: { backgroundColor: 'rgba(6,14,26,.95)', borderColor: '#1E3A5F',
        borderWidth: 1, cornerRadius: 12, padding: 14 } }
  }
});

// ====== THEME TOGGLE ======
var darkMode = true;
document.getElementById('themeToggle').addEventListener('click', () => {
  darkMode = !darkMode;
  const root = document.documentElement;
  if (!darkMode) {
    root.style.setProperty('--bg', '#F0F4F8');
    root.style.setProperty('--bg2', '#E2E8F0');
    root.style.setProperty('--bg3', '#CBD5E1');
    root.style.setProperty('--bg-card', 'rgba(255,255,255,.75)');
    root.style.setProperty('--brd', '#D1D9E6');
    root.style.setProperty('--t1', '#1E293B');
    root.style.setProperty('--t2', '#334155');
    root.style.setProperty('--t3', '#64748B');
    root.style.setProperty('--t4', '#94A3B8');
    root.style.setProperty('--glass', 'rgba(0,0,0,.02)');
    root.style.setProperty('--glass2', 'rgba(0,0,0,.04)');
    document.body.style.background = '#F0F4F8';
    document.getElementById('pc').style.opacity = '.08';
    document.querySelector('.ambient').style.opacity = '.3';
  } else {
    root.style.setProperty('--bg', '#060E1A');
    root.style.setProperty('--bg2', '#0A1628');
    root.style.setProperty('--bg3', '#0D1B2A');
    root.style.setProperty('--bg-card', 'rgba(18,35,61,.65)');
    root.style.setProperty('--brd', '#1E3A5F');
    root.style.setProperty('--t1', '#FFFFFF');
    root.style.setProperty('--t2', '#D1D9E6');
    root.style.setProperty('--t3', '#5A7A9A');
    root.style.setProperty('--t4', '#3D5A80');
    root.style.setProperty('--glass', 'rgba(255,255,255,.03)');
    root.style.setProperty('--glass2', 'rgba(255,255,255,.06)');
    document.body.style.background = '#060E1A';
    document.getElementById('pc').style.opacity = '.3';
    document.querySelector('.ambient').style.opacity = '1';
  }
  const btn = document.getElementById('themeToggle');
  btn.innerHTML = darkMode ? '<i data-lucide="sun"><\\/i>' : '<i data-lucide="moon"><\\/i>';
  lucide.createIcons();
  gsap.fromTo(btn, { rotation: 0 }, { rotation: 360, duration: .5, ease: 'power2.out' });
});

// ====== KPI CARD HOVER ======
document.querySelectorAll('.kpi-mini, .op-kpi').forEach(card => {
  card.addEventListener('mouseenter', () => { gsap.to(card, { scale: 1.03, duration: .3, ease: 'power2.out' }); });
  card.addEventListener('mouseleave', () => { gsap.to(card, { scale: 1, y: 0, duration: .3, ease: 'power2.out' }); });
});

// ====== ENTRANCE ANIMATION ======
const tl = gsap.timeline({ defaults: { ease: 'power3.out' } });
tl.from('.navbar', { y: -60, opacity: 0, duration: .8 })
  .from('.kpi-strip', { y: 40, opacity: 0, duration: .6 }, '-=.4');

// ====== UPDATE FOOTER TIME ======
function updateFooter() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  document.getElementById('footerTime').textContent =
    pad(now.getDate()) + '/' + pad(now.getMonth() + 1) + '/' + now.getFullYear() + ' ' +
    pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
}
setInterval(updateFooter, 1000);
''' + '</' + 'script>' + '''
</body>
</html>'''
    return html


def build_widget(cd, bd):
    total = sum(cd.values())
    bots_ok = sum(1 for b in bd if b['ok'])
    bots_fail = sum(1 for b in bd if not b['ok'])
    total_bots = len(bd)
    rate = round(bots_ok/total_bots*100) if total_bots else 0
    now = datetime.now().strftime('%d/%m/%Y')

    clients_json = json.dumps(cd)
    bots_json = json.dumps([{'name':b['name'],'ok':b['ok']} for b in bd])

    widget = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="theme-color" content="#060E1A">
<meta name="mobile-web-app-capable" content="yes">
<link rel="manifest" href="manifest.json">
<title>NetserGroup Widget</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#060E1A;--bg2:#0A1628;--brd:#1E3A5F;--a:#00B4D8;--a2:#0077B6;--a3:#48CAE4;
--ok:#00B894;--fail:#E17055;--warn:#FDCB6E;--t1:#FFF;--t2:#D1D9E6;--t3:#5A7A9A;--t4:#3D5A80}
html,body{height:100%;overflow:hidden}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--t1);
display:flex;flex-direction:column;-webkit-user-select:none;user-select:none}
#pc{position:fixed;top:0;left:0;width:100%;height:100%;z-index:0;pointer-events:none;opacity:.25}
.wrap{position:relative;z-index:1;flex:1;display:flex;flex-direction:column;padding:12px;gap:10px;overflow-y:auto;-webkit-overflow-scrolling:touch}
.hdr{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;
background:linear-gradient(135deg,#0D1B2A,#1B2838);border-radius:14px;border:1px solid var(--brd);position:relative;overflow:hidden}
.hdr::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
background:linear-gradient(90deg,var(--a),var(--a2),var(--a3),var(--a));background-size:300% 100%;animation:sh 4s ease infinite}
@keyframes sh{0%{background-position:0 50%}50%{background-position:100% 50%}100%{background-position:0 50%}}
.hdr-left{display:flex;align-items:center;gap:10px}
.hdr-logo{width:32px;height:32px}
.hdr h2{font-size:14px;font-weight:800;background:linear-gradient(90deg,#fff,#B8C4D6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.hdr-time{font-size:9px;color:var(--t3)}
.pulse-dot{width:7px;height:7px;background:var(--ok);border-radius:50%;animation:pu 2s infinite;margin-right:6px}
@keyframes pu{0%,100%{box-shadow:0 0 0 0 rgba(0,184,148,.4)}50%{box-shadow:0 0 0 6px rgba(0,184,148,0)}}
.live-badge{display:flex;align-items:center;padding:4px 10px;background:rgba(0,184,148,.08);border:1px solid rgba(0,184,148,.2);border-radius:8px;font-size:9px;color:var(--ok);font-weight:600}
.kpi-strip{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.kpi{background:linear-gradient(145deg,#12233D,#162B4A);border:1px solid var(--brd);border-radius:14px;
padding:14px 10px;text-align:center;animation:fadeUp .5s ease-out both}
@keyframes fadeUp{from{opacity:0;transform:translateY(15px)}to{opacity:1;transform:translateY(0)}}
.kpi:nth-child(2){animation-delay:.1s}.kpi:nth-child(3){animation-delay:.2s}
.kpi-icon{font-size:16px;margin-bottom:4px}
.kpi-val{font-size:28px;font-weight:900;line-height:1;margin-bottom:3px;font-variant-numeric:tabular-nums}
.kpi-val.cyan{color:var(--a)}.kpi-val.green{color:var(--ok)}.kpi-val.warn{color:var(--warn)}
.kpi-label{font-size:8px;color:var(--t3);text-transform:uppercase;letter-spacing:1.2px;font-weight:600}
.section-title{font-size:11px;font-weight:700;color:var(--t2);padding:4px 4px 0;display:flex;align-items:center;gap:6px}
.section-title .bar{width:3px;height:14px;background:linear-gradient(180deg,var(--a),var(--a2));border-radius:2px}
.clients-scroll{display:flex;gap:8px;overflow-x:auto;padding:4px 0 8px;-webkit-overflow-scrolling:touch;scrollbar-width:none}
.clients-scroll::-webkit-scrollbar{display:none}
.cl-card{min-width:100px;flex-shrink:0;background:linear-gradient(145deg,#12233D,#162B4A);
border:1px solid var(--brd);border-radius:12px;padding:12px 10px;text-align:center;animation:fadeUp .4s ease-out both;transition:all .3s}
.cl-card.top{border-color:var(--a);background:linear-gradient(145deg,#0D2847,#163456)}
.cl-icon{font-size:16px;margin-bottom:4px;opacity:.7}
.cl-val{font-size:24px;font-weight:800;color:var(--t1);line-height:1;margin-bottom:3px}
.cl-val.zero{color:var(--t4);font-size:18px}
.cl-name{font-size:9px;color:#7B8DA6;font-weight:500;white-space:nowrap}
.cl-pct{font-size:8px;color:var(--a3);font-weight:700;margin-top:2px}
.bots-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:6px}
.bot-item{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;
background:rgba(18,35,61,.6);border:1px solid var(--brd);border-radius:10px;animation:slideIn .3s ease-out both}
@keyframes slideIn{from{opacity:0;transform:translateX(-10px)}to{opacity:1;transform:translateX(0)}}
.bot-name{font-size:10px;color:var(--t2);font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:85px}
.bot-ok{padding:2px 8px;border-radius:6px;font-size:9px;font-weight:700}
.bot-ok.ok{background:rgba(0,184,148,.12);color:var(--ok);border:1px solid rgba(0,184,148,.2)}
.bot-ok.fail{background:rgba(225,112,85,.12);color:var(--fail);border:1px solid rgba(225,112,85,.2)}
.rate-section{display:flex;align-items:center;gap:16px;padding:14px 16px;
background:linear-gradient(145deg,#12233D,#162B4A);border:1px solid var(--brd);border-radius:14px}
.rate-ring{width:56px;height:56px;position:relative;flex-shrink:0}
.rate-ring svg{transform:rotate(-90deg)}
.rate-ring .val{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:14px;font-weight:900;color:var(--ok)}
.rate-info{flex:1}
.rate-title{font-size:11px;font-weight:700;color:var(--t2);margin-bottom:2px}
.rate-sub{font-size:9px;color:var(--t3)}
.nav{display:flex;justify-content:center;gap:12px;padding:10px 0 6px}
.nav a{padding:8px 20px;border-radius:10px;font-size:11px;font-weight:600;text-decoration:none;transition:all .3s;border:1px solid var(--brd);color:var(--t3)}
.nav a.active{background:rgba(0,180,216,.1);border-color:var(--a);color:var(--a)}
.install-banner{display:none;padding:10px 16px;background:linear-gradient(135deg,#0D2847,#163456);
border:1px solid var(--a);border-radius:12px;text-align:center;animation:fadeUp .5s ease-out}
.install-banner p{font-size:11px;color:var(--t2);margin-bottom:6px}
.install-btn{padding:8px 24px;background:linear-gradient(135deg,var(--a2),var(--a));border:none;
border-radius:8px;color:white;font-weight:700;font-size:12px;cursor:pointer;font-family:'Inter',sans-serif}
</style>
</head>
<body>
<canvas id="pc"></canvas>
<div class="wrap">
<div class="hdr">
  <div class="hdr-left">
    <svg class="hdr-logo" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <defs><linearGradient id="gG" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" style="stop-color:#0077B6"/><stop offset="100%" style="stop-color:#00B4D8"/>
      </linearGradient></defs>
      <circle cx="50" cy="50" r="28" fill="url(#gG)" opacity="0.8"/>
      <ellipse cx="50" cy="50" rx="34" ry="14" fill="none" stroke="#00B4D8" stroke-width="1.5" opacity="0.6" transform="rotate(-20 50 50)"/>
      <ellipse cx="50" cy="50" rx="30" ry="9" fill="none" stroke="#48CAE4" stroke-width="1.5" opacity="0.5" transform="rotate(75 50 50)"/>
    </svg>
    <div><h2>NETSERGROUP</h2><div class="hdr-time" id="clock"></div></div>
  </div>
  <div class="live-badge"><span class="pulse-dot"></span>LIVE</div>
</div>
<div class="kpi-strip">
  <div class="kpi"><div class="kpi-icon">&#128202;</div><div class="kpi-val cyan" id="w-total">0</div><div class="kpi-label">Casos</div></div>
  <div class="kpi"><div class="kpi-icon">&#129302;</div><div class="kpi-val green" id="w-bots">0</div><div class="kpi-label">Bots OK</div></div>
  <div class="kpi"><div class="kpi-icon">&#9889;</div><div class="kpi-val warn" id="w-rate">0%</div><div class="kpi-label">Tasa</div></div>
</div>
<div class="rate-section">
  <div class="rate-ring">
    <svg width="56" height="56" viewBox="0 0 56 56">
      <circle cx="28" cy="28" r="24" fill="none" stroke="#1E3A5F" stroke-width="4"/>
      <circle cx="28" cy="28" r="24" fill="none" stroke="#00B894" stroke-width="4" stroke-dasharray="150.8" id="rate-circle" stroke-linecap="round"/>
    </svg>
    <div class="val" id="rate-val">0%</div>
  </div>
  <div class="rate-info"><div class="rate-title">Estado Global de Bots</div><div class="rate-sub" id="rate-detail">0 / 0 bots operativos</div></div>
</div>
<div class="section-title"><div class="bar"></div>Casos por Cliente</div>
<div class="clients-scroll" id="clients-area"></div>
<div class="section-title"><div class="bar"></div>Estado de Bots</div>
<div class="bots-grid" id="bots-area"></div>
<div class="install-banner" id="install-banner">
  <p>Instala NetserGroup Widget en tu pantalla de inicio</p>
  <button class="install-btn" id="install-btn">Instalar Widget</button>
</div>
<div class="nav"><a href="index.html">Dashboard</a><a href="widget.html" class="active">Widget</a></div>
</div>
''' + '<' + 'script>' + '''
const DATA={clients:''' + clients_json + ''',bots:''' + bots_json + ''',updated:"''' + now + '''"};
const ICONS={"HP Comercial":"\\u{1F5A8}","HPE":"\\u{1F5A8}","Payless":"\\u{1F6D2}","Netapp":"\\u2601",
"Lexmark":"\\u{1F5A8}","Lexmark Kit":"\\u{1F5A8}","CTDI":"\\u{1F527}","Monthly Fee":"\\u{1F4B2}","Lenovo":"\\u{1F4BB}"};
const cv=document.getElementById('pc'),cx=cv.getContext('2d');
cv.width=innerWidth;cv.height=innerHeight;
const pts=[];for(let i=0;i<30;i++)pts.push({x:Math.random()*cv.width,y:Math.random()*cv.height,
vx:(Math.random()-.5)*.3,vy:(Math.random()-.5)*.3,r:Math.random()*1.5+.5,o:Math.random()*.3+.1});
function drw(){cx.clearRect(0,0,cv.width,cv.height);
pts.forEach((p,i)=>{p.x+=p.vx;p.y+=p.vy;
if(p.x<0)p.x=cv.width;if(p.x>cv.width)p.x=0;if(p.y<0)p.y=cv.height;if(p.y>cv.height)p.y=0;
cx.fillStyle='rgba(0,180,216,'+p.o+')';cx.beginPath();cx.arc(p.x,p.y,p.r,0,Math.PI*2);cx.fill();
pts.forEach((q,j)=>{if(j<=i)return;const d=Math.hypot(p.x-q.x,p.y-q.y);
if(d<100){cx.strokeStyle='rgba(0,180,216,'+(0.05*(1-d/100))+')';cx.lineWidth=.5;
cx.beginPath();cx.moveTo(p.x,p.y);cx.lineTo(q.x,q.y);cx.stroke()}})});
requestAnimationFrame(drw)}drw();
addEventListener('resize',()=>{cv.width=innerWidth;cv.height=innerHeight});
function updateClock(){const d=new Date();
document.getElementById('clock').textContent=d.toLocaleDateString('es',{day:'2-digit',month:'short',year:'numeric'})+' '+d.toLocaleTimeString('es',{hour:'2-digit',minute:'2-digit',second:'2-digit'});}
updateClock();setInterval(updateClock,1000);
const total=Object.values(DATA.clients).reduce((a,b)=>a+b,0);
const botsOk=DATA.bots.filter(b=>b.ok).length;
const rate=DATA.bots.length?Math.round(botsOk/DATA.bots.length*100):0;
function animateVal(el,target,suffix){suffix=suffix||'';let c=0;const inc=Math.max(target/40,.5);
(function u(){c+=inc;if(c<target){el.textContent=Math.floor(c)+suffix;requestAnimationFrame(u)}
else el.textContent=target+suffix})();}
animateVal(document.getElementById('w-total'),total);
animateVal(document.getElementById('w-bots'),botsOk);
animateVal(document.getElementById('w-rate'),rate,'%');
const circle=document.getElementById('rate-circle');
const circumference=2*Math.PI*24;
circle.style.strokeDasharray=circumference;circle.style.strokeDashoffset=circumference;
setTimeout(()=>{circle.style.transition='stroke-dashoffset 1.5s ease-out';
circle.style.strokeDashoffset=circumference*(1-rate/100);},300);
document.getElementById('rate-val').textContent=rate+'%';
document.getElementById('rate-detail').textContent=botsOk+' / '+DATA.bots.length+' bots operativos';
const clientsArea=document.getElementById('clients-area');
const sorted=Object.entries(DATA.clients).sort((a,b)=>b[1]-a[1]);
sorted.forEach(([name,val],i)=>{const pct=total?Math.round(val/total*100):0;
const card=document.createElement('div');card.className='cl-card'+(i===0&&val>0?' top':'');
card.style.animationDelay=(0.05+i*0.05)+'s';
card.innerHTML='<div class="cl-icon">'+(ICONS[name]||'\\u{1F4CB}')+'<'+'/div>'+
'<div class="cl-val'+(val===0?' zero':'')+'">'+val+'<'+'/div>'+
'<div class="cl-name">'+name+'<'+'/div>'+(val>0?'<div class="cl-pct">'+pct+'%<'+'/div>':'');
clientsArea.appendChild(card);});
const botsArea=document.getElementById('bots-area');
DATA.bots.forEach((b,i)=>{const item=document.createElement('div');item.className='bot-item';
item.style.animationDelay=(0.03+i*0.04)+'s';
item.innerHTML='<span class="bot-name">'+b.name+'<'+'/span>'+
'<span class="bot-ok '+(b.ok?'ok':'fail')+'">'+(b.ok?'\\u2714':'\\u2718')+'<'+'/span>';
botsArea.appendChild(item);});
let deferredPrompt;window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferredPrompt=e;
document.getElementById('install-banner').style.display='block';});
document.getElementById('install-btn').addEventListener('click',async()=>{if(!deferredPrompt)return;
deferredPrompt.prompt();const r=await deferredPrompt.userChoice;
if(r.outcome==='accepted')document.getElementById('install-banner').style.display='none';deferredPrompt=null;});
if('serviceWorker' in navigator){navigator.serviceWorker.register('sw.js').catch(()=>{});}
''' + '</' + 'script>' + '''
</body></html>'''
    return widget

def main():
    try:
        print("Leyendo Excel...")
        rec = read_data()
        cd = get_clients(rec)
        bd = get_bots(rec)
        html = build_html(cd, bd)
        with open('index.html','w',encoding='utf-8') as f: f.write(html)
        widget = build_widget(cd, bd)
        with open('widget.html','w',encoding='utf-8') as f: f.write(widget)
        total = sum(cd.values())
        bots_ok = sum(1 for b in bd if b['ok'])
        print(f"\n>> Dashboard generado: index.html (UI v2.0 - GSAP + AOS + Lucide)")
        print(f">> Widget generado: widget.html")
        print(f"   Casos: {total} | Bots: {bots_ok}/{len(bd)}")
        for n,v in cd.items(): print(f"   {n}: {v}")
    except FileNotFoundError:
        print("ERROR: No se encontro Reporte_NetserGroup_Final.xlsx")
        return False
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback; traceback.print_exc()
        return False
    return True

if __name__=='__main__': main()
