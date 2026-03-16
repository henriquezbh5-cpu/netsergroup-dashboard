#!/usr/bin/env python3
"""
NetserGroup Dashboard Generator
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

    # Client cards
    cards = ''
    for i,(n,v) in enumerate(cs):
        pct = round(v/total*100) if total else 0
        icon = ICONS.get(n,'&#128203;')
        act = ' active' if i==0 and v>0 else ''
        zero = ' zero' if v==0 else ''
        pp = '<div class="cp">'+str(pct)+'%</div>' if v>0 else ''
        cards += '<div class="cc'+act+'" style="animation-delay:'+str(0.1+i*0.06)+'s">'
        cards += '<div class="ci">'+icon+'</div>'
        cards += '<div class="cv'+zero+' counter" data-target="'+str(v)+'">0</div>'
        cards += '<div class="cn">'+n+'</div>'+pp
        cards += '<div class="cb"><div class="cbf" style="width:'+str(pct)+'%"></div></div></div>\n'

    # Bot rows
    rows = ''
    for i,b in enumerate(bd):
        s = '<span class="sok">✔ OK</span>' if b['ok'] else '<span class="sfail">✘ FAIL</span>'
        rows += '<tr style="animation-delay:'+str(0.05+i*0.06)+'s"><td>'+b['name']+'</td><td>'+s+'</td></tr>\n'

    html = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>NetserGroup Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#060E1A;--bg2:#0A1628;--brd:#1E3A5F;--a:#00B4D8;--a2:#0077B6;--a3:#48CAE4;
--ok:#00B894;--fail:#E17055;--warn:#FDCB6E;--t1:#FFF;--t2:#D1D9E6;--t3:#5A7A9A;--t4:#3D5A80}
html{scroll-behavior:smooth}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--t1);min-height:100vh;overflow-x:hidden}
#pc{position:fixed;top:0;left:0;width:100%;height:100%;z-index:0;pointer-events:none;opacity:.35}
.d{position:relative;z-index:1;max-width:1440px;margin:0 auto;padding:20px 28px 40px}

/* HEADER */
.h{display:flex;justify-content:space-between;align-items:center;padding:18px 28px;
background:linear-gradient(135deg,#0D1B2A,#1B2838);border-radius:18px;border:1px solid var(--brd);
margin-bottom:24px;position:relative;overflow:hidden;animation:sd .8s ease-out;backdrop-filter:blur(12px)}
.h::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
background:linear-gradient(90deg,var(--a),var(--a2),var(--a3),var(--a));
background-size:300% 100%;animation:sh 4s ease infinite}
@keyframes sh{0%{background-position:0 50%}50%{background-position:100% 50%}100%{background-position:0 50%}}
@keyframes sd{from{opacity:0;transform:translateY(-30px)}to{opacity:1;transform:translateY(0)}}
.hl{display:flex;align-items:center;gap:16px}
.logo{width:52px;height:52px;animation:fl 3s ease-in-out infinite}
@keyframes fl{0%,100%{transform:translateY(0)}50%{transform:translateY(-8px)}}
.ht h1{font-size:24px;font-weight:800;letter-spacing:-.5px;background:linear-gradient(90deg,#fff,#B8C4D6);
-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.ht p{font-size:12px;color:var(--t3);margin-top:3px}
.hr{display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.hs{text-align:center;padding:10px 22px;background:rgba(0,180,216,.06);border-radius:12px;
border:1px solid rgba(0,180,216,.12);transition:all .3s}
.hs:hover{background:rgba(0,180,216,.12);transform:translateY(-2px)}
.hs .sv{font-size:26px;font-weight:800;color:var(--a)}
.hs .sl{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:1.5px;margin-top:2px}
.hs.dg .sv{color:var(--fail)}.hs.wn .sv{color:var(--warn)}
.bd{padding:8px 16px;background:rgba(0,184,148,.08);border:1px solid rgba(0,184,148,.2);
border-radius:10px;font-size:11px;color:var(--ok);display:flex;align-items:center;gap:8px}
.pl{width:8px;height:8px;background:var(--ok);border-radius:50%;animation:pu 2s infinite}
@keyframes pu{0%,100%{box-shadow:0 0 0 0 rgba(0,184,148,.4)}50%{box-shadow:0 0 0 8px rgba(0,184,148,0)}}

/* SECTIONS */
.st{display:flex;align-items:center;gap:10px;margin:24px 0 16px;animation:fu .6s ease-out both}
.st .br{width:4px;height:24px;background:linear-gradient(180deg,var(--a),var(--a2));border-radius:2px}
.st h2{font-size:15px;font-weight:600;color:var(--t2);letter-spacing:.3px}
@keyframes fu{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}

/* CLIENT CARDS */
.cg{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:14px;margin-bottom:8px}
.cc{background:linear-gradient(145deg,#12233D,#162B4A);border:1px solid var(--brd);border-radius:16px;
padding:20px 14px 28px;text-align:center;transition:all .4s cubic-bezier(.34,1.56,.64,1);
cursor:pointer;position:relative;overflow:hidden;animation:fu .6s ease-out both}
.cc::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;
background:linear-gradient(90deg,transparent,var(--a3),transparent);opacity:0;transition:opacity .3s}
.cc:hover{transform:translateY(-6px);border-color:var(--a);
box-shadow:0 12px 40px rgba(0,180,216,.15)}
.cc:hover::after{opacity:1}
.cc.active{border-color:var(--a);background:linear-gradient(145deg,#0D2847,#163456)}
.cc.active::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
background:linear-gradient(90deg,var(--a),var(--a3))}
.ci{font-size:22px;margin-bottom:6px;opacity:.7;transition:all .3s}
.cc:hover .ci{opacity:1;transform:scale(1.2)}
.cv{font-size:36px;font-weight:800;color:var(--t1);line-height:1;margin-bottom:6px;font-variant-numeric:tabular-nums}
.cv.zero{color:var(--t4)}.cn{font-size:11px;color:#7B8DA6;font-weight:500}
.cp{position:absolute;top:8px;right:10px;font-size:10px;color:var(--t3);font-weight:600}
.cb{position:absolute;bottom:0;left:14px;right:14px;height:3px;background:rgba(30,58,95,.5);border-radius:2px;overflow:hidden}
.cbf{height:100%;background:linear-gradient(90deg,var(--a),var(--a3));border-radius:2px;
transition:width 1.5s cubic-bezier(.4,0,.2,1)}

/* GRID */
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-bottom:8px}
.g2{display:grid;grid-template-columns:1fr 1.5fr;gap:20px;margin-bottom:8px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.pn{background:linear-gradient(145deg,#12233D,#162B4A);border:1px solid var(--brd);border-radius:18px;
padding:28px;position:relative;overflow:hidden;transition:all .3s;backdrop-filter:blur(8px);animation:fu .7s ease-out both}
.pn:hover{border-color:rgba(0,180,216,.25)}
.pt{font-size:13px;font-weight:600;color:var(--t2);margin-bottom:12px}

/* TOTAL RING */
.tc{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%}
.tr{width:180px;height:180px;border-radius:50%;border:5px solid var(--brd);
display:flex;align-items:center;justify-content:center;flex-direction:column;position:relative;margin:12px 0}
.tr::before{content:'';position:absolute;inset:-5px;border-radius:50%;
border:5px solid transparent;border-top-color:var(--a);border-right-color:var(--a);
border-bottom-color:var(--a2);animation:sp 6s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.tr .rv{font-size:54px;font-weight:900;color:var(--a);line-height:1}
.tr .rl{font-size:12px;color:var(--t3);margin-top:4px;font-weight:500}
.ts{font-size:12px;color:var(--t4);margin-top:8px}

/* CHART */
.ch{position:relative;width:100%;height:260px}

/* TABLE */
.tw{overflow:hidden;animation:fu .7s ease-out both}
table{width:100%;border-collapse:separate;border-spacing:0 5px}
thead th{text-align:left;font-size:10px;color:var(--t3);text-transform:uppercase;
letter-spacing:1.2px;padding:8px 14px;font-weight:600}
thead th:last-child{text-align:center}
tbody tr{background:rgba(26,42,68,.4);transition:all .3s;animation:sl .5s ease-out both}
@keyframes sl{from{opacity:0;transform:translateX(-20px)}to{opacity:1;transform:translateX(0)}}
tbody tr:hover{background:rgba(0,180,216,.08);transform:translateX(4px)}
td{padding:12px 14px;font-size:13px;color:var(--t2)}
td:first-child{border-radius:10px 0 0 10px}td:last-child{border-radius:0 10px 10px 0;text-align:center}
.sok{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;
background:rgba(0,184,148,.1);border:1px solid rgba(0,184,148,.2);border-radius:8px;
font-size:11px;font-weight:600;color:var(--ok);box-shadow:0 0 10px rgba(0,184,148,.15)}
.sfail{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;
background:rgba(225,112,85,.1);border:1px solid rgba(225,112,85,.2);border-radius:8px;
font-size:11px;font-weight:600;color:var(--fail)}

/* KPI */
.kp{background:linear-gradient(145deg,#12233D,#162B4A);border:1px solid var(--brd);border-radius:16px;
padding:24px 16px;text-align:center;position:relative;overflow:hidden;transition:all .3s;animation:fu .7s ease-out both}
.kp:hover{transform:translateY(-4px);box-shadow:0 8px 30px rgba(0,0,0,.3)}
.ki{width:42px;height:42px;border-radius:11px;display:flex;align-items:center;justify-content:center;
font-size:18px;margin:0 auto 12px}
.kp.s .ki{background:rgba(0,184,148,.1);border:1px solid rgba(0,184,148,.2)}
.kp.f .ki{background:rgba(225,112,85,.1);border:1px solid rgba(225,112,85,.2)}
.kp.r .ki{background:rgba(0,180,216,.1);border:1px solid rgba(0,180,216,.2)}
.kp.m .ki{background:rgba(253,203,110,.1);border:1px solid rgba(253,203,110,.2)}
.kv{font-size:36px;font-weight:800;line-height:1;margin-bottom:6px;font-variant-numeric:tabular-nums}
.kp.s .kv{color:var(--ok)}.kp.f .kv{color:var(--fail)}
.kp.r .kv{color:var(--a)}.kp.m .kv{color:var(--warn)}
.kl{font-size:10px;color:#7B8DA6;font-weight:600;text-transform:uppercase;letter-spacing:.8px}
.ks{font-size:10px;color:var(--t4);margin-top:4px}

footer{margin-top:32px;text-align:center;padding:16px;border-top:1px solid var(--brd)}
footer p{font-size:11px;color:var(--t4)}

@media(max-width:1100px){.g2,.g3{grid-template-columns:1fr}.g4{grid-template-columns:repeat(2,1fr)}
.h{flex-direction:column;gap:16px;text-align:center}.hr{justify-content:center}}
@media(max-width:600px){.cg{grid-template-columns:repeat(2,1fr)}.g4{grid-template-columns:1fr 1fr}
.d{padding:12px 14px 30px}}
</style>
</head>
<body>
<canvas id="pc"></canvas>
<div class="d">

<div class="h">
<div class="hl">
  <svg class="logo" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
    <defs><linearGradient id="gG" x1="0%%" y1="0%%" x2="100%%" y2="100%%">
      <stop offset="0%%" style="stop-color:#0077B6"/><stop offset="100%%" style="stop-color:#00B4D8"/>
    </linearGradient></defs>
    <circle cx="50" cy="50" r="28" fill="url(#gG)" opacity="0.8"/>
    <ellipse cx="50" cy="50" rx="34" ry="14" fill="none" stroke="#00B4D8" stroke-width="1.5" opacity="0.6" transform="rotate(-20 50 50)"/>
    <ellipse cx="50" cy="50" rx="37" ry="11" fill="none" stroke="#0077B6" stroke-width="1" opacity="0.4" transform="rotate(30 50 50)"/>
    <ellipse cx="50" cy="50" rx="30" ry="9" fill="none" stroke="#48CAE4" stroke-width="1.5" opacity="0.5" transform="rotate(75 50 50)"/>
    <text x="50" y="88" text-anchor="middle" font-family="Inter,Arial" font-size="7" font-weight="800" fill="white">NETSER</text>
    <text x="50" y="96" text-anchor="middle" font-family="Inter,Arial" font-size="5" fill="#48CAE4">GROUP</text>
  </svg>
  <div class="ht"><h1>NETSERGROUP</h1><p>Dashboard Operacional &mdash; Monitoreo en Tiempo Real</p></div>
</div>
<div class="hr">
  <div class="hs"><div class="sv counter" data-target="'''+str(bots_ok)+'''">0</div><div class="sl">Bots OK</div></div>
  <div class="hs dg"><div class="sv counter" data-target="'''+str(bots_fail)+'''">0</div><div class="sl">Bots Fail</div></div>
  <div class="hs wn"><div class="sv counter" data-target="'''+str(bots_fail)+'''">0</div><div class="sl">Alertas</div></div>
  <div class="bd"><span class="pl"></span>'''+now+'''</div>
</div>
</div>

<div class="st"><div class="br"></div><h2>Casos por Cliente</h2></div>
<div class="cg">
'''+cards+'''</div>

<div class="st"><div class="br"></div><h2>Resumen General</h2></div>
<div class="g3">
  <div class="pn"><div class="tc">
    <div class="tr"><div class="rv counter" data-target="'''+str(total)+'''">0</div><div class="rl">casos activos</div></div>
    <div class="ts">'''+str(active_cl)+''' clientes activos de '''+str(len(cd))+'''</div>
  </div></div>
  <div class="pn"><div class="pt">Distribucion por Cliente</div><div class="ch"><canvas id="barC"></canvas></div></div>
  <div class="pn"><div class="pt">Proporcion de Casos</div><div class="ch"><canvas id="doC"></canvas></div></div>
</div>

<div class="st"><div class="br"></div><h2>Estado de Bots / Flujos</h2></div>
<div class="g2">
  <div class="pn tw"><table><thead><tr><th>Bot / Flujo</th><th>Estado</th></tr></thead><tbody>
'''+rows+'''  </tbody></table></div>
  <div class="pn" style="display:flex;flex-direction:column;align-items:center;justify-content:center">
    <div class="pt">Estado Global de Bots</div>
    <div style="width:200px;height:200px;position:relative">
      <canvas id="botD"></canvas>
      <div style="position:absolute;top:50%%;left:50%%;transform:translate(-50%%,-50%%);text-align:center">
        <div style="font-size:32px;font-weight:800;color:'''+rate_col+'''">'''+str(rate)+'''%%</div>
        <div style="font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:1px">Tasa Exito</div>
      </div>
    </div>
  </div>
</div>

<div class="st"><div class="br"></div><h2>Resumen Operacional</h2></div>
<div class="g4">
  <div class="kp s" style="animation-delay:.3s"><div class="ki">✔</div><div class="kv counter" data-target="'''+str(bots_ok)+'''">0</div><div class="kl">Exitosos</div><div class="ks">'''+str(bots_ok)+''' / '''+str(total_bots)+''' bots</div></div>
  <div class="kp f" style="animation-delay:.4s"><div class="ki">✘</div><div class="kv counter" data-target="'''+str(bots_fail)+'''">0</div><div class="kl">Fallidos</div><div class="ks">'''+str(bots_fail)+''' / '''+str(total_bots)+''' bots</div></div>
  <div class="kp r" style="animation-delay:.5s"><div class="ki">★</div><div class="kv counter" data-target="'''+str(rate)+'''">0</div><div class="kl">Tasa de Exito</div><div class="ks">Rendimiento global</div></div>
  <div class="kp m" style="animation-delay:.6s"><div class="ki">⚡</div><div class="kv counter" data-target="'''+str(total)+'''">0</div><div class="kl">Total Casos</div><div class="ks">Todos los clientes</div></div>
</div>

<footer><p>Ultima Actualizacion: '''+now+''' &mdash; NetserGroup &copy; 2026 &mdash; Humberto Henriquez</p></footer>
</div>

<script>
// PARTICLES WITH NETWORK LINES
const cv=document.getElementById('pc'),cx=cv.getContext('2d');
cv.width=innerWidth;cv.height=innerHeight;
const pts=[];for(let i=0;i<55;i++)pts.push({x:Math.random()*cv.width,y:Math.random()*cv.height,
vx:(Math.random()-.5)*.4,vy:(Math.random()-.5)*.4,r:Math.random()*2+.5,o:Math.random()*.4+.1});
function drw(){cx.clearRect(0,0,cv.width,cv.height);
pts.forEach((p,i)=>{p.x+=p.vx;p.y+=p.vy;
if(p.x<0)p.x=cv.width;if(p.x>cv.width)p.x=0;if(p.y<0)p.y=cv.height;if(p.y>cv.height)p.y=0;
cx.fillStyle=`rgba(0,180,216,${p.o})`;cx.beginPath();cx.arc(p.x,p.y,p.r,0,Math.PI*2);cx.fill();
pts.forEach((q,j)=>{if(j<=i)return;const d=Math.hypot(p.x-q.x,p.y-q.y);
if(d<130){cx.strokeStyle=`rgba(0,180,216,${0.06*(1-d/130)})`;cx.lineWidth=.5;
cx.beginPath();cx.moveTo(p.x,p.y);cx.lineTo(q.x,q.y);cx.stroke()}})});
requestAnimationFrame(drw)}drw();
addEventListener('resize',()=>{cv.width=innerWidth;cv.height=innerHeight});

// COUNTER ANIMATION ON SCROLL
function countUp(el){const t=parseInt(el.dataset.target)||0;if(!t){el.textContent='0';return}
let c=0;const inc=Math.max(t/70,.3);(function u(){c+=inc;if(c<t){el.textContent=Math.floor(c);requestAnimationFrame(u)}
else el.textContent=t})()}
const obs=new IntersectionObserver(en=>{en.forEach(e=>{if(e.isIntersecting){
e.target.querySelectorAll('.counter').forEach(c=>countUp(c));obs.unobserve(e.target)}})},{threshold:.15});
document.querySelectorAll('.cg,.g3,.g4,.h,.g2').forEach(el=>obs.observe(el));

// 3D TILT ON CLIENT CARDS
document.querySelectorAll('.cc').forEach(c=>{
c.addEventListener('mousemove',e=>{const r=c.getBoundingClientRect();
const x=(e.clientX-r.left)/r.width*2-1,y=(e.clientY-r.top)/r.height*2-1;
c.style.transform=`perspective(800px) rotateX(${y*-10}deg) rotateY(${x*10}deg) translateY(-6px) translateZ(15px)`});
c.addEventListener('mouseleave',()=>{c.style.transform=''});
c.addEventListener('click',()=>{document.querySelectorAll('.cc').forEach(x=>x.classList.remove('active'));c.classList.add('active')})});

// CHARTS
Chart.register(ChartDataLabels);
Chart.defaults.color='#5A7A9A';Chart.defaults.font.family='Inter';

// Bar chart
new Chart(document.getElementById('barC'),{type:'bar',
data:{labels:'''+cl+''',datasets:[{data:'''+cv+''',
backgroundColor:['#0D47A1','#1565C0','#1976D2','#1E88E5','#2196F3','#1A237E','#283593','#303F9F','#3949AB'].slice(0,'''+str(len(ac))+'''),
borderColor:['#1565C0','#1976D2','#1E88E5','#2196F3','#42A5F5','#283593','#303F9F','#3949AB','#5C6BC0'].slice(0,'''+str(len(ac))+'''),
borderWidth:2,borderRadius:8,borderSkipped:false}]},
options:{responsive:true,maintainAspectRatio:false,
animation:{duration:1500,easing:'easeOutQuart',delay:function(ctx){return ctx.dataIndex*150}},
plugins:{legend:{display:false},datalabels:{anchor:'end',align:'top',color:'#D1D9E6',
font:{weight:700,size:13},formatter:v=>v>0?v:''},
tooltip:{backgroundColor:'rgba(4,11,20,.95)',borderColor:'#1E3A5F',borderWidth:1,cornerRadius:10,padding:12}},
scales:{x:{grid:{display:false},ticks:{font:{size:10,weight:500}}},
y:{grid:{color:'rgba(30,58,95,.3)'},beginAtZero:true}}},plugins:[ChartDataLabels]});

// Doughnut - corporate dark colors
new Chart(document.getElementById('doC'),{type:'doughnut',
data:{labels:'''+cl+''',datasets:[{data:'''+cv+''',
backgroundColor:['#0D47A1','#1565C0','#1976D2','#1E88E5','#2196F3','#1A237E','#283593','#303F9F','#3949AB'].slice(0,'''+str(len(ac))+'''),
borderColor:'#060E1A',borderWidth:3,hoverOffset:16,hoverBorderColor:'#48CAE4'}]},
options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
animation:{animateRotate:true,duration:2000,easing:'easeOutQuart'},
plugins:{legend:{position:'bottom',labels:{padding:14,usePointStyle:true,pointStyleWidth:10,
font:{size:11,weight:500},color:'#94A3B8'}},
datalabels:{color:'#E2E8F0',font:{weight:700,size:13},
formatter:(v,ctx)=>{const s=ctx.chart.data.datasets[0].data.reduce((a,b)=>a+b,0);return s>0?Math.round(v/s*100)+'%':''},
anchor:'center',align:'center'},
tooltip:{backgroundColor:'rgba(4,11,20,.95)',borderColor:'#1E3A5F',borderWidth:1,cornerRadius:10,padding:12}}},
plugins:[ChartDataLabels]});

// Bot donut
new Chart(document.getElementById('botD'),{type:'doughnut',
data:{labels:['Exitosos','Fallidos'],datasets:[{data:['''+str(bots_ok)+''','''+str(bots_fail)+'''],
backgroundColor:['rgba(0,184,148,.8)','rgba(225,112,85,.8)'],borderWidth:0,hoverOffset:8}]},
options:{responsive:true,maintainAspectRatio:false,cutout:'72%',
animation:{animateRotate:true,duration:1800},
plugins:{legend:{display:false},datalabels:{display:false},
tooltip:{backgroundColor:'rgba(4,11,20,.95)',borderColor:'#1E3A5F',borderWidth:1,cornerRadius:10,padding:12}}}});
</script>
</body>
</html>'''
    return html

def main():
    try:
        print("Leyendo Excel...")
        rec = read_data()
        cd = get_clients(rec)
        bd = get_bots(rec)
        html = build_html(cd, bd)
        with open('index.html','w',encoding='utf-8') as f: f.write(html)
        total = sum(cd.values())
        bots_ok = sum(1 for b in bd if b['ok'])
        print(f"\n>> Dashboard generado: index.html")
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
