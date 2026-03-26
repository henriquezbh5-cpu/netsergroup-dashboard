#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NetserGroup Dashboard Generator v3.0
Lee Reporte_NetserGroup_Final.xlsx y genera index.html futurista.
Solo muestra datos que EXISTEN en el Excel, nada inventado.
"""
import json, os, sys
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(SCRIPT_DIR, "Reporte_NetserGroup_Final.xlsx")
OUTPUT = os.path.join(SCRIPT_DIR, "index.html")

CLIENTS = ["HP Comercial","HPE","Payless","Netapp","Lexmark","Lexmark Kit","CTDI","Monthly Fee","Lenovo"]
BOTS = ["BackUp Mobility","Cierre POs","Cierre Alpha","Cierre HPCM","Tasas Cambio","Encuestas Dell",
        "Respaldo Invoice","Cierre Residencias","Receiving Lab","Reporte Inv HP","HPCM Cenam",
        "HPCM Chile","Licencias FSM","Regularizacion Mobility"]

def safe_int(v):
    if v is None: return 0
    try: return int(float(v))
    except: return 0

def is_ok(v):
    if v is None: return False
    return str(v).strip() in ("\u2714","\u2714\ufe0f","OK","ok","1","TRUE","True","true","\u2713")

def read_data():
    wb = openpyxl.load_workbook(EXCEL, data_only=True, read_only=True)
    ws = wb["Datos"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows or rows[0][0] is None:
        print("ERROR: Sin datos en hoja Datos"); sys.exit(1)
    r = rows[0]
    fecha = str(r[0] or "")
    cases = {c: safe_int(r[1+i]) for i, c in enumerate(CLIENTS)}
    total = safe_int(r[10]) or sum(cases.values())
    update = str(r[11] or fecha)
    bots = {b: is_ok(r[12+i]) for i, b in enumerate(BOTS)}
    wb.close()
    return {"fecha": fecha, "update": update, "total": total,
            "cases": cases, "bots": bots,
            "botsOK": sum(1 for v in bots.values() if v),
            "botsFail": sum(1 for v in bots.values() if not v)}

def generate_html(D):
    # Active clients sorted by cases desc
    active = sorted([(c,v) for c,v in D["cases"].items() if v > 0], key=lambda x: -x[1])
    all_clients = sorted(D["cases"].items(), key=lambda x: -x[1])
    bots_ok = D["botsOK"]
    bots_fail = D["botsFail"]
    bots_total = bots_ok + bots_fail
    tasa = round(bots_ok/bots_total*100, 1) if bots_total > 0 else 0

    # JSON data for charts
    chart_labels = json.dumps([c for c,_ in active])
    chart_values = json.dumps([v for _,v in active])

    # Bot grid HTML
    bot_pills = ""
    for name, ok in D["bots"].items():
        cls = "ok" if ok else "fail"
        icon = "&#10003;" if ok else "&#10007;"
        status_text = "OK" if ok else "FAIL"
        bot_pills += f'<div class="bot-pill {cls}"><span class="bot-icon">{icon}</span><span class="bot-name">{name}</span><span class="bot-tag {cls}">{status_text}</span></div>\n'

    html = f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>NetserGroup — Centro de Operaciones</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
:root{{
  --bg:#0a0e1a;--bg2:#0f1629;--card:rgba(15,23,42,0.65);
  --border:rgba(0,212,255,0.12);--border-hover:rgba(0,212,255,0.3);
  --cyan:#00d4ff;--green:#00f5a0;--red:#ff6b6b;--gold:#fbbf24;--purple:#a78bfa;
  --txt:#e2e8f0;--txt2:#94a3b8;--txt3:#64748b;
}}
body{{font-family:'Inter',system-ui,sans-serif;background:linear-gradient(135deg,var(--bg),var(--bg2));color:var(--txt);min-height:100vh;overflow-x:hidden}}
.wrap{{max-width:1400px;margin:0 auto;padding:20px}}

/* HEADER */
.header{{display:flex;align-items:center;justify-content:space-between;padding:18px 28px;background:var(--card);backdrop-filter:blur(20px);border:1px solid var(--border);border-radius:16px;margin-bottom:20px}}
.logo{{font-size:22px;font-weight:700;letter-spacing:2px}}
.logo span{{color:var(--cyan)}}
.logo-sub{{font-size:12px;color:var(--txt3);margin-top:2px}}
.header-right{{text-align:right;font-size:12px;color:var(--txt2)}}
.clock{{font-size:18px;font-weight:600;color:var(--txt);font-variant-numeric:tabular-nums}}
.status-badge{{display:inline-flex;align-items:center;gap:6px;background:rgba(0,245,160,0.1);border:1px solid rgba(0,245,160,0.3);border-radius:20px;padding:5px 14px;font-size:11px;color:var(--green);font-weight:500}}
.status-dot{{width:7px;height:7px;border-radius:50%;background:var(--green);animation:pulse 2s infinite}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}

/* KPI CARDS */
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}}
.kpi{{background:var(--card);backdrop-filter:blur(20px);border:1px solid var(--border);border-radius:14px;padding:20px 22px;position:relative;overflow:hidden;transition:all .3s}}
.kpi:hover{{border-color:var(--border-hover);transform:translateY(-2px);box-shadow:0 8px 30px rgba(0,0,0,0.3)}}
.kpi-accent{{position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}}
.kpi-val{{font-size:36px;font-weight:700;line-height:1.1;margin-top:6px}}
.kpi-label{{font-size:11px;color:var(--txt3);font-weight:500;letter-spacing:1px;text-transform:uppercase;margin-top:6px}}
.kpi-icon{{position:absolute;top:16px;right:18px;font-size:18px;opacity:.4}}

/* SECTION */
.section-title{{display:flex;align-items:center;gap:10px;font-size:15px;font-weight:600;margin:24px 0 14px;color:var(--txt)}}
.section-title::before{{content:'';width:4px;height:22px;background:var(--cyan);border-radius:4px}}

/* CHARTS */
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.panel{{background:var(--card);backdrop-filter:blur(20px);border:1px solid var(--border);border-radius:14px;padding:22px;transition:border-color .3s}}
.panel:hover{{border-color:var(--border-hover)}}
.panel-title{{font-size:12px;font-weight:600;color:var(--txt2);letter-spacing:1px;text-transform:uppercase;margin-bottom:16px}}
.chart-box{{position:relative;height:320px}}
.donut-center{{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;pointer-events:none}}
.donut-center .num{{font-size:42px;font-weight:700;color:var(--cyan)}}
.donut-center .lab{{font-size:11px;color:var(--txt3);text-transform:uppercase;letter-spacing:1px}}

/* BOTS */
.bots-summary{{display:flex;align-items:center;gap:20px;margin-bottom:16px;flex-wrap:wrap}}
.bots-count{{font-size:18px;font-weight:700}}
.bots-bar{{flex:1;min-width:200px;height:8px;background:rgba(255,255,255,0.06);border-radius:8px;overflow:hidden}}
.bots-bar-fill{{height:100%;border-radius:8px;transition:width 1.5s ease}}
.bots-tasa{{font-size:36px;font-weight:700;text-align:right}}
.bots-tasa-label{{font-size:10px;color:var(--txt3);text-transform:uppercase;letter-spacing:1px}}
.bots-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:10px;margin-top:16px}}
.bot-pill{{display:flex;align-items:center;gap:10px;background:rgba(15,23,42,0.5);border:1px solid rgba(255,255,255,0.05);border-radius:10px;padding:10px 14px;transition:all .25s}}
.bot-pill:hover{{border-color:var(--border-hover);background:rgba(15,23,42,0.8)}}
.bot-pill.ok .bot-icon{{color:var(--green);text-shadow:0 0 8px rgba(0,245,160,0.5)}}
.bot-pill.fail .bot-icon{{color:var(--red);text-shadow:0 0 8px rgba(255,107,107,0.5)}}
.bot-icon{{font-size:16px;font-weight:700;width:24px;text-align:center}}
.bot-name{{flex:1;font-size:12px;font-weight:500;color:var(--txt2)}}
.bot-tag{{font-size:10px;font-weight:600;padding:3px 10px;border-radius:6px;letter-spacing:.5px}}
.bot-tag.ok{{background:rgba(0,245,160,0.12);color:var(--green)}}
.bot-tag.fail{{background:rgba(255,107,107,0.12);color:var(--red)}}

/* FOOTER */
.footer{{text-align:center;padding:30px 0 20px;font-size:11px;color:var(--txt3)}}

/* PARTICLES */
#particles{{position:fixed;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:0}}
.wrap{{position:relative;z-index:1}}

/* ANIMATIONS */
@keyframes fadeUp{{from{{opacity:0;transform:translateY(20px)}}to{{opacity:1;transform:translateY(0)}}}}
.anim{{animation:fadeUp .6s ease both}}
.d1{{animation-delay:.1s}}.d2{{animation-delay:.2s}}.d3{{animation-delay:.3s}}.d4{{animation-delay:.4s}}.d5{{animation-delay:.5s}}

/* RESPONSIVE */
@media(max-width:900px){{
  .kpis{{grid-template-columns:repeat(2,1fr)}}
  .charts-grid{{grid-template-columns:1fr}}
  .header{{flex-direction:column;gap:12px;text-align:center}}
  .header-right{{text-align:center}}
}}
@media(max-width:500px){{
  .kpis{{grid-template-columns:1fr}}
  .chart-box{{height:260px}}
  .bots-grid{{grid-template-columns:1fr}}
}}
</style>
</head>
<body>
<canvas id="particles"></canvas>
<div class="wrap">

<!-- HEADER -->
<div class="header anim d1">
  <div>
    <div class="logo">NETSER<span>GROUP</span></div>
    <div class="logo-sub">Centro de Operaciones</div>
  </div>
  <div class="status-badge"><span class="status-dot"></span> Sistema Operativo</div>
  <div class="header-right">
    <div class="clock" id="clock"></div>
    <div>Actualizado: {D["update"]}</div>
  </div>
</div>

<!-- KPI CARDS -->
<div class="kpis anim d2">
  <div class="kpi">
    <div class="kpi-accent" style="background:var(--cyan)"></div>
    <div class="kpi-icon">&#9776;</div>
    <div class="kpi-val" style="color:var(--cyan)">{D["total"]}</div>
    <div class="kpi-label">Total Casos</div>
  </div>
  <div class="kpi">
    <div class="kpi-accent" style="background:var(--green)"></div>
    <div class="kpi-icon">&#10003;</div>
    <div class="kpi-val" style="color:var(--green)">{bots_ok}</div>
    <div class="kpi-label">Bots OK</div>
  </div>
  <div class="kpi">
    <div class="kpi-accent" style="background:var(--red)"></div>
    <div class="kpi-icon">&#10007;</div>
    <div class="kpi-val" style="color:var(--red)">{bots_fail}</div>
    <div class="kpi-label">Bots Fail</div>
  </div>
  <div class="kpi">
    <div class="kpi-accent" style="background:var(--purple)"></div>
    <div class="kpi-icon">%</div>
    <div class="kpi-val" style="color:var(--purple)">{tasa}%</div>
    <div class="kpi-label">Tasa Exito</div>
  </div>
</div>

<!-- CHARTS -->
<div class="section-title anim d3">Analisis de Casos</div>
<div class="charts-grid anim d3">
  <div class="panel">
    <div class="panel-title">Casos por Cliente</div>
    <div class="chart-box"><canvas id="barChart"></canvas></div>
  </div>
  <div class="panel" style="position:relative">
    <div class="panel-title">Distribucion por Cliente</div>
    <div class="donut-center"><div class="num">{D["total"]}</div><div class="lab">Total</div></div>
    <div class="chart-box"><canvas id="donutChart"></canvas></div>
  </div>
</div>

<!-- BOTS -->
<div class="section-title anim d4">Estado de Bots / Flujos Automatizados</div>
<div class="panel anim d4">
  <div class="bots-summary">
    <div>
      <div class="bots-count">{bots_ok}/{bots_total} Exitosos</div>
      <div style="font-size:12px;color:var(--txt3)">{'Todos los flujos operando con normalidad' if bots_fail==0 else str(bots_fail)+' flujo(s) requieren atencion'}</div>
    </div>
    <div class="bots-bar"><div class="bots-bar-fill" style="width:{tasa}%;background:{'var(--green)' if tasa==100 else 'var(--gold)'}"></div></div>
    <div style="text-align:right">
      <div class="bots-tasa" style="color:{'var(--green)' if tasa==100 else 'var(--gold)'}">{tasa}%</div>
      <div class="bots-tasa-label">Tasa Exito</div>
    </div>
  </div>
  <div class="bots-grid">
    {bot_pills}
  </div>
</div>

<!-- FOOTER -->
<div class="footer anim d5">NetserGroup &copy; 2026 &mdash; Humberto Henriquez</div>

</div>

<script>
// Clock
(function tick(){{
  var d=new Date();
  document.getElementById('clock').textContent=d.toLocaleTimeString('es-SV',{{hour:'2-digit',minute:'2-digit',second:'2-digit'}});
  setTimeout(tick,1000);
}})();

// Particles
(function(){{
  var c=document.getElementById('particles'),x=c.getContext('2d');
  var w,h,pts=[];
  function resize(){{w=c.width=window.innerWidth;h=c.height=window.innerHeight;}}
  resize();window.addEventListener('resize',resize);
  for(var i=0;i<60;i++)pts.push({{x:Math.random()*w,y:Math.random()*h,vx:(Math.random()-.5)*.3,vy:(Math.random()-.5)*.3,r:Math.random()*1.5+.5}});
  function draw(){{
    x.clearRect(0,0,w,h);
    for(var i=0;i<pts.length;i++){{
      var p=pts[i];
      p.x+=p.vx;p.y+=p.vy;
      if(p.x<0||p.x>w)p.vx*=-1;
      if(p.y<0||p.y>h)p.vy*=-1;
      x.beginPath();x.arc(p.x,p.y,p.r,0,Math.PI*2);x.fillStyle='rgba(0,212,255,0.15)';x.fill();
      for(var j=i+1;j<pts.length;j++){{
        var q=pts[j],dx=p.x-q.x,dy=p.y-q.y,dist=Math.sqrt(dx*dx+dy*dy);
        if(dist<120){{x.beginPath();x.moveTo(p.x,p.y);x.lineTo(q.x,q.y);x.strokeStyle='rgba(0,212,255,'+(0.06*(1-dist/120))+')';x.stroke();}}
      }}
    }}
    requestAnimationFrame(draw);
  }}
  draw();
}})();

// Charts
(function(){{
  var labels={chart_labels};
  var values={chart_values};
  var colors=['#00d4ff','#00f5a0','#fbbf24','#a78bfa','#ff6b6b','#38bdf8','#34d399','#f472b6','#818cf8','#fb923c'];
  var bg=labels.map(function(_,i){{return colors[i%colors.length]}});

  Chart.defaults.color='#94a3b8';
  Chart.defaults.font.family="'Inter',sans-serif";
  Chart.defaults.font.size=11;

  // Bar Chart
  new Chart(document.getElementById('barChart'),{{
    type:'bar',
    data:{{labels:labels,datasets:[{{data:values,backgroundColor:bg.map(function(c){{return c+'CC'}}),borderColor:bg,borderWidth:1,borderRadius:6,barPercentage:.7}}]}},
    options:{{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      animation:{{duration:1200,easing:'easeOutQuart'}},
      plugins:{{legend:{{display:false}},tooltip:{{
        backgroundColor:'rgba(10,14,26,0.95)',borderColor:'rgba(0,212,255,0.2)',borderWidth:1,cornerRadius:8,padding:12,
        callbacks:{{label:function(c){{var t={D["total"]};return c.parsed.x+' casos ('+(t>0?Math.round(c.parsed.x/t*100):0)+'%)'}}}}
      }}}},
      scales:{{
        x:{{grid:{{color:'rgba(0,212,255,0.06)'}},beginAtZero:true,ticks:{{font:{{size:10}}}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:11,weight:'500'}},color:'#cbd5e1'}}}}
      }}
    }}
  }});

  // Donut Chart
  new Chart(document.getElementById('donutChart'),{{
    type:'doughnut',
    data:{{labels:labels,datasets:[{{data:values,backgroundColor:bg.map(function(c){{return c+'CC'}}),borderColor:'rgba(15,23,42,0.8)',borderWidth:2,hoverBorderColor:bg,hoverOffset:8}}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,cutout:'65%',
      animation:{{duration:1200,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{position:'bottom',labels:{{padding:12,usePointStyle:true,pointStyle:'circle',font:{{size:10}},color:'#cbd5e1'}}}},
        tooltip:{{
          backgroundColor:'rgba(10,14,26,0.95)',borderColor:'rgba(0,212,255,0.2)',borderWidth:1,cornerRadius:8,padding:12,
          callbacks:{{label:function(c){{var t={D["total"]};return ' '+c.label+': '+c.parsed+' ('+(t>0?Math.round(c.parsed/t*100):0)+'%)'}}}}
        }}
      }}
    }}
  }});
}})();
</script>
</body>
</html>'''
    return html

def main():
    print("=" * 50)
    print("  NetserGroup Dashboard Generator v3.0")
    print("=" * 50)
    D = read_data()
    print(f"  Total: {D['total']} casos | Bots: {D['botsOK']}/{D['botsOK']+D['botsFail']} OK")
    html = generate_html(D)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  index.html generado ({len(html):,} bytes)")
    print(f"  Archivo: {OUTPUT}")

if __name__ == "__main__":
    main()
