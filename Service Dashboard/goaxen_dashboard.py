"""
Go-axen daily dashboard.

Drop the two Excel files into the Dropbox folder (any file names):
  1. Check-in / Prospect file  (has an 'Overall' sheet)
  2. Garage Billed List file   (has a 'Leads Data' sheet)

Run:   python goaxen_dashboard.py "C:\\Users\\<you>\\Dropbox\\GoAxen Drop"
Output (written into the same folder, under  _output/):
  dashboard.html     - open in any browser
"""
import sys, os, re, glob, html, datetime as dt
from collections import defaultdict
import openpyxl

UPLOAD_ONLY = '--upload-only' in sys.argv   # builds a page with no data embedded (for hosting; users upload files in the browser)
_args = [a for a in sys.argv[1:] if not a.startswith('--')]
FOLDER = _args[0] if _args else os.path.join(os.path.expanduser('~'), 'Dropbox', 'GoAxen Drop')
OUT = os.path.join(FOLDER, '_output'); os.makedirs(OUT, exist_ok=True)

GMAP = {'r.k automobiles': 'R.K.Automobiles', 'r.k automobiles ': 'R.K.Automobiles', 'global motors': 'Global Motors'}
SMAP = {'car wash': 'Car Wash - Doorstep', 'car wash - doorstep': 'Car Wash - Doorstep', 'ac service': 'AC Service',
        'spa': 'Complete Car Spa - Doorstep', 'rubbing polish': 'Rubbing Polish - Doorstep'}
STATUSES = ['Delivered', 'Under Service', 'Empty Return', 'Vehicle not reported', 'Inspection Only']

def norm_service(s):
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    return SMAP.get(s.lower(), s)

def norm_garage(g):
    g = re.sub(r'\s+', ' ', str(g or '')).strip()
    return GMAP.get(g.lower(), g)

def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None

def read_sheet_rows(wb, name):
    if name not in wb.sheetnames:
        return None
    ws = wb[name]
    blank = 0; rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            hdr = [str(h).strip() if h is not None else '' for h in row]
            continue
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row[:14]):
            blank += 1
            if blank > 200: break
            continue
        blank = 0
        rows.append(row)
    return hdr, rows

def load_checkin(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    res = read_sheet_rows(wb, 'Overall')
    if not res: return None
    hdr, rows = res
    ix = {h: i for i, h in enumerate(hdr)}
    seen = set(); recs = []
    for r in rows:
        bid = r[ix['Booking ID']]
        if bid is None or bid in seen: continue
        seen.add(bid)
        d = r[ix['Go-axen date']]
        if not isinstance(d, (dt.datetime, dt.date)): continue
        d = d if isinstance(d, dt.datetime) else dt.datetime.combine(d, dt.time())
        recs.append(dict(date=d, service=norm_service(r[ix['service Type']]),
                         garage=norm_garage(r[ix['Workshop']]),
                         status=str(r[ix['Vehicle Delivery Status']] or '').strip(),
                         invoice=num(r[ix['Invoice Amount']])))
    return fix_years(recs)

def load_billed(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    res = read_sheet_rows(wb, 'Leads Data')
    if not res: return None
    hdr, rows = res
    # positions (the file has two 'Status' columns: 8 = billed status, 11 = delivery status)
    seen = set(); recs = []
    for r in rows:
        bid = r[1]
        if bid is None or bid in seen: continue
        seen.add(bid)
        d = r[0]
        if not isinstance(d, (dt.datetime, dt.date)): continue
        d = d if isinstance(d, dt.datetime) else dt.datetime.combine(d, dt.time())
        recs.append(dict(date=d, garage=norm_garage(r[2]), service=norm_service(r[7]),
                         billed=str(r[8] or '').strip().lower() == 'billed',
                         value=num(r[12]) or 0.0))
    return fix_years(recs)

def fix_years(recs):
    if not recs: return recs
    from collections import Counter
    yr = Counter(r['date'].year for r in recs).most_common(1)[0][0]
    for r in recs:
        if r['date'].year != yr:
            try: r['date'] = r['date'].replace(year=yr)
            except ValueError: pass
    return recs

def fmt_rs(x):
    x = float(x or 0)
    s = f'{abs(x):,.0f}'
    # Indian grouping
    if len(s.replace(',', '')) > 3:
        digits = s.replace(',', '')
        head, tail = digits[:-3], digits[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:]); head = head[:-2]
        if head: parts.insert(0, head)
        s = ','.join(parts) + ',' + tail
    return ('-' if x < 0 else '') + 'Rs ' + s


import json

TEMPLATE = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Services Report</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--ink:#1d2433;--mut:#6b7280;--line:#e5e7eb;--brand:#1e3a8a;--card:#fff;--bg:#f5f7fa}
@media (prefers-color-scheme: dark){:root:not([data-theme=light]){--ink:#e5e7eb;--mut:#9ca3af;--line:#374151;--brand:#2563eb;--card:#1f2937;--bg:#111827}}
*{box-sizing:border-box}
body{margin:0;font:14px/1.45 system-ui,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink)}
header{padding:18px 24px;background:linear-gradient(90deg,#1e3a8a,#2563eb,#60a5fa,#93c5fd);color:#fff}
header h1{margin:0;font-size:20px} header p{margin:4px 0 0;opacity:.95;font-size:13px}
main{padding:14px 20px;max-width:1280px;margin:auto}
nav.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 14px}
nav.tabs button{border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:999px;cursor:pointer;font:inherit}
nav.tabs button.on{background:var(--brand);color:#fff;border-color:var(--brand)}
.bar-ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
.bar-ctl label{font-size:12px;color:var(--mut)}
select,input[type=search],input[type=date]{font:inherit;padding:6px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
input[type=search]{min-width:220px}
.btn{font:inherit;padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);cursor:pointer}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px;margin-bottom:14px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:14px}
.kpi{border-top:4px solid var(--c,#2563eb)}
.kpi .l{color:var(--mut);font-size:11px;text-transform:uppercase;letter-spacing:.04em}
.kpi .v{font-size:21px;font-weight:700;margin-top:2px}
.kpi .s{color:var(--mut);font-size:12px}
h2{font-size:15px;margin:18px 0 8px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:10px;background:var(--card)}
table{width:100%;border-collapse:collapse}
th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
th{font-size:12px;color:var(--mut);font-weight:600;cursor:pointer;user-select:none}
th.sorted::after{content:" \25B4"} th.sorted.desc::after{content:" \25BE"}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
.pos{color:#1d4ed8;font-weight:600} .neg{color:#1e3a8a;font-weight:600} .mut{color:var(--mut)}
.chart{overflow-x:auto;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:8px}
.chart svg text{fill:var(--ink);font-size:11px}
.chart svg .ax{fill:var(--mut)}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}
.mini{border-left:6px solid var(--c,#2563eb)} .mini h3{font-size:13px;margin:0 0 4px} .mini .t{color:var(--mut);font-size:12px;margin-bottom:4px}
.drop{border:2px dashed var(--brand);border-radius:12px;padding:28px;text-align:center;background:var(--card);cursor:pointer}
.drop.over{background:#dbeafe}
.files{font-size:13px} .files li{margin:3px 0}
.hidden{display:none}
.note{color:var(--mut);font-size:12px;margin-top:8px}
.legend{font-size:12px;color:var(--mut);margin:0 0 4px 40px}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:4px;vertical-align:middle}
</style></head><body>
<header><h1>Services Report</h1><p id="asof"></p></header>
<main>
<nav class="tabs" id="tabs">
 <button data-t="dropbox" class="on">Drop Box</button>
 <button data-t="overview">Overview</button>
 <button data-t="daily">Date-wise</button>
 <button data-t="compare">Custom comparison</button>
 <button data-t="tables">Garage &amp; service tables</button>
 <button data-t="editable">Editable table</button>
</nav>

<section id="s-overview" class="hidden">
 <div class="bar-ctl">
  <label>Month <select id="f-month"></select></label>
  <label>Garage <select id="f-garage"></select></label>
  <label>Service <select id="f-service"></select></label>
  <label>From <input type="date" id="f-from"></label>
  <label>To <input type="date" id="f-to"></label>
  <label><input type="search" id="f-search" placeholder="Search garage or service..."></label>
  <button class="btn" id="f-reset">Reset</button>
 </div>
 <div class="kpis" id="kpis"></div>
 <h2>Bookings per day (colour = weekday)</h2>
 <div class="chart" id="ch-overview"></div>
</section>

<section id="s-daily" class="hidden">
 <h2>Date-wise report</h2>
 <div class="chart" id="ch-daily"></div>
 <div class="tw" style="margin-top:10px"><table id="t-daily"></table></div>
</section>

<section id="s-compare" class="hidden">
 <div class="bar-ctl">
  <label>Month A <select id="c-a"></select></label>
  <label>vs Month B <select id="c-b"></select></label>
  <label>Garage <select id="c-garage"></select></label>
  <label>Service <select id="c-service"></select></label>
 </div>
 <div class="kpis" id="c-kpis"></div>
 <h2>By service</h2>
 <div class="legend" id="lg-cmp-svc"></div>
 <div class="chart" id="ch-cmp-svc"></div>
 <div class="tw" style="margin-top:10px"><table id="t-cmp-svc"></table></div>
 <h2>By garage</h2>
 <div class="legend" id="lg-cmp-gar"></div>
 <div class="chart" id="ch-cmp-gar"></div>
 <div class="tw" style="margin-top:10px"><table id="t-cmp-gar"></table></div>
</section>

<section id="s-tables" class="hidden">
 <div class="bar-ctl">
  <label>View <select id="tb-view"><option value="garage">By garage</option><option value="service">By service</option></select></label>
  <label><input type="search" id="tb-search" placeholder="Search..."></label>
 </div>
 <div class="tw"><table id="t-tables"></table></div>
</section>

<section id="s-dropbox">
 <div class="card">
  <h2 style="margin-top:0">Upload reports</h2>
  <p class="mut">Drop the Check-in (Prospect) file and the Garage Billed List file here. Each upload replaces that file type for the months it covers and is remembered in this browser.</p>
  <div class="drop" id="drop">Drag the .xlsx reports here, or click to choose files
    <input type="file" id="file" accept=".xlsx" multiple hidden></div>
  <p class="note" id="upstatus"></p>
 </div>
 <div class="card">
  <h2 style="margin-top:0">Files loaded</h2>
  <ul class="files" id="dropfiles"></ul>
  <p class="note" id="dropfolder"></p>
 </div>
</section>
<section id="s-editable" class="hidden">
 <div class="card">
  <div class="bar-ctl">
   <label>Month <select id="e-month"></select></label>
   <label>Rows are <select id="e-dim"><option value="garage">Garages</option><option value="service">Services</option></select></label>
   <label>Add row <select id="e-addrow"></select></label><button class="btn" id="e-addrow-btn">Add row</button>
   <button class="btn" id="e-allrows">Load all</button>
  </div>
  <div class="bar-ctl">
   <label>New formula column <input id="e-colname" placeholder="Column name"></label>
   <label><input id="e-formula" placeholder="e.g. [Billed value] / [Billed leads]" style="min-width:300px"></label>
   <button class="btn" id="e-addcol">Add column</button>
  </div>
  <p class="note">Fields you can use in formulas: [Bookings] [Delivered] [Under service] [Empty return] [Invoices] [Invoiced] [Leads] [Billed leads] [Billed value] [Share]. Operators: + - * / ( ). Example: [Billed value] / [Billed leads].</p>
  <div class="bar-ctl">
   <label>Highlight cells where <select id="r-col"></select></label>
   <select id="r-op"><option>&gt;</option><option>&gt;=</option><option>&lt;</option><option>&lt;=</option><option>=</option></select>
   <input type="number" id="r-val" value="0" style="width:110px">
   <label>colour <select id="r-color"><option value="good">Light blue</option><option value="alert">Dark blue</option></select></label>
   <button class="btn" id="r-add">Add rule</button>
  </div>
  <div id="rules" class="note"></div>
 </div>
 <div class="tw"><table id="t-editable"></table></div>
 <p class="note">Row names, column names and formulas can be edited in place. Settings are saved in this browser.</p>
</section>
<p class="note">Booking IDs are counted once. Customer names and phone numbers are not shown.</p>
</main>

<script>
const PAL=['#1e3a8a','#1d4ed8','#2563eb','#3b82f6','#60a5fa','#93c5fd','#1e40af','#0ea5e9','#38bdf8','#7dd3fc','#0369a1','#1e293b','#334155','#475569'];
const WD=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const STATUSES=["Delivered","Under Service","Empty Return","Vehicle not reported","Inspection Only"];
const MN=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const EMBED = __DATA__;
const SRC_FILES = __FILES__;
const FOLDER = __FOLDER__;
const STORE_KEY = 'goaxen_uploads_v2';
let DATA = [];
let uploads = [];            // [{label, src, months, recs}] applied in order
let garages=[], services=[], months=[], dates=[];
const $ = id => document.getElementById(id);
const esc = s => String(s==null?"":s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
function fmtN(x){ return (x==null||isNaN(x))?"-":Math.round(x).toLocaleString("en-IN"); }
function fmtRs(x){ if(x==null||isNaN(x)) return "-"; const v=Math.round(x); const neg=v<0; let s=String(Math.abs(v)); if(s.length>3){ let h=s.slice(0,-3), t=s.slice(-3), p=[]; while(h.length>2){p.unshift(h.slice(-2)); h=h.slice(0,-2);} if(h) p.unshift(h); s=p.join(",")+","+t; } return (neg?"-":"")+"Rs "+s; }
function fmtPct(a,b){ if(!b) return '<span class="mut">n/a</span>'; const p=(a-b)/b*100; return `<span class="${p>=0?'pos':'neg'}">${p>=0?'+':''}${p.toFixed(0)}%</span>`; }
function dLabel(iso){ const [y,m,d]=iso.split("-"); return `${+d} ${MN[+m-1]}`; }
function mLabel(ym){ const [y,m]=ym.split("-"); return `${MN[+m-1]} ${y}`; }
function wdOf(iso){ const [y,m,d]=iso.split("-").map(Number); return new Date(y,m-1,d).getDay(); }
function fill(sel, items, allLabel, keep){ const v=keep??sel.value; sel.innerHTML = (allLabel!=null?`<option value="">${allLabel}</option>`:"") + items.map(i=>`<option value="${esc(i)}">${esc(i)}</option>`).join(""); if([...sel.options].some(o=>o.value===v)) sel.value=v; }

// ---------- data handling ----------
function applyUpload(u){
  const ms=new Set(u.months);
  DATA = DATA.filter(r=>!(r.src===u.src && ms.has(r.m))).concat(u.recs);
}
function rebuildFromAll(){
  DATA=[]; EMBED.forEach(r=>DATA.push(r));
  uploads.forEach(applyUpload);
  refreshLists();
}
function refreshLists(){
  garages=[...new Set(DATA.map(r=>r.g))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  services=[...new Set(DATA.map(r=>r.s))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  months=[...new Set(DATA.map(r=>r.m))].sort();
  dates=[...new Set(DATA.map(r=>r.d))].sort();
  fill($("f-garage"), garages, "All garages"); fill($("f-service"), services, "All services");
  fill($("c-garage"), garages, "All garages"); fill($("c-service"), services, "All services");
  $("f-month").innerHTML = `<option value="">All months</option>` + months.map(m=>`<option value="${m}">${mLabel(m)}</option>`).join("");
  const prevA=$("c-a").value, prevB=$("c-b").value;
  $("c-a").innerHTML = months.map(m=>`<option value="${m}">${mLabel(m)}</option>`).join("");
  $("c-b").innerHTML = months.map(m=>`<option value="${m}">${mLabel(m)}</option>`).join("");
  if(months.length){ $("c-a").value = months.includes(prevA) ? prevA : months[Math.max(0,months.length-2)]; $("c-b").value = months.includes(prevB) ? prevB : months[months.length-1] }
  if(dates.length){ $("f-from").min=$("f-to").min=dates[0]; $("f-from").max=$("f-to").max=dates[dates.length-1]; if(!$("f-from").value) $("f-from").value=dates[0]; if(!$("f-to").value) $("f-to").value=dates[dates.length-1]; }
  const last=dates[dates.length-1]||"";
  $("asof").textContent = last ? `Data as of ${dLabel(last)} ${last.slice(0,4)} · months loaded: ${months.map(mLabel).join(", ")}` : "No data loaded yet";
}

// ---------- in-browser parsing of uploaded reports (same rules as the Python script) ----------
const GM={'r.k automobiles':'R.K.Automobiles','global motors':'Global Motors'};
const SM={'car wash':'Car Wash - Doorstep','car wash - doorstep':'Car Wash - Doorstep','ac service':'AC Service','spa':'Complete Car Spa - Doorstep','rubbing polish':'Rubbing Polish - Doorstep'};
const cleanTxt = v => String(v==null?"":v).replace(/\s+/g," ").trim();
const nGarage = v => { const t=cleanTxt(v); return GM[t.toLowerCase()]||t; };
const nService = v => { const t=cleanTxt(v); return SM[t.toLowerCase()]||t; };
function toDate(v){
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v==="number"){ return new Date(Math.round((v-25569)*86400*1000)); }  // UTC-based, read with UTC getters
  if(typeof v==="string"){ const d=new Date(v); return isNaN(d)?null:d; }
  return null;
}
// midday trick: SheetJS dates sit at local midnight (or UTC midnight); +12h makes the calendar day unambiguous in any timezone
const isoOf = d => { const t=new Date(d.getTime()+12*3600*1000); return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,"0")}-${String(t.getUTCDate()).padStart(2,"0")}`; };
function fixYears(recs){
  if(!recs.length) return recs;
  const cnt={}; recs.forEach(r=>{ const y=+r.d.slice(0,4); cnt[y]=(cnt[y]||0)+1; });
  const yr=+Object.keys(cnt).sort((a,b)=>cnt[b]-cnt[a])[0];
  recs.forEach(r=>{ if(+r.d.slice(0,4)!==yr){ r.d=yr+r.d.slice(4); r.m=r.d.slice(0,7); } });
  return recs;
}
function parseCheckin(ws){
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
  const hdr=rows[0].map(h=>cleanTxt(h));
  const I=n=>hdr.indexOf(n);
  const iB=I('Booking ID'), iD=I('Go-axen date'), iS=I('service Type'), iW=I('Workshop'), iSt=I('Vehicle Delivery Status'), iI=I('Invoice Amount');
  if([iB,iD,iS,iW,iSt].some(x=>x<0)) throw new Error("not a check-in file");
  const seen=new Set(), out=[];
  for(const r of rows.slice(1)){
    const id=r[iB]; if(id==null||seen.has(id)) continue; seen.add(id);
    const d=toDate(r[iD]); if(!d) continue;
    const iso=isoOf(d);
    out.push({d:iso,m:iso.slice(0,7),g:nGarage(r[iW]),s:nService(r[iS]),src:'c',st:cleanTxt(r[iSt]),inv:typeof r[iI]==="number"?r[iI]:0});
  }
  return fixYears(out);
}
function parseBilled(ws){
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
  const seen=new Set(), out=[];
  for(const r of rows.slice(1)){
    const id=r[1]; if(id==null||seen.has(id)) continue; seen.add(id);
    const d=toDate(r[0]); if(!d) continue;
    const iso=isoOf(d);
    out.push({d:iso,m:iso.slice(0,7),g:nGarage(r[2]),s:nService(r[7]),src:'b',bl:cleanTxt(r[8]).toLowerCase()==='billed'?1:0,v:typeof r[12]==="number"?r[12]:0});
  }
  return fixYears(out);
}
function readFile(file){
  return new Promise((res,rej)=>{ const fr=new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsArrayBuffer(file); });
}
async function processFiles(files){
  const msgs=[];
  for(const f of files){
    try{
      const wb=XLSX.read(await readFile(f),{type:"array",cellDates:true});
      let u=null;
      if(wb.Sheets['Overall']){ const recs=parseCheckin(wb.Sheets['Overall']); u={label:f.name,src:'c',recs}; }
      else if(wb.Sheets['Leads Data']){ const recs=parseBilled(wb.Sheets['Leads Data']); u={label:f.name,src:'b',recs}; }
      else { msgs.push(`${f.name}: not a check-in or billed-list file`); continue; }
      u.months=[...new Set(u.recs.map(r=>r.m))];
      uploads=uploads.filter(x=>!(x.src===u.src && x.months.some(m=>u.months.includes(m))) ).concat([]);
      uploads.push(u);
      msgs.push(`${f.name}: ${u.recs.length} rows (${u.src==='c'?'check-in':'billed list'}), months ${u.months.map(mLabel).join(", ")}`);
    }catch(e){ msgs.push(`${f.name}: could not read (${e.message||e})`); }
  }
  try{ localStorage.setItem(STORE_KEY, JSON.stringify(uploads)); }catch(e){ msgs.push("Could not remember uploads in this browser (storage blocked)."); }
  rebuildFromAll(); renderUploadLog(msgs); renderActive();
}
function renderUploadLog(msgs){
  const base=SRC_FILES.map(f=>`<li>${esc(f)}</li>`).join("");
  const up=uploads.map(u=>`<li><b>Uploaded:</b> ${esc(u.label)} · ${u.recs.length} rows · ${u.src==='c'?'check-in':'billed list'} · ${u.months.map(mLabel).map(esc).join(", ")}</li>`).join("");
  $("dropfiles").innerHTML = up + base;
  $("upstatus").innerHTML = msgs.map(m=>esc(m)).join("<br>");
}

// ---------- aggregation ----------
function passes(r,o){
  if(o.m && r.m!==o.m) return false;
  if(o.g && r.g!==o.g) return false;
  if(o.s && r.s!==o.s) return false;
  if(o.from && r.d<o.from) return false;
  if(o.to && r.d>o.to) return false;
  if(o.q){ const q=o.q.toLowerCase(); if(!(String(r.g).toLowerCase().includes(q)||String(r.s).toLowerCase().includes(q))) return false; }
  return true;
}
function aggregate(rows, keyFn){
  const M=new Map();
  for(const r of rows){
    const k=keyFn(r); if(k==null) continue;
    if(!M.has(k)) M.set(k,{key:k,bookings:0,delivered:0,under:0,empty:0,inv_n:0,invoiced:0,leads:0,billed_n:0,billed:0});
    const a=M.get(k);
    if(r.src==="c"){ a.bookings++; if(r.st==="Delivered")a.delivered++; if(r.st==="Under Service")a.under++; if(r.st==="Empty Return")a.empty++; if(r.inv){a.inv_n++; a.invoiced+=r.inv;} }
    else { a.leads++; if(r.bl) a.billed_n++; a.billed+=r.v||0; }
  }
  return M;
}
function totals(rows){ const t={bookings:0,delivered:0,under:0,empty:0,inv_n:0,invoiced:0,leads:0,billed_n:0,billed:0}; for(const r of rows){ if(r.src==="c"){t.bookings++; if(r.st==="Delivered")t.delivered++; if(r.st==="Under Service")t.under++; if(r.st==="Empty Return")t.empty++; if(r.inv){t.inv_n++; t.invoiced+=r.inv;}} else {t.leads++; if(r.bl)t.billed_n++; t.billed+=r.v||0;} } return t; }
const METRICS={ bookings:{label:"Bookings",fmt:fmtN,get:a=>a.bookings}, invoiced:{label:"Invoiced",fmt:fmtRs,get:a=>a.invoiced},
  billed:{label:"Billed value",fmt:fmtRs,get:a=>a.billed}, leads:{label:"Leads",fmt:fmtN,get:a=>a.leads} };

// ---------- charts (multi-colour) ----------
function barChart(items, opts={}){
  const step=opts.step||44, H=opts.h||220, pad=36;
  const W = opts.fluid ? items.length*step+pad : Math.max(520, items.length*step);
  const max=Math.max(1,...items.map(i=>i.value||0));
  const bw=(W-pad)/Math.max(1,items.length);
  let g="";
  items.forEach((it,i)=>{
    const h=((it.value||0)/max)*(H-pad-18), x=pad+i*bw+bw*0.15, w=bw*0.7, y=H-pad-h;
    const col=it.color||PAL[i%PAL.length];
    const txt=opts.fmt?opts.fmt(it.value):it.value;
    g+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(0,h).toFixed(1)}" rx="3" fill="${col}"><title>${esc(it.label)}: ${esc(txt)}</title></rect>`;
    if(it.value) g+=`<text x="${(x+w/2).toFixed(1)}" y="${(y-4).toFixed(1)}" text-anchor="middle">${esc(txt)}</text>`;
    g+=`<text class="ax" x="${(x+w/2).toFixed(1)}" y="${H-pad+14}" text-anchor="middle">${esc(it.label)}</text>`;
  });
  const sw = opts.fluid ? 'width="100%"' : `width="${W}"`;
  return `<svg ${sw} height="${H}" viewBox="0 0 ${W} ${H}" role="img">${g}</svg>`;
}
const SERIES_A="#93c5fd", SERIES_B="#1e3a8a";
function groupedChart(labels, sA, sB, nameA, nameB, fmt){
  const W=Math.max(520,labels.length*90), H=390, padL=40, padT=24, padB=130;
  const max=Math.max(1,...sA,...sB);
  const plotH=H-padT-padB, y0=H-padB;
  const gw=(W-padL)/Math.max(1,labels.length), bw=gw*0.32;
  let g="";
  labels.forEach((l,i)=>{
    [[sA[i],SERIES_A,0],[sB[i],SERIES_B,1]].forEach(([v,c,k])=>{
      const h=((v||0)/max)*(plotH-16), x=padL+i*gw+gw*0.2+k*bw, y=y0-h;
      g+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" rx="2" fill="${c}"><title>${esc(fmt(v))}</title></rect>`;
      if(v) g+=`<text x="${(x+bw/2).toFixed(1)}" y="${(y-4).toFixed(1)}" text-anchor="middle">${esc(fmt(v))}</text>`;
    });
    const lx=padL+i*gw+gw/2;
    g+=`<text class="ax" transform="translate(${lx.toFixed(1)},${y0+14}) rotate(-45)" text-anchor="end">${esc(l)}</text>`;
  });
  g+=`<line x1="${padL}" x2="${W}" y1="${y0}" y2="${y0}" stroke="var(--line)"/>`;
  const legend=`<span class="sw" style="background:${SERIES_A}"></span>${esc(nameA)} &nbsp; <span class="sw" style="background:${SERIES_B}"></span>${esc(nameB)}`;
  return {legend, svg:`<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${g}</svg>`};
}
function setGrouped(chartId, legId, labels, a, b, na, nb, fmt){
  const r=groupedChart(labels,a,b,na,nb,fmt);
  $(chartId).innerHTML = labels.length? r.svg : '<p class="mut">No data.</p>';
  $(legId).innerHTML = r.legend;
}

// ---------- sortable tables ----------
const CF_METRICS=[["bookings","Bookings"],["delivered","Delivered"],["under","Under service"],["empty","Empty return"],["inv_n","Invoices"],["invoiced","Invoiced"],["leads","Leads"],["billed_n","Billed leads"],["billed","Billed value"],["share","Share of billed"]];
const cfName=k=>{ const m=CF_METRICS.find(x=>x[0]===k); return m?m[1]:k; };
// rule.col holds a metric key; older saved rules used "f_<metric>", so both forms are accepted
function cfStyle(key, v){
  if(!Number.isFinite(v)) return "";
  for(const rule of EDIT.rules){
    if(String(rule.col).replace(/^f_/,"")===key && ruleHit(rule,v)){ const k=RULE_COLORS[rule.color]||RULE_COLORS.good; return `background:${k.bg};color:${k.fg};font-weight:600`; }
  }
  return "";
}
function makeTable(el, cols, rows){
  let sortKey=null, desc=true;
  const useCf = el.id==="t-tables"; // conditional formatting: Garage & service tables only (plus the Editable table)
  const draw=()=>{
    const data=[...rows];
    if(sortKey){ data.sort((a,b)=>{ const x=a[sortKey], y=b[sortKey]; const r=(typeof x==="number"&&typeof y==="number")?x-y:String(x).localeCompare(String(y)); return desc?-r:r; }); }
    el.innerHTML=`<thead><tr>${cols.map(c=>`<th class="${c.num?'n':''} ${sortKey===c.k?'sorted '+(desc?'desc':''):''}" data-k="${c.k}">${esc(c.label)}</th>`).join("")}</tr></thead>
      <tbody>${data.map(r=>`<tr>${cols.map(c=>`<td class="${c.num?'n':''}" style="${useCf?cfStyle(c.cf||c.k, r[c.k]):''}">${c.html?c.html(r):esc(c.fmt?c.fmt(r[c.k]):r[c.k])}</td>`).join("")}</tr>`).join("")}</tbody>`;
    el.querySelectorAll("th").forEach(th=>th.onclick=()=>{ const k=th.dataset.k; if(sortKey===k) desc=!desc; else {sortKey=k; desc=true;} draw(); });
  };
  draw();
}
const pctCell = (a,b) => fmtPct(a,b);

// ---------- filters ----------
function currentFilter(){ return {m:$("f-month").value, g:$("f-garage").value, s:$("f-service").value, from:$("f-from").value, to:$("f-to").value, q:$("f-search").value.trim()}; }

// ---------- renderers ----------
function renderOverview(){
  const o=currentFilter(); const rows=DATA.filter(r=>passes(r,o));
  const t=totals(rows); const pct=t.bookings?(t.delivered/t.bookings*100):0;
  const kp=[["Bookings",fmtN(t.bookings),`${fmtN(t.delivered)} delivered (${pct.toFixed(0)}%)`,"#2563eb"],
            ["Under service",fmtN(t.under),`${fmtN(t.empty)} empty return`,"#60a5fa"],
            ["Invoiced",fmtRs(t.invoiced),`${fmtN(t.inv_n)} invoices`,"#3b82f6"],
            ["Leads",fmtN(t.leads),`${fmtN(t.billed_n)} billed`,"#1e40af"],
            ["Billed value",fmtRs(t.billed),`avg ${fmtRs(t.billed_n?t.billed/t.billed_n:0)} per billed lead`,"#0369a1"]];
  $("kpis").innerHTML=kp.map(([l,v,s,c])=>`<div class="card kpi" style="--c:${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");
  const A=[...aggregate(rows,r=>r.d).values()].sort((a,b)=>a.key.localeCompare(b.key));
  $("ch-overview").innerHTML = A.length ? barChart(A.map(a=>({label:dLabel(a.key),value:a.bookings,color:PAL[wdOf(a.key)]})),{step:36}) : '<p class="mut">No data for this filter.</p>';
}
function renderDaily(){
  const o=currentFilter(); const rows=DATA.filter(r=>passes(r,o));
  const A=[...aggregate(rows,r=>r.d).values()].sort((a,b)=>a.key.localeCompare(b.key));
  $("ch-daily").innerHTML = A.length? barChart(A.map(a=>({label:dLabel(a.key),value:a.bookings,color:PAL[wdOf(a.key)]})),{step:36}) : '<p class="mut">No data.</p>';
  makeTable($("t-daily"),[
    {k:"key",label:"Date",fmt:x=>dLabel(x)},{k:"bookings",label:"Bookings",num:1,fmt:fmtN},{k:"delivered",label:"Delivered",num:1,fmt:fmtN},
    {k:"under",label:"Under service",num:1,fmt:fmtN},{k:"empty",label:"Empty return",num:1,fmt:fmtN},
    {k:"invoiced",label:"Invoiced",num:1,fmt:fmtRs},{k:"leads",label:"Leads",num:1,fmt:fmtN},{k:"billed",label:"Billed value",num:1,fmt:fmtRs}
  ],A);
}
function renderMonthly(){
  const o=currentFilter(); delete o.m; delete o.from; delete o.to;
  const rows=DATA.filter(r=>passes(r,o));
  const A=[...aggregate(rows,r=>r.m).values()].sort((a,b)=>a.key.localeCompare(b.key));
  A.forEach((a,i)=>{ a.chg_b=i?(a.bookings-A[i-1].bookings)/(A[i-1].bookings||1)*100:null; a.chg_v=i?(a.billed-A[i-1].billed)/(A[i-1].billed||1)*100:null; });
  setGrouped("ch-monthly","lg-monthly",A.map(a=>mLabel(a.key)),A.map(a=>a.bookings),A.map(a=>a.billed),"Bookings","Billed value",v=>v>=100000?(v/100000).toFixed(1)+"L":String(v));
  const chg=k=>a=>a[k]==null?'<span class="mut">-</span>':`<span class="${a[k]>=0?'pos':'neg'}">${a[k]>=0?'+':''}${a[k].toFixed(0)}%</span>`;
  makeTable($("t-monthly"),[
    {k:"key",label:"Month",fmt:x=>mLabel(x)},{k:"bookings",label:"Bookings",num:1,fmt:fmtN},{k:"chg_b",label:"Bookings vs prev",num:1,html:chg("chg_b")},
    {k:"delivered",label:"Delivered",num:1,fmt:fmtN},{k:"invoiced",label:"Invoiced",num:1,fmt:fmtRs},
    {k:"leads",label:"Leads",num:1,fmt:fmtN},{k:"billed_n",label:"Billed leads",num:1,fmt:fmtN},
    {k:"billed",label:"Billed value",num:1,fmt:fmtRs},{k:"chg_v",label:"Billed vs prev",num:1,html:chg("chg_v")}
  ],A);
}
function renderCompare(){
  const A=$("c-a").value, B=$("c-b").value;
  const base={g:$("c-garage").value, s:$("c-service").value};
  const ra=DATA.filter(r=>passes(r,Object.assign({m:A},base))), rb=DATA.filter(r=>passes(r,Object.assign({m:B},base)));
  const ta=totals(ra), tb=totals(rb);
  $("c-kpis").innerHTML=[["Bookings",ta.bookings,tb.bookings,fmtN,"#2563eb"],["Delivered",ta.delivered,tb.delivered,fmtN,"#3b82f6"],["Invoiced",ta.invoiced,tb.invoiced,fmtRs,"#60a5fa"],["Billed value",ta.billed,tb.billed,fmtRs,"#0369a1"]]
    .map(([l,a,b,f,c])=>`<div class="card kpi" style="--c:${c}"><div class="l">${l}</div><div class="v">${f(b)}</div><div class="s">${mLabel(B)} · ${mLabel(A)}: ${f(a)} · ${fmtPct(b,a)}</div></div>`).join("");
  const sA=aggregate(ra,r=>r.s), sB=aggregate(rb,r=>r.s);
  const sk=[...new Set([...sA.keys(),...sB.keys()])].sort();
  const sv=sk.map(k=>({key:k,a:sA.get(k)||{bookings:0,billed:0},b:sB.get(k)||{bookings:0,billed:0}}));
  setGrouped("ch-cmp-svc","lg-cmp-svc",sv.map(x=>x.key),sv.map(x=>x.a.bookings),sv.map(x=>x.b.bookings),mLabel(A),mLabel(B),fmtN);
  const cols=name=>[{k:"key",label:name},{k:"a_b",cf:"bookings",label:mLabel(A)+" bookings",num:1,fmt:fmtN},{k:"b_b",cf:"bookings",label:mLabel(B)+" bookings",num:1,fmt:fmtN},
    {k:"d_b",label:"Change",num:1,html:r=>pctCell(r.b_b,r.a_b)},{k:"a_v",cf:"billed",label:mLabel(A)+" billed",num:1,fmt:fmtRs},{k:"b_v",cf:"billed",label:mLabel(B)+" billed",num:1,fmt:fmtRs},
    {k:"d_v",label:"Billed change",num:1,html:r=>pctCell(r.b_v,r.a_v)}];
  const row=x=>({key:x.key,a_b:x.a.bookings,b_b:x.b.bookings,a_v:x.a.billed,b_v:x.b.billed});
  makeTable($("t-cmp-svc"),cols("Service"),sv.map(row));
  const gA=aggregate(ra,r=>r.g), gB=aggregate(rb,r=>r.g);
  const gk=[...new Set([...gA.keys(),...gB.keys()])].filter(Boolean).sort();
  const gv=gk.map(k=>({key:k,a:gA.get(k)||{bookings:0,billed:0},b:gB.get(k)||{bookings:0,billed:0}}));
  setGrouped("ch-cmp-gar","lg-cmp-gar",gv.map(x=>x.key),gv.map(x=>x.a.bookings),gv.map(x=>x.b.bookings),mLabel(A),mLabel(B),fmtN);
  makeTable($("t-cmp-gar"),cols("Garage"),gv.map(row));
}
function renderTrends(){
  const dim=$("tr-dim").value, met=METRICS[$("tr-metric").value], m=$("tr-month").value;
  const rows=DATA.filter(r=>r.m===m);
  const dl=[...new Set(rows.map(r=>r.d))].sort();
  const keys=dim==="garage"?garages:services;
  const cards=[];
  keys.forEach((k,ki)=>{
    const sub=rows.filter(r=>(dim==="garage"?r.g:r.s)===k);
    const byD=aggregate(sub,r=>r.d);
    const total=[...byD.values()].reduce((s,a)=>s+met.get(a),0);
    if(!total) return;
    const col=PAL[ki%PAL.length];
    const items=dl.map(d=>({label:String(+d.slice(8)),value:byD.has(d)?met.get(byD.get(d)):0,color:col}));
    cards.push(`<div class="card mini" style="--c:${col}"><h3>${esc(k)}</h3><div class="t">${mLabel(m)} total: ${esc(met.fmt(total))}</div><div class="chart" style="padding:4px">${barChart(items,{step:26,h:170,fmt:met.fmt,fluid:1})}</div></div>`);
  });
  $("tr-grid").innerHTML = cards.length? cards.join("") : '<p class="mut">No activity in this month.</p>';
}
function renderTables(){
  const o=currentFilter(); const view=$("tb-view").value, q=$("tb-search").value.trim().toLowerCase();
  const rows=DATA.filter(r=>passes(r,{m:o.m})&&(!q||String(r.g).toLowerCase().includes(q)||String(r.s).toLowerCase().includes(q)));
  const M=aggregate(rows, view==="garage"?r=>r.g:r=>r.s);
  const list=[...M.values()];
  const tot=list.reduce((s,a)=>s+a.billed,0);
  list.forEach(a=>a.share=tot?a.billed/tot*100:0);
  makeTable($("t-tables"),[
    {k:"key",label:view==="garage"?"Garage":"Service"},{k:"bookings",label:"Bookings",num:1,fmt:fmtN},
    {k:"delivered",label:"Delivered",num:1,fmt:fmtN},{k:"under",label:"Under service",num:1,fmt:fmtN},{k:"empty",label:"Empty return",num:1,fmt:fmtN},
    {k:"inv_n",label:"Invoices",num:1,fmt:fmtN},{k:"invoiced",label:"Invoiced",num:1,fmt:fmtRs},
    {k:"leads",label:"Leads",num:1,fmt:fmtN},{k:"billed_n",label:"Billed leads",num:1,fmt:fmtN},
    {k:"billed",label:"Billed value",num:1,fmt:fmtRs},{k:"share",label:"Share of billed",num:1,html:r=>r.share.toFixed(1)+"%"}
  ],list);
}
const RENDER={overview:renderOverview,daily:renderDaily,compare:renderCompare,tables:renderTables,editable:renderEditable,dropbox:()=>{}};
function renderActive(){ const t=document.querySelector("#tabs button.on").dataset.t; RENDER[t](); }


// ---------- editable table with formula columns and conditional formatting ----------
const FIELDS={"bookings":"bookings","delivered":"delivered","under service":"under","empty return":"empty","invoices":"inv_n","invoiced":"invoiced","leads":"leads","billed leads":"billed_n","billed value":"billed","share":"share"};
const FIELD_LABEL={bookings:"Bookings",delivered:"Delivered",under:"Under service",empty:"Empty return",inv_n:"Invoices",invoiced:"Invoiced",leads:"Leads",billed_n:"Billed leads",billed:"Billed value",share:"Share"};
const FIELD_FMT={bookings:fmtN,delivered:fmtN,under:fmtN,empty:fmtN,inv_n:fmtN,invoiced:fmtRs,leads:fmtN,billed_n:fmtN,billed:fmtRs,share:v=>v.toFixed(1)+"%"};
const RULE_COLORS={good:{bg:"#dbeafe",fg:"#1e3a8a"},alert:{bg:"#1e3a8a",fg:"#ffffff"}};
let EDIT={month:"",dim:"garage",rows:[],cols:[],rules:[]};
function defaultCols(){ return ["bookings","delivered","invoiced","billed_n","billed"].map(f=>({id:"f_"+f,name:FIELD_LABEL[f],type:"field",field:f})); }
function saveEdit(){ try{ localStorage.setItem("goaxen_editable_v1", JSON.stringify(EDIT)); }catch(e){} }
function loadAllRows(){ const items = EDIT.dim==="garage"?garages:services; EDIT.rows = items.map(x=>({label:x,key:x})); }
function fillAddRow(){ const items=EDIT.dim==="garage"?garages:services; $("e-addrow").innerHTML=items.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join(""); }
function fillRuleCols(){ $("r-col").innerHTML = CF_METRICS.map(([k,n])=>`<option value="${k}">${esc(n)}</option>`).join(""); }
function initEditable(){
  try{ const saved=JSON.parse(localStorage.getItem("goaxen_editable_v1")||"null"); if(saved) EDIT=Object.assign(EDIT,saved); }catch(e){}
  if(!EDIT.cols.length) EDIT.cols=defaultCols();
  if(!EDIT.rows.length) loadAllRows();
  $("e-dim").value=EDIT.dim;
  $("e-month").innerHTML=`<option value="">All months</option>`+months.map(m=>`<option value="${m}">${mLabel(m)}</option>`).join("");
  $("e-month").value=EDIT.month||"";
  fillAddRow(); fillRuleCols();
}
function evalFormula(expr, vals){
  try{
    const e=String(expr).replace(/\[([^\]]+)\]/g,(m,name)=>{ const k=FIELDS[name.trim().toLowerCase()]; const v=k?vals[k]:0; return "("+(Number.isFinite(v)?v:0)+")"; });
    if(!/^[\d\s+\-*/().]*$/.test(e)) return NaN;
    const r=Function('"use strict";return ('+e+');')();
    return Number.isFinite(r)?r:NaN;
  }catch(err){ return NaN; }
}
function ruleHit(rule,v){ if(!Number.isFinite(v)) return false; const x=+rule.val; switch(rule.op){ case ">": return v>x; case ">=": return v>=x; case "<": return v<x; case "<=": return v<=x; default: return v===x; } }
function renderEditable(){
  EDIT.month=$("e-month").value; EDIT.dim=$("e-dim").value;
  const month=EDIT.month;
  const data=EDIT.rows.map((r,ri)=>{
    const rows=DATA.filter(x=>(!month||x.m===month)&&(r.key==="__all__"||(EDIT.dim==="garage"?x.g:x.s)===r.key));
    return {ri,label:r.label,t:totals(rows)};
  });
  const totBilled=data.reduce((s,x)=>s+x.t.billed,0);
  const valsOf=t=>Object.assign({},t,{share: totBilled? t.billed/totBilled*100 : 0});
  const cellVal=(c,vals)=> c.type==="field" ? (c.field==="share"?vals.share:vals[c.field]) : evalFormula(c.formula,vals);
  const fmtCell=(c,v)=> !Number.isFinite(v)? "-" : (c.type==="field"? FIELD_FMT[c.field](v) : (Math.abs(v)>=1000? fmtRs(v) : v.toFixed(2)));
  const styleFor=(c,v)=> cfStyle(c.type==="field"?c.field:c.id, v);
  const head=`<thead><tr><th>Row</th>${EDIT.cols.map((c,i)=>`<th class="n" style="min-width:150px">
      <input class="ed-name" data-i="${i}" value="${esc(c.name)}" style="width:100%">
      ${c.type==="formula"?`<input class="ed-f" data-i="${i}" value="${esc(c.formula)}" style="width:100%;margin-top:4px;font-size:12px">`:`<div class="mut" style="font-size:11px">field: ${esc(FIELD_LABEL[c.field])}</div>`}
      <button class="btn" data-act="delcol" data-i="${i}" style="margin-top:4px;padding:2px 8px;font-size:11px">Remove column</button></th>`).join("")}</tr></thead>`;
  const body=data.map(x=>{
    const vals=valsOf(x.t);
    return `<tr><td><input class="ed-row" data-r="${x.ri}" value="${esc(x.label)}"> <button class="btn" data-act="delrow" data-r="${x.ri}" style="padding:2px 8px;font-size:11px">Remove</button></td>${
      EDIT.cols.map(c=>{ const v=cellVal(c,vals); return `<td class="n" style="${styleFor(c,v)}">${fmtCell(c,v)}</td>`; }).join("")}</tr>`;
  }).join("");
  // totals row
  const sum={}; ["bookings","delivered","under","empty","inv_n","invoiced","leads","billed_n","billed"].forEach(k=>sum[k]=data.reduce((s,x)=>s+(x.t[k]||0),0));
  const sumVals=valsOf(sum);
  const totRow=`<tr style="font-weight:700;background:var(--line)"><td>Total</td>${EDIT.cols.map(c=>{ const v=cellVal(c,sumVals); return `<td class="n">${fmtCell(c,v)}</td>`; }).join("")}</tr>`;
  $("t-editable").innerHTML=head+`<tbody>${body}${totRow}</tbody>`;
  $("rules").innerHTML = EDIT.rules.length ? "Rules: "+EDIT.rules.map((r,i)=>`<span class="pos" style="color:#1e3a8a">${esc(cfName(String(r.col).replace(/^f_/,"")))} ${esc(r.op)} ${esc(r.val)}</span> <a href="#" data-act="delrule" data-i="${i}">remove</a>`).join(" · ")+" (applies to Garage & service tables and Editable table)" : "No conditional formatting rules yet. Rules apply to the Garage & service tables and the Editable table.";
  fillRuleCols();
}
(function wireEditable(){
  const t=$("t-editable");
  t.addEventListener("change",e=>{
    const el=e.target;
    if(el.classList.contains("ed-name")){ EDIT.cols[+el.dataset.i].name=el.value; }
    else if(el.classList.contains("ed-f")){ EDIT.cols[+el.dataset.i].formula=el.value; }
    else if(el.classList.contains("ed-row")){ const r=EDIT.rows[+el.dataset.r]; if(r) r.label=el.value; }
    else return;
    saveEdit(); renderEditable();
  });
  t.addEventListener("click",e=>{
    const b=e.target.closest("[data-act]"); if(!b) return;
    const act=b.dataset.act;
    if(act==="delcol"){ const c=EDIT.cols[+b.dataset.i]; EDIT.cols.splice(+b.dataset.i,1); }
    else if(act==="delrow"){ EDIT.rows.splice(+b.dataset.r,1); }
    saveEdit(); renderEditable();
  });
  $("rules").addEventListener("click",e=>{ const a=e.target.closest('[data-act="delrule"]'); if(!a) return; e.preventDefault(); EDIT.rules.splice(+a.dataset.i,1); saveEdit(); renderEditable(); });
})();

// ---------- wiring ----------
["f-month","f-garage","f-service","f-from","f-to","f-search"].forEach(id=>$(id).addEventListener("input",renderActive));
$("f-month").addEventListener("change",()=>{
  const m=$("f-month").value; const ds=DATA.filter(r=>!m||r.m===m).map(r=>r.d).sort();
  if(ds.length){ $("f-from").value=ds[0]; $("f-to").value=ds[ds.length-1]; }
  renderActive();
});
$("f-reset").onclick=()=>{ $("f-month").value=""; $("f-garage").value=""; $("f-service").value=""; $("f-search").value=""; $("f-from").value=dates[0]; $("f-to").value=dates[dates.length-1]; renderActive(); };
["c-a","c-b","c-garage","c-service"].forEach(id=>$(id).addEventListener("change",renderActive));
$("tb-view").addEventListener("change",renderActive); $("tb-search").addEventListener("input",renderActive);
$("e-month").addEventListener("change",renderActive); $("e-dim").addEventListener("change",()=>{ EDIT.dim=$("e-dim").value; EDIT.rows=[]; fillAddRow(); loadAllRows(); saveEdit(); renderActive(); });
$("e-addrow-btn").onclick=()=>{ const v=$("e-addrow").value; if(v){ EDIT.rows.push({label:v,key:v}); saveEdit(); renderActive(); } };
$("e-allrows").onclick=()=>{ loadAllRows(); saveEdit(); renderActive(); };
$("e-addcol").onclick=()=>{ const f=$("e-formula").value.trim(); if(!f) return; EDIT.cols.push({id:"c"+Date.now(),name:$("e-colname").value.trim()||f,type:"formula",formula:f}); $("e-formula").value=""; $("e-colname").value=""; saveEdit(); renderActive(); };
$("r-add").onclick=()=>{ EDIT.rules.push({col:$("r-col").value,op:$("r-op").value,val:+$("r-val").value,color:$("r-color").value}); saveEdit(); renderActive(); };
document.querySelectorAll("#tabs button").forEach(b=>b.onclick=()=>{
  document.querySelectorAll("#tabs button").forEach(x=>x.classList.toggle("on",x===b));
  document.querySelectorAll("main > section").forEach(s=>s.classList.toggle("hidden", s.id!=="s-"+b.dataset.t));
  renderActive();
});
const drop=$("drop"), fileIn=$("file");
drop.onclick=()=>fileIn.click();
fileIn.onchange=e=>processFiles([...e.target.files]);
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault(); drop.classList.add("over");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault(); drop.classList.remove("over");}));
drop.addEventListener("drop",e=>processFiles([...e.dataTransfer.files].filter(f=>/\.xlsx$/i.test(f.name))));

// start: embedded data + any uploads remembered in this browser
(function init(){
  try{ uploads = JSON.parse(localStorage.getItem(STORE_KEY)||"[]"); }catch(e){ uploads=[]; }
  rebuildFromAll();
  $("f-month").value="";
  $("dropfolder").textContent = "Folder read by the script: " + FOLDER;
  renderUploadLog([]);
  initEditable();
  document.querySelector('#tabs button[data-t="dropbox"]').click();
})();
</script>
</body></html>
"""

def main():
    files = sorted(glob.glob(os.path.join(FOLDER, '*.xlsx')))
    checkin, billed, used = [], [], []
    for f in files:
        if os.path.basename(f).startswith('~$'): continue
        try:
            c = load_checkin(f)
            if c:
                checkin += c; used.append(os.path.basename(f) + '  (check-in, %d bookings)' % len(c)); continue
            b = load_billed(f)
            if b:
                billed += b; used.append(os.path.basename(f) + '  (billed list, %d leads)' % len(b))
        except Exception as e:
            used.append(f'{os.path.basename(f)}  (could not read: {e})')
    if not checkin and not billed and not UPLOAD_ONLY:
        print('No check-in or billed-list files found in', FOLDER); return
    recs = []
    for r in checkin:
        recs.append(dict(d=r['date'].strftime('%Y-%m-%d'), m=r['date'].strftime('%Y-%m'), g=r['garage'], s=r['service'],
                         src='c', st=r['status'], inv=r['invoice'] or 0))
    for r in billed:
        recs.append(dict(d=r['date'].strftime('%Y-%m-%d'), m=r['date'].strftime('%Y-%m'), g=r['garage'], s=r['service'],
                         src='b', bl=1 if r['billed'] else 0, v=r['value'] or 0))
    html_out = (TEMPLATE.replace('__DATA__', json.dumps(recs, separators=(',', ':')))
                        .replace('__FILES__', json.dumps(used))
                        .replace('__FOLDER__', json.dumps(FOLDER)))
    with open(os.path.join(OUT, 'dashboard.html'), 'w', encoding='utf-8') as fh:
        fh.write(html_out)
    print('Dashboard:', os.path.join(OUT, 'dashboard.html'))
    print('Files read:', used)

if __name__ == '__main__':
    main()
