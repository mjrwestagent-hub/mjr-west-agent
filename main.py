"""
MJR West Industrial Property Intelligence Agent - V2
Clean rebuild: direct Supabase HTTP, async processing, no library dependency issues.
"""

import os, json, logging, threading, base64, uuid, imaplib, email
from datetime import datetime
from email.header import decode_header
from flask import (Flask, request, jsonify, render_template_string,
                   redirect, url_for, flash, session)
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from twilio.rest import Client as TwilioClient
import requests as http
import openai
from functools import wraps

log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# ── Config ───────────────────────────────────────────────────────────────────
SUPABASE_URL   = os.getenv("SUPABASE_URL","").rstrip("/")
SUPABASE_KEY   = os.getenv("SUPABASE_KEY","")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY","")
OPENAI_MODEL   = os.getenv("OPENAI_MODEL","gpt-4o-mini")
EMBED_MODEL    = os.getenv("EMBEDDING_MODEL","text-embedding-3-small")
SECRET_KEY     = os.getenv("SECRET_KEY","dev-secret")
ADMIN_USER     = os.getenv("ADMIN_USER","admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD","admin")
UPLOAD_FOLDER  = os.getenv("UPLOAD_FOLDER","/tmp/uploads")
TWILIO_SID     = os.getenv("TWILIO_ACCOUNT_SID","")
TWILIO_TOKEN   = os.getenv("TWILIO_AUTH_TOKEN","")
TWILIO_FROM    = os.getenv("TWILIO_WHATSAPP_FROM","")
WA_BROADCAST   = os.getenv("WA_BROADCAST_LIST","")
GMAIL_USER     = os.getenv("GMAIL_USER","")
GMAIL_PASS     = os.getenv("GMAIL_APP_PASSWORD","")
TIMEZONE       = os.getenv("TIMEZONE","Australia/Melbourne")
BRIEFING_HOUR  = int(os.getenv("BRIEFING_HOUR","7"))
BRIEFING_MIN   = int(os.getenv("BRIEFING_MINUTE","0"))
APP_URL        = os.getenv("APP_URL","")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXT = {"pdf","xlsx","xls","docx","doc","txt","csv","png","jpg","jpeg","eml","msg"}

def allowed_file(fn):
    return "." in fn and fn.rsplit(".",1)[-1].lower() in ALLOWED_EXT

# ── Flask ─────────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = SECRET_KEY
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1)

# ── Supabase direct HTTP ──────────────────────────────────────────────────────
# No supabase-py library — direct REST API calls with service_role key.
# This bypasses ALL RLS and version compatibility issues permanently.

SB_HDR = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation"
}

def sb_select(table, filters=None, order="created_at", limit=200):
    try:
        params = {"select":"*","order":f"{order}.desc","limit":limit}
        if filters:
            params.update(filters)
        r = http.get(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HDR, params=params, timeout=10)
        return r.json() if r.status_code==200 else []
    except Exception as e:
        log.error("sb_select %s: %s", table, e); return []

def sb_insert(table, data):
    try:
        r = http.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HDR, json=data, timeout=10)
        if r.status_code in (200,201):
            rows=r.json(); return rows[0] if rows else data
        log.error("sb_insert %s: %s %s", table, r.status_code, r.text[:150])
        return None
    except Exception as e:
        log.error("sb_insert %s: %s", table, e); return None

def sb_update(table, col, val, data):
    try:
        r = http.patch(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HDR,
                       params={col:f"eq.{val}"}, json=data, timeout=10)
        return r.status_code in (200,204)
    except Exception as e:
        log.error("sb_update %s: %s", table, e); return False

def sb_delete(table, row_id):
    try:
        r = http.delete(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HDR,
                        params={"id":f"eq.{row_id}"}, timeout=10)
        return r.status_code in (200,204)
    except Exception as e:
        log.error("sb_delete %s: %s", table, e); return False

# ── Schema ────────────────────────────────────────────────────────────────────
def init_schema():
    tables = [
        ("properties", """id bigserial primary key, address text, suburb text default 'West Melbourne',
         property_type text default 'Warehouse', size_sqm numeric, land_sqm numeric,
         asking_rent_pa numeric, asking_price numeric, status text default 'Available',
         landlord text, agent_name text, agent_phone text, agent_email text,
         year_built int, zoning text, notes text, source text,
         created_at timestamptz default now()"""),
        ("contacts", """id bigserial primary key, name text, company text, phone text, email text,
         contact_type text default 'Agent', lead_status text default 'Warm',
         notes text, whatsapp_opt_in bool default false, created_at timestamptz default now()"""),
        ("inquiries", """id bigserial primary key, company text, contact_name text, phone text, email text,
         size_min numeric, size_max numeric, budget numeric, location text,
         notes text, source text default 'Web', status text default 'New', score int default 0,
         created_at timestamptz default now()"""),
        ("deals", """id bigserial primary key, tenant text, landlord text,
         deal_type text default 'Lease', size_sqm numeric, rent_pa numeric, sale_price numeric,
         term_years numeric, start_date date, status text default 'Negotiation',
         commission numeric, notes text, created_at timestamptz default now()"""),
        ("vacancies", """id bigserial primary key, address text, suburb text default 'West Melbourne',
         size_sqm numeric, asking_rent_pa numeric, available_date date,
         vacating_tenant text, owner text, agent text, notes text,
         status text default 'Available', created_at timestamptz default now()"""),
        ("requirements", """id bigserial primary key, company text, contact_name text, phone text, email text,
         size_min numeric, size_max numeric, budget_pa numeric,
         preferred_location text default 'West Melbourne', region text default 'W',
         use_type text, term_years numeric, timeline text, rating text, notes text,
         status text default 'Active', created_at timestamptz default now()"""),
        ("market_data", """id bigserial primary key, address text, suburb text default 'West Melbourne',
         deal_type text default 'Lease', size_sqm numeric, rent_pa numeric, sale_price numeric,
         term_years numeric, tenant text, landlord text, date_signed date, source text, notes text,
         created_at timestamptz default now()"""),
        ("documents", """id bigserial primary key, filename text,
         ai_classification text default 'unknown', ai_summary text,
         ai_confidence numeric default 0, ai_urgency text default 'low',
         action_items jsonb default '[]', key_facts jsonb default '{}',
         extracted_properties int default 0, extracted_contacts int default 0,
         extracted_requirements int default 0, extracted_vacancies int default 0,
         mentioned_companies jsonb default '[]',
         processing_status text default 'pending', created_at timestamptz default now()"""),
        ("briefings", """id bigserial primary key, briefing_type text default 'Daily',
         content text, channel text default 'WhatsApp', sent_at timestamptz default now()"""),
        ("email_logs", """id bigserial primary key, sender text, subject text, body text,
         ai_summary text, ai_priority text default 'medium',
         requires_action bool default false, action_items jsonb default '[]',
         reply_status text default 'pending', source text default 'gmail_imap',
         received_at timestamptz default now()"""),
        ("call_logs", """id bigserial primary key, contact_name text, phone text,
         direction text default 'Inbound', duration_secs int, notes text,
         ai_summary text, created_at timestamptz default now()"""),
        ("fees", """id bigserial primary key, deal_id bigint, amount numeric,
         invoice_status text default 'Pending', invoice_date date, paid_date date,
         notes text, created_at timestamptz default now()"""),
    ]
    for name, cols in tables:
        r = http.post(f"{SUPABASE_URL}/rest/v1/rpc/exec_sql",
                      headers=SB_HDR,
                      json={"query": f"create table if not exists {name} ({cols})"},
                      timeout=15)
        if r.status_code not in (200,201,204):
            log.info("Table %s may already exist: %s", name, r.status_code)

# ── OpenAI ────────────────────────────────────────────────────────────────────
def gpt(system, user, max_tokens=2000):
    try:
        client = openai.OpenAI(api_key=OPENAI_API_KEY)
        r = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role":"system","content":system},{"role":"user","content":user}],
            temperature=0.1, max_tokens=max_tokens,
            response_format={"type":"json_object"}
        )
        return r.choices[0].message.content
    except Exception as e:
        log.error("gpt: %s", e); return "{}"

def embed(text):
    try:
        client = openai.OpenAI(api_key=OPENAI_API_KEY)
        r = client.embeddings.create(model=EMBED_MODEL, input=text[:8000])
        return r.data[0].embedding
    except Exception as e:
        log.error("embed: %s", e); return []

# ── File extraction ───────────────────────────────────────────────────────────
def extract_text(filepath, filename):
    ext = filename.rsplit(".",1)[-1].lower() if "." in filename else ""
    text, is_img = "", False
    try:
        if ext in ("txt","csv","eml"):
            text = open(filepath,"r",encoding="utf-8",errors="ignore").read()
        elif ext == "pdf":
            import pypdf
            r = pypdf.PdfReader(filepath)
            text = "\n".join(p.extract_text() or "" for p in r.pages)
        elif ext in ("docx","doc"):
            import docx
            text = "\n".join(p.text for p in docx.Document(filepath).paragraphs)
        elif ext in ("xlsx","xls"):
            import openpyxl
            wb = openpyxl.load_workbook(filepath,read_only=True,data_only=True)
            rows=[]
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rows.append("\t".join(str(c) if c else "" for c in row))
            text = "\n".join(rows[:500])
        elif ext in ("png","jpg","jpeg"):
            is_img = True
            text = base64.b64encode(open(filepath,"rb").read()).decode()
    except Exception as e:
        log.error("extract_text %s: %s", filename, e)
    return text, is_img

AI_PROMPT = """You are an AI for MJR West, an industrial property agency in West Melbourne.
Analyse this document and return ONLY valid JSON:
{
  "classification": "asset_register|vacancy_schedule|requirements_listing|deal_tracker|lease_contract|email|general",
  "confidence": 0.0-1.0,
  "summary": "2-3 sentences",
  "urgency": "low|medium|high",
  "action_items": [],
  "key_facts": {},
  "mentioned_companies": [],
  "table_inserts": {
    "properties": [], "contacts": [], "inquiries": [],
    "deals": [], "vacancies": [], "requirements": [], "market_data": []
  }
}
REGION RULE: Only include requirements for West Melbourne/Footscray/Sunshine/Altona. Set region="W". Exclude all others.
Extract all relevant data into table_inserts arrays as flat objects matching each table's columns."""

def process_document(filepath, filename):
    counts = {k:0 for k in ["properties","contacts","inquiries","deals","vacancies","requirements","market_data"]}
    try:
        text, is_img = extract_text(filepath, filename)
        if not text:
            log.warning("No text from %s", filename); return counts

        if is_img:
            raw = gpt(AI_PROMPT, f"Analyse image: {filename}", max_tokens=3000)
        else:
            raw = gpt(AI_PROMPT, f"Filename: {filename}\n\nContent:\n{text[:12000]}", max_tokens=3000)

        try:
            a = json.loads(raw)
        except Exception:
            log.error("JSON parse failed for %s", filename); return counts

        inserts = a.get("table_inserts", {})

        # Save document record
        sb_insert("documents", {
            "filename": filename,
            "ai_classification": a.get("classification","general"),
            "ai_summary": a.get("summary",""),
            "ai_confidence": a.get("confidence",0),
            "ai_urgency": a.get("urgency","low"),
            "action_items": a.get("action_items",[]),
            "key_facts": a.get("key_facts",{}),
            "mentioned_companies": a.get("mentioned_companies",[]),
            "processing_status": "complete"
        })

        for prop in (inserts.get("properties") or []):
            prop.setdefault("suburb","West Melbourne")
            prop.setdefault("status","Available")
            prop.setdefault("source",f"AI:{filename}")
            if sb_insert("properties", prop): counts["properties"]+=1

        for c in (inserts.get("contacts") or []):
            if sb_insert("contacts", c): counts["contacts"]+=1

        for i in (inserts.get("inquiries") or []):
            i.setdefault("source","AI Intake"); i.setdefault("status","New")
            if sb_insert("inquiries", i): counts["inquiries"]+=1

        for d in (inserts.get("deals") or []):
            if sb_insert("deals", d): counts["deals"]+=1

        for v in (inserts.get("vacancies") or []):
            v.setdefault("suburb","West Melbourne"); v.setdefault("status","Available")
            if sb_insert("vacancies", v): counts["vacancies"]+=1

        for req in (inserts.get("requirements") or []):
            req.setdefault("preferred_location","West Melbourne")
            req.setdefault("region","W"); req.setdefault("status","Active")
            if sb_insert("requirements", req): counts["requirements"]+=1

        for md in (inserts.get("market_data") or []):
            md.setdefault("suburb","West Melbourne")
            if sb_insert("market_data", md): counts["market_data"]+=1

        log.info("AI intake: %s → %s (%.2f) %s", filename, a.get("classification"), a.get("confidence",0), counts)
        return counts
    except Exception as e:
        log.error("process_document %s: %s", filename, e); return counts

# ── WhatsApp ──────────────────────────────────────────────────────────────────
def send_whatsapp(to, body):
    if not (TWILIO_SID and TWILIO_TOKEN and TWILIO_FROM): return False
    try:
        c = TwilioClient(TWILIO_SID, TWILIO_TOKEN)
        to_w = to if to.startswith("whatsapp:") else f"whatsapp:{to}"
        c.messages.create(from_=TWILIO_FROM, to=to_w, body=body[:1500])
        return True
    except Exception as e:
        log.error("send_whatsapp %s: %s", to, e); return False

def broadcast_whatsapp(body):
    if not WA_BROADCAST: return
    for n in WA_BROADCAST.split(","):
        n=n.strip()
        if n: send_whatsapp(n, body)

# ── Briefing ──────────────────────────────────────────────────────────────────
def build_brief():
    props = sb_select("properties",{"status":"eq.Available"},limit=50)
    vacs  = sb_select("vacancies",{"status":"eq.Available"},limit=10)
    reqs  = sb_select("requirements",{"status":"eq.Active"},limit=10)
    deals = sb_select("deals",limit=5)
    lines = [f"🏭 MJR West Daily Brief — {datetime.now().strftime('%a %d %b %Y')}",
             f"\n📊 Properties: {len(props)} | Vacancies: {len(vacs)} | Requirements: {len(reqs)} | Deals: {len(deals)}"]
    if vacs:
        lines.append("\n🔑 VACANCIES")
        for v in vacs[:3]:
            lines.append(f"• {v.get('address','?')} {v.get('size_sqm','?')}sqm ${v.get('asking_rent_pa',0):,.0f}pa")
    if reqs:
        lines.append("\n🔍 REQUIREMENTS")
        for r in reqs[:3]:
            lines.append(f"• {r.get('company','?')} {r.get('size_min','?')}-{r.get('size_max','?')}sqm ${r.get('budget_pa',0):,.0f}pa")
    return "\n".join(lines)

def send_daily_briefing():
    brief = build_brief()
    sb_insert("briefings",{"briefing_type":"Daily","content":brief,"channel":"WhatsApp"})
    broadcast_whatsapp(brief)
    log.info("Daily briefing sent")

def check_gmail():
    if not (GMAIL_USER and GMAIL_PASS): return
    try:
        m = imaplib.IMAP4_SSL("imap.gmail.com")
        m.login(GMAIL_USER, GMAIL_PASS)
        m.select("inbox")
        _, msgs = m.search(None,"UNSEEN")
        for num in (msgs[0].split() or [])[:10]:
            _, data = m.fetch(num,"(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            subj = decode_header(msg["Subject"] or "")[0][0]
            if isinstance(subj, bytes): subj = subj.decode(errors="ignore")
            sender = msg.get("From","")
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type()=="text/plain":
                        body = part.get_payload(decode=True).decode(errors="ignore"); break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")
            sb_insert("email_logs",{"sender":sender,"subject":subj,"body":body[:5000],"source":"gmail_imap"})
        m.logout()
    except Exception as e:
        log.error("check_gmail: %s", e)

scheduler = BackgroundScheduler(timezone=TIMEZONE)
scheduler.add_job(send_daily_briefing, CronTrigger(hour=BRIEFING_HOUR,minute=BRIEFING_MIN), id="brief")
scheduler.add_job(check_gmail, CronTrigger(minute="*/15"), id="gmail")
scheduler.start()

# ── Auth ──────────────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def wrapper(*args,**kwargs):
        if not session.get("logged_in"): return redirect(url_for("login"))
        return f(*args,**kwargs)
    return wrapper

# ── Layout ────────────────────────────────────────────────────────────────────
NAV = [
    ("dashboard","bi-speedometer2","Dashboard"),
    ("properties","bi-buildings","Properties"),
    ("contacts","bi-people","Contacts"),
    ("inquiries","bi-lightning","Lead Pipeline"),
    ("deals","bi-handshake","Deals"),
    ("vacancies","bi-door-open","Vacancies"),
    ("requirements","bi-search","Requirements"),
    ("upload","bi-cloud-upload","AI Intake"),
    ("documents","bi-file-earmark-text","Documents"),
    ("email_page","bi-envelope","Email"),
    ("call_logs","bi-telephone","Call Logs"),
    ("market_data_page","bi-graph-up","Market Data"),
    ("fees_page","bi-cash-coin","Fees"),
    ("briefings_page","bi-broadcast","Briefings"),
    ("whatsapp_page","bi-whatsapp","WhatsApp"),
    ("settings_page","bi-gear","Settings"),
]

STYLE = """
body{background:#0f172a;color:#e2e8f0;font-family:'Inter',sans-serif;}
.sidebar{width:220px;min-height:100vh;background:#1e293b;border-right:1px solid #334155;position:fixed;top:0;left:0;overflow-y:auto;z-index:100;}
.sidebar .brand{padding:1.2rem 1rem;font-weight:700;color:#f59e0b;border-bottom:1px solid #334155;}
.nav-link{color:#94a3b8;padding:.45rem 1rem;font-size:.83rem;border-radius:6px;margin:.1rem .5rem;display:block;}
.nav-link:hover,.nav-link.active{color:#f59e0b;background:rgba(245,158,11,.1);}
.main-content{margin-left:220px;padding:1.5rem;}
.card{background:#1e293b;border:1px solid #334155;border-radius:10px;}
.card-header{background:#334155;border-bottom:1px solid #475569;padding:.75rem 1rem;}
.table{color:#e2e8f0;}
.table thead{color:#94a3b8;font-size:.78rem;text-transform:uppercase;}
.btn-primary{background:#f59e0b;border-color:#f59e0b;color:#0f172a;font-weight:600;}
.btn-primary:hover{background:#d97706;border-color:#d97706;color:#0f172a;}
.stat-card{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:1.2rem;}
.stat-value{font-size:2rem;font-weight:700;color:#f59e0b;}
.flash-success{background:rgba(16,185,129,.15);border:1px solid #10b981;color:#10b981;border-radius:8px;padding:.75rem 1rem;margin-bottom:1rem;}
.flash-error{background:rgba(239,68,68,.15);border:1px solid #ef4444;color:#ef4444;border-radius:8px;padding:.75rem 1rem;margin-bottom:1rem;}
"""

def layout(content, title="MJR West", active=""):
    nav = ""
    for ep, icon, label in NAV:
        try: href = url_for(ep)
        except: href = "#"
        cls = "active" if ep==active else ""
        nav += f'<a href="{href}" class="nav-link {cls}"><i class="bi {icon} me-2"></i>{label}</a>'
    return render_template_string(f"""<!DOCTYPE html>
<html lang="en" data-bs-theme="dark">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title} — MJR West Agent</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">
<style>{STYLE}</style></head>
<body>
<div class="sidebar">
  <div class="brand"><i class="bi bi-buildings-fill me-2"></i>MJR West Agent</div>
  <nav class="py-2">{nav}</nav>
</div>
<div class="main-content">
{{% with messages = get_flashed_messages(with_categories=true) %}}
{{% for cat, msg in messages %}}
<div class="flash-{{{{ cat }}}}"><i class="bi bi-{{'check-circle' if cat=='success' else 'exclamation-circle'}} me-2"></i>{{{{ msg }}}}</div>
{{% endfor %}}
{{% endwith %}}
{content}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body></html>""")

def fc(v):
    try: return f"${float(v):,.0f}" if v else "—"
    except: return "—"

def fd(s):
    try: return datetime.fromisoformat(str(s)).strftime("%d %b %Y") if s else "—"
    except: return str(s) if s else "—"

def sbadge(status):
    c={"Available":"success","Completed":"primary","Negotiation":"warning","Active":"info","New":"warning"}.get(status,"secondary")
    return f'<span class="badge bg-{c} bg-opacity-25 text-{c}">{status or "—"}</span>'

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method=="POST":
        if request.form.get("username")==ADMIN_USER and request.form.get("password")==ADMIN_PASSWORD:
            session["logged_in"]=True; return redirect(url_for("dashboard"))
        flash("Invalid credentials","error")
    return render_template_string("""<!DOCTYPE html>
<html data-bs-theme="dark"><head><meta charset="UTF-8"><title>Login</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">
<style>body{background:#0f172a;}.card{background:#1e293b;border:1px solid #334155;}</style></head>
<body><div class="d-flex align-items-center justify-content-center min-vh-100">
<div class="card p-4" style="width:360px">
<h4 class="text-center mb-4" style="color:#f59e0b"><i class="bi bi-buildings-fill me-2"></i>MJR West Agent</h4>
{% with messages=get_flashed_messages(with_categories=true) %}{% for c,m in messages %}
<div class="alert alert-danger">{{m}}</div>{% endfor %}{% endwith %}
<form method="POST">
<div class="mb-3"><input name="username" class="form-control" placeholder="Username" required></div>
<div class="mb-3"><input name="password" type="password" class="form-control" placeholder="Password" required></div>
<button class="btn btn-warning w-100 fw-bold">Login</button>
</form></div></div></body></html>""")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))

@app.route("/")
@login_required
def index(): return redirect(url_for("dashboard"))

@app.route("/dashboard")
@login_required
def dashboard():
    pc = len(sb_select("properties",{"status":"eq.Available"}))
    vc = len(sb_select("vacancies",{"status":"eq.Available"}))
    rc = len(sb_select("requirements",{"status":"eq.Active"}))
    dc = len(sb_select("deals"))
    vacs = sb_select("vacancies",{"status":"eq.Available"},limit=5)
    reqs = sb_select("requirements",{"status":"eq.Active"},limit=5)
    docs = sb_select("documents",limit=5)
    leads = sb_select("inquiries",{"status":"eq.New"},limit=5)

    vrow = "".join(f"<tr><td>{v.get('address','?')}</td><td>{v.get('size_sqm','?')} sqm</td><td>{fc(v.get('asking_rent_pa'))}</td></tr>" for v in vacs) or "<tr><td colspan=3 class='text-muted p-3 text-center'>No vacancies yet</td></tr>"
    rrow = "".join(f"<tr><td>{r.get('company','?')}</td><td>{r.get('size_min','?')}-{r.get('size_max','?')} sqm</td><td>{fc(r.get('budget_pa'))}</td></tr>" for r in reqs) or "<tr><td colspan=3 class='text-muted p-3 text-center'>No requirements yet</td></tr>"
    drow = "".join(f"<tr><td>{d.get('filename','?')}</td><td><span class='badge bg-secondary'>{d.get('ai_classification','?')}</span></td><td>{fd(d.get('created_at'))}</td></tr>" for d in docs) or "<tr><td colspan=3 class='text-muted p-3 text-center'>No documents yet</td></tr>"
    lrow = "".join(f"<tr><td>{i.get('company','?')}</td><td>{i.get('source','?')}</td><td>{sbadge(i.get('status'))}</td></tr>" for i in leads) or "<tr><td colspan=3 class='text-muted p-3 text-center'>No new leads</td></tr>"

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-4">
  <h2 class="fw-bold mb-0">Dashboard</h2>
  <span class="text-muted">{datetime.now().strftime('%A, %d %B %Y')}</span>
</div>
<div class="row g-3 mb-4">
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{pc}</div><div class="text-muted small">Available Properties</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{vc}</div><div class="text-muted small">Active Vacancies</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{rc}</div><div class="text-muted small">Live Requirements</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{dc}</div><div class="text-muted small">Total Deals</div></div></div>
</div>
<div class="row g-3">
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Recent Vacancies</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Address</th><th>Size</th><th>Rent</th></tr></thead><tbody>{vrow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Active Requirements</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Company</th><th>Size</th><th>Budget</th></tr></thead><tbody>{rrow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Recent Documents</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>File</th><th>Type</th><th>Date</th></tr></thead><tbody>{drow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">New Leads</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Company</th><th>Source</th><th>Status</th></tr></thead><tbody>{lrow}</tbody></table></div></div></div>
</div>"""
    return layout(content, "Dashboard", "dashboard")

@app.route("/upload", methods=["GET","POST"])
@login_required
def upload():
    if request.method=="POST":
        f = request.files.get("file")
        if not f or not f.filename:
            flash("No file selected","error")
        elif not allowed_file(f.filename):
            flash("File type not supported","error")
        else:
            fname = secure_filename(f.filename)
            fpath = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_{fname}")
            f.save(fpath)
            def _bg(fp,fn):
                try: process_document(fp,fn)
                except Exception as e: log.error("bg %s: %s",fn,e)
                finally:
                    try: os.remove(fp)
                    except: pass
            threading.Thread(target=_bg,args=(fpath,fname),daemon=True).start()
            flash(f"⏳ {fname} received — processing in background. Check Documents shortly.","success")

    content = """
<h2 class="fw-bold mb-4">AI Document Intake</h2>
<div class="row">
  <div class="col-md-7">
    <div class="card">
      <div class="card-header fw-semibold">Upload Document</div>
      <div class="card-body">
        <form method="POST" enctype="multipart/form-data">
          <p class="text-muted small">AI will classify and extract data automatically</p>
          <div class="mb-3">
            <input type="file" name="file" class="form-control"
              accept=".pdf,.xlsx,.xls,.docx,.doc,.txt,.csv,.png,.jpg,.jpeg,.eml">
          </div>
          <button type="submit" class="btn btn-primary w-100">
            <i class="bi bi-cpu me-2"></i>Run AI Intake
          </button>
        </form>
      </div>
    </div>
  </div>
  <div class="col-md-5">
    <div class="card">
      <div class="card-header fw-semibold">Recognised Types</div>
      <div class="card-body">
        <div class="mb-2"><span class="badge bg-warning text-dark me-2">REG</span>Asset Register</div>
        <div class="mb-2"><span class="badge bg-success me-2">VAC</span>Vacancy Schedule</div>
        <div class="mb-2"><span class="badge bg-info me-2">REQ</span>Requirements Listing</div>
        <div class="mb-2"><span class="badge bg-primary me-2">DEL</span>Deal Tracker</div>
        <div class="mb-2"><span class="badge bg-secondary me-2">LSE</span>Lease Contract</div>
      </div>
    </div>
  </div>
</div>"""
    return layout(content, "AI Intake", "upload")

@app.route("/documents")
@login_required
def documents():
    docs = sb_select("documents",limit=100)
    rows = "".join(f"""<tr>
      <td>{d.get('filename','?')}</td>
      <td><span class="badge bg-secondary">{d.get('ai_classification','?')}</span></td>
      <td>{int((d.get('ai_confidence') or 0)*100)}%</td>
      <td><span class="badge bg-{'success' if d.get('processing_status')=='complete' else 'warning'}">{d.get('processing_status','?')}</span></td>
      <td class="text-muted small">{(d.get('ai_summary') or '')[:80]}</td>
      <td>{fd(d.get('created_at'))}</td>
    </tr>""" for d in docs) or "<tr><td colspan=6 class='text-muted p-4 text-center'>No documents yet. <a href='/upload'>Upload your first file →</a></td></tr>"

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-4">
  <h2 class="fw-bold mb-0">Documents</h2>
  <a href="/upload" class="btn btn-primary btn-sm"><i class="bi bi-plus-lg me-1"></i>Add Document</a>
</div>
<div class="card"><div class="card-body p-0">
<table class="table table-hover mb-0">
<thead><tr><th>Filename</th><th>Type</th><th>Confidence</th><th>Status</th><th>Summary</th><th>Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Documents", "documents")

def table_page(endpoint, title, cols, headers, rows_fn):
    @app.route(f"/{endpoint.replace('_page','').replace('_','-')}")
    @login_required
    def _view():
        data = sb_select(endpoint.replace("_page","").replace("-","_"),limit=200)
        rows = "".join(rows_fn(d) for d in data) or f"<tr><td colspan={len(headers)} class='text-muted p-4 text-center'>No {title.lower()} yet</td></tr>"
        content = f"""<h2 class="fw-bold mb-4">{title}</h2>
<div class="card"><div class="card-body p-0">
<table class="table table-hover mb-0">
<thead><tr>{"".join(f"<th>{h}</th>" for h in headers)}</tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
        return layout(content, title, endpoint)
    _view.__name__ = endpoint
    return _view

@app.route("/properties")
@login_required
def properties():
    data = sb_select("properties",limit=200)
    rows = "".join(f"<tr><td>{d.get('address','?')}</td><td>{d.get('suburb','?')}</td><td>{d.get('property_type','?')}</td><td>{d.get('size_sqm','?')}</td><td>{fc(d.get('asking_rent_pa'))}</td><td>{sbadge(d.get('status'))}</td><td>{d.get('landlord','?')}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No properties yet</td></tr>"
    content = f"""<div class="d-flex justify-content-between align-items-center mb-4"><h2 class="fw-bold mb-0">Properties</h2>
<button class="btn btn-primary btn-sm" data-bs-toggle="modal" data-bs-target="#addProp"><i class="bi bi-plus-lg me-1"></i>Add</button></div>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Address</th><th>Suburb</th><th>Type</th><th>Size</th><th>Rent pa</th><th>Status</th><th>Landlord</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>
<div class="modal fade" id="addProp" tabindex="-1"><div class="modal-dialog"><div class="modal-content" style="background:#1e293b">
<div class="modal-header"><h5 class="modal-title">Add Property</h5><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
<form method="POST" action="/api/properties"><div class="modal-body">
<div class="mb-2"><input name="address" class="form-control" placeholder="Address" required></div>
<div class="row g-2 mb-2"><div class="col"><input name="suburb" class="form-control" placeholder="Suburb" value="West Melbourne"></div>
<div class="col"><input name="property_type" class="form-control" placeholder="Type" value="Warehouse"></div></div>
<div class="row g-2 mb-2"><div class="col"><input name="size_sqm" type="number" class="form-control" placeholder="Size sqm"></div>
<div class="col"><input name="asking_rent_pa" type="number" class="form-control" placeholder="Rent pa"></div></div>
<div class="mb-2"><input name="landlord" class="form-control" placeholder="Landlord"></div>
<select name="status" class="form-select"><option>Available</option><option>Leased</option><option>Under Offer</option></select>
</div><div class="modal-footer"><button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
<button type="submit" class="btn btn-primary">Add Property</button></div></form></div></div></div>"""
    return layout(content, "Properties", "properties")

@app.route("/api/properties", methods=["POST"])
@login_required
def api_add_property():
    data = {k:v for k,v in request.form.items() if v}
    flash("Property added" if sb_insert("properties",data) else "Failed to add","success" if sb_insert("properties",data) else "error")
    return redirect(url_for("properties"))

@app.route("/contacts")
@login_required
def contacts():
    data = sb_select("contacts",limit=200)
    rows = "".join(f"<tr><td>{d.get('name','?')}</td><td>{d.get('company','?')}</td><td>{d.get('contact_type','?')}</td><td>{d.get('phone','?')}</td><td>{d.get('email','?')}</td><td>{sbadge(d.get('lead_status'))}</td></tr>" for d in data) or "<tr><td colspan=6 class='text-muted p-4 text-center'>No contacts yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Contacts</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Name</th><th>Company</th><th>Type</th><th>Phone</th><th>Email</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Contacts", "contacts")

@app.route("/inquiries")
@login_required
def inquiries():
    data = sb_select("inquiries",limit=200)
    rows = "".join(f"<tr><td>{d.get('company','?')}</td><td>{d.get('contact_name','?')}</td><td>{d.get('size_min','?')}-{d.get('size_max','?')} sqm</td><td>{fc(d.get('budget'))}</td><td>{d.get('source','?')}</td><td>{sbadge(d.get('status'))}</td></tr>" for d in data) or "<tr><td colspan=6 class='text-muted p-4 text-center'>No leads yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Lead Pipeline</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Company</th><th>Contact</th><th>Size</th><th>Budget</th><th>Source</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Lead Pipeline", "inquiries")

@app.route("/deals")
@login_required
def deals():
    data = sb_select("deals",limit=200)
    rows = "".join(f"<tr><td>{d.get('tenant','?')}</td><td>{d.get('landlord','?')}</td><td>{d.get('deal_type','?')}</td><td>{d.get('size_sqm','?')} sqm</td><td>{fc(d.get('rent_pa'))}</td><td>{d.get('term_years','?')} yrs</td><td>{sbadge(d.get('status'))}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No deals yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Deals</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Tenant</th><th>Landlord</th><th>Type</th><th>Size</th><th>Rent pa</th><th>Term</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Deals", "deals")

@app.route("/vacancies")
@login_required
def vacancies():
    data = sb_select("vacancies",limit=200)
    rows = "".join(f"<tr><td>{d.get('address','?')}</td><td>{d.get('suburb','?')}</td><td>{d.get('size_sqm','?')} sqm</td><td>{fc(d.get('asking_rent_pa'))}</td><td>{fd(d.get('available_date'))}</td><td>{d.get('vacating_tenant','?')}</td><td>{sbadge(d.get('status'))}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No vacancies yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Vacancies</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Address</th><th>Suburb</th><th>Size</th><th>Rent pa</th><th>Available</th><th>Vacating Tenant</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Vacancies", "vacancies")

@app.route("/requirements")
@login_required
def requirements():
    data = sb_select("requirements",limit=200)
    rows = "".join(f"<tr><td>{d.get('company','?')}</td><td>{d.get('contact_name','?')}</td><td>{d.get('size_min','?')}-{d.get('size_max','?')} sqm</td><td>{fc(d.get('budget_pa'))}</td><td>{d.get('preferred_location','?')}</td><td>{d.get('timeline','?')}</td><td>{sbadge(d.get('status'))}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No requirements yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Requirements</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Company</th><th>Contact</th><th>Size</th><th>Budget pa</th><th>Location</th><th>Timeline</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Requirements", "requirements")

@app.route("/email")
@login_required
def email_page():
    data = sb_select("email_logs",limit=50)
    rows = "".join(f"<tr><td>{d.get('sender','?')[:30]}</td><td>{d.get('subject','?')[:50]}</td><td><span class='badge bg-{'danger' if d.get('ai_priority')=='high' else 'warning' if d.get('ai_priority')=='medium' else 'secondary'}'>{d.get('ai_priority','?')}</span></td><td class='text-muted small'>{(d.get('ai_summary') or '')[:60]}</td><td>{fd(d.get('received_at'))}</td></tr>" for d in data) or "<tr><td colspan=5 class='text-muted p-4 text-center'>No emails yet</td></tr>"
    content = f"""<div class="d-flex justify-content-between align-items-center mb-4">
<h2 class="fw-bold mb-0">Email Inbox</h2>
<a href="/api/email/check" class="btn btn-primary btn-sm"><i class="bi bi-arrow-clockwise me-1"></i>Check Now</a></div>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>From</th><th>Subject</th><th>Priority</th><th>Summary</th><th>Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Email", "email_page")

@app.route("/api/email/check")
@login_required
def api_email_check():
    threading.Thread(target=check_gmail,daemon=True).start()
    flash("Checking email in background...","success")
    return redirect(url_for("email_page"))

@app.route("/call-logs")
@login_required
def call_logs():
    data = sb_select("call_logs",limit=100)
    rows = "".join(f"<tr><td>{d.get('contact_name','?')}</td><td>{d.get('phone','?')}</td><td>{d.get('direction','?')}</td><td>{d.get('duration_secs','?')}s</td><td class='text-muted small'>{(d.get('ai_summary') or '')[:60]}</td><td>{fd(d.get('created_at'))}</td></tr>" for d in data) or "<tr><td colspan=6 class='text-muted p-4 text-center'>No call logs yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Call Logs</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Contact</th><th>Phone</th><th>Direction</th><th>Duration</th><th>Summary</th><th>Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Call Logs", "call_logs")

@app.route("/market-data")
@login_required
def market_data_page():
    data = sb_select("market_data",limit=100)
    rows = "".join(f"<tr><td>{d.get('address','?')}</td><td>{d.get('suburb','?')}</td><td>{d.get('deal_type','?')}</td><td>{d.get('size_sqm','?')} sqm</td><td>{fc(d.get('rent_pa'))}</td><td>{d.get('tenant','?')}</td><td>{fd(d.get('date_signed'))}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No market data yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Market Data</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Address</th><th>Suburb</th><th>Type</th><th>Size</th><th>Rent pa</th><th>Tenant</th><th>Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Market Data", "market_data_page")

@app.route("/fees")
@login_required
def fees_page():
    data = sb_select("fees",limit=100)
    rows = "".join(f"<tr><td>{d.get('deal_id','?')}</td><td>{fc(d.get('amount'))}</td><td>{sbadge(d.get('invoice_status'))}</td><td>{fd(d.get('invoice_date'))}</td><td>{fd(d.get('paid_date'))}</td></tr>" for d in data) or "<tr><td colspan=5 class='text-muted p-4 text-center'>No fees yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Fee Tracking</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Deal</th><th>Amount</th><th>Status</th><th>Invoice Date</th><th>Paid Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Fees", "fees_page")

@app.route("/briefings")
@login_required
def briefings_page():
    data = sb_select("briefings",limit=20)
    rows = "".join(f"<tr><td>{d.get('briefing_type','?')}</td><td>{d.get('channel','?')}</td><td class='text-muted small'>{(d.get('content') or '')[:100]}</td><td>{fd(d.get('sent_at'))}</td></tr>" for d in data) or "<tr><td colspan=4 class='text-muted p-4 text-center'>No briefings yet</td></tr>"
    content = f"""<div class="d-flex justify-content-between align-items-center mb-4">
<h2 class="fw-bold mb-0">Briefings</h2>
<a href="/api/briefings/send" class="btn btn-primary btn-sm"><i class="bi bi-send me-1"></i>Send Now</a></div>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Type</th><th>Channel</th><th>Preview</th><th>Sent</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Briefings", "briefings_page")

@app.route("/api/briefings/send")
@login_required
def api_briefing_send():
    threading.Thread(target=send_daily_briefing,daemon=True).start()
    flash("Briefing sending...","success"); return redirect(url_for("briefings_page"))

@app.route("/whatsapp", methods=["GET","POST"])
@login_required
def whatsapp_page():
    if request.method=="POST":
        to = request.form.get("to","").strip()
        msg = request.form.get("message","").strip()
        if to and msg:
            ok = send_whatsapp(to,msg)
            flash("Message sent" if ok else "Failed — check Twilio config","success" if ok else "error")
    content = """<h2 class="fw-bold mb-4">WhatsApp</h2>
<div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Send Message</div>
<div class="card-body"><form method="POST">
<div class="mb-3"><input name="to" class="form-control" placeholder="+61412345678" required></div>
<div class="mb-3"><textarea name="message" class="form-control" rows="4" placeholder="Message..." required></textarea></div>
<button type="submit" class="btn btn-primary w-100"><i class="bi bi-whatsapp me-2"></i>Send</button>
</form></div></div></div>"""
    return layout(content, "WhatsApp", "whatsapp_page")

@app.route("/settings")
@login_required
def settings_page():
    def dot(ok): return f'<span class="badge bg-{"success" if ok else "danger"}">{"✓ Connected" if ok else "✗ Not Set"}</span>'
    content = f"""<h2 class="fw-bold mb-4">Settings</h2>
<div class="row g-3">
<div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Integration Status</div>
<div class="card-body"><table class="table table-sm mb-0">
<tr><td>Supabase (Direct HTTP)</td><td>{dot(bool(SUPABASE_URL and SUPABASE_KEY))}</td></tr>
<tr><td>OpenAI</td><td>{dot(bool(OPENAI_API_KEY))}</td></tr>
<tr><td>Twilio WhatsApp</td><td>{dot(bool(TWILIO_SID and TWILIO_TOKEN))}</td></tr>
<tr><td>Gmail</td><td>{dot(bool(GMAIL_USER and GMAIL_PASS))}</td></tr>
</table></div></div></div>
<div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Configuration</div>
<div class="card-body"><table class="table table-sm mb-0">
<tr><td>Version</td><td>V2.0 — Clean Build</td></tr>
<tr><td>Timezone</td><td>{TIMEZONE}</td></tr>
<tr><td>Briefing</td><td>{BRIEFING_HOUR:02d}:{BRIEFING_MIN:02d}</td></tr>
<tr><td>OpenAI Model</td><td>{OPENAI_MODEL}</td></tr>
<tr><td>Embedding Model</td><td>{EMBED_MODEL}</td></tr>
</table></div></div></div></div>"""
    return layout(content, "Settings", "settings_page")

@app.route("/webhook/whatsapp", methods=["POST"])
def webhook_whatsapp():
    from_num = request.form.get("From","")
    body = request.form.get("Body","").strip()
    log.info("WA inbound %s: %s", from_num, body[:80])
    return "",200

@app.route("/health")
def health():
    return jsonify({"status":"ok","version":"2.0","supabase":bool(SUPABASE_URL),"openai":bool(OPENAI_API_KEY)})

if __name__=="__main__":
    init_schema()
    port = int(os.getenv("PORT",8080))
    app.run(host="0.0.0.0",port=port,debug=False)
