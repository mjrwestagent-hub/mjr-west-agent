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
GITHUB_REPO    = os.getenv("GITHUB_REPO","mjrwestagent-hub/mjr-west-agent")

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


# Core indexed columns per table. AI extractions are split:
# known fields → indexed columns, everything else → data{} JSONB.
# This means any document type works without schema changes.
CORE_COLS = {
    "properties":   {"address","suburb","property_type","size_sqm","asking_rent_pa","status","landlord","source"},
    "contacts":     {"name","company","phone","email","contact_type","lead_status","whatsapp_opt_in"},
    "inquiries":    {"company","contact_name","size_min","size_max","budget","location","status","source"},
    "deals":        {"tenant","landlord","deal_type","size_sqm","rent_pa","status","commission"},
    "vacancies":    {"address","suburb","size_sqm","asking_rent_pa","available_date","vacating_tenant","status"},
    "requirements": {"company","contact_name","size_min","size_max","budget_pa","preferred_location","region","status"},
    "market_data":  {"address","suburb","deal_type","size_sqm","rent_pa","tenant","date_signed"},
    "documents":    {"filename","ai_classification","ai_summary","ai_confidence","ai_urgency","processing_status","raw_extraction"},
    "briefings":    {"briefing_type","content","channel"},
    "email_logs":   {"sender","subject","body","ai_summary","ai_priority","requires_action","reply_status","source"},
    "call_logs":    {"contact_name","phone","direction","duration_secs","notes","ai_summary"},
    "fees":         {"deal_id","amount","invoice_status","invoice_date","paid_date","notes"},
    "quotes":       {"client","property_address","quoted_rent_pa","deal_type","notes","status"},
}

def sb_insert(table, data):
    """Smart insert: known columns go to indexed fields, extras go to data{} JSONB.
    AI can return any fields — this handles it gracefully, forever."""
    try:
        core = CORE_COLS.get(table, set())
        if core:
            row = {k: v for k, v in data.items() if k in core and v is not None and v != ""}
            extra = {k: v for k, v in data.items() if k not in core and k != "data" and v is not None}
            if extra:
                row["data"] = extra
        else:
            row = {k: v for k, v in data.items() if v is not None and v != ""}
        if not row:
            return None
        r = http.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=SB_HDR, json=row, timeout=10)
        if r.status_code in (200,201):
            rows=r.json(); return rows[0] if rows else row
        log.error("sb_insert %s: %s %s", table, r.status_code, r.text[:200])
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
    """Create tables with JSONB data column for flexible AI extractions.
    Core indexed fields for querying, data{} for everything else.
    Schema never needs changing — new document types work automatically."""
    sqls = [
        """create table if not exists properties (
            id bigserial primary key, address text, suburb text default 'West Melbourne',
            property_type text default 'Warehouse', size_sqm numeric, asking_rent_pa numeric,
            status text default 'Available', landlord text, source text,
            data jsonb default '{}', created_at timestamptz default now())""",
        """create table if not exists contacts (
            id bigserial primary key, name text, company text, phone text, email text,
            contact_type text default 'Agent', lead_status text default 'Warm',
            whatsapp_opt_in bool default false, data jsonb default '{}',
            created_at timestamptz default now())""",
        """create table if not exists inquiries (
            id bigserial primary key, company text, contact_name text,
            size_min numeric, size_max numeric, budget numeric, location text,
            status text default 'New', source text default 'Web',
            data jsonb default '{}', created_at timestamptz default now())""",
        """create table if not exists deals (
            id bigserial primary key, tenant text, landlord text,
            deal_type text default 'Lease', size_sqm numeric, rent_pa numeric,
            status text default 'Negotiation', commission numeric,
            data jsonb default '{}', created_at timestamptz default now())""",
        """create table if not exists vacancies (
            id bigserial primary key, address text, suburb text default 'West Melbourne',
            size_sqm numeric, asking_rent_pa numeric, available_date date,
            vacating_tenant text, status text default 'Available',
            data jsonb default '{}', created_at timestamptz default now())""",
        """create table if not exists requirements (
            id bigserial primary key, company text, contact_name text,
            size_min numeric, size_max numeric, budget_pa numeric,
            preferred_location text default 'West Melbourne', region text default 'W',
            status text default 'Active', data jsonb default '{}',
            created_at timestamptz default now())""",
        """create table if not exists market_data (
            id bigserial primary key, address text, suburb text default 'West Melbourne',
            deal_type text default 'Lease', size_sqm numeric, rent_pa numeric,
            tenant text, date_signed date, data jsonb default '{}',
            created_at timestamptz default now())""",
        """create table if not exists documents (
            id bigserial primary key, filename text,
            ai_classification text default 'unknown', ai_summary text,
            ai_confidence numeric default 0, ai_urgency text default 'low',
            processing_status text default 'pending', raw_extraction jsonb default '{}',
            created_at timestamptz default now())""",
        """create table if not exists briefings (
            id bigserial primary key, briefing_type text default 'Daily',
            content text, channel text default 'WhatsApp',
            sent_at timestamptz default now())""",
        """create table if not exists email_logs (
            id bigserial primary key, sender text, subject text, body text,
            ai_summary text, ai_priority text default 'medium',
            requires_action bool default false, reply_status text default 'pending',
            source text default 'gmail_imap', data jsonb default '{}',
            received_at timestamptz default now())""",
        """create table if not exists call_logs (
            id bigserial primary key, contact_name text, phone text,
            direction text default 'Inbound', duration_secs int,
            notes text, ai_summary text, data jsonb default '{}',
            created_at timestamptz default now())""",
        """create table if not exists fees (
            id bigserial primary key, deal_id bigint, amount numeric,
            invoice_status text default 'Pending', invoice_date date,
            paid_date date, notes text, created_at timestamptz default now())""",
        """create table if not exists quotes (
            id bigserial primary key, client text, property_address text,
            quoted_rent_pa numeric, deal_type text default 'Lease',
            notes text, status text default 'Draft',
            created_at timestamptz default now())""",
    ]
    for sql in sqls:
        r = http.post(f"{SUPABASE_URL}/rest/v1/rpc/exec_sql",
                      headers=SB_HDR, json={"query": sql}, timeout=15)
        if r.status_code not in (200,201,204):
            log.info("Table may already exist: %s", r.status_code)


def gpt(system, user, max_tokens=2000):
    """Direct HTTP to OpenAI — no library needed."""
    if not OPENAI_API_KEY:
        return "{}"
    try:
        r = http.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={"model": OPENAI_MODEL, "messages": [{"role":"system","content":system},{"role":"user","content":user}],
                  "temperature": 0.1, "max_tokens": max_tokens, "response_format": {"type":"json_object"}},
            timeout=60
        )
        return r.json()["choices"][0]["message"]["content"] if r.status_code==200 else "{}"
    except Exception as e:
        log.error("gpt: %s", e); return "{}"

def embed(text):
    """Direct HTTP to OpenAI embeddings."""
    if not OPENAI_API_KEY:
        return []
    try:
        r = http.post(
            "https://api.openai.com/v1/embeddings",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={"model": EMBED_MODEL, "input": text[:8000]},
            timeout=30
        )
        return r.json()["data"][0]["embedding"] if r.status_code==200 else []
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
    ("quotes_page","bi-file-text","Quotes"),
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

def prop_addr(row):
    a = row.get('address')
    if a and str(a) not in ('','None','none','null'): return a
    data = row.get('data') or {}
    if isinstance(data, str):
        try: import json as _j; data = _j.loads(data)
        except: data = {}
    full = data.get('full_address') or data.get('latitude_longitude','')
    if full and len(str(full)) > 5 and '{' not in str(full): return str(full)[:50]
    num = data.get('number','') or ''
    street = data.get('street','') or ''
    if street: return f"{num} {street}".strip()
    return '?'

def dget(row, *keys):
    """Get value from row or data{} JSONB fallback."""
    for k in keys:
        v = row.get(k)
        if v is not None and str(v) not in ("", "None", "none"): return v
    data = row.get("data") or {}
    if isinstance(data, str):
        try: data = json.loads(data)
        except: data = {}
    for k in keys:
        v = data.get(k)
        if v is not None and str(v) not in ("", "None", "none"): return v
    return None

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

# ── Gmail OAuth2 ─────────────────────────────────────────────────────────────
@app.route("/auth/google")
def auth_google():
    import urllib.parse
    params = urllib.parse.urlencode({
        "client_id": os.environ.get("GOOGLE_CLIENT_ID",""),
        "redirect_uri": os.environ.get("GOOGLE_REDIRECT_URI",""),
        "response_type": "code",
        "scope": "https://www.googleapis.com/auth/gmail.readonly",
        "access_type": "offline", "prompt": "consent"
    })
    return redirect(f"https://accounts.google.com/o/oauth2/auth?{params}")

@app.route("/auth/google/callback")
def auth_google_callback():
    import urllib.parse, json as _j, urllib.request
    code = request.args.get("code","")
    if not code: return "No code", 400
    data = urllib.parse.urlencode({
        "code": code,
        "client_id": os.environ.get("GOOGLE_CLIENT_ID",""),
        "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET",""),
        "redirect_uri": os.environ.get("GOOGLE_REDIRECT_URI",""),
        "grant_type": "authorization_code"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data, method="POST")
    with urllib.request.urlopen(req) as resp:
        tokens = _j.loads(resp.read())
    sb_insert("settings", {"key": "gmail_refresh_token", "value": tokens.get("refresh_token","")})
    sb_insert("settings", {"key": "gmail_access_token", "value": tokens.get("access_token","")})
    log.info("Gmail OAuth2 connected")
    return redirect("/email?status=gmail_connected")

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

QUOTES = [
    ("The secret of getting ahead is getting started.", "Mark Twain"),
    ("Success is not final, failure is not fatal: it is the courage to continue that counts.", "Winston Churchill"),
    ("The only way to do great work is to love what you do.", "Steve Jobs"),
    ("In the middle of every difficulty lies opportunity.", "Albert Einstein"),
    ("Your time is limited, don't waste it living someone else's life.", "Steve Jobs"),
    ("The best time to plant a tree was 20 years ago. The second best time is now.", "Chinese Proverb"),
    ("Don't watch the clock; do what it does. Keep going.", "Sam Levenson"),
    ("Opportunities don't happen. You create them.", "Chris Grosser"),
    ("The harder I work, the luckier I get.", "Gary Player"),
    ("Success usually comes to those who are too busy to be looking for it.", "Henry David Thoreau"),
    ("It does not matter how slowly you go as long as you do not stop.", "Confucius"),
    ("The real estate market is local, and so is success.", "Unknown"),
    ("In real estate, location is everything — but relationships are everything else.", "Unknown"),
    ("A deal is only as good as the people in it.", "Unknown"),
    ("Every property tells a story. Your job is to match it with the right tenant.", "Unknown"),
    ("The best investment on earth is earth.", "Louis Glickman"),
    ("Buy land, they're not making it anymore.", "Mark Twain"),
    ("Real estate cannot be lost or stolen, nor can it be carried away.", "Franklin D. Roosevelt"),
    ("To be successful in real estate, you must always and consistently put your client's best interests first.", "Ron Willingham"),
    ("The most important thing in communication is hearing what isn't said.", "Peter Drucker"),
]

def get_quote_of_day():
    import hashlib
    day_hash = int(hashlib.md5(datetime.now().strftime("%Y-%m-%d").encode()).hexdigest(), 16)
    q, author = QUOTES[day_hash % len(QUOTES)]
    return f'{q} <span style="color:#f59e0b;font-style:normal;font-weight:600;">— {author}</span>'

@app.route("/dashboard")
@login_required
def dashboard():
    pc = len(sb_select("properties"))
    vc = len(sb_select("vacancies",{"status":"eq.Available"}))
    rc = len(sb_select("requirements",{"status":"eq.Active"}))
    dc = len(sb_select("deals"))
    vacs = sb_select("vacancies",{"status":"eq.Available"},limit=5)
    reqs = sb_select("requirements",{"status":"eq.Active"},limit=5)
    leads = sb_select("inquiries",{"status":"eq.New"},limit=5)
    calls = sb_select("call_logs",order="created_at",limit=5)
    today = datetime.now().strftime("%Y-%m-%d")
    emails = sb_select("email_logs",{"requires_action":"eq.true"},limit=5)

    vrow = "".join(f"<tr><td>{v.get('address','?')}</td><td>{v.get('size_sqm','?')} sqm</td><td>{fc(v.get('asking_rent_pa'))}</td></tr>" for v in vacs) or "<tr><td colspan=3 class=\'text-muted p-3 text-center\'>No vacancies yet</td></tr>"
    rrow = "".join(f"<tr><td>{r.get('company','?')}</td><td>{r.get('size_min','?')}-{r.get('size_max','?')} sqm</td><td>{fc(r.get('budget_pa'))}</td></tr>" for r in reqs) or "<tr><td colspan=3 class=\'text-muted p-3 text-center\'>No requirements yet</td></tr>"
    lrow = "".join(f"<tr><td>{i.get('company','?')}</td><td>{i.get('source','?')}</td><td>{sbadge(i.get('status'))}</td></tr>" for i in leads) or "<tr><td colspan=3 class=\'text-muted p-3 text-center\'>No new leads</td></tr>"
    crow = "".join(f"<tr><td>{c.get('contact_name','?')}</td><td>{c.get('direction','?')}</td><td>{c.get('duration_secs','?')}s</td><td class=\'text-muted small\'>{(c.get('ai_summary') or '')[:40]}</td></tr>" for c in calls) or "<tr><td colspan=4 class=\'text-muted p-3 text-center\'>No calls today</td></tr>"
    arow = "".join(f"<tr><td>{e.get('sender','?')[:25]}</td><td>{e.get('subject','?')[:40]}</td><td><span class=\'badge bg-danger\'>Action</span></td></tr>" for e in emails) or "<tr><td colspan=3 class=\'text-muted p-3 text-center\'>No actions pending</td></tr>"

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-3">
  <h2 class="fw-bold mb-0">Dashboard</h2>
  <span class="text-muted">{datetime.now().strftime('%A, %d %B %Y')}</span>
</div>
<div class="mb-4 p-3" style="background:linear-gradient(135deg,rgba(245,158,11,.12),rgba(99,102,241,.08));border:1px solid rgba(245,158,11,.25);border-radius:10px;">
  <div class="d-flex align-items-start gap-3">
    <i class="bi bi-quote" style="font-size:1.8rem;color:#f59e0b;line-height:1;"></i>
    <div>
      <div id="qotd" style="color:#e2e8f0;font-size:.95rem;font-style:italic;line-height:1.6;">{get_quote_of_day()}</div>
      <div class="mt-1 small text-muted">Quote of the Day</div>
    </div>
  </div>
</div>
<div class="row g-3 mb-4">
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{pc}</div><div class="text-muted small">Available Properties</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{vc}</div><div class="text-muted small">Active Vacancies</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{rc}</div><div class="text-muted small">Live Requirements</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card text-center"><div class="stat-value">{dc}</div><div class="text-muted small">Total Deals</div></div></div>
</div>
<div class="row g-3">
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold"><i class="bi bi-exclamation-circle text-warning me-2"></i>Today\'s Actions</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>From</th><th>Subject</th><th>Type</th></tr></thead><tbody>{arow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold"><i class="bi bi-telephone text-info me-2"></i>Today\'s Calls</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Contact</th><th>Direction</th><th>Duration</th><th>Summary</th></tr></thead><tbody>{crow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Recent Vacancies</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Address</th><th>Size</th><th>Rent</th></tr></thead><tbody>{vrow}</tbody></table></div></div></div>
  <div class="col-md-6"><div class="card"><div class="card-header fw-semibold">Active Requirements</div><div class="card-body p-0"><table class="table table-sm mb-0"><thead><tr><th>Company</th><th>Size</th><th>Budget</th></tr></thead><tbody>{rrow}</tbody></table></div></div></div>
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
    rows = "".join(f"<tr><td>{dget(d,'address') or '?'}</td><td>{dget(d,'suburb','city','location') or 'West Melbourne'}</td><td>{dget(d,'property_type','type','asset_type') or 'Warehouse'}</td><td>{dget(d,'size_sqm','size','gla','area') or '?'}</td><td>{fc(dget(d,'asking_rent_pa','rent_pa','rent','annual_rent'))}</td><td>{sbadge(dget(d,'status','occupancy_status') or 'Available')}</td><td>{dget(d,'landlord','owner','landlord_name') or '?'}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No properties yet</td></tr>"
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
    rows = "".join(f"<tr><td>{dget(d,'company','company_name','tenant') or '?'}</td><td>{dget(d,'contact_name','contact','name') or '?'}</td><td>{dget(d,'size_min','min_size') or '?'}-{dget(d,'size_max','max_size') or '?'} sqm</td><td>{fc(dget(d,'budget_pa','budget','max_rent_pa','rent_budget'))}</td><td>{dget(d,'preferred_location','location','suburb') or 'West Melbourne'}</td><td>{dget(d,'timeline','required_by') or '?'}</td><td>{sbadge(dget(d,'status') or 'Active')}</td></tr>" for d in data) or "<tr><td colspan=7 class='text-muted p-4 text-center'>No requirements yet</td></tr>"
    content = f"""<h2 class="fw-bold mb-4">Requirements</h2>
<div class="card"><div class="card-body p-0"><table class="table table-hover mb-0">
<thead><tr><th>Company</th><th>Contact</th><th>Size</th><th>Budget pa</th><th>Location</th><th>Timeline</th><th>Status</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>"""
    return layout(content, "Requirements", "requirements")


def gmail_get_access_token():
    """Get fresh Gmail access token using stored refresh token."""
    import urllib.parse, json as _j, urllib.request as _ur
    rows = sb_select("settings", {"key": "eq.gmail_refresh_token"})
    if not rows: return None
    refresh_token = rows[0].get("value","")
    if not refresh_token: return None
    data = urllib.parse.urlencode({
        "client_id": os.environ.get("GOOGLE_CLIENT_ID",""),
        "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET",""),
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }).encode()
    req = _ur.Request("https://oauth2.googleapis.com/token", data=data, method="POST")
    try:
        with _ur.urlopen(req) as resp:
            return _j.loads(resp.read()).get("access_token")
    except Exception as e:
        log.error("gmail token refresh error: %s", e)
        return None

def gmail_fetch_emails(max_results=20, query=""):
    """Fetch emails from Gmail API."""
    import urllib.request as _ur, json as _j, urllib.parse
    token = gmail_get_access_token()
    if not token: return []
    q = urllib.parse.urlencode({"maxResults": max_results, "q": query or "is:unread"})
    req = _ur.Request(
        f"https://gmail.googleapis.com/gmail/v1/users/me/messages?{q}",
        headers={"Authorization": f"Bearer {token}"}
    )
    try:
        with _ur.urlopen(req) as resp:
            data = _j.loads(resp.read())
        messages = data.get("messages", [])
        results = []
        for msg in messages[:10]:
            req2 = _ur.Request(
                f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{msg['id']}?format=metadata&metadataHeaders=Subject&metadataHeaders=From&metadataHeaders=Date",
                headers={"Authorization": f"Bearer {token}"}
            )
            with _ur.urlopen(req2) as resp2:
                detail = _j.loads(resp2.read())
            headers = {h["name"]: h["value"] for h in detail.get("payload",{}).get("headers",[])}
            results.append({
                "id": msg["id"],
                "subject": headers.get("Subject","(no subject)"),
                "from": headers.get("From",""),
                "date": headers.get("Date",""),
                "snippet": detail.get("snippet","")
            })
        return results
    except Exception as e:
        log.error("gmail fetch error: %s", e)
        return []

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


@app.route("/quotes")
@login_required
def quotes_page():
    data = sb_select("quotes",limit=100)
    rows = "".join(f"""<tr>
      <td>{d.get('client','?')}</td>
      <td>{d.get('property_address','?')}</td>
      <td>{fc(d.get('quoted_rent_pa'))}</td>
      <td>{d.get('deal_type','?')}</td>
      <td>{sbadge(d.get('status'))}</td>
      <td>{fd(d.get('created_at'))}</td>
    </tr>""" for d in data) or "<tr><td colspan=6 class=\'text-muted p-4 text-center\'>No quotes yet</td></tr>"

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-4">
  <h2 class="fw-bold mb-0">Quotes</h2>
  <button class="btn btn-primary btn-sm" data-bs-toggle="modal" data-bs-target="#addQuote">
    <i class="bi bi-plus-lg me-1"></i>Add Quote
  </button>
</div>
<div class="card"><div class="card-body p-0">
<table class="table table-hover mb-0">
<thead><tr><th>Client</th><th>Property</th><th>Rent pa</th><th>Type</th><th>Status</th><th>Date</th></tr></thead>
<tbody>{rows}</tbody></table></div></div>
<div class="modal fade" id="addQuote" tabindex="-1"><div class="modal-dialog"><div class="modal-content" style="background:#1e293b">
<div class="modal-header"><h5 class="modal-title">Add Quote</h5><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
<form method="POST" action="/api/quotes"><div class="modal-body">
<div class="mb-2"><input name="client" class="form-control" placeholder="Client / Company" required></div>
<div class="mb-2"><input name="property_address" class="form-control" placeholder="Property Address"></div>
<div class="row g-2 mb-2">
  <div class="col"><input name="quoted_rent_pa" type="number" class="form-control" placeholder="Rent pa"></div>
  <div class="col"><select name="deal_type" class="form-select"><option>Lease</option><option>Sale</option></select></div>
</div>
<div class="mb-2"><textarea name="notes" class="form-control" rows="2" placeholder="Notes"></textarea></div>
<select name="status" class="form-select"><option>Draft</option><option>Sent</option><option>Accepted</option><option>Declined</option></select>
</div>
<div class="modal-footer">
  <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
  <button type="submit" class="btn btn-primary">Add Quote</button>
</div></form></div></div></div>"""
    return layout(content, "Quotes", "quotes_page")

@app.route("/api/quotes", methods=["POST"])
@login_required
def api_add_quote():
    data = {k:v for k,v in request.form.items() if v}
    flash("Quote added" if sb_insert("quotes",data) else "Failed to add","success" if data else "error")
    return redirect(url_for("quotes_page"))


@app.route("/api/push", methods=["POST"])
def api_push():
    """Autonomous code push endpoint. Accepts code changes and pushes to GitHub.
    Protected by admin auth. Used by Claude to push fixes autonomously."""
    auth = request.headers.get("Authorization","")
    if auth != f"Bearer {ADMIN_PASSWORD}":
        return jsonify({"error": "unauthorized"}), 401
    
    data = request.json or {}
    filename = data.get("filename", "main.py")
    content  = data.get("content", "")
    message  = data.get("message", "auto: fix")
    token    = os.getenv("GITHUB_TOKEN","")
    repo     = os.getenv("GITHUB_REPO", "mjrwestagent-hub/mjr-west-agent")
    
    if not token:
        return jsonify({"error": "GITHUB_TOKEN not set"}), 500
    if not content:
        return jsonify({"error": "no content"}), 400

    try:
        import base64 as b64
        # Get current file SHA
        r = http.get(
            f"https://api.github.com/repos/{repo}/contents/{filename}",
            headers={"Authorization": f"token {token}", "User-Agent": "TurkishAgent"}
        )
        if r.status_code != 200:
            return jsonify({"error": f"get file failed: {r.status_code}"}), 500
        sha = r.json()["sha"]
        
        # Push updated file
        encoded = b64.b64encode(content.encode("utf-8")).decode("ascii")
        r2 = http.put(
            f"https://api.github.com/repos/{repo}/contents/{filename}",
            headers={"Authorization": f"token {token}", "User-Agent": "TurkishAgent", "Content-Type": "application/json"},
            json={"message": message, "content": encoded, "sha": sha}
        )
        if r2.status_code in (200, 201):
            commit = r2.json().get("commit",{}).get("sha","")[:8]
            log.info("Auto-push success: %s -> %s", message, commit)
            return jsonify({"success": True, "commit": commit})
        return jsonify({"error": f"push failed: {r2.status_code}", "detail": r2.text[:200]}), 500
    except Exception as e:
        log.error("api_push: %s", e)
        return jsonify({"error": str(e)}), 500

@app.route("/health")
def health():
    return jsonify({"status":"ok","version":"2.0","supabase":bool(SUPABASE_URL),"openai":bool(OPENAI_API_KEY)})

if __name__=="__main__":
    init_schema()
    port = int(os.getenv("PORT",8080))
    app.run(host="0.0.0.0",port=port,debug=False)
